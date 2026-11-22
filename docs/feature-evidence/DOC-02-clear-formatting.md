# DOC-02 清除格式、文档默认样式与 DOCX 保存保真

证据日期：2026-11-22（Asia/Shanghai）。状态：**未验收 / 待验证**。

本记录覆盖正文和表格内段落的清除格式、文档默认字符/段落样式，以及这些属性的内部快照和 DOCX 读写路径。本轮完成实现与测试草稿的静态审阅，未执行测试、类型检查、前端/Rust/WASM 构建、浏览器或格式检查；运行验证由主任务统一执行。此前 [DOC-02 上下标证据](DOC-02-vertical-align.md)中的历史通过结果不能作为本轮改动的验收结果，也不证明 DOC-02 全量完成。

## 1. 固定源码与入口契约

参考：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地 `E:\rust_test\univer`。静态复核确认本地 HEAD 与该提交一致，下面四个核心参考文件没有工作树改动。

| 来源 | 已核实事实 |
| --- | --- |
| [ui/menu/schema.ts][u-schema]，`menuSchema` | 清除格式位于 START / HISTORY；组内 `order: 3`，格式刷之后；grid 为 `row: 2, column: 2, showLabel: false`。classic 使用同组顺序。 |
| [ui/menu/format-painter.menu.ts][u-menu]，`ClearFormattingMenuItemFactory` | 按钮命令为 `ui.command.clear-formatting`；禁用状态来自 `!session.canClear()`，不是是否曾点击格式刷。 |
| [ui/controllers/format-painter.controller.ts][u-controller]、[format-painter-session.service.ts][u-session] | 共享命令委托 session；`canClear()` 要求活动 adapter 存在 `clear` 且 `canStart()` 成立；执行还有 applying 检查，并取消格式刷会话。 |
| [docs-ui/controllers/format-painter.controller.ts][u-doc-controller]，`DocFormatPainterController` | 注册 `doc-text` adapter；使用默认字符格式和空 `paragraphStyle` 调用文档格式刷命令。活动文档、有效文本范围、编辑权限影响可执行性。 |
| [docs-ui/services/doc-menu-style.service.ts][u-defaults]，`getDefaultStyle` | 合并 UI 默认字符属性和文档 `documentStyle.textStyle`；主题提供默认字色；页眉/页脚使用独立默认字号。 |
| [docs-ui/commands/commands/format-painter.command.ts][u-command]，`captureTextFormat`、`buildTextFormatActions`、`ApplyTextFormatPainterCommand` | 光标扩展到词范围；字符属性用 REPLACE 更新；相交段落清除样式 ID/段落格式/列表；保留文本、段落和实体；通过 `RichTextEditingMutation` 返回扩展后的文本范围。 |

本切片核查的常驻根入口为 **1**；classic/grid 是同一入口的两种呈现。没有新增插件扩展根项、溢出专用项或上下文根项；共享溢出容器可以承载该入口，不另计功能。全量 Sheets/Docs、插件及上下文的清点和分类口径继续见 [顶部清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)和[工具栏对照](../UNIVER_TOOLBAR_PARITY.md)。上述计数不作为功能完成证明。

本项目保留 Ribbon 的 `ui.command.clear-formatting` 别名，实际映射到显式的 **`doc.command.clear-formatting`**。菜单/UI 和 `FDocument.executeCommand` 使用同一文档命令；没有以捕获固定示例格式、提示消息或空操作充当清除。

## 2. 模型、归一化与默认值

- [domain/model.ts](../../packages/office-studio/src/domain/model.ts) 增加可选 `DocumentModel.defaultStyle`，结构为 `{ version: 1, text, paragraph }`。`text` 包含 Latin/East Asian 字体、字号、字色、四种基本强调、高亮和上下标；`paragraph` 复用 `DocumentParagraphStyle`。`InlineRun.highlight` 为显式领域字段，旧 `metadata.highlight` 仍兼容。
- `createDocumentDefaultStyle()` 为新建文档生成独立对象：Arial 11、`#1f2937`、正常强调、左对齐、1.5 倍行距和段后 8px。新建文档显式保存默认样式；旧快照没有字段时继续使用原工作台默认值，不把兼容回退标记为已提取的 Office 默认样式。
- [document-default-style.ts](../../packages/office-studio/src/features/document/model/document-default-style.ts) 校验版本、必要字段、有限数值、行距/段距、缩进、上下标以及 `wordWrap` / `snapToGrid`。字色限制为 RGB 或 `auto`；高亮支持 Office 名称、RGB、空值和 `none`。非法默认属性在清除预检中拒绝；返回副本，避免修改默认样式对象。
- [document-highlight.ts](../../packages/office-studio/src/features/document/model/document-highlight.ts) 统一 16 种 Office 高亮名称与 HEX。默认样式预检和导入归一化共用该模型 helper，Canvas 使用 RGB；`highlight=none` 可以与独立 RGB shading 同时存在，显式 nil shading 由 Rust 适配恢复。
- [document_default_style.rs](../../crates/office-parser/src/document_default_style.rs) 复用 duke-pages 的 `parse_style_sheet`，只向样式引擎提供已保留的 styles/theme。解析时规范化 Word/Drawing 命名空间；`officeSource` 内的原 XML 不被改写。默认段落样式的继承、`basedOn`、字体主题槽交给已有样式引擎解析，`wordWrap` 复用项目 sidecar 的继承解析。
- Rust DTO 增加 `defaultStyle` 和 `documentDefaultStyle` 能力声明。[documentNormalizer.ts](../../packages/office-studio/src/parser/documentNormalizer.ts) 将 DTO 转为领域默认样式；字号使用 pt，段落尺寸使用 px，行距区分 auto 倍数与 exact/atLeast 尺寸。原 DTO 留在 `metadata.source`。DOC/ODT 或解析未能提供默认样式时删除 blank constructor 的字段，不伪装成已解析结果。

duke-pages 区分缺失 `rPrDefault`（其 Word 默认字号 12pt）与已声明但没有字号的 `rPrDefault`（10pt）；缺失段落默认与显式空段落默认也不同。对应差异已写入原生/归一化测试草稿，尚未运行验证。源样式解析失败保留诊断，不作为完整保真结果。

## 3. 清除命令、权限和历史

调用链：`document-ribbon.ts` → `DocumentWorkbench.run` → `FDocument.executeCommand` / `executeDocumentCommand` → `prepareDocumentClearFormatting` → 缓存编号标签变换 → `applyDocumentClearFormatting` → 已有文档内容历史与保存路径。

| 行为 | 当前实现 |
| --- | --- |
| 非空选区 | 保持半开范围和反向选区；只重置选中文字的字符格式。跨段选区以末段 offset=0 结束时排除末段。 |
| 折叠光标 | 复用 `documentPainterTarget` 扩展到词；空白且无法扩词时无操作。 |
| 段落 | 所有相交段落重置为文档默认段落格式及普通正文；清除标题级别、命名样式、列表/清单和旧 HTML 缓存。表格只修改地址指向的段落，并使其容器 HTML 缓存失效。 |
| 实体与来源 | 保留链接、来源 metadata、未选 run、零长度脚注/字段引用；只拆分发生字符格式变化的文本。重复清除相同默认格式保持 run 身份和 no-op。 |
| 内容锁 | 段落格式影响整段，因此检查整段中的内容锁，包括选词之外的控件。缓存编号前缀删除另检查 `controlLocked` 包装锁。 |
| 原子性 | 对全部相交目标预检权限、合法块、控件变换和默认格式后再写入；混合正文/表格后段拒绝时不得先清除前段或缓存编号。 |
| 缓存编号 | 用户明确清除时移除缓存标签及重放标记；同步变换评论/控件的偏移和选区，保留零长度引用 ID。沿用 DOC-05 的编号实现。 |
| 撤销、取消、保存 | 公共命令走文档隔离的统一内容历史；取消/拒绝不发布内容或历史，重复 no-op 不新增内容历史。默认样式随内部快照保存。以上已编写契约草稿，运行结果待主任务提供。 |
| 下一次输入 | 替换选区继承第一个选中字的格式，光标输入继续继承左侧字符；避免清除后重新继承未选前缀的格式。 |

实现文件：[document-clear-formatting.ts](../../packages/office-studio/src/features/document/services/document-clear-formatting.ts)、[document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts)、[document-format-painter.ts](../../packages/office-studio/src/features/document/services/document-format-painter.ts)、[document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts)。新增命令加入标签表，公共注册复用已有机制；不增加第二套内容历史。

Workbench 仅在清除成功后清空 `pendingMarks` 和格式刷会话；失败/取消保留 pending，并安排返回 Canvas 输入焦点。段落格式清除本身不改变正文文本；移除既有缓存编号标签属于明确的编号结构变换。

## 4. DOCX 写出与重读

- [docx_default_style.rs](../../crates/office-writer/src/docx_default_style.rs) 将领域默认样式写入受控模板的 `w:docDefaults`；无字段时保留旧模板，非法字段使导出失败。保留生成的 Normal/Heading/引用样式及 DOC-05 的源样式依赖重放路径。
- [docx.rs](../../crates/office-writer/src/docx.rs) 共用 `run_properties_xml` 写出正文和默认字符属性。显式 `false` 写为 `w:b/i/strike val=0`，下划线关闭写 `w:u val=none`，避免重新继承文档默认强调。
- 文档默认含上下标时，领域模型的正常正文显式写 `w:vertAlign=baseline`。默认含高亮时，无高亮正文显式写 `w:highlight=none` 和 `w:shd val=nil`。正文、表格正文及含 runs 的注释正文使用相同默认效果标记。
- 标准 RGB 高亮写 Office `ST_HighlightColor` 名称；自定义 RGB 写 `highlight=none + shd clear/fill`，覆盖继承的 Office 高亮。run 顶层高亮优先于旧 metadata。
- [document_format_sidecar.rs](../../crates/office-parser/src/document_format_sidecar.rs) 收集显式 `w:shd val=nil` 字符范围；仅移除该范围内 Duke 错误继承的 `backgroundColor`。按 UTF-8 安全边界拆分并重新计算 UTF-16 start/end，保留范围外样式、链接和空实体；无法安全映射时给出诊断。
- 段落默认写出对齐、缩进、行距规则、段前后、wordWrap 和 snapToGrid。首行缩进与悬挂缩进归为有效净偏移；不把这种归一化声称为原属性组合逐字保留。

这些为读写实现和测试草稿说明。本轮未生成新的样本文件、下载哈希、Worker 记录或独立 Office 重开结果。

## 5. UI 与测试草稿

清除按钮沿用 [RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 的 `antdv-next` Button/Tooltip、共享溢出和主题配置；没有新建原生表单控件。正文输入沿用 [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的 `Input.TextArea` 事件桥，内容绘制继续使用 Canvas。只静态核对本次路径；主题、焦点恢复、真实输入法和实际溢出交互仍待运行验证。

| 测试文件 | 已编写、尚未运行的场景 |
| --- | --- |
| [document-clear-formatting.contract.test.mjs](../../packages/office-studio/tests/document-clear-formatting.contract.test.mjs) | 扩词、反向半开范围、默认格式、重复 no-op、来源/链接/空引用、表格、权限、范围外内容锁、非法默认、颜色归一化、缓存标签的评论/控件变换与包装锁、正文跨表格原子拒绝、清除后输入。 |
| [document-editor-commands.contract.test.mjs](../../packages/office-studio/tests/document-editor-commands.contract.test.mjs) | `FDocument` 命令/事件、一次历史、撤销重做、取消/权限/锁拒绝、保存重开默认样式、重复 no-op、下一次输入及跨正文/表格拒绝后不发布内容与历史。 |
| [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | classic/grid 的真实清除入口、主题、焦点、立即输入、撤销重做和重开；取消保留 pending，权限禁用。该文件同时包含 DOC-05 场景，不用其历史结果替代本切片验证。 |
| `document_default_style.rs`、`docx_default_style.rs` 单元草稿 | 默认样式继承与主题字体、缺失/空声明、命名空间/非法资源、生成默认 XML、非法颜色和布尔字段、标准/自定义 RGB 高亮。 |
| `document_format_sidecar.rs` 单元草稿 | Unicode 范围上的 nil shading 恢复，未选背景、链接、空实体与其他属性保留。 |
| [document_format_roundtrip_tests.rs](../../crates/office-parser/src/document_format_roundtrip_tests.rs) | 原生 DOCX 默认格式、正文显式关闭强调/上下标/高亮、自定义 shading、DTO 重导出后二次重读。 |
| [document-numbering-roundtrip.spec.ts](../../tests/browser/document-numbering-roundtrip.spec.ts) | 生产预览中由真实 writer WASM 生成 DOCX → 上传 → 清除词格式 → 撤销重做 → 两次下载/独立 parser Worker 重读/重导入。默认 Georgia/宋体 15、RGB `#ABCDEF`、右对齐及 exact 24px；原文 Arial 26、基本强调、黄色高亮、上标；保留 `Hello `，清除 `world`。 |

生产测试复用 [document-numbering-roundtrip.playwright.config.ts](../../tests/browser/document-numbering-roundtrip.playwright.config.ts) 的构建产物预览和真实 WASM，不以 mock Writer/Parser 验证 Office 互操作。独立 Worker 是另一次本项目读取，不等同于独立 Office 程序。

主任务待执行的相关验证命令示例（**不是已执行结果**）：

```powershell
pnpm exec vitest run --config vitest.config.ts packages/office-studio/tests/document-clear-formatting.contract.test.mjs packages/office-studio/tests/document-editor-commands.contract.test.mjs
pnpm exec vitest run --config vitest.components.config.ts packages/office-studio/tests/document-home-colors.browser.test.mjs
cargo test --manifest-path crates/office-writer/Cargo.toml
cargo test --manifest-path crates/office-parser/Cargo.toml
pnpm check:types
pnpm build:wasm
pnpm build:writer
pnpm build
pnpm exec playwright test --config tests/browser/document-numbering-roundtrip.playwright.config.ts
```

运行后需要记录通过/失败场景、环境、真实文件 hash 和诊断，再由主任务更新两份汇总清单。所有相关验收当前保持未验收。

## 6. 实现差异与待完成边界

- 本项目用明确的 clear 命令及可序列化默认格式，代替上游通过格式刷 adapter 传入空段落样式的内部做法。保留扩词、相交段落、清除列表、实体、返回范围和共享入口的可观察语义；字符/段落重置采用已解析 Office 文档默认样式。
- 清除成功后才取消 pending/格式刷；与上游 session 在调用 adapter 前取消会话的内部顺序不同。取消/失败后保持输入意图属于本项目选择，需浏览器场景核验。
- 对整段内容锁和缓存编号包装锁的预检比上游当前 adapter 的单位级检查更具体；需要与真实锁定控件样本核验，不能仅用按钮禁用作为证据。
- **代码块及页眉/页脚清除未接入。** 本项目当前对这些入口显示不可用原因；上游可按 segment 处理，并对页眉页脚默认字号作区分。它们仍为完整交付目标，不以当前禁用验收。
- 高级字符属性（outline、caps/smallCaps 等）、复杂脚本字体槽、主题颜色 tint/shade、语言/字距等未形成完整清除和读写闭环；原始来源仍保留，不能据此声称可编辑保真。
- DOC/ODT 的默认样式提取、完整字段/SDT/注释/嵌套表格 Office 保真及所有正文容器的范围模型尚待补齐。本切片没有证明 DOC-11、DOC-12、DOC-13 或 DOC-22 全量完成。
- 尚无真实系统 IME、独立 Microsoft Office/兼容程序、WebGPU/Canvas2D 各自闭环以及真实长文档性能/内存/取消测量。本轮没有性能优化宣称。

## 7. 文件归属与主任务交接

本切片修改 Docs 默认样式模型、命令、归一化、写出/sidecar、Workbench 清除状态、Ribbon 映射及对应测试。共享 `domain/model.ts` 和 Rust `model.rs` / `parser.rs` 只改 Docs 段；Sheets 改动保留。高亮 helper 最终位置为 `features/document/model/document-highlight.ts`，先前未跟踪的 `parser/documentHighlight.ts` 已迁走，引用已更新。

DOC-05 的编号/源样式实现沿用此前交接，不在此重做。公共装配 `facade.ts`、`office-studio.ts`、`index.ts` 与 [功能清单](../UNIVER_FEATURE_PARITY.md)、[工具栏清单](../UNIVER_TOOLBAR_PARITY.md)的汇总更新由主任务负责。本证据页本身不提升功能验收状态。

[u-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/schema.ts
[u-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/format-painter.menu.ts
[u-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/controllers/format-painter.controller.ts
[u-session]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/format-painter/format-painter-session.service.ts
[u-doc-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/format-painter.controller.ts
[u-defaults]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/doc-menu-style.service.ts
[u-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/format-painter.command.ts
