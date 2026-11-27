# DOC-02 局部字符格式、清除、输入继承与 Office 来源保真

证据日期：2026-11-22（Asia/Shanghai）。状态：**未验收 / 待验证**。

本切片覆盖正文、表格内段落和页眉页脚 story 的局部字符格式、混合工具值及继续输入继承，保留代码块局部清除、DOC/DOCX/ODT 默认样式、字段缓存、输入桥和 DOCX 来源读写路径。DOC/ODT 全局默认定义、原始样式资源与 DOCX 历史来源 sidecar 已接入代码；下文仅说明静态实现进度，不能据此认定实际交互或文件保真已经通过。共享命令、输入会话与内容历史的完整接线由主任务记录在 [COMMON-03 证据](COMMON-03-input-selection-axis-history.md)。

**本轮运行验证为零。** 没有运行测试、类型检查、前端/Rust/WASM 构建、浏览器或格式检查，没有生成本轮文件样本、哈希或性能记录。所有新增及修改的测试均为未执行草稿；以前的通过结果不能作为本轮结果，也不证明 DOC-02、DOC-03、DOC-13 或 DOC-22 全量完成。

## 1. 固定源码、入口与差异

参考为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地只读路径 `E:\rust_test\univer`。全量顶部清点继续以[顶部入口表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)、[工具栏对照](../UNIVER_TOOLBAR_PARITY.md)及[源码规则](../UNIVER_REFERENCE_RULES.md)为准；本页不重新定义 OSS 或 Pro 总入口数。

| 固定来源与符号 | 静态核实的契约 |
| --- | --- |
| [ui/menu/schema.ts][u-schema]，`menuSchema` | 清除格式在 START / HISTORY，组内 `order: 3`，格式刷之后。grid 为 `row: 2, column: 2, showLabel: false`；classic 按同组顺序显示。 |
| [ui/menu/format-painter.menu.ts][u-menu]，`ClearFormattingMenuItemFactory` | 根按钮使用 `ui.command.clear-formatting`，禁用状态来自 `!session.canClear()`。 |
| [ui/controllers/format-painter.controller.ts][u-controller]、[format-painter-session.service.ts][u-session] | 共享命令委托活动 adapter；`canClear()` 要求 `clear` 和 `canStart()`，执行检查 applying 状态并取消格式刷会话。 |
| [docs-ui/controllers/format-painter.controller.ts][u-doc-controller]，`DocFormatPainterController` | `doc-text` adapter 用默认字符属性和空 `paragraphStyle` 执行清除；活动文档、有效文本范围、焦点与权限参与可执行性判断；捕获可取正文或 header/footer segment。 |
| [docs-ui/services/doc-menu-style.service.ts][u-defaults]，`getDefaultStyle` | 上游 UI 回退为 Arial 11，并合并文档 `documentStyle.textStyle` 与主题字色；页眉页脚覆盖字号为 9pt。 |
| [docs-ui/commands/commands/format-painter.command.ts][u-command]，`captureTextFormat`、`buildTextFormatActions`、`ApplyTextFormatPainterCommand` | 光标扩展到词；字符属性以 REPLACE 更新；相交段落替换段落格式、样式 ID 与列表属性；同一 segment 的范围通过 `RichTextEditingMutation` 进入历史，并保留文本与非格式实体。 |
| [docs-ui/menu/header-footer-ribbon.ts][u-header-footer] | header/footer 是独立上下文，具有导航、插入、选项、位置与结束编辑入口；本页只补 clear/输入/格式保留路径，不把整个上下文分类标为完成。 |
| [docs-ui/menu/schema.ts][u-doc-schema]，`menuSchema` | 字号在 START / FORMAT，`order: 6`，grid 为 `row: 1, column: 7, columnSpan: 2, width: 84, minWidth: 80`；字体同组 `order: 7`，grid 为 `row: 1, column: 3, columnSpan: 4, width: 140, minWidth: 124`。classic 按同组 order，字号先于字体。 |
| [docs-ui/menu/menu.ts][u-doc-menu]，`FontFamilySelectorMenuItemFactory`、`FontSizeSelectorMenuItemFactory` | 两者是 selector；非活动 Docs 隐藏，`disableMenuWhenNoDocRange` 合并无文本/矩形范围、相交代码块与编辑权限条件。字号输入配置 1–400；当前值来自 `getFontStyleAtCursor`，汉字可显示东亚字体槽。动态 factory 与真实控件行为尚未执行。 |
| [docs-ui/commands/commands/inline-format.command.ts][u-inline]，`SetInlineFormatCommand`、`getStyleInTextRange` | 字体 ID 为 `doc.command.set-inline-format-font-family`，字号为 `doc.command.set-inline-format-fontsize`；折叠范围缓存下一次输入格式。expanded 范围实际取默认值与所选字号的最大值，并取首个 explicit 字体、携带东亚字体槽；注释中的 min 不代表实现取最小值。 |
| [docs/services/doc-selection-manager.service.ts][u-selection] | 选区转换显式携带 `preserveInputStyle`；缓存输入偏好与内容 mutation 分开。本项目用文档隔离的 Pinia 会话实现相应边界，不复制上游内部 selection service。 |

本页逐项展开的根入口分类如下；此表只统计 clear、字体和字号，不代替 DOC-02、完整 Ribbon 或 Pro 的入口清点。

| 本页详细展开的根项 | 常驻 | 扩展专用 | 溢出专用 | 上下文专用 |
| --- | --- | --- | --- | --- |
| 清除格式 | 1 | 0 | 0 | 0 |
| 字体、字号 | 2 | 0 | 0 | 0 |
| 合计 | 3 | 0 | 0 | 0 |

classic/grid 是同一入口集合的布局变体；共享 More 和页眉页脚复用不另计根项。加粗、斜体、下划线、删除线、上下标及颜色控件同样改为依据活动字符上下文判断禁用，但其逐项位置/命令仍沿用总清单，不能把本页 3 项写成全部字符入口或功能完成数。全部 Pro 入口仍属于完整目标；[代码块规格][doc-code]与[Docs Office 转换规格][doc-io]另按对应功能 ID 验收，公开导航来自既有基线，本轮没有新演示观察结果。

本项目保留 `ui.command.clear-formatting` 别名，映射到显式 `doc.command.clear-formatting`。新建文档用 **Aptos 11**，与固定上游 Arial UI 回退不同；导入文档优先使用已解析的 Office 默认样式，页眉页脚依照固定上游字号契约使用 9pt。没有独立 Office/兼容程序验证，不能把这些选择写成已经核实的所有 Office 默认行为。

本项目字体/字号及其他字符入口共用 `doc.command.format-text`，新增可选 `story` 目标；控件值从真实所选 runs/cache 归并，混合字体或字号显示空值与混合提示，区别于上游上述最大字号/首个字体行为。UI 字号仍为 1–400，显式格式 patch 要求有限且大于 0、不超过 400pt；合法导入值的保存和继续输入继承不按新 patch 上限裁剪。Office 对应混合态、脚本字体和超出 UI 范围字号的最终交互尚无独立验证。

## 2. 默认样式、story 模型与协议

| 层 | 当前代码 |
| --- | --- |
| [domain/model.ts](../../packages/office-studio/src/domain/model.ts) | 可选 `DocumentModel.defaultStyle = { version: 1, text, paragraph }`。新建文档独立保存 Aptos / East Asian Aptos、11pt、`#1f2937`、正常强调、左对齐、1.5 倍行距、段后 8px。旧快照缺少字段时回退，不冒称提取到了 Office 默认样式。 |
| [document-default-style.ts](../../packages/office-studio/src/features/document/model/document-default-style.ts)、[document-highlight.ts](../../packages/office-studio/src/features/document/model/document-highlight.ts) | 校验版本、必需字段、有限尺寸、缩进、行距/段距、wordWrap/snapToGrid、上下标和颜色。字色为 RGB 或 `auto`；16 种 Office 高亮与 RGB 共用 helper。返回副本，非法默认属性在修改前拒绝。 |
| [document_default_style.rs](../../crates/office-parser/src/document_default_style.rs) | 复用 duke-pages `parse_style_sheet`，只提供已保留的 styles/theme。规范化 Word/Drawing 名称供样式引擎解析，保留 `officeSource` 原 XML；继承、basedOn、主题字体与段落默认交给现有引擎，wordWrap 复用本项目 sidecar。 |
| [document_legacy_default_style.rs](../../crates/office-parser/src/document_legacy_default_style.rs) | DOC 从同一次 `parse_doc` 生成的 package 读取 styles，另保留 FIB 选择的 Table 中 STSHF/STTBF_FFN 原二进制范围；合成 XML 明确标记。ODT 只取 metadata 默认 ID 与唯一默认定义，保留 styles.xml/content.xml，未采样首段或把 ODF XML 当成 Word part。资源与定义错误不产生 defaults。 |
| [parser.rs](../../crates/office-parser/src/parser.rs)、[documentNormalizer.ts](../../packages/office-studio/src/parser/documentNormalizer.ts) | DOCX/DOC/ODT 仅在真实默认定义成功时输出 `defaultStyle` 并声明 `documentDefaultStyle`；来源合同错误原子拒绝。领域字号为 pt，段落尺寸为 px，行距区分 auto 与 exact/atLeast；原 DTO 留在 `metadata.source`。未解析到 defaults 时删除 blank constructor 默认字段，保留旧快照缺省兼容。 |
| [document_legacy_source.rs](../../crates/office-io-common/src/document_legacy_source.rs)、[contracts.ts](../../packages/office-studio/src/parser/contracts.ts)、[documentOfficeSource.ts](../../packages/office-studio/src/parser/documentOfficeSource.ts) | parser/writer 共用历史来源类型、严格 JSON codec、去重及预算；TS 校验当前来源与解析格式一致，并校验去重后的整体资源/诊断/JSON 大小。未知来源协议字段明确拒绝，不静默反序列化丢弃；其余解析 DTO 属性继续完整留在 metadata.source。 |
| [document-story.ts](../../packages/office-studio/src/features/document/model/document-story.ts) | `DocumentStoryRange = { sectionId, story, start, end }`；story 为 header/footer、首页或偶数页的六种片段。范围是输入 buffer 的 UTF-16，每个字段占一个 U+FFFC；literal runs/paragraphs 使用不含字段的文本偏移。校验 cache、段落和锁，拒绝拆分代理对。 |
| [document-run-range.ts](../../packages/office-studio/src/features/document/model/document-run-range.ts)、[document-text-formatting.ts](../../packages/office-studio/src/features/document/services/document-text-formatting.ts)、[document-story-formatting.ts](../../packages/office-studio/src/features/document/services/document-story-formatting.ts) | 字符 patch、UTF-16 边界、局部拆分、正文控件锁及 story 权限/锁共用可测试服务。正文先准备全部目标再写 runs；story 分别准备 literal runs 和所选字段 cache，不修改休眠正文选区。 |
| [document-input-session.ts](../../packages/office-studio/src/features/document/model/document-input-session.ts)、[document-input-store.ts](../../packages/office-studio/src/features/document/stores/document-input-store.ts) | 文档隔离的瞬态 `location/marks/modelRevision/generation/valid`；body 与 story 采用不同 location。getter 保留 own property 的清色/清高亮 `undefined`，克隆嵌套 metadata；不写入快照、dirty 或内容历史。共享装配及公共 API 由主任务维护。 |

沿用已经落盘的版本化协议，本阶段没有新增 `DocumentPageField` 顶层字段：

```ts
{ id: 'doc.command.clear-formatting'; story?: DocumentStoryRange }
{ id: 'doc.command.format-text'; patch: RunPatch; clearFields?: ('color' | 'highlight')[]; story?: DocumentStoryRange }
{ id: 'doc.command.set-section-header-footer'; sectionId: string; patch: SectionHeaderFooterPatch; input?: DocumentStoryInputCommit }

field.metadata.officeStoryFieldCache = { version: 1, runs: DocumentRun[] }
```

Rust story DTO 可选 `cachedRuns` 经归一化进入该缓存。指令、`cachedResult` 和缓存格式分开保存；cache run 文本拼接必须等于 `cachedResult`。字段始终为一个输入对象，显示可包含多种字号、颜色、链接和换行片段。

新增来源字段仍属于 version 1 的可选扩展：

```ts
type Origin = {
  format: 'doc' | 'odt'
  synthesized: boolean // DOC=true，ODT=false
  resources: (
    | { kind: 'xml'; partName: string; xml: string }
    | { kind: 'binary'; streamName: string; offset: number; dataHex: string }
  )[]
  resourceErrors?: string[]
}
interface DocumentOfficeSource {
  version: 1
  styles?: { partName: string; xml: string } | null
  theme?: { partName: string; xml: string } | null
  resourceErrors?: string[]
  origin?: Origin
  legacyOrigins?: Origin[]
}
```

`origin` 只表示当前 DOC/ODT 导入。经 DOCX 写出与重开后，当前 `styles`/`theme` 来自新的原生 Word parts，当前 `origin` 缺省，原 DOC/ODT 资源放入 `legacyOrigins`。来源自身的 `resourceErrors` 随历史保留；它们不会替代 DOCX 当前 Word 资源的顶层诊断或阻止其默认定义解析。

来源继续用 `metadata.source.sourceFormat = { version: 1, propertiesXml?, namespaces?, paragraphStyleId? }`、run/field 的 `officeStoryStyleBaseline` 与 paragraph 的 `officeStoryParagraphStyleBaseline`。clear 写入 `officeStoryTextFormatCleared` / `officeStoryParagraphFormatCleared`，保留来源但禁止旧源格式重放。Rust `DocumentOfficeSourcePayload.resource_errors` 对应可选 JSON `resourceErrors: string[]`；已声明却丢失的 styles/theme 会形成诊断，不能当成合法缺省。

当前 duke-pages Word 默认区分仍保留：缺少 rPrDefault 时为 12pt，已声明但没有字号时为 10pt；缺少段落默认与显式空默认也不同。这是当前解析库契约，草稿未运行，不是独立 Office 验证。

## 3. 局部格式、清除、权限与继续输入

调用链为 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) → [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) → 公共 FDocument/编辑命令 → [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) → 字符格式/clear 服务 → 统一内容历史与保存。clear 使用 [document-clear-formatting.ts](../../packages/office-studio/src/features/document/services/document-clear-formatting.ts)，普通字符 patch 使用上述 run-range、text-formatting 与 story-formatting 服务；组件只装配上下文、焦点与输入意图。

| 上下文或边界 | 静态实现 |
| --- | --- |
| 正文/表格局部字符格式 | 字体、字号或其他字符 patch 只更新指定半开范围内所给属性，保留其余字符属性、链接、source、空实体及范围外 ID；不替换段落格式。反向范围归一化，末段 offset=0 不改该段，全部目标预检后才写。普通字体操作不采用 clear 的扩词规则。 |
| story 局部字符格式 | 字符入口由实际 header/footer 选区、UTF-16、节/精确 story/通用 header-footer 权限及内容锁决定可用性。所选 literal runs 与字段对象分开处理；字段 cache 的每个运行只更新请求属性，保留不同字号/字色、链接和空实体。字段对象拒绝 href 或 metadata 的整体替换。 |
| 汉字与混合工具值 | [document-toolbar-state.ts](../../packages/office-studio/src/components/document/document-toolbar-state.ts) 读取活动正文或 story 的真实所选文字和字段显示 cache；零长实体不制造混合字符值，字段空显示使用其样式回退。字体 patch 在选中文字或字段 cache 含汉字时同步东亚槽；不把一个 U+FFFC 当作字段全部脚本。混合字体/字号为空值提示，其余混合字符属性仍逐项记录。 |
| 折叠范围的字符偏好 | 公共 API 和 standalone 均只改变临时输入 marks，不改运行 ID、段落、内容历史或保存点。颜色/高亮的显式清除保留属性意图；权限、正文控件锁与 story 锁和实际输入共用判断。重复相同偏好不制造内容历史。 |
| 正文、标题、引用、列表和表格内段落 | 保留反向、半开范围，只清除选中文字的字符格式；相交段落恢复文档默认段落格式及普通正文，清除标题/命名样式/列表/清单和 HTML 缓存。末段 offset=0 不包含该段；表格只改指定段落并使容器 HTML 缓存失效。 |
| 折叠光标 | 复用 [document-format-painter.ts](../../packages/office-studio/src/features/document/services/document-format-painter.ts) 扩词，无法扩词时无内容变化。使用 Intl.Segmenter；与上游和 Office 多语言词边界的差异尚未实测。 |
| 来源与实体 | 保留未选 run、链接、来源、零长度引用及实体 ID，只拆分发生变化的字符范围。重复清除相同默认格式不新增内容历史。 |
| 页眉页脚 | 预检节/精确 story/通用 header 或 footer 的权限、节是否存在、前节关联、UTF-16 与全部相交段落锁。正文模型和正文选区不变；字符范围只映射到对应 literal runs，段落恢复该文档默认格式，字符字号使用 9pt。 |
| 字段缓存 | 选中一个字段时逐 cache run 清除，保留每个 ID、文本、链接和零长度实体；指令、offset、缓存文本不变。显示与 writer 同样逐 run clear，避免合成单 run 丢失 hyperlink。即使主样式原本等于默认，也清除其他 cache run 的不同属性。 |
| clear 后再局部格式 | 已物化的 cache 保留 source clear 标记，更新字段主样式与对应 baseline；后续只改字体或字号时，其他 cache 区间不重新压成默认值。显示与 writer 应继续抑制旧来源 rPr，同时写出后来合法的 mixed 属性；TS 和 Rust 对照草稿均未执行。 |
| 正文/表格锁与原子性 | 段落格式影响整段，因此检查词外锁；旧缓存编号前缀删除另检查包装锁。所有目标验证后才写入，后段拒绝不能先改前段。编号、评论/控件偏移和选区变换复用 DOC-05。 |
| story 锁与直接 API patch | [document-story-locks.ts](../../packages/office-studio/src/features/document/services/document-story-locks.ts)、[section-header-footer.ts](../../packages/office-studio/src/features/document/services/section-header-footer.ts) 检查文本、字段、段落与源格式。等价 metadata 键顺序不误判；锁前编辑平移范围，content 锁拒绝内容/格式修改，wrapper-only 锁允许内部编辑并保留空包装，直接 patch 不能绕锁。下一段起点的空锁不阻止前段，末段 endpoint 仍保护段落标记。 |
| 代码块局部 clear | 确认 runText 等于 code source；只对所选范围设置 `metadata.codeSyntaxHighlight = false`，保留语言、source、lineNumbers 与未选范围。后续 [code-blocks.ts](../../packages/office-studio/src/features/document/code-blocks/code-blocks.ts) 输入/删除同步 source 与 runs，不重新压成单 run；Canvas 按范围抑制语法色。 |
| 取消、拒绝、no-op | 取消/拒绝不提交模型或内容历史，重复 no-op 不增加历史。现有 `executeCommandSync` 仍发布成功返回的 CommandExecuted，即使 `changed=false`；事件条数不等于内容变化次数。 |
| 历史与保存 | [document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts) 和公共 API 复用同一服务，内容历史按文档 ID 隔离；story、cache、代码 source/格式范围和 defaults 随快照保存。非活动文档调用的隔离已写契约草稿。 |
| 下一次输入与焦点 | run-range 的替换继承第一个选中字，插入继承左侧 run，保留 clear/syntax 标记而不复制 note/公式引用。字段替换继承第一个非空 cache 字符，字段后的光标继承最后一个；不将字段链接或 source metadata 复制为新文字实体。成功普通输入保留捕获的输入偏好，成功 clear 清偏好/格式刷，取消或拒绝保留原输入意图。 |

以上接入不等于完整覆盖 DOC-03、DOC-13 或全部 Pro 节点。未实现容器、特殊字段与完整资源映射继续保留在第 7 节，不以提示、空操作或永久禁用代替实现。

story 最终输入携带源 range、caret、modelRevision 与 generation；共享 API 先核对真实拼接和权限/锁，再把内容写入一次历史，随后只授权一次 `advance-input-marks`。Workbench 复查结果 revision、generation 与精确 caret；`selectionCurrent: false` 或提交观察者已导航/重新格式化/关闭时，读取当前有效会话并停止旧结果投影，缺少会话时不再读取已关闭 Facade 的模型。body 最终输入采用共享命令和同样的会话版本边界。

Workbench 的 `undo` / `redo` 事件新增可选来源 `unitId`，活动且 facade/model ID 一致时发送当前文档 ID，由宿主复用公共历史路径。工作台本身不维护另一个内容栈，也不在发出 undo 前清掉输入偏好；取消保留偏好，成功 replay 的清理与 Before/完成事件、历史发布次序由主任务的 [COMMON-03 历史证据](COMMON-03-input-selection-axis-history.md#5-历史观察事件与取消)说明。生产接线和相应草稿由主任务处理，本页不宣称其运行结果。

## 4. DOCX 写出与共享源样式

- [docx_default_style.rs](../../crates/office-writer/src/docx_default_style.rs) 将领域默认属性写入受控模板 w:docDefaults；非法属性拒绝，无字段时用旧模板。[docx.rs](../../crates/office-writer/src/docx.rs) 共用正文/表格/注释 run 属性，显式关闭强调/上下标/高亮避免错误继承。标准 RGB 高亮写 Office 名称，自定义 RGB 用 highlight=none + shd。
- [document_format_sidecar.rs](../../crates/office-parser/src/document_format_sidecar.rs) 按范围恢复 nil shading，安全拆分 Unicode runs，保留范围外背景、链接与空实体。段落读写覆盖对齐、缩进、行距/段距、wordWrap、snapToGrid；净首行/悬挂缩进归一化不冒称原 XML 组合逐字保留。
- [story_fields.rs](../../crates/office-parser/src/story_fields.rs) 提取 literal runs/paragraphs、受支持字段的 mixed cache、源 rPr/pPr、命名空间、关系及可映射 SDT 锁。DTO、领域模型、显示布局和输入对象索引分别保存。
- [docx_story.rs](../../crates/office-writer/src/docx_story.rs) 先验证全部文本、段落、字段、cache、锁和偏移再写 story。cache 参与源 baseline 比较；未改且资源可恢复时重放源 XML，可支持编辑后按模型生成。fldSimple 指令保留，clear 后每个 cache run 生成清除后的 rPr 并保留 hyperlink/关系。
- [docx_source_styles.rs](../../crates/office-writer/src/docx_source_styles.rs)、[docx_source_story.rs](../../crates/office-writer/src/docx_source_story.rs)、[docx_source_xml.rs](../../crates/office-writer/src/docx_source_xml.rs) 与 DOC-05 共用 SourceStyles，恢复 default/style/direct 属性、依赖闭包、冲突 ID remap 和主题；只 overlay 改变的已建模属性，保留未改高级属性，明确 clear 则抑制原 rPr/pPr。
- 完整 story 遍历携带祖先 namespace scope；属性拆出和 tabs/borders/paragraph mark 合并保留各自子作用域。strict/alternate prefix 转 transitional，序列化前拒绝绑定冲突。**完整 prefix remap 未实现**，同一元素合并冲突 extension attributes 仍可能拒绝。
- 同一 registry 供 story 和原编号 marker 使用；DOC-05 编号结构、序列与前缀清除不在本阶段重做。

源格式抑制不等于 XML 永远没有 pStyle/lang。第二次合法重放可能物化普通 Normal 或 writer 的普通 zh-CN 默认；草稿要求旧自定义高级属性和源字符语言不返回。语言/脚本/主题的完整编辑和默认保真仍未闭环。

此链路使用本项目 parser/writer，不依赖 Univer Server、Exchange 或私有协议。Rust 路径与原生往返草稿尚未编译，也没有本轮 WASM/独立 Office 重开结果。

## 5. Canvas 输入与 antdv-next

clear、字体/字号及其余字符入口沿用 [RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 与 [RibbonEditableSelect.vue](../../packages/office-studio/src/platform/ribbon/RibbonEditableSelect.vue) 的 antdv-next 控件、主题和共享溢出。正文和 story 都使用 [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的 Input.TextArea，未添加原生表单控件；内容仍为 Canvas，未以整页 contenteditable 或组件库 Table 替代内核。

[use-document-story-input.ts](../../packages/office-studio/src/components/document/use-document-story-input.ts) 只保存受控 draft。beforeinput 捕获真实范围与 marks，重复文字替换和同文改写按捕获范围解释；拒绝后恢复 buffer/选区、清掉重复提交抑制并允许相同重试。composition 捕获固定范围、模型与 marks，只预览，结束提交一次。恢复任务带 epoch、文档/页面/story scope；模型替换、文档切换、只读、story 切换和 KeepAlive deactivation 取消旧 composition 并替换 TextArea。

`captureInputSession(sourceModel?: DocumentModel)` 的无参调用是被动 observer getter；只有正文 composition 开始、story beforeinput/composition 或字段插入等真实输入开始时才传入 Canvas/Workbench 模型。clear、undo 或内部重开后缺少有效会话时，通过公共 `setInputSelection` 建立空 marks 会话，并再次检查 facade、活动状态、模型身份及公共快照，拒绝 Before/观察者重入后的旧模型。story 还须核对 section/story 与真实 DOM 范围；公共 composition 不再以缺少 token 为本地 draft fallback，standalone 不传公共 callback。

Vue 尚未把更新模型传给 Canvas 时，旧模型捕获不能重建公共会话。正文/story 在捕获失败时立即阻止旧 TextArea、恢复缓冲并递增 key；后续旧 composition/input/history/keyboard 事件不能提交或撤销新内容。已有 composition 也在公共 generation/revision、同位置字符偏好或同 ID Facade 重绑后取消。新增快照 JSON 比较及克隆成本未测量，不记为性能优化。

旧元素在 keyed 替换前即被阻止；正文/story 的 beforeinput/input/composition/keyboard/selection/focus 检查真实目标与活动状态，旧 historyUndo 不能影响新模型，正文同样过滤旧剪贴板事件。focus/blur 传入实际 FocusEvent，避免旧元素改变新上下文的光标显示。

[story-rendering.ts](../../packages/office-studio/src/features/document/services/story-rendering.ts)、[document-layout.ts](../../packages/office-studio/src/features/document/layout/document-layout.ts)、[document-scene.ts](../../packages/office-studio/src/features/document/rendering/document-scene.ts) 共用字段原子范围：mixed/跨行字段 glyph 保持单一输入对象，尾光标落在最终片段，空段按插入字体度量；代码语法色按 run 范围应用，lineNumbers 来自持久化代码模型。

浏览器草稿装配真实 ConfigProvider、KeepAlive、Workbench、OfficeStudio/FDocument 和 antdv-next 1.8.2 TextArea；视口 1440×900、DPR 1.25、Chromium msedge、classic/grid、明暗主题。**尚未启动 fixture**。DOM 合成 composition 不代表真实 Windows 中文 IME、候选框位置、系统剪贴板、主题/焦点或 WebGPU/Canvas2D 验收。

## 6. 已写入但未执行的测试

| 文件 | 草稿场景与限制 |
| --- | --- |
| [document-clear-formatting.contract.test.mjs](../../packages/office-studio/tests/document-clear-formatting.contract.test.mjs) | 正文扩词/半开反向范围/defaults/no-op/链接/空引用/表格/权限/词外锁/非法默认/颜色、DOC-05 缓存编号与评论/控件变换、跨模块原子拒绝、clear 后输入。 |
| [document-partial-text-formatting.contract.test.mjs](../../packages/office-studio/tests/document-partial-text-formatting.contract.test.mjs) | 局部字体/字号只改请求属性、半开反向/跨段落、坏 patch 原子拒绝、source/链接/空 note 保留及拆分归属、多行输入、权限/锁；field-only 混合值、逐 cache 字体、clear 后 mixed 与可选属性清除、字段首/尾字符继承、导入大字号继承、重复同文/Unicode 原生范围、真实 Facade 一次历史/取消/内部重开。 |
| [document-input-session.contract.test.mjs](../../packages/office-studio/tests/document-input-session.contract.test.mjs) | 主任务草稿：折叠偏好与内容分离、own undefined/嵌套 metadata getter、连续正文/CJK/索引桥、无效位置/权限/锁、导航/格式/输入取消、clear/undo/redo、Before/提交观察者重入、story 拼接及一次 advance。 |
| [history-publication.contract.test.mjs](../../packages/office-studio/tests/history-publication.contract.test.mjs) | 主任务草稿：历史应用完成后发布、回调失败保留重试、同 unit 重入拒绝、跨 unit 独立、settled listener 在提交结束后启动下一次编辑；不等于持久版本服务。 |
| [document-story-formatting.contract.test.mjs](../../packages/office-studio/tests/document-story-formatting.contract.test.mjs) | story UTF-16/字段对象/正文不变、锁与直接 patch/空锁段落边界、DTO cache、逐 run clear 保留链接/空实体、主样式已为默认的 mixed cache、跨行字段命中/尾光标、空段字体、代码局部 clear/输入/Canvas 色。已修正 page 字段，未执行。 |
| [document-editor-commands.contract.test.mjs](../../packages/office-studio/tests/document-editor-commands.contract.test.mjs) | 公共正文命令/事件、取消/权限/锁、一次历史、undo/redo、save/reopen defaults、no-op 与下一次输入。 |
| [document-story-api.contract.test.mjs](../../packages/office-studio/tests/document-story-api.contract.test.mjs) | SSR 载入真实 OfficeStudio/FDocument：story/cache clear/edit/history/save/reopen、正文选区不变、取消/权限/代理对/词外锁、直接 patch 锁平移、非活动文档隔离、代码输入重开。内部快照不是 DOCX 验收。 |
| [document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) | 内嵌 HTML 装配真实 TextArea/Workbench：classic/grid 字体/字号连续输入、Canvas 正反拖选/clear/history/内部重开、重开后同范围键入继承清除后的格式；field-only 混合缓存、外部 body/story 选区与清色；DOM composition 单提交、clear 后原范围重建有效会话、Vue 更新前旧 Canvas 模型捕获失败及旧事件阻止；五种宿主取消、公共导航/同位置格式/同 ID 重绑、提交观察者导航/关闭、同文取消恢复重试、canceled undo 保留偏好和来源 unitId。仅验证草稿中的合成事件意图，未运行，不是系统 IME、生产历史接线或 Office 文件验收。 |
| [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | 已有正文 clear、主题/焦点/立即输入/history/reopen、取消保留 pending、权限禁用；含 DOC-05，历史结果不替代本切片。 |
| [docx_source_story_tests.rs](../../crates/office-writer/src/docx_source_story_tests.rs) | story/编号共享 remap/依赖、strict/alternate prefix/theme、overlay/clear、paragraph mark、同名 extension prefix、tabs/borders 合并、坏 source/resourceErrors、合法缺少 styles、numPr 拒绝。小主题夹具只测 registry，不是完整 Office 样本。 |
| [document_story_roundtrip_tests.rs](../../crates/office-parser/src/document_story_roundtrip_tests.rs) | 主任务合成真实 ZIP parts：A😀 word + 换行 Tail、offset=3 DATE、Georgia 18 红色和 Aptos 12 蓝色链接 cache、高级 rPr/pPr/主题/Tail 锁；两次其他文字编辑/导出/parser 重读；clear 后两次保存保留缓存链接，拒绝 cache 不一致/声明 styles 丢失。新增 `cleared_field_cache_retains_later_mixed_local_edits_and_links_on_two_docx_saves`：DATE 的 `2026/` + `11/22` cache 清除后再改 17/9pt、bold/italic、yellow 高亮和链接，经过两次本项目 native DOCX 写出/parser 重开。helper 构造领域协议，**未执行、未调用 TS normalizer/WASM 或独立 Office**，不是全浏览器转换链。 |
| `document_default_style.rs`、`docx_default_style.rs`、`document_format_sidecar.rs` | 默认继承/主题/缺省差异、非法资源/颜色/布尔值、高亮 XML、nil shading Unicode 范围。 |
| [document_format_roundtrip_tests.rs](../../crates/office-parser/src/document_format_roundtrip_tests.rs) | 默认格式、显式关闭强调/上下标/高亮、自定义 shading、DTO 重导出后二次读取。 |
| [document_legacy_default_style_tests.rs](../../crates/office-parser/src/document_legacy_default_style_tests.rs) | 草稿构造真实 CFB、FIB、STSH/FFN/PAPX 记录和 ODT ZIP：全局继承与 42pt 首段不同、三种字体槽、单位/段距、FIB 范围溢出/越界/零长、内部截断/字体槽坏值、Table 回退、ODT 空/缺失/重复/不一致默认定义、坏 XML 与必要资源缺失。未编译，夹具语义也未验收。 |
| [document_legacy_source_tests.rs](../../crates/office-io-common/src/document_legacy_source_tests.rs) | JSON 字面 XML/hex/顺序/诊断保留与稳定重复编码、不同资源不误合并、未知字段/重复 JSON key/版本/形状拒绝、UTF-8 名称与 u32 范围、整体来源/资源/诊断/原始字节预算、JSON 控制字符转义膨胀拒绝。大边界草稿不是性能测量。 |
| [docx_legacy_source_tests.rs](../../crates/office-writer/src/docx_legacy_source_tests.rs) | 当前来源与历史合并及诊断迁移、当前 ODT 不注入 Word 缺省、历史 ODT 不阻止 Word 样式继承/冲突 remap、非法来源在返回 DOCX 字节前拒绝。复用 DOC-05 registry，未重做编号。 |
| [document_legacy_source_sidecar_tests.rs](../../crates/office-parser/src/document_legacy_source_sidecar_tests.rs) | 专用 part/关系/内容类型、当前 Word defaults 与历史诊断隔离、无 sidecar 的旧 DOCX、声明资源缺失/错误类型/外部或越界路径/重复关系/命名空间属性冒充/多 part/坏或超预算 JSON 的拒绝，以及 URI 解码边界。 |
| [document_legacy_source_roundtrip_tests.rs](../../crates/office-parser/src/document_legacy_source_roundtrip_tests.rs) | 原 CFB DOC/ODT ZIP 草稿经本项目 native writer 两次 DOCX 导出/parser 重开，核对原 binary hex、含未知 ODF 扩展的 XML、历史去重、当前 Word 来源与默认字号、正文文本。helper 构造领域 DTO，**未调用 TS normalizer，也没有真实 Office 程序或 WASM 链路结果**。 |
| [document-legacy-source.contract.test.mjs](../../packages/office-studio/tests/document-legacy-source.contract.test.mjs) | DOC/ODT DTO 归一化→真实 OfficeStudio/FDocument 公共 clear→no-op/一次历史/undo/redo→内部 save/reopen→下一次输入，原资源不变；DOCX 原生 defaults 与历史诊断隔离；来源形状、UTF-8/u32 与整体预算拒绝。SSR 与内部快照不能替代 Office 文件或浏览器验收。 |
| [document-numbering-roundtrip.spec.ts](../../tests/browser/document-numbering-roundtrip.spec.ts)、[配置](../../tests/browser/document-numbering-roundtrip.playwright.config.ts) | 既有生产预览/真实 writer WASM/上传/clear/history/两次下载/独立 parser Worker 重读/重导入草稿，本轮未执行。本项目另一个 Worker 不等于独立 Office 程序。 |

当前所有结果均为**未运行**。恢复运行验证后须分别完成类型、前端、原生 Rust/WASM、文件、浏览器、系统 IME 和性能检查，记录通过/失败、环境、样本 hash、诊断与未覆盖场景，再更新验收。测试通过也不代表整个功能目录已经覆盖。

## 7. DOC/ODT 来源持久化、高级字符与剩余目标

静态对照依赖 `duke-pages 0.6.0`，本机源码根：`C:/Users/15487/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/duke-pages-0.6.0/src`。DOC/ODT 默认来源读取与历史 sidecar 已实现代码和测试草稿，**运行验证仍为零；高级字符领域/Canvas 能力未补齐**。

| 依赖文件与符号 | 对照及实现差异 |
| --- | --- |
| `lib.rs::parse_document_model_with_options`、`doc/build.rs:151` | DOC/DOCX 先 Word package，再 build_doc_model_with_options；ODT 直接 parse_odt。当前 Docs parser 显式一次 `parse_doc(bytes)`，同一 package 建模型与默认定义，不第二次完整 DOC 转换。 |
| `doc/stsh.rs::StyleSheet::to_styles_xml`、`doc/build.rs:1795` | STSHI 标准字体槽生成 docDefaults，`to_styles_xml(&self.fonts, 20)` 用 10pt 合成基线。当前默认链复用它并解析 Normal/basedOn；原二进制与合成 XML 分离。该基线/继承是依赖库契约，尚未用独立 Office 样本核实。 |
| `doc/cfb.rs::CompoundFile::parse/try_stream`、`doc/fib.rs::Fib::fc_lcb/table_stream_name`、`doc/fonts.rs`、`doc/stsh.rs` | 当前捕获 STSHF=1、STTBF_FFN=15，严格 u32 加法和流边界；结构坏但范围完整时保留原字节并诊断。检查完整 STSHI/STD/FFN 和非空字体槽；FIB 指定 Table 缺失才回退并记录实际来源。公开 CFB reader 会再次读取/复制流，成本未测。 |
| `odt.rs::Builder::format`、`:1881 style_definitions`、`:1944 默认 ID`、`:2049 style_definition` | 使用 StyleSheet.resolve_paragraph/resolve_text；Standard 映射 Normal，无 Standard 时由 family defaults 建立 Normal。当前只从默认 ID 和唯一 `is_default=true` 定义取基本 run/paragraph 属性；styles.xml 可缺省、content.xml 必需，XML 结构错误或默认标记矛盾不产生 defaults。 |
| `model.rs::TextStyle` | 已有 caps/smallCaps、shadow/outline/doubleStrike、Latin/HAnsi/EastAsian/CS 与 theme/resolved 字体槽、language/bidi/RTL/complexScript、字距/缩放/kerning/ligatures/contextualAlternates/runBorder。本项目仍主要映射基本属性，高级 DTO/XML 留在 source。 |

[docx_legacy_source.rs](../../crates/office-writer/src/docx_legacy_source.rs) 在任何 package 写入前准备来源，随后由 [docx.rs](../../crates/office-writer/src/docx.rs) 注册：

| 包契约 | 值 |
| --- | --- |
| Part | `word/office-studio/legacy-sources.json` |
| Content type | `application/vnd.office-studio.document-legacy-sources+json` |
| document relationship type | `urn:office-studio:relationships:document-legacy-sources:1` |
| document relationship target | `office-studio/legacy-sources.json`，内部关系，rId 由现有 context 分配 |
| JSON | `{ version: 1, legacyOrigins: Origin[] }`；仅历史来源数据 |

[document_legacy_source_sidecar.rs](../../crates/office-parser/src/document_legacy_source_sidecar.rs) 校验一个当前 Office document owner、唯一专用关系与已注册 part、内容类型解析的一致性、命名空间与非外部路径；兼容合法 Default 或 Override 的内容类型声明。包资源不存在且无任何声明时沿用旧 DOCX 路径；已有声明或专用 part 却缺失/损坏/歧义/超预算时返回解析错误，不能当作合法缺省或丢弃历史继续成功。

合并顺序是当前 DOC/ODT `origin`，然后已有 `legacyOrigins`。身份由格式、合成标记、资源顺序、名称/offset 和内容共同决定；XML 按字面文本，binary hex 按字节值比较、忽略 hex 大小写。保留首次资源文本/顺序，合并不同诊断；资源字节、offset 或顺序改变时保留独立来源。当前顶层错误只合入对应当前 DOC/ODT origin，不迁移 DOCX 原生 Word 错误。重开只还原历史，不把它当成当前 Word 样式；`SourceStyles` 只跳过当前 ODT 来源。

整体边界在去重后校验：最多 32 个来源、1024 个资源；origin 资源的 XML 每 part 16 MiB，当前/历史 origin 中原始资源合计 32 MiB，其中 binary 合计 16 MiB；资源名最多 1024 UTF-8 字节。这些 origin 的诊断合计最多 1024 条、每条 4096 UTF-8 字节、总计 1 MiB。DOCX 当前 Word styles/theme 各自沿用 16 MiB part 限制，顶层当前资源诊断另受同样的诊断限制。专用 JSON 的实际编码/读取上限为 64 MiB，计入 JSON 转义膨胀；数组/字符串反序列化受限，输出超限返回错误，不写截断资源。TS 使用同一规则，拒绝孤立 Unicode 代理项。这些是协议限制，**不是内存/性能达标证据**。

sidecar 保存原始样式资源快照；编辑后的当前 DOCX 内容由领域模型写出。DOC 当前保留的二进制范围为 STSHF/STTBF_FFN，ODT 保留上述 XML；**这不等于完整原文件、全部关系资源或所有高级语义已能编辑/渲染/重放**。没有独立 Office 重开/另存结果，尚未确认外部程序是否持续保留此应用专用 part。

固定版 [i-document-data.ts][u-note-model] 的 `IDocumentNote.body` / `IDocxNoteStory.body` 和 [IBaseComment.text][u-comment-model] 均为 `IDocumentBody`；[ThreadCommentEditor.tsx][u-comment-editor] 使用 RichTextEditor 及评论自身选区/焦点。本项目 [domain/model.ts](../../packages/office-studio/src/domain/model.ts) 的 `DocumentNote.text` 和 `DocumentCommentMessage.text` 仍是字符串，尚未补富文本模型、Canvas story、独立输入/格式会话、历史及 Office 保真。页眉页脚接入不能被写成脚注、尾注或评论格式能力已经完成。

仍需完成：

- 高级字符、复杂脚本字体、主题 tint/shade、语言/字距的完整领域模型、编辑、Canvas shaping/测量和 Office 写出；来源保留及 clear 抑制不能替代这些能力。
- DOC/ODT defaults、历史来源 sidecar 的类型/原生/WASM/浏览器与真实 Office 样本验证；当前只有静态实现和未执行夹具草稿。DOC/ODT 原始资源重新解释、完整样式语义与其他格式写出仍需补齐。
- story 原 numPr/tblStyle、内嵌关系资源、复杂 SDT、未知字段与表格等完整可编辑重放。受支持编辑之外当前拒绝写出并保留资源/诊断，不能把拒绝作为最终完成。
- 字段重算后 mixed cache 的自动迁移；目前 cachedResult/cache 必须同时更新，不一致拒绝。只更新日期缓存文本的 UI 刷新路径尚未补齐 mixed cache 同步。
- 全部正文/高级容器的来源格式保真、内容控件/脚注尾注/评论/嵌套表格/Pro 节点跨模块闭环；本切片不证明 DOC-11/12/13/22 完整支持。
- 六种页眉页脚 story 的全量布局与交互覆盖；公共 story 选区在 modern 模式或不存在对应 layout story 时，当前 Workbench 导航仍直接返回，尚未实现完整外部 UI 定位。脚注、尾注和评论富文本/Canvas 编辑属于明确待开发项，不改变最终范围。
- 独立 Office/兼容程序、公开转换演示行为对照、真实 Windows IME/系统剪贴板、主题/焦点/缩放及 WebGPU/Canvas2D 各自运行验收。
- 长文档首屏、滚动、编辑、缩放、Worker/WASM、传输、取消和内存测量；附着 namespace scope、额外 CFB 全流读取/复制、binary hex、ODT XML、sidecar 编解码、输入捕获的公共快照克隆与 JSON 比较成本/峰值内存均未测量。本轮无优化/达标宣称。

## 8. 文件归属与交接

本阶段涉及 Docs defaults/story 模型、clear 服务、story 锁/编辑/显示、代码范围、Canvas 布局/场景、Workbench/TextArea composable；Rust parser DTO/default/story/source、writer default/story/SourceStyles；对应草稿见以上链接。共享 domain/model.ts、Rust model.rs/parser.rs/lib.rs 只修改已协调 Docs 段，保留其他任务的 Sheets 改动。

本次续做另新增 office-io-common 的 Docs 来源模块与草稿，只在 common lib.rs 增加该模块声明；parser 增加当前/历史来源字段、默认来源和 sidecar helper；writer 只局部装配 Docs preflight/part/relationship 与当前 ODT 判断；TS 只增加 parser 契约/来源校验和专项草稿。没有新增 UI 控件，沿用前述 antdv-next/Canvas 输入实现；主题、焦点及真实 IME 仍未验收。

本轮局部字符格式与输入续接修改 `DocumentWorkbench.vue`、`DocumentCanvasViewport.vue`、`use-document-story-input.ts`、`document-ribbon.ts`、`document-toolbar-state.ts`、`document-commands.ts`、`document-run-range.ts`、`document-story.ts`、`document-story-locks.ts`、`story-editor.ts`，新增 text-formatting/story-formatting 服务及局部格式契约草稿，补既有 story/API/浏览器草稿。静态回读最新会话捕获、旧元素阻止和浏览器断言后，补入本页；没有执行任何代码验证。

接口增量为 `DocumentCommand.format-text.story?`、`set-section-header-footer.input?`、story 编辑 helper 的可选 marks/sourceRange、`DocumentRibbonContext.textFormattingDisabled?`、Canvas 的 `captureInputSession(sourceModel?: DocumentModel)` 回调，以及 Workbench 的 `undo: [unitId?: string]` / `redo: [unitId?: string]`。公共装配 facade.ts、office-studio.ts、index.ts、输入/历史 store 与[功能清单](../UNIVER_FEATURE_PARITY.md)、[工具栏清单](../UNIVER_TOOLBAR_PARITY.md)汇总继续由主任务负责；本轮复用 FDocument 的 clear/executeCommand 与输入 state/selection/advance 接口，不覆盖这些共享文件。DOC-05 编号和共享源样式方案继续沿用；本页不提升验收状态。

<a id="home-format-composition"></a>

## 9. Home 字符格式与清除的组字状态续核

续核日期：2026-11-27（Asia/Shanghai）。本批修正已写入代码，**运行验证为零，状态仍为未验收**。本节的静态检查结果只适用于本批文件，不替代前述实现与草稿的运行验证。

固定版 [docs-ui/menu/schema.ts][u-doc-schema]、[menu.ts][u-doc-menu] 和 [inline-format.command.ts][u-inline] 确认以下既有入口；新增根入口为 **0**，不重新计算完整 Ribbon / Pro 分类总数。

| 既有入口 | 固定上游布局及命令 | 本项目命令 |
| --- | --- | --- |
| 下标、上标 | START / FORMAT，classic `order: 4/5`；grid 下标 `row: 2, column: 6`，上标 `row: 2, column: 5`；`doc.command.set-inline-format-subscript/superscript` | `doc.command.format-text` 的 `verticalAlign` patch |
| 文字颜色、高亮 | START / FORMAT，classic `order: 8/9`；grid 文字颜色 `row: 2, column: 8`，高亮 `row: 2, column: 7`；`doc.command.set-inline-format-text-color/text-background-color`，另有 reset 子项 | `doc.command.format-text` 的 `color/highlight` patch 或 `clearFields` |
| 清除格式 | START / HISTORY，classic `order: 3`；grid `row: 2, column: 2`；`ui.command.clear-formatting` | 共享别名转为 `doc.command.clear-formatting` |

上游 `disableMenuWhenNoDocRange` 检查文本/矩形范围、相交代码块及编辑权限；`SetInlineFormatCommand` 用 `va` 翻转上下标，以 `cl/bg` 写颜色，折叠范围缓存下一次输入样式，展开范围进入 `RichTextEditingMutation`。本地继续使用既有 patch、权限、锁、文档隔离输入偏好与历史链路；本批没有更改这些数据语义。本地输入法守卫来自真实 TextArea 状态，不把上游上述菜单条件写成上游已经证明了本地组字行为。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 原先的 `textFormattingState`、`clearFormattingState` 只检查 `storyComposing`。正文、表格与富内容 story 使用 [DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的普通 `antdv-next` `Input.TextArea`，其组字状态通过既有 `isBodyComposing()` 暴露；旧式 flat story 使用独立 TextArea 和 `storyComposing`。本批仅修改三个区段：两处可执行状态统一读取既有 `documentInputComposing`，`execute` 对 `doc.command.format-text` 复用同一活动输入/组字守卫；clear 执行继续检查修正后的 `clearFormattingState`。延迟的 Ribbon 回调因此在 Workbench 输入边界返回，不进入模型 mutation、输入偏好或内容历史。公共 API、领域模型、保存协议及 Worker/WASM 接口增量均为 **0**。

| 本批浏览器草稿 | 计划用例及断言意图 | 本批结果 |
| --- | --- | --- |
| [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | 正文 2 例：classic 明色、grid 暗色；组字时所有字符按钮、字体/字号、两个颜色主按钮及菜单箭头、clear 禁用；用真实 Ribbon 组件事件发送延迟格式/clear，模型、历史及 `CommandExecuted` 不变；中文提交一次，之后下标、clear 各一次且恢复输入焦点 | 未运行 |
| [document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) | flat/rich header × classic 明色/grid 暗色，共 4 例；flat 保留 UTF-16 字段对象并替换文字，rich 检查实际 `{ kind: 'story', sectionId: 'first', story: 'header' }` 宿主；同样检查禁用、延迟回调、单次提交、格式/clear 恢复、正文不变与宿主焦点 | 未运行 |

本批共 **6 个计划用例、0 个运行结果**。表格和代码块的输入来源已静态追到同一正文 TextArea，但本批未新增这两类宿主的专用浏览器用例；不能据此声称六种 story、全部容器或代码格式权限已经覆盖。沿用既有 `antdv-next` 控件与统一 ConfigProvider，没有新增 UI；主题、焦点、真实 Windows IME、WebGPU/Canvas2D、保存/重开、独立 Office 样本和性能均未新增运行验收。第一批富 story 导航证据仍见 [DOC-03 的独立记录](DOC-03-header-footer-ribbon-sizing.md#rich-story-navigation)，不能与本批 6 例混计。

本批限定静态检查：两个 `.mjs` 的 Node 语法、内嵌 HTML module 语法、Workbench Vue SFC / TypeScript AST 语法，以及两个草稿各 11 项本地 import 路径通过。DOC-02 / DOC-03 的本地链接、固定提交源路径、引用锚点及五个归属文件的限定 `git diff --check` 通过，未发现冲突标记。静态解析未执行测试、类型检查、应用构建、Rust/WASM、dev server、浏览器或基准，不提升功能验收状态。

<a id="fixed-vertical-shortcuts"></a>

## 10. 固定 Docs 上下标快捷键续核

续核日期：2026-11-27（Asia/Shanghai）。本批补齐已有上下标入口的快捷键接线，**未运行、未验收**；常驻、扩展、溢出、上下文根入口增量均为 **0**。

| 固定来源 | 已静态确认的契约与差异 |
| --- | --- |
| [docs-ui/shortcuts/toolbar.shortcut.ts][u-toolbar-shortcuts] 的 `SubscriptShortCut` / `SuperscriptShortCut` | 下标绑定 `CTRL_COMMAND + COMMA`，上标绑定 `CTRL_COMMAND + PERIOD`，两者均无 Shift。不能把 Sheets 的 Ctrl/Cmd+Shift+Comma/Period 字号步进登记为 Docs 同一契约。 |
| [docs-ui/controllers/ui.controller.ts][u-doc-ui-controller] 的快捷键注册、[shortcuts/utils.ts][u-shortcut-utils] 的 `whenDocAndEditorFocused` | 两个快捷键均实际注册；要求 Docs 与编辑器焦点，排除 common drawings。权限与内容范围继续由字符命令承担。 |
| [inline-format.command.ts][u-inline] 的 `SetInlineFormatCommand` | 上下标共享 `va`，相同值再次执行回到基线；折叠范围保留输入偏好，展开范围写入真实内容历史。 |

[DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的 flat `storyKeys` 与普通 `keydown` 原先只有 B/I/U 字符快捷键；rich story 复用普通 body bridge，[use-document-story-input.ts](../../packages/office-studio/src/components/document/use-document-story-input.ts) 提供 flat 当前输入元素、范围和组字状态。两条键盘路径现在识别物理 `event.code = Comma/Period`，要求 Ctrl 或 Meta、无 Shift/Alt、实际对应 TextArea 焦点且无图片选区，沿用现有宿主与 composition 校验。flat 继续传递源模型与实际 story 范围，保留 stale source、公共输入会话与权限约束；未新增 helper 或输入协议。

两条路径都复用既有 `toggle-format` 事件。[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 根据 `currentMarks.verticalAlign` 构造 `doc.command.format-text` patch，再走已有组字/活动输入、权限、锁、历史与焦点恢复链路。事件名、参数形状、命令 ID、公共 API、模型、保存及 Worker/WASM 协议均未新增；**既有内部事件的 `field` 联合扩展 `subscript/superscript` 两个取值**，并同步修改唯一 Workbench 调用方类型。基线 patch 按 [document-model.ts](../../packages/office-studio/src/features/document/model/document-model.ts) 的 `applyRunPatch` 删除显式 `verticalAlign`，草稿断言使用 `undefined`，不要求领域 run 保留字符串 `baseline`。

本批浏览器草稿新增 **6 个计划用例、0 个运行结果**：正文 classic 明色/grid 暗色 2 例，flat/rich header × 两种布局主题 4 例。两份既有文件检查 Ctrl/Meta 的相同值切换与上下标互斥、字号及正文/字段不变、Shift/Alt 和失焦无操作、真实 undo/redo、折叠偏好无内容历史及随后输入继承、权限变化阻止 mutation；story 另检查公共内部 save/reopen。前节 6 个组字草稿追加快捷键阻止断言，不新增该节用例数量。正文用例未新增保存/重开路径；内部快照也不等于 Office 文件验收。

本批没有新增 UI，继续使用统一 ConfigProvider 与 `antdv-next` TextArea。图片选区排除已静态核对，尚无本批专用 drawing 浏览器用例；真实 Windows/Mac 快捷键、系统 IME、全部容器与六种 story、主题焦点、独立 Office、WebGPU/Canvas2D 及性能仍未验收。本批两份 Vue SFC / TypeScript AST、两个 `.mjs` 及内嵌 module 语法、两个草稿各 11 项本地 import、DOC-02 / DOC-03 本地与固定版源链接及显式锚点、六个归属文件的限定 `git diff --check` 均通过；仅执行静态检查，测试、类型、构建和浏览器继续后置。

本节提示元数据续补：[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的 `verticalFormat` 原先没有 `shortcut`，现按同一固定绑定登记下标 `Ctrl+,`、上标 `Ctrl+.`，无 Shift。[document-toolbar-state.contract.test.mjs](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs) 在既有真实 Ribbon/命令用例增加两条对应断言，未新增用例、未运行。入口 ID、布局、激活/禁用判断、patch、模型和命令均沿用既有实现。

[ShortcutPanel.vue](../../packages/office-studio/src/platform/ribbon/ShortcutPanel.vue) 从可见 Ribbon 项的 `shortcut` 生成面板列表，并在 Mac 显示时将 `Ctrl+` 替换为 `⌘+`；[ribbon-command-menu.ts](../../packages/office-studio/src/platform/ribbon/ribbon-command-menu.ts) 也读取该元数据绘制菜单项快捷键。后者是通用消费者，本批上下标仍为既有根按钮，不据此声明新增菜单入口。[RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 的 Tooltip 只读 `tooltip/title`，不会自动追加 `shortcut`；本次只补既有元数据及面板提示接线，不声称悬浮 Tooltip 已新增快捷键。共享 Ribbon 组件未修改；面板实际显示、Mac 文案与焦点行为仍未运行验收。

本提示续补只执行静态检查：Ribbon TypeScript AST、既有契约 `.mjs` 的 Node/AST 语法、27 项本地 import 与 9 项 SSR 源引用、DOC-02 的 86 项本地链接/18 项固定版源链接/2 项引用锚点，以及三个涉及文件的限定 `git diff --check` 均通过。新增断言未执行，浏览器计划用例仍为前述 14 例、运行结果为 0；未启动测试、类型检查、构建、浏览器或基准。

<a id="home-format-layout-shortcuts"></a>

## 11. Home 字符、对齐和列表快捷键续核

续核日期：2026-11-27（Asia/Shanghai）。固定版 [toolbar.shortcut.ts][u-toolbar-shortcuts] 的 START / FORMAT、START / LAYOUT 共 **12 条绑定**：B/I/U、删除线、上下标、四种对齐和两种列表。本地此前已有 5 条，本批补入剩余 **7 条**；这里只说明快捷键静态接线，**运行结果为 0，状态仍为未验收**。

| 既有入口与新增绑定 | 固定分组、classic 顺序和 grid 位置 | 上游命令 → 本项目命令 |
| --- | --- | --- |
| 删除线，Ctrl/Cmd+Shift+X | START / FORMAT，`order: 3`；`row: 2, column: 4` | `doc.command.set-inline-format-strikethrough` → `doc.command.format-text` 的 `strike` patch |
| 左 / 居中 / 右 / 两端对齐，Ctrl/Cmd+Shift+L/E/R/J | START / LAYOUT，对齐 selector `order: 2`；`row: 1, column: 1`；四个选项按左、居中、右、两端顺序 | 根 `doc.command.align-action` 与 `doc.command.align-left/center/right/justify` → `doc.command.format-paragraph` 的显式 `align` patch |
| 有序列表，Ctrl/Cmd+Shift+7 | START / LAYOUT，`order: 7`；`row: 2, column: 1` | `doc.command.order-list` → `doc.command.set-block-type`，`type: numbered` |
| 无序列表，Ctrl/Cmd+Shift+8 | START / LAYOUT，`order: 8`；`row: 1, column: 2` | `doc.command.bullet-list` → `doc.command.set-block-type`，`type: bullet` |

入口布局来自固定 [schema.ts][u-doc-schema]；[menu.ts][u-doc-menu] 的 `StrikeThroughMenuItemFactory`、`AlignMenuItemFactory`、`HORIZONTAL_ALIGN_OPTIONS`、`OrderListMenuItemFactory`、`BulletListMenuItemFactory` 确认按钮/selector、选项标签与顺序、真实字符/段落/列表当前值、非活动 Docs 显隐和 `disableMenuWhenNoDocRange`。本批对应 **4 个既有常驻根项、4 个对齐子选项**，7 条绑定不等于 7 个新根项；常驻、扩展、溢出和上下文根入口增量均为 **0**。classic/grid、More 和页眉页脚复用沿用既有入口清点，完整 Pro 范围继续按总清单验收。

| 固定行为来源 | 本批复用与差异 |
| --- | --- |
| [inline-format.command.ts][u-inline]，`SetInlineFormatStrikethroughCommand`、`SetInlineFormatCommand` | 上游翻转 `st.s`，折叠范围缓存输入格式，展开范围执行真实富文本 mutation。本地翻转 `currentMarks.strike` 并走既有字符 patch、权限/锁、输入偏好及历史。 |
| [paragraph-align.command.ts][u-paragraph-align]，四种对齐与 `AlignOperationCommand` | 上游按当前 segment 和真实段落范围执行 `RichTextEditingMutation`；所选段落已经为同一对齐时改为 `UNSPECIFIED`。本地保留既有显式赋值与幂等行为，重复相同对齐保持模型和内容历史不变；这是已记录的实现差异，未做独立 Office 行为核实。 |
| [list.command.ts][u-list-command]，`OrderListCommand`、`BulletListCommand`、`ListOperationCommand` | 默认列表按钮和快捷键进入 `BuildTextUtils.paragraph.bullet.switch` 与真实 mutation；指定 preset 另走 change-list-type。本地复用 `set-block-type` 已有选区全量 toggle-off、默认 `ORDER_LIST/BULLET_LIST`、相邻序列和编号来源保留逻辑，不新增列表算法。 |

[DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue) 的两条键盘路径按物理 `event.code` 识别绑定，要求 Ctrl 或 Meta、Shift、无 Alt、实际对应 TextArea 焦点并排除图片选区，沿用当前输入会话、composition 和 229 守卫。删除线沿用 `toggle-format`；六个段落/列表键共用本地命令映射，普通正文与 rich story 使用已有 `dispatchCommand` 携带块宿主。flat story 同步 UTF-16 选区并确认源模型/当前事件，然后交给 Workbench 的既有命令入口。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 对 `format-paragraph`、`set-block-type`、`change-list-type` 复用活动输入/组字守卫，Ribbon 的 `editingDisabled` 同步此状态。flat story 首次段落命令通过现有 `activateStoryBlockEditing` 和 [document-block-host.ts](../../packages/office-studio/src/features/document/services/document-block-host.ts) 将 U+FFFC 输入范围转换为实际块/字段缓存范围，转入 `{ kind: 'story', sectionId: 'first', story: 'header' }` 宿主；后续按同一正文桥执行。已有 `documentModelForBlockHost` / `applyDocumentBlockHostModel` 将编辑写回根模型、literal story、字段投影及锚点，继续使用同一权限、锁、内容历史和保存路径。字段投影会补入既有块来源元数据，草稿按字段类型、literal offset、缓存文本和字体字号核对保真。

本批计算域为已有 Docs UI 输入与命令编排；复用领域服务，未新增计算引擎。**既有内部 `toggle-format.field` 联合新增 `strike` 一个取值**，唯一 Workbench 调用方同步；事件名及参数形状、公共 API、命令 ID、领域模型、保存和 Worker/WASM 协议增量均为 **0**。生产增量相对上一批为 `+41/-5` 行，净增 36 行，没有新增控件或样式，继续使用 `antdv-next` TextArea 和统一 ConfigProvider。

[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 为删除线、两种列表主按钮和四个对齐子项登记同一绑定的 `shortcut`；Workbench 的既有快捷键面板另显式列出四种对齐。菜单快捷键消费沿用 [ribbon-command-menu.ts](../../packages/office-studio/src/platform/ribbon/ribbon-command-menu.ts)，面板沿用 [ShortcutPanel.vue](../../packages/office-studio/src/platform/ribbon/ShortcutPanel.vue)。共享 [RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 的 Tooltip 仍只读取 `tooltip/title`，本批没有补悬浮快捷键文案；面板、菜单和 Mac 显示仍未运行验收。

| 本批草稿 | 新用例数量与断言意图 | 运行结果 |
| --- | --- | --- |
| [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | 2 例：classic 明色/grid 暗色；两段正文的删除线开关、四种对齐、重复对齐幂等、有序/无序列表开关与共享序列、10 次内容历史及 undo/redo、无 Shift/Alt/失焦阻止、权限变化、内部 JSON 重开和字体正文保留 | 0 |
| [document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) | 4 例：flat/rich header × 两种布局主题；初始 flat/rich 输入守卫、Ctrl/Meta 操作、首次段落命令宿主转换、后续正文桥、单段列表/对齐、10 次历史及 undo/redo、字段与未选段落/正文/footer 保留、权限变化、公共内部 save/reopen | 0 |
| [document-toolbar-state.contract.test.mjs](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs) | 既有真实 Ribbon/菜单/命令用例内补 7 组快捷键及命令映射断言，新增用例 0 | 0 |

前述 **6 个 composition 草稿**追加全部 7 条 Ctrl/Meta+Shift 组合、对齐/列表根控件禁用及真实 Ribbon 组件的延迟段落、列表和 preset 回调阻止，不另增加用例数。本批新增 **6 个浏览器计划用例**；2026-11-27 续核累计 **20 个浏览器计划用例、0 个运行结果**，包括此前 2 个 rich story 导航、6 个组字和 6 个上下标用例。快照重开不等于 Office 文件验收。表格/代码块与其他 story、专用 drawing 排除、真实 Windows/Mac 键盘与系统 IME、主题焦点、WebGPU/Canvas2D、独立 Office 文件和性能仍未验收。

本批限定静态检查已通过：两个 Vue SFC 与六份源文件的 TypeScript AST 语法、三个 `.mjs` 的 Node 语法、两个内嵌 HTML module 语法、219 项本地 import 路径和 9 项 SSR 源引用；DOC-02 的 96 项本地链接/20 项固定版源链接/2 项引用锚点、DOC-03 的 42 项本地链接/20 项固定版源链接/2 项引用锚点，以及七个归属文件的 `git diff --check` 均通过，未发现冲突标记。检查仅解析文件和核对路径，没有执行测试、类型检查、构建、Rust/WASM、dev server、浏览器或基准，不提升验收状态。

<a id="paragraph-alignment-unspecified"></a>

## 12. 对齐重复执行的真实差异与待确认方案

核对日期：2026-11-27（Asia/Shanghai）。**本节仅记录修复方案，未实施模型、命令或读写变更，运行结果为 0。** 第 11 节的显式赋值、重复对齐幂等及 10 次历史草稿仍是当前代码；不能将下述方案写成已修复。

| 固定版真实来源 | 已核实行为 |
| --- | --- |
| [paragraph-align.command.ts][u-paragraph-align]，`AlignOperationCommand` | 从活动 segment 的实际范围取段落；只有全部段落的直接 `horizontalAlign` 等于请求值才切换。同值写入数值 `HorizontalAlign.UNSPECIFIED = 0`，保留属性，不删除它。 |
| [core/paragraph-style.ts][u-paragraph-resolution]，`_definedStyle`、`resolveDocumentParagraphStyle`、`_assignParagraphStyle` | 直接样式中的 0 不会被过滤，最后覆盖命名样式和文档默认值。省略属性才继承；0 与省略不是同一状态。 |
| [line-adjustment.ts][u-paragraph-line-adjustment]，`horizontalAlignHandler` | 普通 Docs 的 0 不执行居中、右对齐或两端对齐偏移，呈现左定位；这不代表模型直接赋成 `LEFT`。命令中的旧 `layout-ruler.ts` 注释路径已失效，以本表实际文件为准。 |
| [menu.ts][u-doc-menu]，`getParagraphStyleAtCursor`、`AlignMenuItemFactory` | 菜单读取直接段落样式。selector 的值为 `horizontalAlign ?? LEFT`，保留 0；0 使用左侧根图标回退，四个子选项均不激活。不能依据画布左定位把 0 标成已激活左对齐。 |

本地阻碍已逐文件确认：[DocumentParagraphStyle](../../packages/office-studio/src/domain/model.ts) 的 `align` 是必需的四值字段；[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) 拒绝 `align: undefined` 和第五种值；[document-body-style.ts](../../packages/office-studio/src/api/document-body-style.ts) 读入 0 时沿用默认对齐，写出时总输出 1–4；[documentNormalizer.ts](../../packages/office-studio/src/parser/documentNormalizer.ts) 与 [docx.rs](../../crates/office-writer/src/docx.rs) 将未知/缺失对齐归为 left。仅在 Ribbon 写 `left` 或删除 patch 字段都不能修复该契约。

**待确认的最小建议**：在现有 `align` 字段表达六种直接意图 `inherit | unspecified | left | center | right | justify`，不另建一套对齐 mutation。`inherit` 是本项目内部对省略的显式表示；公共 body 序列化时省略 `horizontalAlign`，`unspecified` 写 0，四种显式值写 1–4。解析 DTO、直接意图和布局结果继续分离，文档默认值使用不含 `inherit` 的已解析对齐类型。旧四值快照保留原有显式含义，不凭呈现效果反推它曾经省略。

| 意图 | 布局及菜单 | 保存/重开 |
| --- | --- | --- |
| `inherit`，公共属性省略 | 复用当前命名样式链、文档默认及已保留来源解析有效对齐；Ribbon 按上游直接值缺省提示，不用继承后的视觉值伪造直接设置 | 内部快照保留意图和既有样式来源；公共 body 继续省略属性 |
| `unspecified`，公共值 0 | 明确覆盖继承；复用现有行定位，无居中/右/两端偏移；根图标回退，四子项均不激活 | 内部快照保存 `unspecified`，公共 body 保存 0，不能写成 1 后声称保真 |
| 四种显式值 | 复用已有四种定位和菜单选项 | 保留现有四值快照及 body 1–4 映射 |

命令建议给既有 `doc.command.format-paragraph` 增加可选对齐执行方式 `set | toggle`，默认 `set`。Ribbon 和四条快捷键使用 `toggle`；[Facade](../../packages/office-studio/src/api/facade.ts) 的 `formatParagraph` / `FDocumentParagraph.setParagraphStyle` 沿用默认显式赋值，相同赋值继续不新增历史。`toggle` 在命令内从真实 `selectedTargets` 比较直接意图；全部同值时统一写 `unspecified`，混合或继承时统一写请求值。复用原先半开范围、caret、表格单元格、整段锁预检、批量提交和一次内容历史，不在 UI 用 toolbar 的混合值替代目标段落检查。

正文、rich story 和 flat story 共用该命令；flat 首次段落操作仍经 `activateStoryBlockEditing` 和现有块宿主转换。公共 API、菜单、快捷键的权限判断及 body/story 输入、composition、焦点恢复继续走当前链路。布局在现有段落排版入口解析有效对齐，再交给已有 `positionLine`，不增加计算引擎；工具栏保留直接意图与混合态，设置面板不把 `unspecified` 显示为显式 left。可替换控件仍使用 `antdv-next`，本方案不新增控件。

公共 body 的读入必须在既有命名样式解析前记录直接属性是省略、0 还是显式值；[document-body-paragraph.ts](../../packages/office-studio/src/api/document-body-paragraph.ts) 已保留 raw 段落和 inherited 来源，但当前分页 baseline 只覆盖四个分页字段，不能声称对齐来源已经完整。应扩展既有来源记录，保证修改直接意图后旧 raw 对齐不重新覆盖它；[document-body-source.ts](../../packages/office-studio/src/api/document-body-source.ts) 继续排除派生源，不保存第二棵权威块树。默认样式、命名样式应用和 clear 也需核对同一对齐意图更新，保留原调用方的其他格式行为。

**DOCX 必须一起处理**：当前 writer 总写 `<w:jc>`；省略继承应保留样式引用并省略直接 `<w:jc>`，显式四值按既有路径写出。`UNSPECIFIED(0)` 与原生 DOCX 的可用 `<w:jc>` 值没有已核实的一一对应，不能把写成 left 当成状态保真。建议在本项目转换链路添加版本化的段落对齐来源标记：原生部分表达左定位，附加来源保存显式 0，按实际 body/table/story 部件及段落标识绑定；重开仅在标识和原生对齐仍匹配时恢复该状态。parser/writer 复用一个共享来源 codec，外部编辑改变原生值后以原生数据为准。现有 `customXml/office-numbering.xml` 仅服务编号，并非完整模型 sidecar；本方案需新增对齐 payload 和读写接线，不能声称已有自动保留能力。标识选择、原生表达、Office 保留附加来源及外部同值编辑的处理仍需公开规格和文件样本核实，全部保留未验收。

下表是按已读调用链给出的拟修改范围，**尚未获得实施确认**。行数为待审阅的代码增量粗估，不包含本节文档，也不代表完成量；共享 Facade、模型、解析和来源协议需由主任务协调归属。

| 归属与实际文件 | 预计行数 |
| --- | --- |
| 模型：[domain/model.ts](../../packages/office-studio/src/domain/model.ts)、[document-model.ts](../../packages/office-studio/src/features/document/model/document-model.ts)、[document-default-style.ts](../../packages/office-studio/src/features/document/model/document-default-style.ts)；直接/有效对齐类型、默认及既有样式应用 | 40–70 |
| 命令：[document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts)、[document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts)；复用原事务，增加对齐方式和真实同值比较 | 35–60 |
| Docs UI：[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)、[DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue)、[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts)、[document-toolbar-state.ts](../../packages/office-studio/src/components/document/document-toolbar-state.ts)、[document-paragraph-settings-state.ts](../../packages/office-studio/src/components/document/document-paragraph-settings-state.ts)、[DocumentParagraphSettingsPanel.vue](../../packages/office-studio/src/components/document/DocumentParagraphSettingsPanel.vue)；切换入口与直接/混合态 | 45–75 |
| 公共 body：[document-body-style.ts](../../packages/office-studio/src/api/document-body-style.ts)、[document-body-paragraph.ts](../../packages/office-studio/src/api/document-body-paragraph.ts)、[document-body-source.ts](../../packages/office-studio/src/api/document-body-source.ts)；省略/0/显式值、来源和重开 | 45–75 |
| 布局：[document-paragraph-layout.ts](../../packages/office-studio/src/features/document/layout/document-paragraph-layout.ts)、[document-text-layout.ts](../../packages/office-studio/src/features/document/layout/document-text-layout.ts)；复用现有行定位 | 10–20 |
| Office：[documentNormalizer.ts](../../packages/office-studio/src/parser/documentNormalizer.ts)、[document_paragraph_source.rs](../../crates/office-parser/src/document_paragraph_source.rs)、[parser.rs](../../crates/office-parser/src/parser.rs)、[docx.rs](../../crates/office-writer/src/docx.rs)、[docx_source_story.rs](../../crates/office-writer/src/docx_source_story.rs)；保留来源意图、原生属性与附加对齐来源接线 | 110–180 |
| 来源协议：拟新增 `crates/office-io-common/src/document_alignment_source.rs`，在 [lib.rs](../../crates/office-io-common/src/lib.rs) 注册；共享版本化对齐 payload、段落绑定及预算 | 60–100 |
| 合计 | 23 个生产文件（22 个既有文件，包含 `lib.rs` 注册；1 个新来源 codec），约 345–580 行；新增模型取值、命令参数和 Office 来源协议，属于需确认的较大改动 |

确认后所需草稿至少覆盖：直接/继承/0 的公共 body 往返、命名样式和默认为 center/right 时显式 0 的覆盖、四种同值 toggle 与混合范围、半开/表格/锁原子拒绝、Facade 同值 set 幂等、直接菜单状态、布局定位、撤销重做及 body/flat/rich story 内部保存重开。复用已有 [设置契约](../../packages/office-studio/tests/document-settings.contract.test.mjs)、[段落设置状态契约](../../packages/office-studio/tests/document-paragraph-settings-state.contract.test.mjs)、[工具栏契约](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs)、[布局契约](../../packages/office-studio/tests/document-paragraph-layout.contract.test.mjs)、[Facade 契约](../../packages/office-studio/tests/document-facade.contract.test.mjs)、两份 Home/story 浏览器草稿及 [Rust 文件往返草稿](../../crates/office-parser/src/document_format_roundtrip_tests.rs) / [story 来源草稿](../../crates/office-writer/src/docx_source_story_tests.rs)，预计增改约 220–360 行，不新增空测试。现有重复 justify 幂等断言和 10 次历史预期必须同步修订；该次操作变成真实 mutation 后预计增加一次历史，最终状态和计数仍需实际验证。

本方案使用现有同步模型/命令与段落布局计算域，未涉及新增 Worker/WASM 热路径。实施后的类型、前端/Rust/WASM、浏览器、Office 样本、主题/焦点/IME 和大样本性能验证继续按项目标准执行；本阶段未运行，不据此声明功能覆盖或性能达标。

本节仅修改本证据文件。133 项本地链接、22 项固定提交源路径、24 项引用定义、2 项引用锚点及新增稳定锚点唯一性均通过静态核对，限定 `git diff --check` 通过；生产代码及测试草稿与本节前一致，浏览器计划仍为 20 例、运行结果仍为 0。

<a id="home-list-numbering-context"></a>

## 13. START 列表画廊与编号上下文续核

续核日期：2026-11-27（Asia/Shanghai）。本批修正两处既有编号动作的上下文/输入接线，**运行结果为 0，状态仍为未验收**。第 12 节对齐模型和 Office 来源方案仍待确认，本批未实施它。

| 固定来源 | 本批核对的入口与行为 |
| --- | --- |
| [schema.ts][u-doc-schema]、[menu.ts][u-doc-menu] 的 `OrderListMenuItemFactory` / `BulletListMenuItemFactory` | START / LAYOUT；classic `order: 7/8`；grid 有序为 `row: 2, column: 1`，无序为 `row: 1, column: 2`。两者是 `BUTTON_SELECTOR`，当前值按活动 range 的 `segmentId` 读取实际正文/header/footer 段落并解析列表，非活动 Docs 隐藏，禁用使用 `disableMenuWhenNoDocRange`。 |
| [list-type-picker/Picker.tsx][u-list-picker]、[index.ts][u-list-picker-index] | 各画廊固定六项：默认项及 `_1`–`_5`，不把 quick-typing 预设加入画廊；三列 72×80，四行预览依次使用 level 0/1/2/0，当前预设按真实值标记。 |
| [list.command.ts][u-list-command] 的 `ListOperationCommand` / `ChangeListTypeCommand` | 主按钮按真实段落范围走 bullet switch，预设按同一 segment 走 bullet set / mutation；正文和 header/footer 共用命令路径。当前两条 Ctrl/Cmd+Shift+7/8 绑定仍见第 11 节，本批没有新增绑定。 |

本地两根项、两套六卡画廊继续使用 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts)、[DocumentListPicker.vue](../../packages/office-studio/src/components/document/DocumentListPicker.vue) 和 [document-lists.ts](../../packages/office-studio/src/features/document/model/document-lists.ts)。重新编号、起始编号表单和继续前一编号属于本项目既有扩展，不是固定版两画廊的额外预设；本批没有新增根项、卡片或编号动作。常驻、扩展、溢出及上下文根入口增量均为 0，完整 Pro 范围继续按总清单验收，不以本批小修替代列表或 Docs 全量完成。

两处实际改动都位于 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)：编号 `canRestart/canContinue` 改为使用既有 `paragraphContext.model/selection`，不再从 flat story 读取休眠正文。该上下文已用于段落设置和缩进，正文/rich story 使用活动块宿主，flat story 使用当前 UTF-16 范围与已有块投影。普通页眉不能因为正文是列表而启用编号，页眉自身的可编号段落也不再因为正文普通而失去能力。

既有 `execute` 的活动输入/composition 守卫增加 `doc.command.restart-list-numbering` 和 `doc.command.continue-list-numbering`，在 flat 宿主转换之前阻止延迟 Ribbon 回调。正常执行仍使用 `activateStoryBlockEditing`、真实 [编号计划与能力判断](../../packages/office-studio/src/features/document/commands/document-commands.ts)、现有源编号/序列、权限和一次内容历史；没有另写模型 mutation。生产仅修改两处接线并换行，净增 1 行；内部事件、命令参数、公共 API、模型、保存及 Worker/WASM 协议增量均为 0。

本批未新增控件或样式，继续使用统一 ConfigProvider 和 `antdv-next` 的 Button、Form/Form.Item、InputNumber 及 Canvas TextArea。编号输入自身的 composition / 229 处理沿用 `DocumentListPicker`，此次补的是工作台活动文档输入的守卫，两者各保护真实输入边界。

| 既有草稿文件 | 本批新增/续补内容 | 运行结果 |
| --- | --- | --- |
| [document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) | 新增 4 例：flat/rich header × classic 明色/grid 暗色；正文列表/页眉普通及正文普通/页眉编号两种相反上下文、真实 InputNumber/六卡、重新编号、继续前一编号、一次历史及 undo/redo、宿主与正文/footer 保留、有效编号段落上的延迟 composition 回调、公共内部 save/reopen | 0 |
| 同一 story 草稿及 [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | 原有 6 个 body/flat/rich composition 用例追加两条真实 Ribbon 编号回调，验证没有内容、历史或命令事件变化；不另增加用例数 | 0 |

本批新增 4 个浏览器计划用例；2026-11-27 续核累计 **24 个浏览器计划用例、0 个运行结果**。新 fixture 通过已有 `projectDocumentStoryBlocks` 生成真实 flat/rich 内容及保留的编号来源，复用 OfficeStudio/FDocument，没有 Mock 编号实现。此处保存/重开是内部模型验收草稿，不是 DOCX 或独立 Office 文件验证。其他 story、表格/代码块、全部混合预设及导入列表菜单状态、失活回调、真实系统 IME、主题焦点、WebGPU/Canvas2D 和大样本性能继续保留未验收。

三份代码/草稿文件的 Vue SFC、5 段 TypeScript/JavaScript AST（包含两份 fixture module）、1 份模板、两份 `.mjs` Node 语法及 137 项本地 import 路径均通过静态检查，限定 diff 检查通过。未执行测试、类型检查、构建、Rust/WASM、dev server、浏览器或基准；语法和来源检查不提升功能验收状态。

本批 DOC-02 的 140 项本地链接、24 项固定提交源路径、26 项引用定义、2 项引用锚点和两个方案/续核稳定锚点唯一性均通过；DOC-03 的 42 项本地链接、20 项固定提交源路径和 2 项引用锚点也通过。本批四个修改文件的限定 diff 检查通过，代码及草稿保持稳定。

<a id="home-paragraph-indent-spacing"></a>

## 14. 缩进、行距与段距续核

续核日期：2026-11-27（Asia/Shanghai）。本批补齐缩进动作在活动输入/composition 期间的守卫，正常执行继续使用既有命令。**运行结果为 0，状态仍为未验收**；第 12 节对齐模型及完整 Picker 仍待确认，本批未实施。

| 固定来源 | 入口及可观察契约 |
| --- | --- |
| [docs-ui/menu/schema.ts][u-doc-schema]、[paragraph-menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/paragraph-menu.ts) | 固定 START / LAYOUT 没有直接缩进、行距或段距根。缩进位于段落浮动菜单的 indent 组，组 order 1；增加/减少 order 0/1，ID 为 `doc.menu.paragraph-t.indent.increase/decrease`。两个 factory 只定义按钮和非 Docs 隐藏，没有定义本地的边界 disabled 规则。 |
| [ParagraphMenu.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/ParagraphMenu.tsx) | 缩进读取实际 formatting range 首段的直接 indentFirstLine，按 ±20 修改，最低为 0，再执行段落设置命令；该分支未截断到 100。 |
| [doc-paragraph-setting.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/doc-paragraph-setting.command.ts)、[controller](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/doc-paragraph-setting.controller.ts) | `doc-paragraph-setting.command` 按 range 的 segmentId 读取正文/header/footer，更新选中段落，进入一次 RichTextEditingMutation / FormatParagraph 历史；面板宽 300。 |
| [line-spacing.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/paragraph-setting/line-spacing.ts)、[hook/utils.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/paragraph-setting/hook/utils.ts) | auto 范围 1–5、step 0.1；atLeast/exact 范围 1–100。模式转换使用真实 glyph/line metrics；fixed 显示保留 2 位、存储保留 4 位。hook 在 16ms 内合并 mode/value 后调用同一段落命令。 |

完整标签、grid 行列、classic 顺序、六个行距选项、段前/段后 14 个叶及面板字段沿用 [DOC-02 段落缩进与间距清点](DOC-02-paragraph-spacing-indent.md)。本批涉及 **0 个固定常驻顶部根、4 个既有本地扩展顶部根**：减少缩进、增加缩进、行距、段落间距，仍位于 Home / `doc-layout-extensions`；本批新增根、上下文叶及独立溢出功能均为 0。宽度不足时共享 Ribbon 呈现同一入口，不增计功能。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 `execute` 原有活动输入/composition 守卫缺少 `doc.command.change-paragraph-indent`；本批将它加入已有列表，在 flat story 宿主转换和实际 mutation 之前拒绝延迟回调。生产只改这一行，没有修改模型、命令参数、Facade、历史或 Office/Worker/WASM 协议。

缩进 capability 已使用 `paragraphContext.model/selection`，本批没有改写它。正文/rich story 使用活动块宿主，flat story 使用真实 UTF-16 范围及既有投影；正常执行仍走 `activateStoryBlockEditing`、[prepareDocumentParagraphIndent](../../packages/office-studio/src/features/document/commands/document-commands.ts)、权限及整段锁预检，然后统一提交。普通段仍从各自首行缩进加减 20px、范围 0–100；列表仍修改真实 level 并保留相对编号定义的缩进差值。

行距和段距使用同一 `doc.command.format-paragraph`，已受原守卫保护；[段落面板](../../packages/office-studio/src/components/document/DocumentParagraphSettingsPanel.vue) 也沿用该命令及自身 composition 暂存。规则切换读取 [已有段落状态与 metrics](../../packages/office-studio/src/components/document/document-paragraph-settings-state.ts)，历史与内部保存仍由 OfficeStudio / FDocument 管理，没有并行状态或第二套历史。

差异保持明确：固定缩进分支取首段直接值并统一 patch，本地每段相对增减且设 100 上限；fixed 精度、本地 px 与公开 scalar 的换算、无 metrics 的字号回退、默认段后 8/12、完整继承及 signed 导入缩进仍按前述证据保留差异或未核实。本批不改这些模型或 Office 单位契约，也不以新增守卫宣称完整功能对齐。

| 既有草稿 | 本批内容 | 运行结果 |
| --- | --- | --- |
| [document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) | 新增 4 例：flat/rich header × classic 明色/grid 暗色；正文已到缩进上限、页眉仍可增减的相反上下文，真实缩进按钮、fixed 24px 状态转 auto 1.5、段前/后真实菜单、每次内容历史、undo/redo、活动 host、正文/footer/其他段保留、有效目标上的延迟 composition 回调、内部 save/reopen | 0 |
| 同一 story 草稿及 [document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) | 原有 6 个 body/flat/rich composition 用例追加缩进 ±1、行距/段距延迟回调及四根控件禁用断言；不另增用例数。Home 草稿的既有入口查找补齐 grid 与共享溢出选择器 | 0 |

本批新增 4 个浏览器计划用例；2026-11-27 续核累计 **28 个浏览器计划用例、0 个运行结果**。草稿直接复用真实 Workbench、`antdv-next` 的 Button/Select/Menu/TextArea、OfficeStudio/FDocument 和 story 投影，没有 Mock 格式实现。生产未新增可替换控件、原生输入桥或样式；统一主题路径保留。正文/页脚及其他 story 的完整交互、全部锁/权限与范围组合、真实系统 IME、主题与焦点、独立 Office 文件、WebGPU/Canvas2D 及性能/内存继续未验收。

静态检查已通过：5 段脚本（含两份 fixture module）、1 份模板、137 项本地 import、两份 `.mjs` Node 语法；本证据的 147 项本地链接、30 项固定提交源路径、26 项引用定义、2 项引用锚点及 3 个稳定锚点唯一性均通过。固定仓库 HEAD 与引用提交一致，四个修改文件的限定 diff 检查通过。未运行测试、类型检查、构建、Rust/WASM、dev server、浏览器或基准；静态通过不提升功能验收状态。

<a id="home-character-input-session"></a>

## 15. Home 字符入口的工作台活动会话

续核日期：2026-11-27（Asia/Shanghai）。本批只补工作台失活时的字符入口状态，**运行结果为 0，DOC-02 仍未验收**。第 14 节的 28 个浏览器计划用例是此前 2026-11-27 续核累计快照，并非单个文件总数；本批另增 2 个计划用例，续核累计 30 个，不能用这些数量证明全量覆盖。

固定依据继续使用 Univer v1.6.0、commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`：[Docs schema 的 FORMAT 分组][u-doc-schema] 中，本次核对字体、字号、B/I/U/删除线、上下标及两个颜色常驻根，共 10 项；段落样式为该组另一个根，本批新增入口为 0。grid 第一行按段落样式、字体、字号排列；第二行依次为 B/I/U/删除线、上标、下标、高亮、字色。classic 中上下标按 order 4/5、字号/字体按 order 6/7 排列，两种布局保持原声明。

[menu.ts 的 FontFamilySelectorMenuItemFactory、FontSizeSelectorMenuItemFactory 及字符按钮 factory][u-doc-menu] 共用 disableMenuWhenNoDocRange，合并实际文本/矩形范围、相交代码块与编辑权限；状态读取 getFontStyleAtCursor，并经 shouldSuppressDocMenuStateRefresh 避免宿主交互抑制期间刷新。[SetInlineFormatCommand][u-inline] 拒绝缺少活动范围/模型/body，折叠范围只缓存下一次输入样式，展开范围提交同一 RichTextEditingMutation。固定上游未定义本项目的 Vue KeepAlive 生命周期，本地通过已有输入会话实现对应活动边界，不复制其内部服务。

[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的 textFormattingDisabled 使用字符专用值优先于 editingDisabled，使 story 格式能力可以独立判断。此前 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 只传 !textFormattingState.available，导致 KeepAlive 失活时字符控件仍显示可用，而真实 execute 已返回 false。本批在专用值中合并 !inputStateActive；ribbonTabs 显式读取现有 storyActionVersion，由 onActivated/onDeactivated 的原递增触发控件重算，恢复不依赖额外 focus 或 selection。相邻段落入口、只读可用的 UI operation 及每条命令的独立权限/锁条件保持原行为。

生产修改两处表达式，净增 1 行；继续使用原 FDocument/Studio 输入会话、权限、内容命令、一次历史和保存路径。API、命令参数、DTO、模型、Office 写出及 Worker/WASM 协议变更均为 0；没有新增可替换控件或原生输入桥，统一 ConfigProvider 与 antdv-next 的 Button、Select/Input、Canvas TextArea 沿用既有主题路径。

[document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) 的既有真实 Workbench/FDocument 夹具增加 KeepAlive 隐藏/恢复；一个 test.each 声明提供 light/classic 和 dark/grid 两例。草稿核对 10 个真实 Ribbon 状态和 12 个 Ant 输入/按钮，失活时检查全部禁用并发送 10 种合法字符回调，要求模型、历史与命令事件不变；恢复后在重新 focus/selection 前检查全部启用，再通过真实字号输入提交一次内容历史，执行 undo/redo 与内部 JSON 模型重开。该重开沿用原夹具 replaceFixtureModel，不将领域快照宣称为公开 FDocument.save 数据或 Office 文件验证。

该 Home 浏览器文件当前有 29 处 test 声明、37 个计划案例：21 个普通案例、4 处 test.each 展开 8 例、4 处数据 for-of 展开 8 例。本批新增 1 处声明、2 例，文件总数与前述跨 Home/story 的续核累计分别记录，运行结果均为 0。限定静态检查通过：三个脚本（Vue、浏览器草稿及解码后的内嵌夹具）、一个模板及三文件 diff；没有执行产品模块或启动 fixture。

完整 flat/rich 及其他 story、全部范围/锁/权限组合、真实 Windows IME、弹层主题与焦点、独立 Office 文件、WebGPU/Canvas2D、性能与内存继续未验收；前述完整对齐模型及字符 Picker 方案仍待确认，本批未实施。

<a id="home-list-preset-mixed-state"></a>

## 16. Home 混合列表预设的激活状态

续核日期：2026-11-27（Asia/Shanghai）。本批按既有 Home 清点继续核对 LAYOUT 中的有序、无序与清单三个常驻根，修正前两根的混合预设激活状态；新增入口为 0，清单命令和状态保持。**运行结果为 0，DOC-02 / DOC-05 仍未验收**。

固定参考仍为 Univer v1.6.0、commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`：[schema.ts][u-doc-schema] 定义三根 order 7/8/9，grid 为有序 r2c1、无序 r1c2、清单 r2c2。两种列表各沿用六卡弹层，清单是单按钮，扩展/溢出/上下文入口没有增量。[menu.ts][u-doc-menu] 的 listValueFactory$ 读取实际 segment 中所选段落，经 resolveDocumentParagraphBullet 解析直接或样式列表；只有全部段落的 listType 相同才给出值，三个 factory 再按 ORDER_LIST、BULLET_LIST、CHECK_LIST 前缀计算 activated。不同 preset 即使同属有序或无序，根状态也不激活。

[ListOperationCommand / ChangeListTypeCommand][u-list-command] 仍使用真实段落范围、segment 和同一 RichTextEditingMutation。固定 [switchParagraphBullet](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/docs/data-model/text-x/build-utils/paragraph.ts#L44) 的同类前缀判断与激活状态的精确预设比较不同：混合有序/无序预设点主按钮仍可整体移除列表；选择一张卡则统一应用该预设。本批保留本地 [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 已有类别 toggle、显式 preset、编号来源/层级、标题身份、输入会话、权限/锁、一次内容历史和保存行为，没有把激活判断写入 mutation。

[document-toolbar-state.ts](../../packages/office-studio/src/components/document/document-toolbar-state.ts) 此前用 undefined 同时表示不同预设和缺少已识别预设；[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 随后回退 blockStyle，导致普通编号/无序段的混合预设被误亮。本批新增可选的 UI 状态字段 listPresetMixed，通过已有 blockListPreset 的实际结果计算，并复用于 listPreset 汇总；仅在明确混合时阻止这两个根的 activated 回退。无列表状态、旧编号默认预设、调用方未传该字段、清单及所有 disabled、弹层 capability 保持原行为。生产仅改这两个文件；公开 API、领域模型、命令参数、保存、Office/Worker/WASM 协议增量均为 0，内部工具栏状态新增一个可选字段。

[document-toolbar-state.contract.test.mjs](../../packages/office-studio/tests/document-toolbar-state.contract.test.mjs) 增加三个参数案例（有序、无序、编号标题）和一个旧编号案例。草稿通过真实命令建立不同预设，检查反向选区、结束于下一段 offset 0 的排除、三根状态与无选中卡，并在真实模型副本上确认混合主按钮仍整体移除列表；通过真实 Ribbon adapter / command 统一预设、恢复激活再移除列表，保持标题身份与正文。旧编号无直接 list 记录时仍使用默认 ORDER_LIST。

[document-home-colors.browser.test.mjs](../../packages/office-studio/tests/document-home-colors.browser.test.mjs) 增加 light/classic、dark/grid × 有序/无序四例，装配现有 Workbench、OfficeStudio/FDocument 和 antdv-next。草稿检查主按钮 aria-pressed 与六卡状态、菜单打开不修改模型/历史/事件；真实选卡统一预设、一次历史、undo/redo、内部 JSON 模型重开及主按钮移除列表，保留段落/run ID 和文字。没有新建夹具文件、Mock、原生控件或样式，继续使用统一 ConfigProvider；可替换输入仍沿用 Ant Button、Form、InputNumber 和 Canvas TextArea。

第 15 节的 Home 文件 29 声明/37 计划案例及续核累计 30 属上一批快照；本批另增四个浏览器计划案例、四个契约计划案例。Home 文件当前 30 处声明、41 个计划案例（21 普通、5 处 test.each 展开 12 例、4 处 for-of 展开 8 例）；toolbar-state 契约文件当前 19 处声明、21 个计划案例（18 普通、1 处 test.each 展开 3 例）。2026-11-27 浏览器续核累计为 34 例，所有运行结果为 0。静态语法/引用/diff 由主线统一收尾。本次没有补足全部自定义 Office 编号定义之间的状态区分、完整 story/表格、全部锁/权限/范围、原生 IME、主题/焦点、浏览器/Office 文件或性能验收；完整字符 Picker 与段落继承方案仍待确认，本批未实施。

[u-list-picker]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/list-type-picker/Picker.tsx
[u-list-picker-index]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/list-type-picker/index.ts
[u-paragraph-resolution]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/docs/data-model/paragraph-style.ts
[u-paragraph-line-adjustment]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/engine-render/src/components/docs/layout/block/paragraph/line-adjustment.ts
[u-paragraph-align]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/paragraph-align.command.ts
[u-list-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/list.command.ts
[u-toolbar-shortcuts]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/shortcuts/toolbar.shortcut.ts
[u-doc-ui-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/ui.controller.ts
[u-shortcut-utils]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/shortcuts/utils.ts
[u-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/schema.ts
[u-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/menu/format-painter.menu.ts
[u-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/controllers/format-painter.controller.ts
[u-session]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/ui/src/services/format-painter/format-painter-session.service.ts
[u-doc-controller]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/format-painter.controller.ts
[u-defaults]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/doc-menu-style.service.ts
[u-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/format-painter.command.ts
[u-header-footer]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/header-footer-ribbon.ts
[u-doc-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts
[u-doc-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts
[u-inline]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/inline-format.command.ts
[u-selection]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/services/doc-selection-manager.service.ts
[u-note-model]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/types/interfaces/i-document-data.ts
[u-comment-model]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/thread-comment/src/types/interfaces/i-thread-comment.ts
[u-comment-editor]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/thread-comment-ui/src/views/ThreadCommentEditor.tsx
[doc-code]: https://docs.univer.ai/zh-CN/guides/docs/features/code
[doc-io]: https://docs.univer.ai/zh-CN/guides/docs/features/import-export
