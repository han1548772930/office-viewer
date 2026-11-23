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
