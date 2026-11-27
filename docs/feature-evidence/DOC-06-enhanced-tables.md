# DOC-06 表格与增强表格

状态：**未验收**。本批按固定上游完成静态入口与命令追踪，再推进本项目独立实现。未运行测试、类型检查、构建、浏览器、Rust/WASM 或基准；测试草稿和静态读取不代表功能、像素或 Office 往返已通过。

更新日期：2026-11-27。当前是候选实现与待统一验证的草稿记录；本次 DOC-06 切片收束不代表全部 Docs / Sheets / Pro 目标结束。

## 固定来源与证据边界

- OSS：`E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。
- Pro：本项目 `.research.local/univer-1.6.0/packages/docs-table`、`docs-table-ui` 的 v1.6.0 公开发布类型与 `lib/es/index.js`。仅静态读取，未导入、执行、安装商业 SDK；单行发布代码以字符串位置引用，不能视为可维护源码或运行证据。
- HTML 转换：固定 OSS `packages/docs-ui/src/services/clipboard/html-to-udm/converter.ts` 的 `buildHtmlTableGrid()`（726）、`collectHtmlTableColumnWidths()`（785）、`createTableCellFromHtml()`（821）；Pro `docs-table/lib/types/common/clipboard.d.ts` 声明 detached selection、TSV、HTML parse/serialize/paste 契约。发布 ES 静态可定位 plain-text builder `Na`、parse `Pa/Fa`、crop `ja`；`buildPasteHtmlTableActions` / `serializeTableSelectionAsHtml` 同名实现未在该 ES export 目录中核实，不能把类型声明当作已执行的 Pro HTML 行为。
- 公开规格：[Docs 表格](https://docs.univer.ai/zh-CN/guides/docs/features/table)。Microsoft Office 可观察行为为合并、拆分、固定尺寸、重复表头、嵌套表格与分页的对照目标，复杂样本仍待浏览器与文件复查。
- 入口清点遵循 [上游引用规则](../UNIVER_REFERENCE_RULES.md)、[执行标准](../../PROJECT_EXECUTION_STANDARD.md)、[渲染标准](../RENDERING_AND_COMPUTE_STANDARD.md)。所有 Pro 能力仍在目标范围内。

## 入口清点

分类计数是静态注册目录的计数，不是完成证明。替换后的 Insert 根不与 OSS 原根重复相加；selector/submenu 的根与子项分别列出；移动和尺寸 Canvas handle、移动端复用入口、浮动工具栏不混进常驻 Ribbon 根数。

| 分类 | 可复查计数 | 来源与说明 |
| --- | ---: | --- |
| OSS 常驻表格根 | 1 | `packages/docs-ui/src/menu/schema.ts:345`：Insert / MEDIA，order 2，grid r1 c2，`doc.menu.table` |
| Pro 扩展替换常驻根 | 1 | `docs-table-ui/lib/es/index.js` 的 `Ya` 替换同一个 Insert 根；不是额外常驻根 |
| Pro 上下文 Ribbon 根 | 17 | `Ja` 声明 `docs-table-ui.ribbon.table`、order 100、contextual；`Za` 为 Style 5、Layout 6、Data 2、Options 4 |
| 独立 overflow 根 | 0 | 该表格 schema 未注册独立 overflow；容器自适应溢出不是新增功能 |
| OSS 右键结构子项 | 7 | `schema.ts:457,477`：insert 4、delete 3，根各 1 |
| Pro border 叶项 | 10 | `common/border.d.ts`；同一目录在 Ribbon、右键和浮动区复用 |
| Pro insert / delete / sort / column-type 叶项 | 4 / 3 / 2 / 5 | `views/menu.d.ts`、`common/type.d.ts`、`Qa/Za` 静态注册 |
| Canvas trigger 类型 | 7 | `views/canvas/docs-table-trigger-model.d.ts`：table-select、row-header、column-header、row-insert-dot、column-insert-dot、resize-row、resize-column |
| Pro 内容命令 ID | 29 | `docs-table/lib/types/commands/*.d.ts` 与发布代码；其中 clipboard 命令不保证改变模型 |

### 上下文 Ribbon（Pro）

共同显隐：当前为 Docs、存在可解析表格选区（factory `U`）；共同按钮/参数由 `H` 构造。禁用还必须结合当前命令参数、表格拓扑、模型权限和内容控件锁，不能仅由“光标在表格内”判断。

| 分组 / order | 标签 / 根 ID | order、grid | 命令或叶目录 | 状态与上游来源 |
| --- | --- | --- | --- | --- |
| Style / 0 | 背景 `set-table-background` | 0，r1 c1 | `docs-table.command.set-table-background` | `Za`、background factory |
| Style / 0 | 边框 `docs-table-ui.menu.border` | 1，r2 c3 | all、inner、outer、top、bottom、left、right、none、innerHorizontal、innerVertical | `Za/Qa`；`common/border.d.ts` |
| Style / 0 | 边框颜色 | 2，r1 c2 | `set-table-border-color` | `Za`；允许清除显式颜色 |
| Style / 0 | 边框宽度 | 3，r2 c2 | `set-table-border-width` | UI 单位 pt，候选 0、.5、.75、1、1.5、2.25、3、4.5、6；本项目持久化需转换 CSS px |
| Style / 0 | 边框样式 | 4，r1 c3 | `set-table-border-style`：solid / dot / dash | `Za` 与 style factory |
| Layout / 1 | 插入 `docs-table-ui.ribbon.insert` | 0，r1 c1 | 上方行 / 下方行 / 左侧列 / 右侧列 | `Ui` 复用 4 个 insert command |
| Layout / 1 | 整表对齐 `docs-table-ui.ribbon.align` | 1，r2 c1 | `set-table-align`：START / CENTER / END | `Vi`、`Wa`；上游 modern 禁用；命令也验证 documentFlavor===TRADITIONAL |
| Layout / 1 | 合并 `docs-table-ui.menu.merge-cells` | 2，r1 c2 | `merge-cells` | `Ga` 调用 canMergeCells，覆盖合并范围须为矩形闭包 |
| Layout / 1 | 拆分 `docs-table-ui.menu.unmerge-cells` | 3，r2 c2 | `unmerge-cells` | `Ka` 调用 canUnmergeCells |
| Layout / 1 | 平均分布行 | 4，r1 c3 | `distribute-rows` | `Za`、`commands/distribute-rows.command.d.ts` |
| Layout / 1 | 平均分布列 | 5，r2 c3 | `distribute-columns` | `Za`、`commands/distribute-columns.command.d.ts` |
| Data / 2 | 排序 `docs-table-ui.menu.sort` | 0，r1 c1，showLabel | ascending / descending → `sort-table` | `Ki`；按选择列、列类型和 header 元数据处理 |
| Data / 2 | 列类型 `docs-table-ui.menu.column-type` | 1，r2 c1，showLabel | text / number / date / checkbox / dropdown → `set-column-type` | `common/type.d.ts`，dropdown 带 id/label/color 选项 |
| Options / 3 | 标题行 `docs-table-ui.menu.insert-title-row` | 0，r1 c1，showLabel | `insert-title-row` | `Ua`：传统模式显示 |
| Options / 3 | 固定表头 `docs-table-ui.menu.pin-header` | 1，r2 c1，showLabel | `set-header-row-count` | `Ua`：传统模式显示；不是仅冻结视觉覆盖层 |
| Options / 3 | 删除 `docs-table-ui.ribbon.delete` | 2，r1 c2，showLabel | 行 / 列 / 表格 | `Gi` selector；本项不是 Options panel |
| Options / 3 | 浮动工具栏开关 | 3，r2 c2，showLabel | `docs-table-ui.operation.toggle-floating-toolbar` | 默认 false，操作状态不进入内容历史 |

### 其他入口、复用关系和未核实项

- OSS Insert 根的 grid 为 r1 c2；子 `DocCreateTableOperation` order 0。`getTableDisabledObservable` 要求单个 collapsed 文本选区、有效 offset/model/skeleton、非 custom range；固定 OSS 对路径含 cells 的选区禁止插入嵌套表格。本项目仍实现嵌套目标，不把 OSS 临时限制改为范围删除。
- OSS 右键 insert 根 order 1，叶项上/下/左/右 order 1–4；delete 根 order 2，行/列/表 order 1–3。Pro `Qa` 在 paragraph、empty、insert-content、top-bottom 位置复用 Insert；cell-style 组放背景、颜色、宽度、样式；structure 组 order 1.5 复用 insert/delete；Others 放 title/header。隐藏旧根与移动端组是兼容注册，不重复累计功能。
- Pro 插表 picker 的公开尺寸目录为 14 列 × 10 行，手工尺寸 1–20；本项目可以使用 `antdv-next` 数字表单输入，保留可观察插入契约，不能把自绘方格 UI 当作 Canvas 内核。
- 浮动区公开 `DocsTableFloatingToolbarControls.d.ts` 包含粗体、斜体、删除线、下划线、文本样式、文字水平对齐、单元格垂直对齐、前景/背景色、边框预设/颜色/宽度/样式、reset；这些复用正文格式或表格命令。触发包含选区、悬停、删除目标与 anchorRect，默认关闭；具体 anchor/occlusion/触屏交互未运行核实。
- Pro 移动行/列、移动整表与尺寸拖拽通过 Canvas trigger 与 `docs-table-move-drag.d.ts`、`docs-table-resize-drag.d.ts`、guides 进入命令，不是 17 个 Ribbon 根中的遗漏按钮。需要分别验收命中、预览、提交、取消、跨页与覆盖层焦点。
- 发布代码字符串位置 71356 的 `DocsTableOptionsPanel` 只见带 min=0 的 header-count 输入，没有绑定 handler；其后的 `DocsTableToolbar` 只见 3 个按钮示意。此静态片段不能证明功能闭环，本项目不以空操作复刻。header 数量通过公开命令/Facade、真实布局与持久化实现。
- mobile/context 同一操作应接本项目统一命令；移动端像素、触控尺寸、浮动滚动条、跨页粘附表头均未核实。

## 命令、模型、历史和保存链路

| 行为目录 | 固定上游契约 | 本项目归属与差异 | 验收状态 |
| --- | --- | --- | --- |
| 插表、插删行列/整表 | OSS `doc-table-insert.command.ts`、`doc-table-delete.command.ts`：TextX 正文与 JSONX tableSource 同事务，删光行列转删表，插列保持既有总宽 | DocumentTable / compact anchors → detached planner → `DocumentCommand` → OfficeStudio 文档快照历史；逻辑列与存储索引分离 | 未验收 |
| 合并/拆分 | Pro merge range closure、canMerge/unmerge；`merge-cells` 发布位置 81338 只在 mutation 成功后发布结构选区 | 真实跨度和被合并段落归属；保留段落/字段/控件/批注 ID，拆分创建实际单元格 | 未验收 |
| 移动行、列、整表 | from/to/count/position；整表 targetOffset 或 deltaX/deltaY；`move-rows` 发布位置 82515 | 在本项目网格重排 anchors、列宽/列类型及行样式；不可把跨跨度网格切碎；全表正文位置或缩进偏移同历史 | 未验收 |
| 行高/列宽、均分 | resize-row height/hRule，resize-column width，distribute 含选中范围 | CSS px；固定宽度应进入 Canvas 几何，exact 行高应有可见裁剪与可编辑命中，不能等比放满页后声称尺寸有效 | 未验收 |
| 背景/边框/垂直对齐/边距 | range params；border 10 presets，width/style/color patch；垂直对齐独立于整表对齐 | typed style + effective table/cell defaults → Canvas 四边绘制；左右文字对齐复用 paragraph，整表对齐单独改变表位置 | 未验收 |
| 标题、表头、列类型、排序 | metadata titleRow/headerRowCount/columnTypes；sort column+direction；title/header 传统模式入口 | 标题行、重复表头、列类型和真实排序进入模型/命令；本地 checkbox / dropdown 覆盖层编辑使用 antdv-next，并以 option ID 区分同名选项。上游列类型的主要已核实用途是排序提示，不能冒称覆盖层为其公开内部实现 | 未验收 |
| 剪贴板 | copy-selection、paste-without-formatting；公开 HTML clipboard 类型 | 本地结构 copy 提供 TSV 与 HTML，crop compact anchors 并材料化有效格式；单一富表按源矩形写目标 footprint，普通 blocks / 1×1 源表替换首格正文；source / destination 模式复用现有 paste-blocks 命令，越界或切开已有 merge 原子拒绝 | 未验收 |
| 内容 mutation | `apply-table-actions` 发布位置 36756，校验 disabled/body，并通过 RichTextEditingMutation；不把 menu operation 加到内容历史 | 本项目 `commitDocument` 在 detached 全文草稿后校验 changedBlocks ACL，再按 unit ID 提交一个内容历史项 | 未验收 |
| 保存/重开 | tableSource 与 Pro metadata resource 参与保存；Facade mutations 返回同步 boolean | JSON 快照保存表格、续格正文和稳定 ID；原生 table/row/cell 属性、嵌套、vMerge、legacy hMerge、条件样式来源和 writer 候选见 [DOC-22 原生表格证据](DOC-22-native-tables.md)，其 Rust / DTO 草稿也未执行 | 未验收 |

29 个内容命令目录：`apply-table-actions`、`delete-columns`、`delete-rows`、`delete-table`、`distribute-columns`、`distribute-rows`、`insert-columns`、`insert-rows`、`insert-table`、`insert-title-row`、`merge-cells`、`move-columns`、`move-rows`、`move-table`、`resize-column`、`resize-row`、`set-column-type`、`set-header-row-count`、`set-table-align`、`set-table-background`、`set-table-border-color`、`set-table-border-style`、`set-table-border-width`、`set-table-border`、`set-table-vertical-align`、`sort-table`、`unmerge-cells`、`copy-selection`、`paste-without-formatting`。前缀均为 `docs-table.command.`。

## 公共 Facade 目录

来源：`docs-table/lib/types/facade/f-document.d.ts`、`f-document-table.d.ts`、`f-document-table-cell.d.ts`、`types.d.ts`。

- 文档：getTables、getTable、getTableAt、getTableAtSelection、findTableByText、findTables、insertTable、insertTableFromData。
- 表格读/选区：info/source/description、row/column count、text/content ranges、margin/override、metadata；selectTable/Rows/Columns/Cell/Range。
- 表格写：insert/delete/move table/rows/columns，merge/unmerge、cell text/background/margin/border、table background/border/style、row height/column width/distribute、header/pin/title、column type/sort。
- 单元格：getTable、row/column、text/range/insertOffset、effective margin/override；setText/background/border/margin、clearMarginOverride、mergeTo/unmerge。
- 本项目独立 `FDocumentTable`、`FDocumentTableCell` 持有 unit+table ID 并每次读取现模型；模型历史替换后不能继续编辑旧 object。Facade、Ribbon、右键、键盘和 Canvas 均需走同一 `DocumentCommand`，坐标参数采用逻辑网格；内部 TextTarget.col 仍为 anchor cell 的存储索引。

## 本批协议与后续证据

`DocumentCell.columnIndex` 明确逻辑起始列；compact rows 仅存 anchor，跨行覆盖行仍保留空 cells。`rowSpan/colSpan` 正整数。`DocumentTableRow.height/heightRule/allowSplit` 和 table/cell margin、borders、width、align、indent、offsetY 均持久化；CSS px 为领域单位，Office 边界转换 twips，UI pt 边框宽度转换 px。保留未知 metadata，非法或重叠网格原子拒绝，不能丢弃后纠正为“已支持”。

## 已落盘的候选实现与模块归属

| 模块 | 文件 | 当前实际职责与待验证边界 |
| --- | --- | --- |
| 网格 / 范围 / 导航 | `features/document/model/document-table-{grid,selection,navigation,text}.ts` | compact anchor 的逻辑覆盖、合法 covered rows、合并闭包、对象格结构选择、真实正文 targets、按 cell ID 重基、Shift 扩展/收缩和 TSV。读/选区不制造缺失段落 |
| 续格容器 | `model/document-table-continuations.ts`、`document-model.ts`、`api/document-indexed-input.ts` | `nativeContinuationContent[offset]` 保留独立可编辑正文数组；查找、输入、字段、批注、复制、merge/unmerge、插删移动都按真实容器与稳定 paragraph/row/cell ID 处理；TextTarget 指向 owner 的存储位置与实际 `cellBlockId` |
| 候选 mutation | `commands/table-{commands,mutations,command-types,insertion}.ts` | 插删移动、尺寸、跨度、边框/边距/背景/对齐、标题/表头/列类型/排序、范围内容和富文本 paste 在 detached 候选上校验，再发布；锁定控件、空锁字段、跨段字段与共享边邻格 ACL 纳入候选检查 |
| 属性 patch | `model/document-table-style.ts`、`DocumentTablePropertiesDialog.vue` | 行高/规则/拆行、四边 margin 和 border 支持 mixed 值与局部 patch；省略字段或显式 undefined 保留各项旧值；auto 布局改变为 fixed 时使用当前已测列宽。整表对齐和单元格/文字对齐各自有语义 |
| 剪贴板 | `services/document-block-paste.ts`、`document-clipboard.ts`、`document-model.ts` | 借用的 normalized blocks 不变；复制 remap block/run/row/cell/checklist/field ID、field endBlockId 和 continuation 身份；HTML 保留阅读顺序、多段、nested、跨度、尺寸/样式与独立续格 marker；copy 同时输出 HTML/plain |
| Facade / 输入会话 | `api/document-tables.ts`、`stores/document-input-store.ts` | FDocumentTable / Cell 每次回查当前快照；模型替换时存活 paragraph ID 回查、结构范围重基并清派生 pageAffinity；旧 typing marks 保持 invalid，只有 checked completion 可恢复。Root 已在模型发布入口调用 revalidateModel |
| 排版 / 场景 | `layout/document-table-{measurement,intrinsic-widths,layout,interaction}.ts`、`rendering/document-table-scene.ts` | 相同测量数据驱动 fixed/auto 列宽、rowSpan 高度约束、exact 裁剪、nested、整行或整行文字跨页碎片、重复表头与命中；resize 使用分页 gap 前 `contentHeight`，不是隐藏溢出正文的完整 intrinsic height |
| 工作台 / 覆盖层 | `DocumentWorkbench.vue`、`document-ribbon.ts`、`document-table-ribbon.ts`、`DocumentTable{InsertDialog,PropertiesDialog,Overlay,ColorControl}.vue`、`DocumentCanvasViewport.vue` | 17 个 Pro 上下文根、右键、属性、handle / insert dots、移动/resize 预览/取消、typed controls、floating toolbar、copy/paste 共用命令。所有可替换控件使用 antdv-next；主题、焦点、真实 IME 和跨页手势仍未验证 |

表中 `features/…`、`model/…`、`commands/…`、`services/…`、`layout/…`、`rendering/…` 均指 `packages/office-studio/src/features/document` 对应目录；API 位于 `src/api`，UI 位于 `src/components` / `src/components/document`。Office 原生 IO 归 Root，见独立证据，不由 UI 写出 HTML 代替 DOCX。

## 新增接口、粘贴与来源契约

- `copyDocumentBlocksForPaste(blocks: readonly DocumentBlock[]): DocumentBlock[]` 作为可复用复制服务；检查嵌套/块/逻辑格预算和字段端点后再 remap。每次 paste 使用新 IDs；存活目标 anchor 保留其原 cell ID，拆出的新 anchor 使用新复制 ID。
- `prepareDocumentTablePasteMutation(before, selection, sourceBlocks, context, mode)` 产生完整候选；`mode` 为 source / destination。目标表格的尺寸、列类型仍属目标状态，正文/拓扑按实际 footprint 改变；单格 replacement 保留原 rowspan/colspan，不把仅选中首格的操作扩大为清空全部旧选择。
- `doc.command.paste-blocks` 新增可选 `tablePasteMode?: 'source' | 'destination'`；没有增加平行 command ID。`documentTablePasteCommand()` 的第 5 参数接同一 mode；匹配目标格式按每个实际目标格获取段落、文字、fill/margin/borders。
- `documentTableClipboardData(model, view): DocumentClipboardData` 提供 HTML/plain；`writeDocumentClipboard(writer, data): Promise<boolean>` 返回是否写入 rich payload。只成功 writeText 时，rich cut 保留原内容并恢复 Canvas 焦点。真实系统权限/能力失败行为待验证。
- 续格 HTML 使用 `<div data-office-table-continuation="v1" data-office-row-offset="1">…</div>`；解析时重新绑定当前 ownerCellId、rowId、colSpan。它表示独立正文容器，不是原生 OOXML 来源的替代物。
- 横向合并成员只属 `nativeHorizontalMergeSource` provenance，**不参与正 offset 的 vertical continuation 协议**。复制不会以旧 member/block IDs 恢复旧内容；其 source refs 在跨文档 IO 中是否需要额外 remap 尚未核实，待与 native IO 来源协议共同验证。
- HTML reader / 单次块复制候选使用 64 层、250000 blocks / logical cells 预算；单表 reader 使用 10000 rows / 512 cols / 1000000 slots。原生 writer 有独立 32 层 / 整篇预算，跨域深度与错误行为尚未统一验收，不能称任意深度已可往返。

## 测试草稿与统一验证计划

以下均已编写、**未执行**；也未执行格式、语法或 diff 检查。

- [表格编辑合同草稿](../../packages/office-studio/tests/document-table-editing.contract.test.mjs)：合法 covered rows、预算、object-only select、clone/Shift contraction、结构格式/clear/TSV/search、mixed patch/undefined、邻格 ACL、独立续格输入/索引/unmerge/插删/copy 重绑、字段/批注/控件移动、rich paste source 不变/稳定目标 ID/拓扑/nested/destination formatting/越界与 merge-cut/空锁/实际 footprint 权限、同名 dropdown ID、输入 store 重验与 guarded completion、Facade/history/JSON reopening。
- [排版合同草稿](../../packages/office-studio/tests/document-table-layout.contract.test.mjs)：固定尺寸/整表位置/内边距、exact 溢出裁剪、rowSpan 可调整行、nested targets/deepest hit、重复表头与 page affinity、合并表头 closure、长行跨页完整文字、分页 gap 与 resize contentHeight。
- [工作台浏览器草稿](../../packages/office-studio/tests/document-table-workbench.browser.test.mjs)：handles / insert dots、resize 延后提交与 Escape、autoscroll / 跨页行移动、mixed 属性/焦点、同名选项、主题/classic/grid、HTML+TSV 复制、merged/nested/continuation 粘贴与重开、rowspan=0、多段 HTML、合成 composition。用例可暴露待修复实现；未启动 Vite 或浏览器。
- [旧表格工具栏合同](../../packages/office-studio/tests/docs-table-toolbar.contract.test.mjs) 已修订过时断言：当前上下文 17 根分组、Form 尺寸预算、命令发布后重新读取 table 对象、显式逻辑 columnIndex、非法 overlap/hole 与合法 vertical/mixed-horizontal 区分。
- Root 的 [DOC-22 原生表格草稿](DOC-22-native-tables.md) 另覆盖 Rust writer → Duke parser、DTO、source 身份、legacy merges 与 native widths / borders。真实 Office 样本与独立重开仍未执行。

## 真实剩余缺口与未核实项

- 页眉/页脚及 footnote/endnote 已有辅助块 host 与作用域命令；本批只复用其参数传递。完整 `storyContent.blocks` 表格目标继续开发和验收，不能以 body/nested 表格或 host 接线替代 story 编辑、选择、布局、权限及原生 IO 验收。
- 通用 HTML 不保存完整字段定义、SDT、媒体/其他资源依赖与原始 Office provenance；正文、跨度与续格结构保留不能证明完整跨文档资源粘贴。Normalized blocks 的字段 ID/remap 与跨文档共享资源导入须分开验证。
- tracked row/cell/property revisions、gridBefore/gridAfter 不规则格、零/塌缩宽度或 exact height、原生条件样式/编号与复杂字段/SDT 相交等仍需 native IO 与域模型补齐，具体 IO 缺口见 DOC-22。
- mobile/触控、floating occlusion、跨页拖动 source 控件生命周期、真实剪贴板权限、主题弹层、焦点与中文 Windows IME 尚未验收；合成事件草稿不等同真实 IME 或系统剪贴板验证。
- UI 及算法候选没有真实大样本 Worker/Rust/WASM、Canvas2D/WebGPU 对照、内存/复制/主线程长任务或同数据集前后基准。不得声称已优化或性能达标。

跨页碎片与重复表头、嵌套内容、控件锁/字段跨段/批注、ACL、undo/redo、JSON/DOCX 重开、Canvas2D/WebGPU、主题/焦点/IME与大样本均留待统一验证，DOC-06 继续保持未验收。

## 2026-11-27 顶部表格插入操作接线

本轮回读固定 [schema.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts)、[menu.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts)、[doc-create-table.operation.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/doc-create-table.operation.ts) 和 [doc-table-create.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/table/doc-table-create.command.ts)。表菜单是一个常驻根，INSERT / MEDIA，order 2、grid r1c2；插入表格为一个子项，order 0，执行 `doc.operation.create-table`。操作打开默认 3 行、5 列的尺寸弹窗，确认转发 `{ rowCount, colCount }` 到 `doc.command.create-table`。根内子项、classic/grid 和 More 副本不增加常驻计数。

[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 的 `doc-table` 和 `doc-table-insert` 已由本地 open-table 改接真实 create-table ID；[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 接受真实 ID，并保留旧 open-table 调用。沿用 [DocumentTableInsertDialog.vue](../../packages/office-studio/src/components/document/DocumentTableInsertDialog.vue) 的 Ant Modal/Form/InputNumber 和 3×5 默认值，经 `insertTableFromDialog` 将确认尺寸作为 `doc.command.create-table` 的 rowCount/colCount 提交，保留已有插入选项并进入现有模型与内容历史。默认页宽接线见下文；没有新增 UI、内容模型或 Worker/WASM 协议。

[docs-table-toolbar.contract.test.mjs](../../packages/office-studio/tests/docs-table-toolbar.contract.test.mjs) 同步真实根与子项 ID；[document-toolbar.browser.test.mjs](../../packages/office-studio/tests/document-toolbar.browser.test.mjs) 新增一个真实 Ribbon → 菜单 → Ant 尺寸表单 → 插入的草稿，检查默认 3×5、确认 4×7、保留作者文字、撤销/重做和内部快照重开。全部草稿未运行，不把源码读取断言当成实际 UI 或 Office 验收。

本段记录顶部操作接线；公开内容命令与尺寸预算的后续接线见下节。从 Studio 注册表转发到 UI 操作的完整宿主协议、结构 gap 和完整 story 插表仍需继续补齐；实际主题、焦点恢复、输入法、浏览器、Office 样本和大样本保持未验收。常规根数与完整功能目标没有因本批改名增加或缩减。

## 2026-11-27 公开创建命令与尺寸预算

本批补读固定 [doc-table-create.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/table/doc-table-create.command.ts)、[doc-table-create.command.spec.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/__tests__/doc-table-create.command.spec.ts) 和 [TableCreate.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/views/table/create/TableCreate.tsx)。公开内容命令只要求 `rowCount: number`、`colCount: number`；上游优先消费段落菜单插入范围，否则读取当前 Docs 文本选区，并以当前 segment 和 skeleton 的页内容宽度构建表格，合成一次 RichTextEditingMutation。测试明确覆盖当前选区、无活动选区时的菜单插入范围、表格/高级块内部与 custom block 拒绝、表格前后边界和缺少 skeleton 的拒绝。上游数字表单上限是 20 行、20 列，不能据此推断本地已有大尺寸插入链路只能处理 20×20。

- [table-command-types.ts](../../packages/office-studio/src/features/document/commands/table-command-types.ts) 增加 `DocumentTableCreateCommand`，用必填 rowCount/colCount 接入现有 DocumentCommand；本地可选插入属性沿用 `DocumentTableInsertOptions`。[document-editor-commands.ts](../../packages/office-studio/src/api/document-editor-commands.ts) 增加创建命令标签。
- [table-insertion.ts](../../packages/office-studio/src/features/document/commands/table-insertion.ts) 在已有插入校验中映射公开尺寸，保留旧 insert-table 的 rows/cols、尺寸推导、1×1 默认值、初始数据和样式等选项。新增共用尺寸校验供插入服务与 [document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 适配器复用：正安全整数、最多 10000 行 / 512 列 / 250000 格。10000 行 / 512 列来自现有网格构建范围，250000 格是原单次插入内存预算；Ant 数字表单也使用上述行列范围。原参数校验只限制正整数与格数，但真实插表还经过网格限制，本批将相同拒绝提前到生成单元格之前，没有收窄旧 API 的实际可用范围，也没有建立第二套表格引擎。
- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 的 create-table / insert-table 共用原插表分支，并共用候选模型、实际位置权限、字段与批注/控件锚点处理和发布路径；没有二次进入 command bus 或追加逐格内容命令。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 单独注册创建命令。公开 flat `{ rowCount, colCount }` 读取当前活动 Docs 的有效正文输入会话，保留已有块 host 和 revision/history/inputGeneration，直接进入同一个 commitDocumentEditorCommand，返回 changed:boolean；原 `{ unitId, command, selection, ...guards }` envelope 保留 DocumentEditorCommandResult。没有有效输入选区、输入会话修订不匹配或当前活动 Sheets 时，flat 返回 false；尺寸非法时抛出现有 RangeError，不发布部分内容历史。

[docs-table-toolbar.contract.test.mjs](../../packages/office-studio/tests/docs-table-toolbar.contract.test.mjs) 新增草稿覆盖 21×21、合法预算边界、缺失/错误类型/非整数/行列与格数越界、old/create 在既有 10000 行 / 512 列边界共同拒绝且不改变模型/历史、旧默认值与初始数据/几何保留、真实 Studio flat/envelope 注册、一次 CommandExecuted/内容历史、撤销重做和内部快照重开；已有 footnote 块 host 的接线另覆盖正文不变与 note 内容保存。这些草稿均未执行，未启动服务或浏览器，也未运行类型、构建、Rust/WASM 或性能检查。

[document-table-grid.ts](../../packages/office-studio/src/features/document/model/document-table-grid.ts) 的原有 10000 行 / 512 列构建限制保留。10001×1 / 1×513 虽曾能通过旧尺寸参数校验，真实 createDocumentTableBlock 仍会被该网格限制拒绝；本批只是提前拒绝，未宣称这些尺寸此前可插入或已获得真实性能改善。

本地插入继续使用原位置/范围替换与嵌套能力；后续小批已补 canonical 默认页宽，旧 insert-table 保留 100 CSS px/列，具体规则和物理页差异见下节。菜单插入范围消费和结构 gap 尚未对齐。现有块 host 参数传递不等同完整页眉页脚及 footnote/endnote 表格编辑、布局、权限与 Office IO 已通过。实际弹层主题、焦点/IME、浏览器、DOCX 文件重开和大样本测量保持未验收，DOC-06 总体状态不变。

## 2026-11-27 创建表格的默认页面宽度

固定 [doc-table-create.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/table/doc-table-create.command.ts) 从插入位置的 glyph 取得其实际物理页，并将 `pageWidth - marginLeft - marginRight` 传给 `genTableSource`。该值是整页正文宽度，不是节内单栏或父表格单元格宽度。[doc-table-create.command.spec.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/__tests__/doc-table-create.command.spec.ts) 的 2×3 样本使用 540 px 页宽、左右各 90 px 页边距，得到三列各 120 px。

- [table-insertion.ts](../../packages/office-studio/src/features/document/commands/table-insertion.ts) 的 `createDocumentTableBlock` 增加可选插入点和来源文档参数。canonical 未设置 width 时复用 `documentEffectiveSectionSetup`，按目标段落的顶层祖先定位所在节，并以该节整页正文宽度均分列；不减栏间距或除栏数。辅助 host 复用其作用域模型中已记录的 section ID，从来源文档读取节几何，保留后续节的宽度与页边距。
- [document-commands.ts](../../packages/office-studio/src/features/document/commands/document-commands.ts) 在原插表分支先解析 offset、documentEnd 或 selection，再将 `orderedSelection(...).start` 交给表格构建。反向选区按真实替换起点计算，offset 优先于 documentEnd；范围替换、块插入、权限、锚点和原子发布路径继续共用。已有第 5 个来源文档参数命名扩为 sourceModel，沿递归 host 接线保留；原超链接目标解析行为不变。
- [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的顶部确认已调用 canonical 命令；当当前布局属于同一个模型且不是 inspection 虚拟页时，复用 `lineForPoint` 定位当前选区起点，并将 `pages[line.page].content.width` 传入 width。显式对话框 width 优先；没有可用物理页时由内容命令按有效节设置计算，不为创建默认宽度同步重排整篇文档。

单列 `columnWidths[column]` 优先于 `width / colCount`；canonical 缺省 width 使用上述页面宽度，旧 insert-table 缺省仍为 100 px/列。已有 style.columnWidths 随后整体替换列宽，style.width 再按解析后的比例缩放；初始内容、标题行、行高、边距、列类型及扩展元数据保持原处理顺序。传入命令对象不会因默认宽度计算被修改。

[docs-table-toolbar.contract.test.mjs](../../packages/office-studio/tests/docs-table-toolbar.contract.test.mjs) 新增 4 个声明，文件共 35 个声明：固定 540/90/90 样本与旧默认、节级覆盖和多栏的整页宽度、offset/documentEnd/反向选区、显式宽度/部分与完整列宽/style 优先级、嵌套表格所在节。已有 flat 内容历史及 footnote flat/envelope 草稿补列宽断言，后者使用第二节的 720/90/120 几何，连同撤销重做与内部快照重开检查来源节设置未丢失。根任务同步补顶部浏览器草稿；这些声明均未运行，也未执行类型、构建、Rust/WASM、服务或基准。

已知差异：headless canonical 未收到实际 width 时使用 nominal 节几何。本地 [document-section-flow.ts](../../packages/office-studio/src/features/document/layout/document-section-flow.ts) 的连续分节可沿用前一节物理页的边距，到新页才采用当前节设置，因此不同页边距的 continuous 场景可能与实际物理页宽度不同；精确页定位仍待核实。顶部已接当前实际页宽度，但完整辅助 host/inspection、跨页光标与布局版本行为尚未通过浏览器验证。嵌套表格沿用本地增强能力，不能由默认宽度补齐推断上游限制、跨页布局或 Office 往返全部等价。

本批未新增可替换 UI，沿用 Ant Modal/Form/InputNumber；真实主题、焦点恢复、Windows IME、Canvas2D/WebGPU、DOCX 文件样本与性能保持未验收。DOC-06 状态仍为未验收。

<a id="table-insert-live-disabled"></a>

## 2026-11-27 插表弹窗的实时禁用状态

复读固定 `schema.ts:345–353`、`menu.ts:1156–1193`、`doc-create-table.operation.ts` 和 `TableCreate.tsx` 后，确认顶部仍为一个表格根、一个插入子项、默认 3×5，尺寸确认进入 `doc.command.create-table`。没有新增常驻、扩展、溢出或上下文根；本轮根增量为 0。

[DocumentTableInsertDialog.vue](../../packages/office-studio/src/components/document/DocumentTableInsertDialog.vue) 增加可选 disabled，并传给 Ant Form、Modal 的插入按钮与既有提交函数。[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的现有调用同步输入会话失效、编辑器组合输入、legacy story 与当前选区权限状态。表单全部输入控件通过 Form 继承禁用态；指定宽度输入已有显式 disabled，本轮合并总状态，避免 `InputNumber` 的显式 false 覆盖 Form 上下文。实际安装的 `antdv-next/dist/input-number/index.js`、`input/TextArea.js`、`switch/index.js` 均优先读取控件的 disabled，再读取上下文；`form/Form.js` 提供该上下文。

禁用期间保留表单草稿和取消操作，尺寸不合法的现有提示仍只由尺寸与数据校验决定。恢复权限或结束编辑器组合输入后可继续编辑，沿用原 create-table、候选模型、权限、内容历史与保存链路。本轮没有修改业务命令、WASM/Worker 协议或保存格式；组件接口仅增加可选 disabled。

[document-table-workbench.browser.test.mjs](../../packages/office-studio/tests/document-table-workbench.browser.test.mjs) 增加一个 test.each 声明，共四个案例：light/classic、dark/grid 下分别在打开弹窗后通过真实 FDocument 权限 API 收回编辑权，或对现有 Canvas 输入桥发送合成 composition。草稿检查五个数字输入、两个开关、初始数据及插入按钮禁用，指定宽度保留、取消仍可用、Form submit 不发布模型/历史，恢复后继续使用同一表单。全部未运行；合成事件不是 Windows 原生 IME 或弹窗文本输入的验收证据。

相邻的[分隔线](DOC-02-horizontal-line.md)与[分隔符](DOC-03-section-breaks.md)本轮仅复核现有注册、默认值与执行路径。结构 gap、非空段落下边框绘制、分隔符完整辅助 host 状态、插表模式/位置和物理页差异继续保留缺口。实际弹层尺寸、主题、hover、关闭后焦点、原生 IME、浏览器、Office 文件样本与性能均未验收。

本轮限定静态检查通过：两个 Vue 脚本/模板、一个浏览器草稿及一段夹具脚本、两个样式块、126 处本地引用；两份专项文档的 48 条本地链接均存在，五个修改文件无尾随空白或冲突标记，限定 diff 检查通过。首次解析器入口未在根依赖中暴露，随后改用实际安装的 `vue/compiler-sfc` 与 Vite 的 PostCSS 依赖完成检查；没有安装依赖、执行产品模块或运行测试。

<a id="table-workbench-input-session"></a>

## 2026-11-27 表格命令的工作台输入会话边界

本批沿用已完成的表格清点：一个常驻 Insert 根及一个插入子项；Pro 替换同一根，不重复计数；17 个 Pro 上下文 Ribbon 根，Style 5 / Layout 6 / Data 2 / Options 4；OSS 右键两个结构根、七个叶项；独立 overflow 根为 0。本次新增入口数为 0，这些计数不代表功能完成。

固定来源继续使用 v1.6.0、commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`：

- [schema.ts:345–353、457–492](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts#L345) 定义 Insert / MEDIA、order 2、grid r1c2 和右键插入四向、删除行/列/表；[menu.ts:1156–1193](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/menu.ts#L1156) 与 [context-menu.ts:93–150、235–397](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/context-menu.ts#L93) 根据真实 Docs 文本/表格范围、模型、skeleton 与编辑权限判断入口状态。OSS 根隐藏于页眉页脚；Pro 发布代码的 Insert factory `ni`（字符串位置 7660）及 `ai` 使用当前表格选区显隐，不能直接套用 OSS 的辅助 story 限制。
- [doc-create-table.operation.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/doc-create-table.operation.ts) 默认 3×5，确认进入 [doc-table-create.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/table/doc-table-create.command.ts)；后者使用插入范围、segment、实际物理页宽，合成正文和 tableSource 的一次 mutation。[doc-table-insert.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/table/doc-table-insert.command.ts) 的四向包装器复用插行/插列，[doc-table-delete.command.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/table/doc-table-delete.command.ts) 在行列全部删除时转删表；均更新同一正文/tableSource 事务。
- [doc-state-change-manager.service.ts:214–257](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/services/doc-state-change-manager.service.ts#L214) 按 unitId 聚合 undo/redo mutations；[f-document.ts:348](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs/src/facade/f-document.ts#L348) 经资源加载器保存文档及资源。Pro 的 `H` / `U`（21846 / 22640）和 `Ja` / `Ya` / `Za`（23688）、`Qa`（26117）静态记录共享命令 factory、表格选区显隐和上下文复用；发布类型或代码读取不构成运行验收。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 tableRibbonContext 将输入会话失活并入已有 composing 状态，供按命令能力判断的表格 Ribbon 项、七个基础右键项和附加属性入口使用。execute 的既有入口守卫加入 `isDocumentTableCommand`，并显式加入不属于该谓词的 `doc.command.create-table` / `doc.command.insert-table`；失活或编辑器组合输入期间，延迟内容回调在辅助 host 转换及候选提交之前返回 false。后续表格能力分支继续调用真实 documentTableCommandState，移除已由入口处理的重复组合输入判断。

inputStateActive 沿用原普通 let；tableRibbonContext 显式读取既有 storyActionVersion，使用 onActivated / onDeactivated 原有的递增保证失活和恢复均重新计算 composing。失活时 clearInputMarks 经 FDocument / Studio 进入 [document-input-store.ts](../../packages/office-studio/src/features/document/stores/document-input-store.ts) 的 clearMarks，即使 marks 已为空也替换 units shallowRef、设 valid:false，pendingMarks → toolbarState 原链路也会失效。恢复时 applyInputState 会拒绝该失效 session，因此控件重算不能只依赖重新应用输入状态；本批使用生命周期版本，不改全局活动标志或 API/session 规则。

本次改动属于 Vue 工作台输入会话边界。表格命令、detached planner、真实跨度/预算、模式、ACL、内容控件锁、锚点更新、一次发布和按文档隔离历史均沿用原链路；标准 FDocument / Studio envelope 内容命令仍可在没有 UI 输入会话时调用。公开 flat 创建命令原有的有效正文输入要求保持。接口、领域模型、保存格式及 Worker/WASM 协议变更均为 0；没有新建计算路径或可替换 UI。

[document-table-workbench.browser.test.mjs](../../packages/office-studio/tests/document-table-workbench.browser.test.mjs) 在原夹具加入最小 KeepAlive 隐藏/恢复，并补一个 test.each 声明、四个计划案例：light/classic 和 dark/grid 分别覆盖合成 composition 与工作台失活。草稿读取真实 Ribbon 状态，检查已打开右键菜单的实时禁用；向真实 Ribbon 发送合法 create、旧 insert、插行和背景回调，检查模型、内容历史及模型变化事件不变。恢复后在任何 focus、selection 或右键操作前检查六个表格控件重新启用，再通过实际右键插行，检查原单元格 ID、一次历史、Canvas 输入焦点、撤销/重做和公开保存后内部重开的领域 blocks；不将 FDocument.save() 的公开数据误读为领域模型。既有插表弹窗草稿同步实际 grid 选择器。

本批结束时该文件共 12 个 test 声明、18 个计划案例，本批新增四个；运行结果为 0，下一节继续补充。所有控件继续使用既有 antdv-next / ConfigProvider；当时限定静态检查通过：三个脚本（Vue、浏览器草稿及已解码夹具）、一个模板、123 处本地 import、35 条本地链接、18 处固定上游链接源路径、两个稳定锚点，以及该 mjs 的 Node 语法和三文件 diff 检查。实际主题弹层、hover、焦点恢复、原生 Windows IME、浏览器与 Office 文件样本仍未验收；合成事件不代表原生 IME 已通过。

本批没有修改打开属性、create/open-table 或浮动切换，其后续会话小修见下一节；通用右键回调与打开时选区的绑定契约仍待核实。本地增强嵌套能力未改成 OSS 限制。完整 story 表格、结构 gap、原生 Office 往返及大样本 Canvas2D/WebGPU、Worker/Rust/WASM 与内存/复制基准继续保持未验收。DOC-06 总体状态不变。

<a id="table-ui-operation-sessions"></a>

## 2026-11-27 表格属性与 UI 操作的会话状态

本批继续使用前节固定 v1.6.0 来源和已有入口分类，新增入口为 0。Pro `docs-table-ui/lib/types/commands/toggle-docs-table-floating-toolbar.operation.d.ts`、`services/docs-table-floating-menu.service.d.ts`、`views/menu.d.ts` 声明切换操作、enabled$ 与菜单 factory；发布 ES 的 `pi`（11701）为 `docs-table-ui.operation.toggle-floating-toolbar`，仅调用 service.toggle()，`va`（19608）将 enabled$ 映射到 activated$，没有 edit 权限或 composition 条件。该操作是 UI 偏好，不是内容 mutation；本地保留活动只读和组合输入时的开关行为。Insert factory `ni`（7660）及 schema `Za`（23688）继续对应既有常驻根与上下文选项位置。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 openTableProperties 在原组合输入/legacy story 入口守卫中加入输入会话失活；create/open-table 在打开前复用既有插表弹窗的活动会话、组合输入、legacy story 与 canEditContextSelection 判断。浮动切换只拒绝失活后的延迟回调，活动只读或组合输入时仍可修改 UI 偏好，不发布模型或内容历史。属性弹窗的 composing 改用 tableRibbonContext.composing，复用前节生命周期版本依赖，同时覆盖失活、组合输入和已有 linked story 状态。

[DocumentTablePropertiesDialog.vue](../../packages/office-studio/src/components/document/DocumentTablePropertiesDialog.vue) 的四处显式 disabled 合并 composing：表格/单元格背景 ColorPicker、行高 InputNumber 和四边 margin InputNumber 的共用表达式。现有清除背景、自动行高和继承边距条件继续生效。实际安装的 antdv-next `ColorPicker.js:32` 使用 `props.disabled ?? contextDisabled.value`，`input-number/index.js:86` 使用 `customDisabled.value ?? disabledContext.value`；显式 false 会覆盖 Form 上下文，本批直接修正这四处表达式，保留 Form、Modal、取消操作和表单草稿。

属性通用编辑权限、requestModel/requestRange、原 applyTableProperties 的 host 比对、host 改变时关闭属性、创建确认使用当前选区/host、session 与内容命令规则均保持。属性提交继续进入 [table-commands.ts](../../packages/office-studio/src/features/document/commands/table-commands.ts) 的真实能力判断和 [table-mutations.ts](../../packages/office-studio/src/features/document/commands/table-mutations.ts) 的 detached 候选、ACL/锁/锚点与一次发布，内容历史按文档隔离；没有为只读查看添加通用 edit 门槛，也没有新增创建目标快照。

保存仍经 [document-output.ts](../../packages/office-studio/src/api/document-output.ts) / [document-body-table.ts](../../packages/office-studio/src/api/document-body-table.ts) 的公开 body/tableSource 与扩展元数据。项目原生 [docx_table_styles.rs](../../crates/office-writer/src/docx_table_styles.rs) 读取真实宽度、对齐、偏移、内边距、背景/边框、行高规则、跨页拆分及重复表头；[DOC-22](DOC-22-native-tables.md) 记录 writer、作者扩展和原生来源边界。本批没有改保存/写出链路，静态读取不证明实际 Office 往返。

[document-table-workbench.browser.test.mjs](../../packages/office-studio/tests/document-table-workbench.browser.test.mjs) 新增两个 test.each 声明、12 个计划案例：属性声明八例，分别在 light/classic、dark/grid 下覆盖 table、row、cell、background，每例顺序检查 composition 和 KeepAlive 失活两轮；另一声明四例检查失活/组合输入期间的延迟属性、create、旧 open 与浮动切换，并通过真实 FDocument 权限 API 验证只读创建拒绝和浮动偏好仍可切换且无内容历史。

属性草稿使用真实 Ant 色值/数字输入准备可提交的修改，读取实际控件和按钮禁用状态，检查取消保留、Form submit 不改变模型/历史/事件；恢复后在重新 focus/selection 前检查控件启用，再通过真实应用提交一次历史，撤销/重做及公开保存后内部重开领域 blocks。控件数量有明确断言，不允许空集合使禁用检查通过。没有新建夹具文件、Mock、API、DTO、领域模型或 Worker/WASM 协议。

该浏览器文件当前共 14 个 test 声明、30 个计划案例；本批新增 2 个声明、12 个案例，属性声明的两轮检查仍属各自同一个计划案例。运行结果为 0，主线统一静态检查与后续运行验证另行记录。本批沿用现有 antdv-next 与主题 ConfigProvider；真实主题弹层、hover、焦点/Windows 原生 IME、完整辅助 host 表格、创建/右键目标绑定、浏览器/Office 文件样本和性能继续保持未验收，DOC-06 总体状态不变。
