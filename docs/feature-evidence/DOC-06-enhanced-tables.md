# DOC-06 表格与增强表格

状态：**未验收**。本批按固定上游完成静态入口与命令追踪，再推进本项目独立实现。未运行测试、类型检查、构建、浏览器、Rust/WASM 或基准；测试草稿和静态读取不代表功能、像素或 Office 往返已通过。

更新日期：2026-11-24。当前是候选实现与待统一验证的草稿记录；本次 DOC-06 切片收束不代表全部 Docs / Sheets / Pro 目标结束。

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

- 页眉/页脚及 footnote/endnote 目前保留原段落 story 契约，完整 `storyContent.blocks` 表格目标继续开发；不能以 body/nested 表格实现替代 story 编辑、选择、布局、权限及原生 IO 验收。
- 通用 HTML 不保存完整字段定义、SDT、媒体/其他资源依赖与原始 Office provenance；正文、跨度与续格结构保留不能证明完整跨文档资源粘贴。Normalized blocks 的字段 ID/remap 与跨文档共享资源导入须分开验证。
- tracked row/cell/property revisions、gridBefore/gridAfter 不规则格、零/塌缩宽度或 exact height、原生条件样式/编号与复杂字段/SDT 相交等仍需 native IO 与域模型补齐，具体 IO 缺口见 DOC-22。
- mobile/触控、floating occlusion、跨页拖动 source 控件生命周期、真实剪贴板权限、主题弹层、焦点与中文 Windows IME 尚未验收；合成事件草稿不等同真实 IME 或系统剪贴板验证。
- UI 及算法候选没有真实大样本 Worker/Rust/WASM、Canvas2D/WebGPU 对照、内存/复制/主线程长任务或同数据集前后基准。不得声称已优化或性能达标。

跨页碎片与重复表头、嵌套内容、控件锁/字段跨段/批注、ACL、undo/redo、JSON/DOCX 重开、Canvas2D/WebGPU、主题/焦点/IME与大样本均留待统一验证，DOC-06 继续保持未验收。
