# Sheets Ribbon 源码入口清点（2026-11-21）

参考源码固定为 Univer `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地固定检出位于 `E:\rust_test\univer`。本清点只追踪 desktop Sheets preset 的 Ribbon、上下文菜单注册和本地 command dispatcher，不包含浏览器或像素验收。

2026-11-22 使用 `scripts/audit-univer-oss-ribbon.mjs` 重新读取固定 v1.6.0 的 OSS 注册 schema：常规根声明为 `30/9/12/6/6`，共 63 项。2026-11-26 人工回读当前 `createSheetRibbon`，包含范围主题、图片与背景根，为 `30/7/11/17/6`，共 71 项；69/68 是此前清单快照。[局部移位](SHEET-02-cell-shift-transactions.md)、[九区段编辑](SHEET-04-range-theme-editor.md)、[图表面板 route](SHEET-22-chart-edit-panel-routing.md)、[P18 线型](SHEET-22-line-and-area-styles.md)与[共享绘图工作台](SHEET-21-drawing-workbench-affine.md)不新增常规根。factory、动态显隐与实际装配尚未执行，不作为完整运行态入口数或验收结论。

## 口径

- 一个“根入口”是合并菜单后位于 Ribbon Tab / Group 下的顶层菜单节点。Split、menu/select 的子项不重复计数；动态公式分类按实际生成的分类菜单计数。Context menu、浮动对象菜单和 Footer 菜单单列。
- Univer 菜单由 `packages/ui/src/services/menu/menu-manager.service.ts` 提供 Tab/Group 顺序，再由 `examples/src/sheets/mount.ts` 装配 preset，以及 sheets-ui、numfmt、formula、filter、sort、table、drawing、hyperlink、thread-comment、data-validation、conditional-formatting 和共享 find-replace 插件合并注册。
- 本地读取 `createSheetRibbon()` 的根数组，children 不重复计数；公式 8 个根分类由当前 117 项 catalog 的类别静态展开。常规 schema 为 `30/7/11/17/6`，包括条件项。日期时间已作为快速日期子项；插入函数作为常用函数菜单子项；透视表、切片器和时间线保留在 Data，图片/迷你图位于 Insert，背景图位于 View。
- 固定 OSS 顶层父菜单声明为 `30/9/12/6/6`。View 的十字高亮由 `sheets-crosshair-highlight/src/menu/schema.ts` 注入，属于固定版入口；此前将其记为本地独有扩展不准确。
- 根入口数只比较 schema 规模，不能说明功能完整。子菜单、上下文入口、权限隐藏、动态状态、溢出菜单、Pro 规格入口和实际命令效果分开核对。

## 标签与分组

| 标签 | Univer 根入口与来源顺序 | 本地根入口与现状 | 主要差异 |
| --- | --- | --- | --- |
| START / 开始 | 30。HISTORY 5、FORMAT 11、LAYOUT 6、NUMBER 5、OTHERS 3。入口来自共享 ui、sheets-ui、sheets-numfmt-ui 的 `menu/schema.ts`，组顺序来自 `menu-manager.service.ts`。 | 30。clipboard 5、font 11、layout 6、number 5、tools 3。 | 水平/垂直对齐、换行、旋转、合并均已收为对应父菜单；数字五项、颜色/最近颜色/重置、字体/字号与边框已有共享命令切片，均待统一验证。Start tools 保留保护范围、本地清内容和范围主题；主题为第 3 根，grid r1c3 跨两行，快捷键/功能搜索仍需核对。 |
| INSERT / 插入 | 9。MEDIA 4：图片、超链接、评论面板、添加评论；CELL 5：复选框、下拉、日期、快捷数据条/图标集。 | 7。media 4：图片、超链接、图表、迷你图；cells 3：复选框、下拉、日期父项。 | 图片根 order 0、r1c1 跨两行，有浮动/单元格图片两个子项；超链接 r1c2。日期时间归日期子项。线程评论及两类快捷条件格式仍缺完整入口闭环，图片/图表/迷你图所有真实行为与 Office 仍未验收。行列结构编辑由上下文承载。 |
| FORMULAS / 公式 | 12。BASIC 为常用加十类函数菜单，OTHERS 为计算选项；来源 `sheets-formula-ui/src/menu/schema.ts`。 | 11。formula-edit 8 个 catalog 分类，calculation-options 3 项：模式、编辑公式、重算。 | 插入函数已收进常用菜单；完整函数目录、分类、参数/返回类型和求值契约仍未齐。仅复制/粘贴公式已在本地选择性复制/粘贴菜单，仍须按固定命令与历史核对。 |
| DATA / 数据 | 6。RULES 2、ORGANIZATION 3、OTHERS 1；来源 validation、conditional-formatting、sort、filter、table、find-replace 的 schema。 | 17。organization 2、rules 2、pivot 4、outline 8、data-tools 1。 | 排序父菜单、筛选 split、规则及查找已有入口；筛选复用切换/清条件/重应用命令。仍缺结构化表格选择器及全部 Pro 分组/分析语义。透视表、透视图、切片器与时间线不再重复于 Insert。 |
| VIEW / 视图 | 6。DISPLAY 4：冻结、缩放、背景图、十字高亮；VISIBILITY 2：网格线、全屏。 | 6。display 4：冻结、背景图、缩放、十字高亮；visibility 2。 | 背景根 order 0.5、r1c2 跨两行，设置/删除两子项接图片资源与命令；缩放/十字高亮在 c3。背景历史/Office 未验收；十字高亮颜色、冻结/缩放与视图历史按 SHEET-20 继续验收。 |

### 本地 handler 核对

`sheet-ribbon.ts` 的根项与子项由 `ExcelWorkbench.vue:executeRibbon()` 分派。格式、行列结构、合并、筛选/排序、验证、条件格式、保护、透视分析、迷你图、大纲、公式、冻结、缩放、十字高亮、超链接及全屏已有处理函数或面板；注册命令、权限、历史、保存及文件互操作仍逐项追踪，存在 handler 不等于功能验收。

此前的三水平对齐根项、Home 查找、Insert/Data 重复透视入口、日期时间独立根项已经调整。大纲与分析属于 Pro 扩展，须按 v1.6.0 发布包再核对位置与行为，不能只按 OSS schema 判为无对应功能。图片和背景已有真实入口/命令，见[图片证据](SHEET-19-images.md)；评论、表格选择器与打印仍缺完整顶部编辑闭环。超链接和十字高亮不再列为无 handler；所有功能仍未验收。

本地 `sheet.command.clear-content`、`clear-format`、`clear-all` 在右键清除父菜单中分开派发，开始仍保留独立清格式和扩展清内容；[合并/清除证据](SHEET-02-merge-clear.md)追踪共享命令、确认、权限、历史和旧模型适配。文字色/底色有独立重置及完整选择器。筛选已有切换、清条件、重应用及条件面板。上述各入口的成功/失败原子性、所有权限和历史闭环仍未统一验证。

| 入口组 | Univer 菜单注册/调用 ID | 本地 ID / 处理器核对 |
| --- | --- | --- |
| START | `UndoCommand`、`RedoCommand`、`ActivateFormatPainterOperation`、`ClearFormattingCommand`、`PASTE_RIBBON_MENU_ID` 下的 `SheetPasteCommand`/Value/Format/ColWidth/BesidesBorder；字体/样式注册为 `SetRangeFont*Command`、`SetRangeBoldCommand`、`SetRangeTextColorCommand`、`SetBackgroundColorCommand`、`SetBorderBasicCommand`、`SetHorizontalTextAlignCommand`、`SetVerticalTextAlignCommand`、`AddWorksheetMergeCommand`/All/Vertical/Horizontal、`RemoveWorksheetMergeCommand`、`AddRangeProtectionFromToolbarCommand`；数字格式由 `OpenNumfmtPanelOperator`、`SetPercentCommand`、`SetCurrencyCommand`、`AddDecimalCommand`、`SubtractDecimalCommand` 注入。 | `sheet.command.undo/redo`、`sheet.operation.format-painter`、`sheet.command.clear-format`、`sheet.command.paste*`、字体/颜色、边框、对齐、合并、保护范围与数字格式 handlers 已存在。固定 `sheets/src/commands/commands/set-style.command.ts` 还定义 `ResetTextColorCommand` / `ResetBackgroundColorCommand`，本地重置已注册；完整颜色控件与共享字体路径本批仍在补齐。`clear-content` 是本地额外入口，不能替代清格式。 |
| INSERT | `SHEETS_IMAGE_MENU_ID` → `InsertFloatImageCommand`/`InsertCellImageCommand`；`InsertHyperLinkToolbarOperation`；`ToggleSheetCommentPanelOperation`/`ShowAddSheetCommentModalOperation`；`InsertQuickSheetDataValidationCommand`/`QUICK_DROPDOWN_MENU_ID`/`QUICK_DATE_MENU_ID`；`AddDataBarConditionalRuleCommand`/`AddIconSetConditionalRuleCommand`。 | `sheet.menu.image` → `sheet.command.insert-float-image` / `sheet.command.insert-cell-image`，接 Ant picker、资源准备及统一事务；`sheet.operation.insert-hyperlink` → 弹窗及共享 set/remove；复选框、下拉、日期父项复用 quick validation。图表/迷你图是本地 Pro 入口，透视/切片器/时间线仅保留在 Data。线程评论与两类快捷条件格式仍缺完整顶部闭环；图片全部保真/Office 未验收。 |
| FORMULAS | `InsertFunctionOperation.id` 的 common/financial/logical/text/date/lookup/math/statistical/engineering/information/database 分类、`MoreFunctionsOperation.id`、`CalculationOptionsOperation.id`；`SheetCopyFormulaOnlyCommand` 与 `SheetOnlyPasteFormulaCommand` 注册在 copy/paste special 上下文菜单。 | 本地函数分类/完整函数面板入口、`start-formula`、`recalculate`、`set-calculation-mode` 有处理器。上下文 `copy-formulas` / `paste-formulas` 均已声明，仍须核对固定命令参数、引用转换与权限/历史边界。 |
| DATA | `OpenFindDialogOperation.id`；`SmartToggleSheetsFilterCommand.id`，children `ClearSheetsFilterCriteriaCommand.id`/`ReCalcSheetsFilterCommand.id`；`SortRangeAsc/AscExt/Desc/DescExt/CustomCommand.id`；`OpenTableSelectorOperation.id`；`OpenConditionalFormattingOperator.id`；`OpenValidationPanelOperation.id`/`AddSheetDataValidationAndOpenCommand.id`。 | 筛选 split 已派发 `toggle-auto-filter` / `clear-filter` / `reapply-filter`；条件面板另走 `open-data-panel`。升降序菜单已有 `{ extendSelection: true }` 子项及自定义排序。验证/条件格式在 RULES，查找替换在 Data 的 data-tools。表格选择器仍缺等价顶部入口。 |
| VIEW | `SHEET_FROZEN_TOOLBAR_MENU_ID`、`ZOOM_RATIO_MENU_ID`、`ToggleGridlinesCommand.id`、`ToggleFullscreenOperation.id`、`ToggleCrosshairHighlightOperation.id`、`WORKSHEET_BACKGROUND_IMAGE_MENU_ID` 下的 Add/Delete background commands。 | 冻结、zoom、网格线、全屏与十字高亮已有处理器；背景根和 `sheet.command.add-worksheet-background-image` / `sheet.command.delete-worksheet-background-image` 进入既有图片 API 与内容模型，存在性/编辑能力控制禁用；实际保存/重开未验收。 |

以上本地 IDs 由组件级 `executeRibbon()` 分发，并非都注册到共享应用 command bus。对当前可见 Ribbon 项未发现空回调/占位消息式 no-op；明确缺少的是表中未注册的源入口/处理器以及不等价的替代行为。命令有 handler 只证明入口已连接，不代表 mutation、撤销、权限、持久化、重开和公开规格完整。

## 上下文与工作表栏

固定源码在 `sheets-ui/src/menu/schema.ts` 注册独立 MAIN_AREA、COL_HEADER、ROW_HEADER、FOOTER_TABS 菜单，并由功能插件向上下文组注入排序、评论、超链接、表格等入口。

MAIN_AREA 源核心 IDs 包括 `SheetCopyCommand`/`SheetCutCommand`/`SheetPasteCommand`、`COPY_SPECIAL_MENU_ID`/`PASTE_SPECIAL_MENU_ID`、`CLEAR_SELECTION_MENU_ID` 下 content/format/all、`TEXT_TO_NUMBER_CONTEXT_MENU_ID`、`CELL_INSERT_MENU_ID`、`DELETE_RANGE_MENU_ID`、`SHEET_FROZEN_MENU_ID`、`SHEET_PERMISSION_CONTEXT_MENU_ID`。行列头另注册 multi-insert、hide/show、width/height、auto-width/auto-height 和各自 freeze/permission IDs；表签注册 `RemoveSheetConfirmCommand`、`CopySheetCommand`、`RenameSheetOperation`、`SetTabColorCommand`、`SetWorksheetHideCommand`、`ShowMenuListCommand` 及 sheet permission IDs。功能插件还在 MAIN_AREA 加 sort/filter、hyperlink、thread comments 和 table contexts。

| 上下文位置 | Univer 可见操作类别 | 本地实现与差异 |
| --- | --- | --- |
| MAIN_AREA | 快速复制/剪切/粘贴；仅复制/粘贴特殊类型；清内容/格式/全部；文本转数字；插入/删除单元格并可移动邻格；冻结；范围权限；插件注入排序、评论、超链接等。 | [共享移位](SHEET-02-cell-shift-transactions.md)接四 confirm ID/固定子项顺序、部分合并/筛选尾部、候选/资源/权限/历史；清除、特殊剪贴板、冻结/保护及已有链接有接线。排序/筛选、评论、链接编辑/删除与完整 table context 仍缺，所有行为未验收。 |
| COL_HEADER | 剪贴板/特殊粘贴；插入多列、隐藏/显示列、删除列、列宽/自动列宽、列冻结、范围权限。 | 结构/显隐入口传 sheetId/ranges，尺寸入口传 sheetId/index；命中选区、批量结构、spans/过滤/折叠隐藏组合与宿主保护见[轴事务](COMMON-03-input-selection-axis-history.md#3-行列命令与-ui)。完整剪贴板、字体/auto-height/Office 和性能未验收。 |
| ROW_HEADER | 剪贴板/特殊粘贴；插入多行、隐藏/显示行、删除行、行高/适应内容、行冻结、范围权限。 | 同列头按实际命中轴/选区传参，共享结构、显隐和尺寸命令/候选；[公开 h/ah/ia/w/hd 导入](SHEET-03-public-axis-import.md)及来源变换已落盘，真实组合场景和 Office/性能未验收。 |
| FOOTER_TABS | 删除/复制/重命名/标签色/隐藏/显示列表/工作表权限等。 | 生命周期/标签/保护和普通隐藏列表已有接线；[复制资源](SHEET-01-worksheet-copy.md)已处理可识别 chart/pivot/table 等依赖及身份重映射，不能继续记为仅复制 cells。全部未知资源/veryHidden/权限/文件重开仍未验收，详见[生命周期](SHEET-01-worksheet-lifecycle.md)。 |
| Drawing / 浮动对象 | 图片保存、绘图评论、组合/排列/对齐和对象上下文工具。商业图表/形状上下文需按公开规格核验。 | [图片入口/浮条](SHEET-19-images.md)、[图表四根](SHEET-22-chart-object-context-actions.md)、标题/绘图区和[Drawer route](SHEET-22-chart-edit-panel-routing.md)已有接线；[共享绘图](SHEET-21-drawing-workbench-affine.md)接 typed 选区、组合/排列/对齐三父项及 14 叶、另一本地删除、四控件浮条和真实 Canvas/native owner 命中。八个对齐 operation 前缀已修正；[切片器/时间线 affine 与 Ant 控件](SHEET-25-slicer-timeline-affine.md)和 [Pro 多目标连接](SHEET-25-slicer-connections.md)已接未验收；shape 完整作者/渲染、组 clipboard、交叉筛选成员状态/Pro API 和完整直接编辑仍缺。 |

2026-11-26 静态现状：行列插删与显隐按命中或已有合法整轴选区传 sheetId/ranges，尺寸/适应按 sheetId/index；共享命令、候选、资源和内容/选择历史接线见[轴事务](COMMON-03-input-selection-axis-history.md#3-行列命令与-ui)。隐藏使用点与 spans，手动、筛选、outline、零尺寸原因分别保留，不再把显示描述成无条件清除整轴。Canvas 字体度量、实际右键选区/焦点、全部资源/样本、Worker/Rust/WASM 大表路径及 Office 保真仍需统一验证；当前没有新增执行结果。

## 溢出与布局入口

嵌入工作台缺省 classic，可运行示例缺省 grid。`RibbonGridToolbar.vue` 按完整 group 测量并收纳到 More，classic 按项溢出；二者共享 schema、命令和状态。布局比较须采用相同变体，More 内的根项不增加统计，也不能用溢出替代缺失入口。

当前本地 Grid 的分组顺序、`gridLayout` 与语义组映射仍按上方逐 Tab 列表有偏差。完整宽度下出现于 schema 的额外根入口仍然是实际视觉布局的一部分，溢出只影响可见位置，不会补足缺失入口或修正错组。

## Pro 能力边界与状态

固定 OSS 检出不含 Pro 的原始实现；项目缓存中的 60 个 v1.6.0 ES 发布包已静态读取注册与 schema，完整目录见[Pro 入口记录](../UNIVER_TOOLBAR_PARITY.md#根入口统计与布局)。重复挂载、上下文、替换、浮动控件和未解析表达式仍待追踪。下表同时采用发布包、公开规格与 feature ID，不冒称已掌握闭源内部实现；完整目标及验收状态仍在 `UNIVER_FEATURE_PARITY.md`。

| 目标 | 当前入口/handler 状态 | 入口缺口或已知范围 |
| --- | --- | --- |
| SHEET-10 高级公式 | Formula catalog 入口和插入 handler 存在。 | 117 项本地 catalog、有限求值支持；完整函数、数组传播等未实现/未验收。 |
| SHEET-21 形状及绘图 | Insert 有图片/图表入口；共享选区、组合/取消组合、混合排列/对齐/transform/Delete、Canvas 浮条/右键和七项 FWorksheet wrappers 已接；后续方向键/API move-drawing 共用实际 skeleton 边界和 1000ms 连续内容历史。 | [slicer/timeline affine、五类叶/组、Ant 控件与 98/182 草稿](SHEET-25-slicer-timeline-affine.md)及 [Pro 多目标连接](SHEET-25-slicer-connections.md)已补入但未运行；shape/connector 完整作者和渲染、组 clipboard、交叉筛选成员状态/Pro API、Office affine/嵌套组、极端边界及大样本仍缺；原 82/98 主任务草稿仍独立保留。 |
| SHEET-22/23/24 图表、透视表、透视图 | 对应 Ribbon 入口与面板 handler 存在。 | 入口连接不等于范围绑定、编辑能力、刷新、持久化和 Office 往返完整；均未验收。 |
| SHEET-25 切片器和时间线 | Data 两个既有根保持；[连接批](SHEET-25-slicer-connections.md)已接格式四条件根、Ant 快速样式/设置/报表连接、shared cache 与 linked/owned 多目标筛选，常规根增量 0。 | [原 98/182 草稿](SHEET-25-slicer-timeline-affine.md)、[连接 UI](SHEET-25-slicer-connections-ui.md)、[复制/结构生命周期](SHEET-25-slicer-connections-lifecycle.md)及核心新草稿未运行；完整交叉筛选成员状态、日期/Pro API、缺依赖跨 workbook、Office 与真实 UI/性能仍未验收。 |
| SHEET-26 迷你图 | Insert 入口映射到 SparklinePanel。 | 完整图型/组行为/引用变换/导出未验收。 |
| SHEET-27 大纲分组 | Data 中有行列组合、折叠和汇总位置 handler。 | 与源 Data 分组位置不同；嵌套/筛选/手动隐藏组合未验收。 |
| SHEET-28 打印与 PDF | 无 Ribbon 入口及本地 print command handler。 | 完整缺失。 |

图片、超链接、批注/线程评论、结构化表格、工作表背景、十字高亮均在范围内，分别对应 SHEET-16/17/18/19/20；固定 OSS 源码有明确入口。图片、背景、超链接与十字高亮已有本地入口，所有完整工作流/权限/文件互操作仍按证据验收；评论和表格选择器等缺口不能归因于 Pro 源码不可得。

## 源码复核路径

- 固定源共享 Ribbon 结构和位置：`packages/ui/src/services/menu/menu-manager.service.ts`、`packages/ui/src/services/menu/types.ts`、`packages/ui/src/menu/schema.ts`。
- Sheets 根菜单：`packages/sheets-ui/src/menu/schema.ts`、`packages/sheets-numfmt-ui/src/menu/schema.ts`、`packages/sheets-formula-ui/src/menu/schema.ts`。
- Sheets 插件扩展：`packages/sheets-sort-ui/src/menu/schema.ts`、`packages/sheets-filter-ui/src/menu/schema.ts`、`packages/sheets-table-ui/src/menu/schema.ts`、`packages/sheets-drawing-ui/src/menu/schema.ts`、`packages/sheets-hyper-link-ui/src/menu/schema.ts`、`packages/sheets-thread-comment-ui/src/menu/schema.ts`、`packages/sheets-data-validation-ui/src/menu/schema.ts`、`packages/sheets-conditional-formatting-ui/src/menu/schema.ts`、`packages/find-replace/src/menu/schema.ts`。
- 固定 desktop 装配：`examples/src/sheets/mount.ts`。商业插件来源边界：固定版本 `README.md` 与 `UNIVER_FEATURE_PARITY.md` Pro 条目。
- 本地 schema / handler / context / overflow：`packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts`、`packages/office-studio/src/components/ExcelWorkbench.vue`、`packages/office-studio/src/components/spreadsheet/sheet-context-menu.ts`、`packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue`、`packages/office-studio/src/components/spreadsheet/SheetFooter.vue`、`packages/office-studio/src/platform/ribbon/RibbonGridToolbar.vue`、`packages/office-studio/src/platform/ribbon/ribbon-schema.ts`。

本清单不改变任何 feature acceptance 状态；未做浏览器、样本文件、权限、历史/重开或 Pro 互操作验收。所有功能 ID 继续按总清单保持未验收。
