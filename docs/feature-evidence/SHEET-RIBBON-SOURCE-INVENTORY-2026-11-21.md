# Sheets Ribbon 源码入口清点（2026-11-21）

参考源码固定为 Univer `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`，本地固定检出位于 `E:\rust_test\univer`。本清点只追踪 desktop Sheets preset 的 Ribbon、上下文菜单注册和本地 command dispatcher，不包含浏览器或像素验收。

## 口径

- 一个“根入口”是合并菜单后位于 Ribbon Tab / Group 下的顶层菜单节点。Split、menu/select 的子项不重复计数；动态公式分类按实际生成的分类菜单计数。Context menu、浮动对象菜单和 Footer 菜单单列。
- Univer 菜单由 `packages/ui/src/services/menu/menu-manager.service.ts` 提供 Tab/Group 顺序，再由 `examples/src/sheets/mount.ts` 装配 preset，以及 sheets-ui、numfmt、formula、filter、sort、table、drawing、hyperlink、thread-comment、data-validation、conditional-formatting 和共享 find-replace 插件合并注册。
- 本地按当前工作树的 `createSheetRibbon()` schema 统计。拆分项/子菜单的 children 不计根项。公式 Ribbon 的 8 个根分类由当前 117 项 `SHEET_FUNCTION_CATALOG` 动态生成。2026-11-21 本轮归组修正及十字高亮加入后，按 START、INSERT、FORMULAS、DATA、VIEW 顺序，本地根项为 `31/8/12/19/5`，共 75 项。条件格式从 Home 移到 Data 的 RULES 组，因此跨 Tab 总数不变；Insert 根栏移除行列插删四项，它们保留在 ROW_HEADER/COL_HEADER 上下文菜单。View 中十字高亮为本地扩展入口。
- 固定源码按顶层父菜单 schema 计数为上游 `30/9/12/6/5`，本地当前根项为 `31/8/12/19/5`。
- 根入口数只比较 schema 规模，不能说明功能完整。子菜单、上下文入口、权限隐藏、动态状态、溢出菜单、Pro 规格入口和实际命令效果分开核对。

## 标签与分组

| 标签 | Univer 根入口与来源顺序 | 本地根入口与现状 | 主要差异 |
| --- | --- | --- | --- |
| START / 开始 | 30。HISTORY：全局撤销、重做、格式刷、清除格式，加 Sheets 粘贴 split。FORMAT：字体/字号/字号增减、粗斜下划线、删除线、文本色、底色、边框。LAYOUT：水平/垂直对齐、换行、缩小适应、旋转、合并。NUMBER：数字格式面板、百分比、货币、小数位增减。OTHERS：范围保护、快捷键面板、功能搜索。顺序来自 `menu-manager.service.ts`，入口来自 `ui/src/menu/schema.ts`、`sheets-ui/src/menu/schema.ts`、`sheets-numfmt-ui/src/menu/schema.ts`。 | 31。clipboard 5、font 11、layout 8、number 4、tools 3。主要入口均接入 `ExcelWorkbench.vue` 的 `executeRibbon()`。 | 基础格式覆盖大体对应，但水平对齐由一个菜单拆成 3 个常驻按钮；单独清除文本色/底色入口缺失，六色静态调色板不足以上游颜色选择器语义。清除内容和查找仍在 Start tools，查找是本地便捷入口；固定 Univer 的 Find/Replace 属 Data/OTHERS。条件格式已移至 Data/RULES。源粘贴 split 含值、格式、列宽、忽略边框；本地补了公式粘贴和忽略边框，属于扩展，不代表源菜单完全一致。 |
| INSERT / 插入 | 9。MEDIA：图片、超链接、评论面板、添加评论。CELL：复选框/下拉/日期快速验证，数据条/图标集快速条件格式。来源分别为 `sheets-drawing-ui/src/menu/schema.ts`、`sheets-hyper-link-ui/src/menu/schema.ts`、`sheets-thread-comment-ui/src/menu/schema.ts`、`sheets-data-validation-ui/src/menu/schema.ts`、`sheets-conditional-formatting-ui/src/menu/schema.ts`。 | 8。objects 4（图表、透视表、切片器、时间线）、quick-validation 4（复选框、下拉、日期、日期时间）。 | 图片、超链接、评论和两类快速条件格式缺入口。错误放置的行列插删已从 Ribbon 移除，由行/列标题上下文菜单承载。透视/切片器/时间线在本地 Insert 重复出现于 Data，属于本地扩展，不是源结构。快速验证多出日期时间项。 |
| FORMULAS / 公式 | 12。BASIC：常用、财务、逻辑、文本、日期、查找、数学、统计、工程、信息、数据库函数，共 11 个根入口；OTHERS：计算选项。来源 `sheets-formula-ui/src/menu/schema.ts`。 | 12。插入函数、8 个本地 catalog 分类、编辑公式、重新计算、计算模式。函数子项映射 `sheet.operation.insert-function`，编辑/重算/模式均由 `executeRibbon()` 处理。 | 数量相同但分类契约不一致。源有 10 类函数菜单，本地以 catalog 当前 8 个中文分类分组；本地函数目录 117 项且不能代表上游完整函数列表。源的复制公式/仅粘贴公式位于复制/特殊粘贴上下文菜单，本地“仅复制公式”只在自定义右键菜单出现。 |
| DATA / 数据 | 6。默认组顺序 FORMULAS、RULES、ORGANIZATION、ANALYSIS、OTHERS。可见入口为数据验证、条件格式管理、排序、筛选、结构化表格选择器、查找替换。来源 `sheets-data-validation-ui/src/menu/schema.ts`、`sheets-conditional-formatting-ui/src/menu/schema.ts`、`sheets-sort-ui/src/menu/schema.ts`、`sheets-filter-ui/src/menu/schema.ts`、`sheets-table-ui/src/menu/schema.ts`、`find-replace/src/menu/schema.ts`。 | 19。validation 1、rules 1、sort/filter 4、pivot/data analysis 5、outline 8。条件格式已移入 RULES；排序与筛选组成 ORGANIZATION 对应组；本地增强分析和大纲仍为额外组。 | 分组顺序已调整为验证、规则、排序筛选、分析、大纲。筛选入口现在是 split，主项切换自动筛选，子项提供清除条件、重应用和本地筛选条件面板，分别连接 `sheet.command.toggle-auto-filter`、`sheet.command.clear-filter`、`sheet.command.reapply-filter` 和现有面板操作。仍缺结构化表格选择器及 Find/Replace 的 Data/OTHERS 入口。排序本地有升/降序和自定义面板，未呈现源扩展选择区域排序命令入口。查找替换仍作为 Home 便捷工具。透视/切片器/时间线与 Insert 重复。 |
| VIEW / 视图 | 5。DISPLAY：冻结、缩放、工作表背景图；VISIBILITY：网格线、全屏。背景图父菜单内有添加/删除子命令。来源 `sheets-ui/src/menu/schema.ts`、`sheets-drawing-ui/src/menu/schema.ts`。 | 5。DISPLAY 有冻结、缩放；VISIBILITY 有网格线、全屏、十字高亮。 | 缺工作表背景图父菜单及添加/删除处理器；十字高亮是本地扩展，不属于固定版根入口。 |

### 本地 handler 核对

`sheet-ribbon.ts` 的大部分可见命令在 `ExcelWorkbench.vue:1413` 起的 `executeRibbon()` 中有同 ID handler。格式、行列增删、合并、筛选/排序面板、验证、条件格式、范围保护、透视/切片器/时间线/迷你图、大纲、公式、冻结、缩放、网格线、全屏都有实际处理函数或面板入口，因此本表把它们记为“已接 handler”，不据此宣称功能验收完成。

下列 schema 入口没有对应的上游概念，或放错组，属于偏差而非 no-op：独立三水平对齐按钮、Home 查找、Insert/Data 重复的透视类入口、日期时间快速验证、大纲扩展组。行列插删已从 Insert Ribbon 移除，按固定版上下文菜单分配到 ROW_HEADER/COL_HEADER。`executeRibbon()` 未发现图像插入、超链接、评论、工作表背景图、表格选择器、十字高亮和打印命令 handler。

`sheet.command.clear-content` 只对应本地清内容流程，不等同于源上下文菜单中的“仅清内容 / 仅清格式 / 全部清除”三种操作。源颜色控件含单独 reset 文本色/底色，本地只暴露整体清格式和固定色板。源数据筛选入口含切换、清条件和重算，本地仅有 filter panel。入口名称相近不能替代这些具体操作。

| 入口组 | Univer 菜单注册/调用 ID | 本地 ID / 处理器核对 |
| --- | --- | --- |
| START | `UndoCommand`、`RedoCommand`、`ActivateFormatPainterOperation`、`ClearFormattingCommand`、`PASTE_RIBBON_MENU_ID` 下的 `SheetPasteCommand`/Value/Format/ColWidth/BesidesBorder；字体/样式注册为 `SetRangeFont*Command`、`SetRangeBoldCommand`、`SetRangeTextColorCommand`、`SetBackgroundColorCommand`、`SetBorderBasicCommand`、`SetHorizontalTextAlignCommand`、`SetVerticalTextAlignCommand`、`AddWorksheetMergeCommand`/All/Vertical/Horizontal、`RemoveWorksheetMergeCommand`、`AddRangeProtectionFromToolbarCommand`；数字格式由 `OpenNumfmtPanelOperator`、`SetPercentCommand`、`SetCurrencyCommand`、`AddDecimalCommand`、`SubtractDecimalCommand` 注入。 | `sheet.command.undo/redo`、`sheet.operation.format-painter`、`sheet.command.clear-format`、`sheet.command.paste*`、`sheet.command.set-font-*`、`toggle-*`、`set-color/background/borders/align/vertical-align`、`merge/unmerge`、`open-range-protection` 与 `set-number-format`/decimal handlers 均存在。没有上游 color reset 对应项；`sheet.command.clear-content` 不对应 Ribbon 清格式入口。 |
| INSERT | `SHEETS_IMAGE_MENU_ID` → `InsertFloatImageCommand`/`InsertCellImageCommand`；`InsertHyperLinkToolbarOperation`；`ToggleSheetCommentPanelOperation`/`ShowAddSheetCommentModalOperation`；`InsertQuickSheetDataValidationCommand`/`QUICK_DROPDOWN_MENU_ID`/`QUICK_DATE_MENU_ID`；`AddDataBarConditionalRuleCommand`/`AddIconSetConditionalRuleCommand`。 | 本地 `sheet.command.insert-row/column`、`delete-row/column`、`operation.insert-chart`、`open-pivot`、`open-slicer`、`open-timeline`、`quick-data-validation` 可执行；以上图像/链接/评论/条件格式源 IDs 均没有本地 schema 或 handler。 |
| FORMULAS | `InsertFunctionOperation.id` 的 common/financial/logical/text/date/lookup/math/statistical/engineering/information/database 分类、`MoreFunctionsOperation.id`、`CalculationOptionsOperation.id`；`SheetCopyFormulaOnlyCommand` 与 `SheetOnlyPasteFormulaCommand` 注册在 copy/paste special 上下文菜单。 | 本地 `sheet.operation.open-function`/`insert-function`、`start-formula`、`sheet.command.recalculate`、`set-calculation-mode` 均有处理器。仅复制公式有自定义 context ID `sheet.command.copy-formulas`，没有固定源的粘贴公式上下文 menu schema ID 对照。 |
| DATA | `OpenFindDialogOperation.id`；`SmartToggleSheetsFilterCommand.id`，children `ClearSheetsFilterCriteriaCommand.id`/`ReCalcSheetsFilterCommand.id`；`SortRangeAsc/AscExt/Desc/DescExt/CustomCommand.id`；`OpenTableSelectorOperation.id`；`OpenConditionalFormattingOperator.id`；`OpenValidationPanelOperation.id`/`AddSheetDataValidationAndOpenCommand.id`。 | `sheet.operation.find`、`sheet.operation.open-data-panel`、`sort-ascending/descending`、`open-data-validation`、`open-conditional-formatting` 有处理器。API 新增 `sheet.command.toggle-auto-filter` 和 `sheet.command.reapply-filter`，并保留 `sheet.command.clear-filter`；尚未接入 Ribbon dispatcher。条件格式已归 Data/RULES。`find-replace` 以本地 find dialog 承载但位于 Home tools；source ext-sort 与 table selector 仍无本地等价入口/handler。 |
| VIEW | `SHEET_FROZEN_TOOLBAR_MENU_ID`、`ZOOM_RATIO_MENU_ID`、`ToggleGridlinesCommand.id`、`ToggleFullscreenOperation.id`、`WORKSHEET_BACKGROUND_IMAGE_MENU_ID` 下的 Add/Delete background commands。 | 冻结、zoom、网格线、全屏都有本地处理器；背景图菜单和 Add/Delete handlers 缺失。 |

以上本地 IDs 由组件级 `executeRibbon()` 分发，并非都注册到共享应用 command bus。对当前可见 Ribbon 项未发现空回调/占位消息式 no-op；明确缺少的是表中未注册的源入口/处理器以及不等价的替代行为。命令有 handler 只证明入口已连接，不代表 mutation、撤销、权限、持久化、重开和公开规格完整。

## 上下文与工作表栏

固定源码在 `sheets-ui/src/menu/schema.ts` 注册独立 MAIN_AREA、COL_HEADER、ROW_HEADER、FOOTER_TABS 菜单，并由功能插件向上下文组注入排序、评论、超链接、表格等入口。

MAIN_AREA 源核心 IDs 包括 `SheetCopyCommand`/`SheetCutCommand`/`SheetPasteCommand`、`COPY_SPECIAL_MENU_ID`/`PASTE_SPECIAL_MENU_ID`、`CLEAR_SELECTION_MENU_ID` 下 content/format/all、`TEXT_TO_NUMBER_CONTEXT_MENU_ID`、`CELL_INSERT_MENU_ID`、`DELETE_RANGE_MENU_ID`、`SHEET_FROZEN_MENU_ID`、`SHEET_PERMISSION_CONTEXT_MENU_ID`。行列头另注册 multi-insert、hide/show、width/height、auto-width/auto-height 和各自 freeze/permission IDs；表签注册 `RemoveSheetConfirmCommand`、`CopySheetCommand`、`RenameSheetOperation`、`SetTabColorCommand`、`SetWorksheetHideCommand`、`ShowMenuListCommand` 及 sheet permission IDs。功能插件还在 MAIN_AREA 加 sort/filter、hyperlink、thread comments 和 table contexts。

| 上下文位置 | Univer 可见操作类别 | 本地实现与差异 |
| --- | --- | --- |
| MAIN_AREA | 快速复制/剪切/粘贴；仅复制/粘贴特殊类型；清内容/格式/全部；文本转数字；插入/删除单元格并可移动邻格；冻结；范围权限；插件注入排序、评论、超链接等。 | `sheet-context-menu.ts` 有剪切、复制、复制公式、粘贴类型、清内容、插删整行列、冻结、范围保护。缺清格式/全部、文本转数字、单元格位移操作、插件项。 |
| COL_HEADER | 剪贴板/特殊粘贴；插入多列、隐藏/显示列、删除列、列宽/自动列宽、列冻结、范围权限。 | 已接入左/右插入、删除命中列、隐藏/显示全部隐藏列、按单元格显示文本自适应列宽、冻结命中列和保护命中列。命令参数显式携带 sheetId/index；仍缺剪贴板/特殊粘贴、批量插入与手动设置列宽。 |
| ROW_HEADER | 剪贴板/特殊粘贴；插入多行、隐藏/显示行、删除行、行高/适应内容、行冻结、范围权限。 | 已接入上/下插入、删除命中行、隐藏/显示全部隐藏行、基于文字换行与列宽自适应行高、冻结命中行和保护命中行。命令参数显式携带 sheetId/index；仍缺剪贴板/特殊粘贴、批量插入与手动设置行高。 |
| FOOTER_TABS | 删除/复制/重命名/标签色/隐藏/显示列表/工作表权限等。 | 已加入复制工作表、标签颜色选择和模型持久字段。复制会复制单元格、格式、合并/尺寸/验证等工作表快照，不会复制 chart/pivot/table 关联对象；有明确提示且不视为完整复制功能。仍缺范围/工作表权限菜单入口和 veryHidden 语义；隐藏工作表列表可恢复普通隐藏表。 |
| Drawing / 浮动对象 | 图片保存、绘图评论和对象上下文工具。商业图表/形状上下文需按公开规格核验。 | 当前 Excel Ribbon 无 contextual tab schema；画布仅有图表 Overlay。没有图片/图形对象菜单或工具栏入口。 |

行列结构操作曾错误依赖旧的单元格选区。现由 `sheet-context-menu.ts` 将右键命中的 `sheetId/index` 传入行列插删、隐藏和自适应命令；冻结和范围保护也带入命中轴/范围。`ExcelWorkbench.vue` 据参数修改对应 sheet/axis，并将变更放入工作簿历史。已补充的隐藏/显示以模型 `hiddenRows` / `hiddenColumns` 为持久数据，显示操作当前一次清除该轴全部手动隐藏项。自动行高/列宽采用浏览器 Canvas 字体度量和本地模型尺寸上限；它不是上游 Worker/WASM 大表格布局实现，也未做字体、溢出、绘图对象或跨样本 Office 互操作验证。行列头右键仍不会像上游那样把命中区域设成当前选区；命令目标独立传递，选区只在命令成功后移动到结构编辑后的单元格。

## 溢出与布局入口

生产工作台在 `ExcelWorkbench.vue` 使用 Grid Ribbon。`RibbonGridToolbar.vue` 按完整 group 测量宽度，窄屏收纳整组到 More，内部保留 item 顺序和命令；此点与源 `RibbonGrid.tsx` 的组级溢出语义一致。当前审计没有把溢出实现判为缺项。源另支持 classic 变体并按项溢出；比较应与当前生产传入的 grid 变体一致，不能拿 classic 控件数量或布局行列与生产实现混算。

当前本地 Grid 的分组顺序、`gridLayout` 与语义组映射仍按上方逐 Tab 列表有偏差。完整宽度下出现于 schema 的额外根入口仍然是实际视觉布局的一部分，溢出只影响可见位置，不会补足缺失入口或修正错组。

## Pro 能力边界与状态

固定 Univer 检出中未发现 `@univerjs-pro` 商业插件的 Ribbon/schema 源码。README 将商业插件标为单独分发/授权，因此下表“Pro 目标”使用项目公开规格与现有 feature parity ID，不冒称来自 Pro 源码。完整目标仍在 `UNIVER_FEATURE_PARITY.md`，工具栏审计只判断入口存在与否。

| 目标 | 当前入口/handler 状态 | 入口缺口或已知范围 |
| --- | --- | --- |
| SHEET-10 高级公式 | Formula catalog 入口和插入 handler 存在。 | 117 项本地 catalog、有限求值支持；完整函数、数组传播等未实现/未验收。 |
| SHEET-21 形状及绘图 | 当前 Insert 有图表入口和 ChartOverlay。 | 缺形状/连接线/图片插入入口及对象 contextual toolbar。 |
| SHEET-22/23/24 图表、透视表、透视图 | 对应 Ribbon 入口与面板 handler 存在。 | 入口连接不等于范围绑定、编辑能力、刷新、持久化和 Office 往返完整；均未验收。 |
| SHEET-25 切片器和时间线 | Insert/Data 有入口，映射到 SlicerPanel。 | 入口重复；目标表格、交叉筛选、完整布局和保存仍未验收。 |
| SHEET-26 迷你图 | Data 入口映射到 SparklinePanel。 | 完整图型/组行为/引用变换/导出未验收。 |
| SHEET-27 大纲分组 | Data 中有行列组合、折叠和汇总位置 handler。 | 与源 Data 分组位置不同；嵌套/筛选/手动隐藏组合未验收。 |
| SHEET-28 打印与 PDF | 无 Ribbon 入口及本地 print command handler。 | 完整缺失。 |

图片、超链接、批注/线程评论、结构化表格、工作表背景、十字高亮属于当前完整范围，分别对应 SHEET-16/17/18/19/20；这些在固定 OSS 源码有明确入口或操作，本地 Ribbon/handler 中缺失。该缺口不是 Pro 源码不可得造成。

## 源码复核路径

- 固定源共享 Ribbon 结构和位置：`packages/ui/src/services/menu/menu-manager.service.ts`、`packages/ui/src/services/menu/types.ts`、`packages/ui/src/menu/schema.ts`。
- Sheets 根菜单：`packages/sheets-ui/src/menu/schema.ts`、`packages/sheets-numfmt-ui/src/menu/schema.ts`、`packages/sheets-formula-ui/src/menu/schema.ts`。
- Sheets 插件扩展：`packages/sheets-sort-ui/src/menu/schema.ts`、`packages/sheets-filter-ui/src/menu/schema.ts`、`packages/sheets-table-ui/src/menu/schema.ts`、`packages/sheets-drawing-ui/src/menu/schema.ts`、`packages/sheets-hyper-link-ui/src/menu/schema.ts`、`packages/sheets-thread-comment-ui/src/menu/schema.ts`、`packages/sheets-data-validation-ui/src/menu/schema.ts`、`packages/sheets-conditional-formatting-ui/src/menu/schema.ts`、`packages/find-replace/src/menu/schema.ts`。
- 固定 desktop 装配：`examples/src/sheets/mount.ts`。商业插件来源边界：固定版本 `README.md` 与 `UNIVER_FEATURE_PARITY.md` Pro 条目。
- 本地 schema / handler / context / overflow：`packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts`、`packages/office-studio/src/components/ExcelWorkbench.vue`、`packages/office-studio/src/components/spreadsheet/sheet-context-menu.ts`、`packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue`、`packages/office-studio/src/components/spreadsheet/SheetFooter.vue`、`packages/office-studio/src/platform/ribbon/RibbonGridToolbar.vue`、`packages/office-studio/src/platform/ribbon/ribbon-schema.ts`。

本清单不改变任何 feature acceptance 状态；未做浏览器、样本文件、权限、历史/重开或 Pro 互操作验收。所有功能 ID 继续按总清单保持未验收。
