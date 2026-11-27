# SHEET-11 / SHEET-12 顶部排序与筛选核对

状态：2026-11-27 静态核对固定版 Data 入口，修正独立筛选配置 mutation 的派生行重置；续批将重新应用筛选改为不写内容历史，并接入弹层主题 token。本次补齐[折叠大纲行的排序位置](#sort-collapsed-outline-rows-review)。新草稿未运行，功能未验收。真实功能 ID 为 SHEET-11（筛选）和 SHEET-12（排序）；SHEET-07 属于冻结、网格线、视口和缩放。固定来源为 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

## 固定入口、顺序与状态

来源为 [排序 schema](../../../univer/packages/sheets-sort-ui/src/menu/schema.ts)、[排序 factories](../../../univer/packages/sheets-sort-ui/src/menu/sheets-sort.menu.ts)、[筛选 schema](../../../univer/packages/sheets-filter-ui/src/menu/schema.ts)、[筛选 factories](../../../univer/packages/sheets-filter-ui/src/menu/sheets-filter.menu.ts) 和各插件的 [排序中文文案](../../../univer/packages/sheets-sort-ui/src/locale/zh-CN.ts)、[筛选中文文案](../../../univer/packages/sheets-filter-ui/src/locale/zh-CN.ts)。

这两根同属 `RibbonDataGroup.ORGANIZATION`：排序 order 0，grid r1c1；筛选 order 1，grid r2c1。它们在同一列上下排列；classic 顺序也由相同 order 决定，不能描述成两个横向大按钮。

| 位置 | 标签 | 固定 ID | 状态或派发 |
| --- | --- | --- | --- |
| 常驻根 r1c1 | 排序 | `sheet.menu.sheets-sort` | SUBITEMS；工作簿、工作表和范围编辑/排序权限控制禁用 |
| 排序子项 0 | 当前区域升序 | `sheet.command.sort-range-asc` | `triggerSortDirectly(true, false)`；命中 exclusive range interest 时隐藏 |
| 排序子项 1 | 拓展区域升序 | `sheet.command.sort-range-asc-ext` | `triggerSortDirectly(true, true)` |
| 排序子项 2 | 当前区域降序 | `sheet.command.sort-range-desc` | `triggerSortDirectly(false, false)` |
| 排序子项 3 | 拓展区域降序 | `sheet.command.sort-range-desc-ext` | `triggerSortDirectly(false, true)` |
| 排序子项 4 | 自定义排序 | `sheet.command.sort-range-custom` | `triggerSortCustomize()`；先选择保持/拓展范围，再打开配置 |
| 常驻根 r2c1 | 筛选 | `sheet.command.smart-toggle-filter` | BUTTON_SELECTOR；已有 filter model 时激活；权限和 exclusive range 控制禁用 |
| 筛选子项 0 | 清除筛选条件 | `sheet.command.clear-filter-criteria` | 没有 filter model 或没有条件时禁用；图表工作表禁用 |
| 筛选子项 1 | 重新计算 | `sheet.command.re-calc-filter` | 同样要求已有 filter model 和条件；图表工作表禁用 |

[menu-util.ts](../../../univer/packages/sheets-ui/src/menu/menu-util.ts) 的 `getCurrentRangeDisable$` 还处理无工作簿/工作表、单元格或公式栏编辑、图表工作表、绘图焦点及权限更新。`getCurrentExclusiveRangeInterest$` 检查选区命中的 exclusive range interest group，不等同于“只允许一个矩形选区”。[filter-model.ts](../../../univer/packages/sheets-filter/src/models/filter-model.ts) 的 `hasCriteria$` 来自 `_filterColumnByIndex.size > 0`，不能仅以已创建筛选区域推断。

本次可复查计数：固定版常驻根 2、子项 7；MAIN_AREA 的排序上下文另有 1 根和 5 个 `-ctx` 命令，不计入常驻根。窄屏溢出复用已有入口，不增加独立功能数。此计数仅覆盖排序/筛选这两个插件的上述 schema，不代表完整 Data 或 Pro 入口总数。

## 固定命令、模型与历史

[排序 UI 命令](../../../univer/packages/sheets-sort-ui/src/commands/commands/sheets-sort.command.ts) 转入 [SheetsSortUIService](../../../univer/packages/sheets-sort-ui/src/services/sheets-sort-ui.service.ts)。`_detectSortLocation` 对显式 extend=true 调用 [expandToContinuousRange](../../../univer/packages/sheets/src/basics/expand-range.ts)，按四个方向扩展连续区域并更新选区；false 保持原范围；自定义排序先显示保持/拓展/取消选择。排序键始终使用 `selection.primary.actualColumn`，扩展范围后也保留原活动列。

[SheetsSortService](../../../univer/packages/sheets-sort/src/services/sheets-sort.service.ts) 检查单行、空范围、数组公式和合并区域，再派发 [SortRangeCommand](../../../univer/packages/sheets-sort/src/commands/commands/sheets-sort.command.ts)。它按多关键字比较实际单元格记录，跳过筛选和手动隐藏的行，生成 reorder 映射。[ReorderRangeCommand](../../../univer/packages/sheets/src/commands/commands/reorder-range.command.ts) 包含内容、相关拦截器和自动行高的 undo / redo。固定版允许完整等尺寸合并块等特定场景，本地目前拒绝与任意合并区域相交的排序。

[筛选命令](../../../univer/packages/sheets-filter/src/commands/commands/sheets-filter.command.ts) 的 smart-toggle 在已有模型时删除筛选；新建时单格四向扩展、单行只向下扩展，多行使用原范围。清除条件保留筛选范围，并记录撤销；重新计算调用 mutation，固定源码明确不为此创建 undo / redo。本地此前经快照历史提交，续批已改用既有工作簿 operation 发布派生结果，不改变原内容历史或 modified 标记；实际运行验收后置。

[SheetsFilterService](../../../univer/packages/sheets-filter/src/services/sheet-filter.service.ts) 注册 `SHEET_FILTER_PLUGIN` 资源，按工作簿将各工作表 `FilterModel.serialize()` 保存为 JSON，onLoad 调用 `FilterModel.deserialize` 重建模型并更新活动状态；卸载时释放模型。排序结果通过既有 reorder 内容写入工作表单元格，不依赖额外的排序快照资源。

## 本地真实链路与未接通部分

[sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 已有同位置排序菜单和筛选 split。排序五个子项使用升序/降序、`{ extendSelection: true }` 与自定义排序；筛选有清除、重新应用和额外的“筛选条件”面板入口。因此本地常驻根仍为 2，子项为 8，其中 1 个是本地额外入口。未新增根或 UI。

本轮开始时的 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 存在三项入口差异：排序键采用范围最左列，`sortSelection` 收到 `_extendSelection` 却未使用，筛选 clear / reapply 的禁用判断没有检查条件。主任务已将排序键改为 `parseCellKey(selectedKey.value)!.col`，与 `syncApiSelection` 的 primary 使用同一真实活动单元格来源；已接入 `filterHasCriteria`，按真实条件列数量控制清除和重新应用的状态。

[sheet-number-format.browser.test.mjs](../../packages/office-studio/tests/sheet-number-format.browser.test.mjs) 已新增 classic / grid 草稿：在跨列选区内使用非最左活动列排序，检查整行关联、撤销/重做和原生重开；另检查有/无条件及清除后的筛选菜单状态。草稿未运行。本轮入口接线仍待确认的三项为 `_extendSelection` 的连续区域计算、单格/单行 smart filter、自定义排序的保持/扩展/取消范围选择；活动列排序和条件状态已落盘，但未验收。

内容路径经 [FWorksheet](../../packages/office-studio/src/api/facade.ts) 和 [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的注册命令：`sort-range`、`set-auto-filter`、`toggle-auto-filter`、`set-filter-criteria`、`clear-filter`、`reapply-filter`、`configure-filter`。[sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts) 保留完整单元格记录、稳定排序、非参与行位置和方向无关的空白末尾；公式引用在清空/写回前统一转换。纯数值排序符合既有派发阈值时使用 [NumericIndexService](../../packages/office-studio/src/features/sheet-compute/numeric-index-service.ts)，其余使用既有 JS 比较路径；本轮没有性能实测或优化结论。

[SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue) 使用 antdv-next Form / Input / Select / Checkbox / Switch / Button / Modal，支持现有值/条件筛选和多关键字排序。其模型提交复用 [sheet-data-panel-state.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-data-panel-state.ts)，一次 `commitWorkbook` 持有内容历史。公开 `commitWorkbook` 在发布前统一刷新工作表 autoFilter；`getFilteredRows` / `getAutoFilter` 查询也会重算，不能以这些查询结果推断独立 mutation 已保持派生数据一致。

原生 WorkbookModel 保存包含筛选范围、条件、filteredRows、手动隐藏及排序后的单元格记录。[excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 区分筛选与手动隐藏；[workbookAutoFilter.ts](../../packages/office-studio/src/parser/workbookAutoFilter.ts)、[xlsx_filter.rs](../../crates/office-writer/src/xlsx_filter.rs)、[sheet_filter_import.rs](../../crates/office-parser/src/sheet_filter_import.rs) 已有本项目 DTO 和 Office 筛选读写路径。原生重开与已有读写模块不能替代第三方 XLSX 样本和实际 Office 互操作验收。

## 本轮小修与草稿

`configureFilterFromPanel` 先在隔离 plan 上调用 `setAutoFilter`，同范围时保留并重算原条件；随后直接 `columns = {}`。传入空条件时循环不执行，旧 filteredRows 留在该独立 mutation 的输出中。现在复用 `clearFilter(plan)`，同时清空条件和派生隐藏行，再应用输入中的新条件。生产改动仅为 import 和一行调用，接口及已公开的内容命令没有变化。

公开 `commitWorkbook` 此前已在发布前刷新 filteredRows，因此本轮修正针对独立 mutation 输出的一致性，不宣称已观察到公开清除流程留下隐藏行。

[sheet-data-panel.contract.test.mjs](../../packages/office-studio/tests/sheet-data-panel.contract.test.mjs) 新增真实隔离 mutation 草稿：先通过公开命令建立有效筛选，再在原生保存副本上调用真实配置函数，直接断言空条件、空 filteredRows、范围不变、手动隐藏不变、源单元格不变和原工作簿不变。另在既有公开清除草稿中，增加 getter 重算前的实际模型 filteredRows 断言。全部草稿未运行。

本轮仅静态复读差异、固定来源、文档文件链接及限定 `git diff --check`；未运行测试、类型、构建、Rust/WASM、dev、浏览器或性能检查。

### 重新应用筛选、主题与现有草稿接线续核

- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的 `sheet.command.reapply-filter` 复用 `applyWorkbookOperation`，仍由同一 `refreshAutoFilter` 计算派生行；已有权限及模型发布边界继续执行，保留原公开命令 ID 与 boolean 返回值。移除本次重算产生的内容撤销记录，保留既有 undo/redo 队列和 modified 标记。
- [SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue) 的值列表边框和错误文本改为当前 ConfigProvider 的 `colorBorderSecondary` / `colorError`，其余 Form、Checkbox、Select、Modal 和计算路径不变。
- [sheet-data-panel.contract.test.mjs](../../packages/office-studio/tests/sheet-data-panel.contract.test.mjs) 的既有真实 Studio 草稿改为验证重新应用后的派生行、undo/redo 队列及 modified 不变，并继续执行清除条件的 redo/undo。
- [sheet-data-toolbar.spec.ts](../../tests/browser/sheet-data-toolbar.spec.ts) 的三个既有浏览器草稿现在经过实际排序菜单或筛选 split 的“筛选条件”子项打开设置；首次筛选先执行真实启用命令，再进入条件面板。[sheet-toolbar-format.contract.test.mjs](../../packages/office-studio/tests/sheet-toolbar-format.contract.test.mjs) 的既有声明同步当前筛选主动作、条件子动作及 Data 下的 Pro 分析位置。

上述都是已落盘小修，未执行草稿、类型、构建或浏览器。范围计算方案仍待确认；实际主题、焦点、历史/重开和 Office 验收不由静态接线代替。模块仍为既有应用命令与 Vue 视图；未改变领域计算引擎、Worker/WASM 协议或内存所有权。

## 连续区域复用与待确认方案

固定 [expand-range.ts](../../../univer/packages/sheets/src/basics/expand-range.ts) 的实际判定如下。公式是否算内容由缓存值或正文决定，不能直接按 `f` 是否存在判断。

| 场景 | 固定源码语义 | 本地方案及差异 |
| --- | --- | --- |
| 空值、样式、公式 | `cellHasValue`（41–47）仅检查 `v` 非 undefined/null/空串或 `p !== undefined`；0、false、空格字符串算内容，样式和 `f` 本身不参与 | 读取模型 `value`；不以 style、display 或 formula 存在性扩展。公式缓存为空时，只有保留的富文本正文可使其算内容 |
| 富文本存在性 | `p` 存在即算内容，包括空正文和 `p: null` | 非空 richText runs、已保存 publicBody 可识别；导入时 absent 与空 runs 已合并，公开 p:null 会清除正文，不能声称完全等价 |
| 合并内部格 | `getMatrixWithSpanInfo` / `hasValueFromMatrixWithSpanInfo` 只借合并主格内容，内部残留值不参与 | 直接使用领域 `mergedRanges` 定位主格；复用矩形命中与合并闭包，不使用 A1 view 的合并适配 |
| 相邻边和空合并 | Left 158 / Right 212 / Up 266 / Down 319 检查相邻整条边；一格有内容即可扩一格并收集合并范围。全边为空时，仅保留与当前范围相交的合并；最后 union | 新函数负责边探测与合并纳入条件，现有合并 helper 只处理已命中集合的矩形闭包，不能把相邻空合并全部纳入 |
| 斜角相邻 | `getExpandedRangeByFastPath`（376）先检查四个角；对应两个方向都启用且角格有内容时，同时扩两个方向，合并角格取锚范围 | 四条边探测前保留角探测；只向下时不检查左右下角 |
| 空种子 | `expandToContinuousRange`（483）没有“种子必须有值”的提前返回；孤立空格不扩，旁边有内容仍可扩 | 保留该实际控制流，不按注释添加空种子提前返回 |
| 方向及稳定性 | 先角探测，再按 up/down/left/right 逐次扩展到稳定，最后 union 收集的合并 | 显式扩展排序、单格筛选四向；单行筛选只向下；多行筛选和不扩展排序保持原范围 |

现有复用点为 [sheet-selection.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-selection.ts) 的 `SheetSelectionState.ranges / activeRangeIndex / primary` 和 `normalizeSheetSelection`，以及 [sheet-range-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-command.ts) 的 `sheetRangeContainsCell / sheetRangesIntersect / sheetRangesEqual`。[sheet-clipboard-data.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-data.ts) 的 `expandSheetClipboardRange` 已实现相交合并的反复 union，可复用其合并闭包；它不负责连续数据区域探测。模型范围使用 startCol/endCol，UI 范围使用 startColumn/endColumn，沿用既有选区转换并保留 primary，不重新建立选区模型。

[sheet-selection-navigation.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-selection-navigation.ts) 的 `nextSheetSelectionCell` 是可见格导航；Facade 的 `FSelection.getNextDataRange(direction)` 实际调用该导航。按实际调用链检索，尚未找到等价于固定 `expandToContinuousRange` 的现有 helper。

富文本差异有明确边界来源：[domain/model.ts](../../packages/office-studio/src/domain/model.ts) 的 richText 可选，但 [workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts) 无条件归一化为数组；Rust [model.rs](../../crates/office-parser/src/model.rs) 使用 Vec，[parser.rs](../../crates/office-parser/src/parser.rs) 使用 unwrap_or_default，均丢失缺失与显式空 runs 的区别。[sheet-image-public.ts](../../packages/office-studio/src/features/sheet-images/sheet-image-public.ts) 保存对象正文到 metadata.publicBody，包括空 dataStream；p:null 清除该字段。[model-utils.ts](../../packages/office-studio/src/api/model-utils.ts) 的 applyCellValue 先清除旧 publicBody，再应用 p。因此不能用 `richText !== undefined` 把导入的纯样式空格判为有内容。完整存在性对齐还需在这六个真实文件的 DTO、归一化、公开写入与保存边界保留类型化存在性字段，预计另加 60–100 行及边界契约草稿；现存数据无法从空数组反推出该字段，差异仍待验收。

以下为待确认的连续区域接线方案，未实施。路径相对 `packages/office-studio/src/`；生产代码约 215–305 行，富文本存在性保真及草稿另计。

| 文件 | 行数 | 具体改动 |
| --- | --- | --- |
| 新增 `features/sheet-sort-filter/sheet-data-range.ts` | 135–180 | 同一同步纯领域函数供排序与筛选复用，输入 sheet、CellRange 和方向；按缓存值/正文、主格、斜角和整条边探测稀疏内容，复用命中、范围比较及已命中合并集合的闭包，不构造整表稠密矩阵 |
| `components/ExcelWorkbench.vue` | 40–60 | `_extendSelection` 接入共享计算，扩展后同步选区但保留真实活动列；单格/单行筛选按对应方向计算并更新禁用判断；自定义排序保留来源选区和 primary 传入现有面板 |
| `features/sheet-sort-filter/sheet-data-panel-state.ts` | 15–25 | 复用共享计算准备自定义排序的保持/扩展目标，返回既有 CellRange；沿用单次入口校验和既有目标/内容版本协议，不修改排序引擎 |
| `components/spreadsheet/SheetDataPanel.vue` | 25–40 | 用现有 antdv-next Modal / Form 控件表达保持/扩展/取消；确认目标后再进入既有配置，取消不写内容，关闭后按既有工作台焦点链恢复画布 |

后置契约草稿直接覆盖共享真实函数的 0/false/空串、空缓存公式、空种子、合并内部残留值/空合并、仅斜角内容和向下限定；浏览器草稿覆盖三个入口的实际范围、原活动列、取消零写入、历史/重开和主题/焦点。方案未落盘，不能以现有草稿代替这些场景的验证。

上述范围计算属于 sheet-sort-filter 领域层，UI 仅使用结果；大样本计算域、取消、跨 Worker/WASM 传输与内存预算仍需同数据集基准决定。不能以未运行的稀疏方案声称满足性能标准。

列头筛选下拉、颜色与完整日期条件、全部比较/locale、合并和数组公式、关联资源/结构变换、完整权限与协作、主题/焦点/IME、第三方文件和大样本均保持未验收。SHEET-11 / SHEET-12 仍为部分实现。

### 搜索结果的全选与提交

固定 [SheetsFilterByValuesPanel.tsx](../../../univer/packages/sheets-filter-ui/src/views/components/SheetsFilterByValuesPanel.tsx) 显示当前搜索项的勾选状态和计数；[ByValuesModel](../../../univer/packages/sheets-filter-ui/src/services/sheets-filter-panel.service.ts) 的 `onCheckAllToggled` 和 `apply` 都读取 `_filterItems`，即搜索结果，而非全部 `_rawFilterItems`。[既有固定契约](../../../univer/packages/sheets-filter-ui/src/services/__tests__/sheets-filter-panel.service.spec.ts) 的 `should toggle all applied to searched results` 直接核对这个范围。

本地 [SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue) 原本仅用 `matched` 限制显示，点击全选和提交仍读取全部 `values`。搜索“东”后全选提交会把搜索外的“西”也纳入。本批只改两行、三处数据来源：全选使用 `matched`，计数显示匹配项中的已选数量/匹配项总数，值条件仅提交匹配项中已选的原始类型值。没有搜索时 `matched` 等于原候选，条件筛选和现有命令、版本、权限、撤销/保存接口继续沿用。搜索、勾选与全选仍是面板草稿，只有应用才提交模型。

[现有浏览器草稿](../../tests/browser/sheet-data-toolbar.spec.ts) 的搜索后全选、提交范围及提交前零内容/历史变更断言由主线追加，未运行。面板继续使用已有 antdv-next Input、Button、Checkbox 和 Form；本批没有新增控件、业务引擎或 Worker/WASM 协议。SFC、脚本和样式语法、证据链接与限定空白检查已通过，真实主题、焦点、输入法及大样本性能未验收。

保留差异：本地仍是同步单串包含搜索、扁平候选和每页 20 项；固定版使用分词树搜索、500ms 节流和树叶统计。固定版无勾选时禁止应用，本地已有空选择隐藏全部数据行的行为继续保留。清除搜索后的勾选重建，以及 `selectionMode:'exclude' / recordSelection` 在面板中的回读与保存保真尚未补齐，本次不将这些行为标为完成。

<a id="sort-collapsed-outline-rows-review"></a>

## 2026-11-27 折叠大纲行的排序位置续修

本次按 DATA / Organization 从排序、筛选到组合及本地分组显示/设置逐项复读。入口、常驻/扩展/上下文分类和命令沿前文及 [SHEET-27 Organization 清点](SHEET-27-outline.md#2026-11-27-organization-续核)；排序/筛选仍为两根、八个本地子项，本次新增根及子项均 0。`_extendSelection`、单格/单行 smart filter、自定义保持/扩展/取消仍归上述约 215–305 行的待确认方案，没有实施或缩减这些目标。

[固定 sort-range 命令第 73–80 行](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-sort/src/commands/commands/sheets-sort.command.ts#L73) 分别跳过 `getRowFiltered` 与 `getRowRawVisible=false` 的行。[固定 Worksheet 第 923 行](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/sheets/worksheet.ts#L923) 将 raw visible 委托 RowManager。已读取的 [Pro v1.6.0 outline core ES](../../.research.local/univer-1.6.0/packages/sheets-outline/lib/es/index.js) 中 `Je/Ye/Xe` 在行组折叠/展开时生成 SetRowHiddenMutation / SetRowVisibleMutation；因此固定版折叠行进入上述排序排除链，不能只按普通手动隐藏理解该状态。

本地大纲保留独立 `outline` 与隐藏原因，不把折叠写入 `hiddenRows`。原 [sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts) 的 `getSortableRowOffsets` 只查筛选、手动数组/spans 和零高度，遗漏折叠行。这会让隐藏的组内记录参加数值/文本排序，并在可见行位置出现。

现复用 [getSheetOutlineHiddenSpans](../../packages/office-studio/src/features/sheet-outline/sheet-outline.ts)，在唯一行索引入口排除该轴的折叠范围；extent 使用领域 rowCount 或既有 Office 缺省上限。数值输入的 destinationOffsets 和一般比较的完整 permutation 共用该结果，非参与行仍保留原位置。展开行组后重新参与；列组折叠不屏蔽行。没有新建大纲模型、隐藏数组或第二套排序算法。

真实写入沿 FWorksheet.sortRange → sheet.command.sort-range → 原范围权限/revision 校验 → sortRange → commitWorkbook。组状态、手动隐藏、源单元格完整记录和原内容历史/内部保存链保留；本批未处理其他关联资源变换的既有缺口。生产只改 sort-filter-engine.ts，增量 **+7/-1**；现有行循环展开以保持条件可读，无新增模型、公共 API、命令 ID、Worker 控制 DTO 或 WASM 内存协议。

[原排序契约草稿](../../packages/office-studio/tests/sheet-sort-filter-engine.contract.test.mjs) 新增 1 声明，直接使用真实构造器与实现，覆盖折叠行组/列组、数值准备与一般比较、完整索引、同时存在的筛选/手动/零高度隐藏、组内公式和 metadata、展开后参与。[原数字格式浏览器草稿](../../packages/office-studio/tests/sheet-number-format.browser.test.mjs) 沿已存在的排序 fixture 新增 1 个 test.each 声明、classic/grid 两个计划用例：通过真实排序菜单执行，检查可见整行重排、折叠/手动隐藏行记录与组状态不变，一次内容历史、undo/redo 和内部模型保存重开。没有新增测试文件、fixture 或产品测试接口。

本批四文件为排序引擎、上述两份原草稿和本页。当前排序契约 **10 声明/10 计划**；数字格式浏览器 **32 声明/39 计划**（25 个直接声明、6 个两例 test.each、1 个两例外层 for），本批合计新增 2 声明/3 计划，全部运行 0。数值契约检查真实索引准备和展开，没有执行 Worker/Rust/WASM；浏览器草稿使用小样本，不能据此宣称大样本或性能达标。

本批没有 UI 变更，实际入口、菜单和 Canvas 输入沿用现有 antdv-next / ConfigProvider。只完成源码/diff 回读、来源和声明/计划计数；统一语法/引用收尾由主线执行。测试、类型检查、构建、dev、浏览器、Rust/WASM、Office 文件及基准全部后置。真实主题/焦点/IME、两渲染后端、全部权限与关联资源、Office/协作保真和大样本仍 **未验收**；内部保存重开草稿不替代 XLSX 独立重开。
