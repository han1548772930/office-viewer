# SHEET-11：筛选表头入口与颜色条件续接

日期：2026-11-28。固定参考 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本记录续接[工作台记录](WORKBENCH-CONTINUATION-2026-11-28.md)和[连续区域/面板 owner](SHEET-11-12-continuous-range-2026-11-28.md)，复用既有入口清点，状态为**已落盘、未验收**。

## 固定源码与入口

| 来源 | 已核实的契约 |
| --- | --- |
| `packages/sheets-filter-ui/src/menu/schema.ts` | DATA / ORGANIZATION 下筛选根 `order:1`，grid `row:2,column:1`；清除/重算分别为已有子项 `order:0/1`，不是本批新增根 |
| `packages/sheets-filter-ui/src/views/widgets/filter-button.shape.ts` | 图标尺寸 16、padding 1，激活态区分已有条件 |
| `packages/sheets-filter-ui/src/views/render-modules/sheets-filter.render-controller.ts` | 在实际筛选范围首行、每列右侧放置控件；纵向位置取单元格对齐；文本 `rightOffset:16`；不是工作表 A/B/C 列标按钮 |
| `packages/sheets-filter-ui/src/commands/operations/sheets-filter.operation.ts` | `sheet.operation.open-filter-panel` 参数 `{unitId,subUnitId,col}`；打开前结束单元格输入；`sheet.operation.close-filter-panel` 只改变面板状态 |
| `packages/sheets-filter/src/models/types.ts`、`models/filter-model.ts` | VALUES / COLORS / CONDITIONS；颜色条件支持 `cellFillColors:(string\|null)[]` 与 `cellTextColors:string[]`，填充字段存在时优先；null 是默认无填充，默认文字黑色 |
| `packages/core/src/sheets/worksheet.ts` | `getComposedCellStyle` 合成缺省/行列/主题/单元格样式，并通过 `getCell` 读取拦截后的单元格 |
| `packages/sheets-filter-ui/src/services/sheets-filter-panel.service.ts::ByColorsModel` | 候选颜色排除其余列已过滤的行；填充和文字勾选互斥；同类允许多色；无勾选提交时清除此列条件 |
| `packages/sheets-filter/src/models/filter-model.ts::FilterColumn.calc`、`packages/core/src/sheets/worksheet.ts::iterateByColumn` | 合并单元格按锚点判定，失败时隐藏完整 rowSpan；横向合并内部列仍读取左侧锚点；从合并内部行开始遍历时跳过该组 |
| `packages/sheets-filter/src/facade/f-range.ts::createFilter`、`commands/commands/sheets-filter.command.ts::SetSheetFilterRangeCommand`、`commands/mutations/sheets-filter.mutation.ts::SetSheetsFilterRangeMutation` | 公开 createFilter 原样传递 range，set-range 只拒单行、不自动扩展合并边界；因此合并组跨 ref.endRow 时 cachedFilteredOut 可以包含 ref 外的尾部行 |

固定筛选根、清除与重算子项已在[原顶部清点](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)和[Sheets 来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)登记。本批常规 Ribbon 根增量 **0**，使用原 72 根基线；表头控件数量由当前筛选列和可见 pane 派生，单列在其所属 pane 呈现一次，不另计作常驻 Ribbon 根。

## 已落盘链路

- [sheet-filter-header.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-filter-header.ts) 从同一 Canvas 几何派生可见按钮、缩放、冻结、裁剪和条件激活态。[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 使用 `antdv-next` Button/Tooltip，点击保留原选区，组字/权限/宿主失效时禁用。编辑中的按钮位于输入覆盖层之上；填充柄保持原更晚的同级层序，沿用原捕获处理。
- [scene.ts](../../packages/office-studio/src/rendering/spreadsheet/scene.ts) 对实际筛选表头保留 16px 的文本空间，标签/文本均使用收窄后的内容矩形；筛选表头不再向右溢出绘制内容。正文/其他行不改变。
- [sheet-filter-panel-commands.ts](../../packages/office-studio/src/api/sheet-filter-panel-commands.ts) 注册固定 open/close operation，支持公开 `subUnitId/col` 和本地 `sheetId/column` 一致性校验。editor owner 按 Studio/unit 隔离；替换、关闭文档、卸载与 dispose 释放；无实际 host 时返回 false。句柄执行绑定版本、历史 generation、用户与 owner，拒绝观察事件重入后的旧请求。打开和按钮禁用态共用 `sheetFilterPermissionRanges`，覆盖 ref 外相交合并尾部的保护范围。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 绑定真实 editor，点击相应列后打开该字段。公共 operation 和鼠标入口均通过现有 `finishEdit`；组字或提交拒绝时保留输入，不打开新面板。面板关闭和打开不写内容历史。
- [SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue) 声明并使用 `initialColumn`，以现有 Ant Form/Select/Checkbox 增加颜色模式，填充/文字互斥、多色勾选、无勾选清除。组字期间 Enter 不重新读取并清空字段，字段切换与提交被守卫；文本输入仍交给 Ant Input 的原生组字桥。
- [sheet-filter-colors.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-filter-colors.ts) 合成继承/主题/直接样式，复用原 `conditionalFormatStyleAt`、`conditionalFormatFormulaAt` 和 `evaluateFormula` 读取派生颜色，不另建条件格式或公式引擎，不把显示颜色写入单元格。
- [sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts) 颜色规则参与已有行可见性谓词，保留 fill 优先、null、默认黑色与空 API 数组的语义。getter、排序准备、刷新与设置入口接受可选整本 Workbook 上下文，避免读回时丢失工作簿缺省样式；共享调用点由主线接入。
- [sheet-data-panel-state.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-data-panel-state.ts) 值/颜色候选排除其余字段条件，忽略当前列自身条件。确认仍用已有 `sheet.command.configure-filter` 和候选模型，一次提交范围及全部列条件，复用权限、内容事件、撤销和 native 保存。
- [sheet-filter-column.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-filter-column.ts) 复用一份列遍历契约，读合并锚点、跳过垂直内部行、返回完整 rowSpan。引擎按各列失败行求并集，保留手动隐藏等其他原因；横向合并内部列读取左锚点，不把内部脏值当实际内容。值候选只要组内有一行未被其他列排除就保留；颜色候选按固定 `skipNonTopLeft=true` 仅取顶左候选及其首行的其余列状态。

跨 ref 末行的合并组仍保持原 `filter.range`，隐藏行缓存保留完整组，不缩短或悄悄改写数据区域。主线同步 TypeScript/Rust 精确缓存边界，仍拒标题行、重复行和 Office 绝对越界，writer 继续检查显式 rowCount；criteria 列和数据 ref 的边界没有放宽。

颜色模型与 parser 归一化、Rust/Office 原生 `colorFilter`/dxf、精确来源保留和共享导出由主任务协调；该部分不以本 UI 记录代替 Office 往返验收。

## 测试草稿和实际检查

新增 [sheet-filter-color-header.contract.test.mjs](../../packages/office-studio/tests/sheet-filter-color-header.contract.test.mjs) 十一个声明：原颜色/表头/快捷键六项，再补垂直/横向合并、跨 header/ref 边界、组内候选差异和独立权限 footprint 五项。新增 [sheet-filter-panel-commands.contract.test.mjs](../../packages/office-studio/tests/sheet-filter-panel-commands.contract.test.mjs) 四个声明：真实列/无内容副作用、owner/目标校验与销毁、观察事件重入拒绝及受保护合并尾部拒绝打开。现有 [sheet-data-panel.contract.test.mjs](../../packages/office-studio/tests/sheet-data-panel.contract.test.mjs) 加颜色、跨 ref 合并两条历史/重开草稿；[sheet-data-toolbar.spec.ts](../../tests/browser/sheet-data-toolbar.spec.ts) 加一条真实表头点击、多色/互斥、历史及原选区草稿。

同批公式 UI 接线新增两个浏览器声明，见下节。合计本子任务新增 **20 个 test 声明**（17 个模型/命令/契约、3 个浏览器），并调整既有公式菜单及参数表单草稿；循环展开的计划场景不混作已执行数。

已执行局部 `git diff --check`，当时通过；最终全局语法/引用/文档检查由主任务在交接后统一记录。**运行测试、类型检查、构建、Rust/WASM、dev、浏览器、Office 和性能均未执行**，不提升功能验收状态。

## 同批 SHEET-09 公式 UI 接线

主线实现自动范围 planner 和批量命令后，本子任务将分类菜单改接固定 `formula-ui.operation.insert-function` / `{value:name}`。原 `sheet.operation.insert-function` UI alias 复用相同 editor 句柄；“插入函数”参数 Form 继续保留为明确入口，不再替代分类选择的自动引用行为。

真实 host `registerSheetFunctionEditor` 在 open 时恢复并精确核对当前逻辑范围/主格，映射合并主格后创建现有公式输入会话；未闭合 `=FUNC(ref` 保留给用户编辑。batch 后同步当前模型视图，已有输入会话仅在编辑格内容变化等真实冲突时失效，其他格的 batch 不因 revision 改变被取消。原内容历史/保存与 Input/IME 桥沿用。

[sheet-function-shortcuts.ts](../../packages/office-studio/src/components/spreadsheet/sheet-function-shortcuts.ts) 按 `QuickSumShortcut` 和 `whenSheetEditorFocused` 实现 Windows Alt+=、Mac Cmd+Alt+=，仅在 Canvas 获焦且单元格编辑器未激活时派发，排除组字和其他修饰键。既有快捷键面板和 SUM 菜单元数据同步。

[sheet-ui-ribbon.contract.test.mjs](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 核对固定命令/对象 DTO；[sheet-toolbar.spec.ts](../../tests/browser/sheet-toolbar.spec.ts) 保留显式参数 Form 的求值/历史/重开场景，另加分类批量总计的一次历史和 Alt+= 未提交输入/确认草稿。没有新增唯一函数算法或分类根；主线 planner/API 的测试和差异单独记录。

## 仍需继续的目标

- 真实浏览器中的按钮命中、冻结/滚动/缩放、明暗/弹层主题、焦点与 Windows IME，以及前轮填充柄修复的交叉回归尚未验收。
- 本地使用现有排序/筛选 Modal；固定版为列旁浮动面板。位置及外观有明确差异，不能声称已复刻浮层交互。
- 值候选日期分组树、完整复合条件、结构化表自身的全部表头控件及跨资源行为继续保留缺口；合并的范围外缓存、权限及 Office 交叉场景尚未运行验收。上游未用的 DynamicFilter/Top10 类型不能作为已支持证明。
- 颜色比较当前明确覆盖常用 hex/RGB 与原始同名字符串；全部 CSS 颜色等价、源主题/条件格式各分支和外部 Office 修改的往返必须继续补样本。
- 颜色投影复用已有 TypeScript 计算路径；大型筛选/条件规则热点的 Worker/Rust 分派、取消与同数据集性能测量尚未完成，不宣称已优化。
- 完整函数/数组/高级公式、Pro 深层行为和其余 65 个功能组继续按全局清单推进，本批不是完整功能交付。
