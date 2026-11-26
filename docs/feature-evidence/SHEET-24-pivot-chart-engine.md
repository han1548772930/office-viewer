# SHEET-24 独立与关联透视图字段引擎

日期：2026-11-26。状态：实现与契约草稿已落盘，**未验收**。本记录覆盖字段引擎；公开 Facade、Studio 命令、模型迁移、Panel 与浏览器交互由对应主任务记录，不用本记录代替其验收。

后续更新：共享 Show Values As 的固定 15 项、历史兼容项、字段设置与独立/关联报告编辑已落盘，见 [SHEET-23 / SHEET-24 值显示方式证据](SHEET-23-24-pivot-show-data-as.md)。本记录的 30 条声明 / 51 个草稿用例属于前一批字段引擎，后续草稿和验证缺口单列；全部仍未验收。

## 固定依据与模块归属

- 固定 OSS 仓库：`E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。
- Pro 依据使用项目内已保留的 v1.6.0 发布类型和可访问包装实现，不安装或依赖商业 SDK；类型可以证明公开契约，不能证明私有计算实现或运行结果。
- [types.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts)：range/table source union、initialFieldsConfig、14 类 `IPivotChartEdit`、filter query DTO、字段配置与来源同步。
- [f-pivot-chart.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts)：开头说明 owned/linked 所有权，`setFieldsConfig/resetFields/addField/updateFieldPosition`、值位置、Show Values As、分组/排序/筛选/折叠和视觉入口。
- [pivot-chart-edit.controller.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/controllers/pivot-chart-edit.controller.d.ts)：共享 report edit 与 detached PivotTable 包装路径。
- [pivot-chart-source.service.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/services/pivot-chart-source.service.d.ts) 与 [sheets-pivot-chart.model.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/models/sheets-pivot-chart.model.d.ts)：有效报告解析与 source/report 边界。
- UI 契约同时对照 [menu/schema.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/menu/schema.d.ts) 与 [pivot-chart-field-view.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/pivot-chart-field-view.d.ts)。本批不新增顶部工作区入口，不改其分类计数。

模块属于完整工作台架构的 `SHEET-24 → SHEET-23/SHEET-22`。领域字段、排序、过滤和分组继续由已有 `features/pivot/services/pivot-engine.ts` 计算；新 `pivot-chart-engine.ts` 只编排所有权、字段实例解析、校验、单次刷新与原子提交。组件、Facade、store 不另建聚合算法。

## 公开入口清点与本批对应

上游 edit union 共 **14 类**。本地 union 在后续值显示方式更新后共 **17 类**：11 类对照上游共享引擎路径，另加 resetFields、setNumberGroup、setDateGroupingMode、setManualGroups、refresh、setShowDataAs 共 6 类。这个计数属于协议类型清点，不能作为能力完成率。

| 上游 edit 类型 | 本地路径与实际作用 | 差异、未实现或未核实 |
| --- | --- | --- |
| setFieldsConfig | 校验真实 source/local ID、区域位置、字段属性、options；复用所有已有 setter 后单次刷新 | 本地 DTO 为 fields/options/dateGroupingMode；不直接接收上游 snapshot/collection JSON。未实现属性显式拒绝 |
| setOptions | 复用报告选项：重复行标签、行/列小计及总计、值放行/列 | 上游其他选项尚未逐项实现，不能吞掉未知键 |
| addField | source ID 或 source 顺序索引；复用字段放置 setter | source 索引解析成稳定 cache ID；重复来源的新增实例需完整配置，addField 沿用已有放置语义 |
| moveField | canonical source ID 解析到唯一实例，明确 local ID 精确移动 | 多个相同 source 实例必须消歧；不默认移动第一个实例 |
| removeFields | source ID 删除该来源的实例，local ID 删除明确实例 | canonical ID 优先；切片器/时间线或 surviving value filter 依赖会拒绝 |
| updateValuePosition | 真正变更共享 report 的值轴行/列并重新计算 | 多值虚拟字段任意 `index` 尚未实现；带 index 的编辑显式拒绝 |
| setFieldSetting | displayName、format、11 类 aggregation 与后续 showDataAs 规则 | 值显示方式见后续独立证据；标签字段不能配置 value aggregation/show rule |
| setFilter | 原始 source 成员的字符串键、isAll、clearAllFilters；清除已有 label/value filter | partialItems、日期条件、hasData、三态过滤查询尚未实现 |
| setValueFilter | 8 类数值比较；valueFieldId 解析成真实 local value instance | 上游更多条件、Top-N 等扩展未实现；无依据不宣称支持 |
| setDateFilter | 无 | 日期条件筛选尚未实现，显式拒绝，不用提示消息充当操作 |
| setCollapse | 无 | 共享 report 折叠、指定成员/层级和 Canvas 字段按钮尚未实现 |
| setCategoryLevel | 无 | 类别层级钻进/上卷尚未实现 |
| setDateGroup | 复用 year/quarter/month/day 分组及明确 Excel-like mode | 仅现有分组语义；上游秒/分钟/小时等完整粒度、分组字段结构尚未对齐 |
| setSort | 标签 ascending/descending/null；真实修改聚合输出顺序 | 上游 sort-info、值排序/手工序列等完整排序语义尚未实现 |

Facade 中的完整字段及来源入口也纳入目标：

| 入口族 | 清点内容 | 本批对应与缺口 |
| --- | --- | --- |
| 所有权、来源与报告读取 | getSourceKind、isLinkedToPivotTable、getSourceRangeInfo、getPivotTableReference、getPivotTable、getFieldsConfig、getSourceFieldsInfo、getFieldSetting、getFieldIdsByArea | 本地有效 owned XOR linked 报告解析与 source fields 可供主 Facade 包装；读取返回类型仍需逐一对照 |
| 字段编排与报告选项 | setFieldsConfig、resetFields、addField、removeField、updateFieldPosition、updateValuePosition、setFieldSetting、renameField、getOptions、setOptions | 复用本批 edit driver；报告 options 与既有 Chart 作者 options 由主任务分开 |
| Show Values As | getShowDataAs、setShowDataAs、resetShowDataAs | 后续共享规则、计算与 edit 已落盘，见值显示方式证据；API/UI/Office 验收仍未执行 |
| 筛选、排序、日期与层级 | setLabelSort、clearLabelSort、setLabelManualFilter、setDateFilter、getValueFilter、getValueFilters、setValueFilter、clearValueFilter、setDateGroupType、setCollapse、setCategoryLevel | label/value 部分可包装；日期条件、collapse、category level 和完整 filter query 仍缺 |
| 来源替换 | setSourceRange | owned/linked 来源替换与 cache/connection 原子迁移尚未由本批实现；不得改变 ownership mode |
| 图表、控件、绘图 | getInfo、toBuilder、update、getId、getType、getControls；type/title/subtitle/legend/theme/palette/appearance、auto-gradient、invalid-value strategy；controls/field buttons/expand-collapse buttons；position/absolute position/size/z-order/remove | 统一复用图表/绘图主链路；本批不把已有视觉按钮等同上述全部 API 已完成。控件字段按钮仍需 Canvas 与 UI 闭环 |

来源类型包装包所引用的底层 Pro 包在本地不存在，仍保留完整目标。没有公开内部源码不构成范围缩减；后续使用已公开类型、文档/演示及真实 Office 样本补行为与证据。

## 已落盘行为

### 真正独立的范围报告

`createDetachedPivotReport` 创建 fields=[]、默认 options、revision=0 的实际领域报告。`createPivotChartReport` 校验唯一 report ID、真实来源和宿主，登记/复用 retained source cache 后返回报告，供主命令放入 `chart.ownedPivot`。

- 不加入 `workbook.pivotTables`，不创建虚拟工作表，不预留宿主单元格。
- anchor 记录真实 chart host；row/col=0 是宿主身份数据，data-only 计算不会把它当输出地址。
- 未配置字段允许创建、读取成员和继续编辑；无 value 时共享计算返回空矩阵。缺少 row/value 的图表展示空数据路径由主任务更新 `chart-pivot-data.ts`。
- retained cache 明确保留 source field ID 和 sourceColumn；report-local field ID 可独立于 source ID。

### 共享计算、单次刷新与 no-op

现有 setter 增加末尾 `PivotRefreshMode`：默认 worksheet 保留普通透视表写单元格路径；data-only 使用同一 `buildPivotOutput`，不找输出宿主、不写 cells；deferred 只用于 detached batch 校验与编辑，最终调用者统一刷新。

`applyPivotChartEdit` clone 当前 report，复用共享 setter，以 fields/options/dateGroupingMode 判定真实内容变化。相同 options/setting、相同标签选择集合、等价区域顺序以及不存在字段的删除返回 false，不增加 report/chart revision。

一次真实编辑将 draft revision 还原到当前 revision 后调用一次最终 `refreshPivotTable`：owned 使用 data-only，linked 使用 worksheet。显式 refresh 是重新计算操作，单次计算并增加一次 revision；没有为 owned 报告持久化第二份派生矩阵。

linked 输出在候选 cells map 中预计算/写入，成功后才提交。先预检 output conflict 再移除旧 generated cells；冲突保留旧报告、单元格、source mirrors 和 chart revisions。关联同一 PivotTable 的多个图表共享报告修改，但保留各自视觉配置。

本 helper 不再次调用 `normalizeSlicerCaches` 或 `synchronizeSlicerSelections`。主任务共同 `prepareWorkbookChange` 在 workbook 候选上负责 selection/cache fan-out、权限及统一内容历史，避免重复扇出。跨命令撤销重做和 UI 监听由主任务集成草稿覆盖，本记录不声称已浏览器验证。

### 实例身份与保护

- shared `setPivotField` 新增最后可选 `instanceId`，精确移动已配置实例并验证 sourceColumn 一致；既有调用默认语义保留。
- 创建字段时 local ID 与已有实例冲突会分配独立 ID，sourceFieldId 保持稳定。
- 共享聚合和 drill-down 的 group/member keys 以 report-local ID 查找领域字段，再计算相对于 source.range.startCol 的偏移，修正 local ID 与 source ID 不同以及非零源列时的真实匹配。
- 带多个实例的 canonical source 字段设置必须消歧。raw source member 查询可以直接使用 canonical ID，不需要已有报告实例。
- 绑定的最后一个标签实例删除/转 value 会抛 `PivotChartEditError('field-connected')`。有效连接与 retained cache refs 均保护字段；兄弟报告暂时缺失或 control view 暂时不存在不绕过保护。
- 仍被存活 valueFilter 引用的 value instance 不可删除或转标签。一次 reset 全部相关字段允许清空报告，不留下悬空 valueFieldId。
- 未实现 edit、字段属性或参数键显式拒绝，保留原数据；不静默保存一个计算器忽略的 Pro 配置。

### 读取接口

```ts
createPivotChartReport(workbook, source, hostSheetId, id?): PivotTableModel
getPivotChartReport(workbook, chart): PivotTableModel
getPivotChartSourceFields(workbook, chart): PivotSourceFieldInfo[]
getPivotChartFieldItems(workbook, chart, fieldId: string): string[]
applyPivotChartEdit(workbook, chart, edit: PivotChartReportEdit): boolean
```

field items 从实际 source 列读取 source 顺序的原始去重值，再按共享 label-filter 的 `String(value ?? '')` 生成键。canonical source ID 优先；可读未配置 source field，也兼容明确 local ID。底层 `getPivotSourceFieldItems` 保留 `CellScalar[]`，空值和布尔值不会被伪造为 plot output。查询不改筛选、不读图表投影、不返回 grouped/hasData/date-condition/tri-state 结果。

## 计算域与协议

本批没有新的 Worker/WASM 协议，也不建立第二套权威计算引擎。当前字段 setter 和 report output 继续复用已有 TypeScript pivot engine；这不等于完成大数据 retained Rust 聚合。

后续大计算仍按 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 纳入现有 Worker/Rust 引擎，使用协议版本、unit/baseRevision、取消、预算和 lease/复制统计。没有同数据集实测，不声称本次提高性能、零拷贝或达到大表标准。

## 文件、草稿与验证状态

本批文件归属：

- [pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts)：detached report、共享 refresh modes、no-op、实例定位、local ID 聚合/drill-down、依赖保护与来源边界校验。
- [pivot-chart-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts)：owned/linked 报告编辑驱动、成员读取和 atomic staging。
- [pivot-chart-engine-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-chart-engine-fixture.mjs)：真实非零 source offset、retained cache、owned/linked 报告及连接样本。
- [pivot-chart-engine.contract.test.mjs](../../packages/office-studio/tests/pivot-chart-engine.contract.test.mjs)：独立契约草稿。
- 本证据文件；不编辑 model、Facade、Studio、index、Panel 或主功能清单。

草稿计划涵盖：创建前后 cells/sheets/pivotTables 身份、未配置来源查询、一次完整配置、11 类汇总、worksheet/data-only 等价、多个 value instance、canonical/local ID collision、真实 drill-down、numeric/date/manual groups、label/value filter、sort/value position、no-op、explicit refresh、linked mirrors、output conflict、连接及 valueFilter 依赖拒绝、非法/未实现参数、来源/宿主/报告缺失、快照重开。

静态清点 **30 条 test/test.each 声明**；按三组 test.each 参数展开为 **51 个草稿用例**（11 类汇总、2 种所有权、11 种非法/未实现编辑）。内部循环覆盖多个断言场景，未另计为运行用例或通过结果。

按照用户当前“先实现、后统一测试”要求，**未运行任何测试、类型检查、构建、Rust/Cargo/WASM、dev、浏览器、文件样本或性能基准，也未运行 git diff --check**。静态读写及草稿声明清点不属于执行验证，不称为测试通过。

仍需验收或实现：

1. 主命令/API/Panel 的权限、selection fan-out、撤销重做、跨宿主和 UI/IME/焦点/主题浏览器闭环；本批不修改 UI，不能替 UI 迁移验收。
2. 真实 XLSX pivotChart/pivotCache/chart relationship 写出、Office 原生读写和独立重开；内部快照草稿不能证明 Office 往返。
3. Show Values As 的真实闭环验收见后续证据；date-condition、partial filter query、collapse/category level、任意多值字段位置、来源替换、完整 Pro fields/options/controls 等上述未实现入口继续保留完整范围。
4. retained Rust/WASM 大样本聚合、layout/render coupling、任务取消、内存和真实前后基准。

SHEET-24 与所有上述 Pro 行为继续保持完整交付目标和“未验收”，本批字段引擎不能替代全部透视图能力完成证明。
