# SHEET-22 图表类型源码对照

状态：未验收。开源参考 checkout 为 `E:\rust_test\univer` tag `v1.6.0`，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro UI 证据来自本地研究缓存的 `@univerjs-pro/chart-ui@1.5.2` 和 `@univerjs-pro/engine-chart@1.5.2` 发布包，不能视为 v1.6.0 Pro 行为已核实。

## Pro 发布包证据

- `.research.local/univer/packages/chart-ui/lib/types/components/ChartTypeOptions.d.ts` 导出的 `CHART_TYPE_CATALOG` 声明 Line、Column、Bar、Pie、Area、Radar、Scatter、Combination、WordCloud、Funnel、Bubble、Relation、Waterfall、Pareto、Sankey、Heatmap、Boxplot、Candlestick、Histogram、Treemap、Sunburst、Gauge 和 Chord；另有堆积/百分比堆积/3D 变体、Pie of Pie 与 Bar of Pie。
- `.research.local/univer/packages/engine-chart/lib/types/enum.d.ts` 的 `ChartTypeBits` 声明类型位。其 `ChartAttributeBits` 将 ThreeD、Stack、PercentStack、Horizontal 分别编码为独立属性位。
- `.research.local/univer/packages/chart-ui/lib/es/index.js` 的图表样式面板按 chart type 渲染 Radar、Funnel、Treemap、Sunburst、Gauge、Chord、Waterfall、Pareto、Pie、WordCloud、Heatmap 和 Relation 专属配置 section。发布包同时注册 `ChartSetupPanel`、`ChartStyleSection`、`ChartTypeSpecificStyleSection`、`ChartSeriesSection`、`ChartTitlesSection`、轴、图例、趋势线及误差线 section。
- `.research.local/univer/packages/chart-ui/lib/types/components/ChartSemanticDataFields.d.ts` 描述 Candlestick 的 category/open/high/low/close 字段、Histogram 数值字段和 Treemap 层级/数值字段。`ChartTypeOptions.d.ts` 的 `IChartTypeOption` 与 `IChartSelectOption` 保留 type bit、子饼图类型、id 与 label 的映射。
- 固定版 v1.6.0 OSS checkout 未找到上述 Pro 图表 UI 源码。固定版 `packages/sheets-drawing-ui/src/services/canvas-float-dom-manager.service.ts` 的 `applySheetChartTransformerConfig` 只证明通用图表对象的选中、变换框交互，不能证明 Pro 图表类型或编辑面板语义。

## 本地实现与差异

- `packages/office-studio/src/domain/model.ts` 的 `ChartType` 当前为 `bar | line | area | pie | doughnut | stacked-column`。`ChartModel` 保存 range、columns/rows orientation、categoryField、valueFields、轴和图例状态，没有 chart type bit、组合系列类型、散点 X/Y 映射或 Pro 类型专属配置。
- `packages/office-studio/src/features/charts/chart-engine.ts` 的 `validateChartType` 只接受上述六类。`readChartData` 以一个分类字段和其余数值字段构造类别/系列；`chartLayout` 为类别均匀分配 X 坐标；`chartRenderCommands` 分别绘制柱、线、面积、饼和圆环。没有散点数值 X 轴、组合系列、百分比/3D 堆积或 Pro 专属图形布局。
- `packages/office-studio/tests/chart-engine-rendering.contract.test.mjs` 当前只覆盖折线图轴标签/网格和对数轴，未覆盖六种类型各自的几何、边界及拒绝条件。
- `packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue` 提供六种本地类型选项。`ExcelWorkbench.vue` 普通图表对话框当前只提供柱形与折线入口；其可见范围不能视为 Pro 图表编辑器的等价实现。

## 范围与未核实项

此记录证明 Pro 1.5.2 已发布 bundle 的类型声明及部分可观察 UI 组件，不证明闭源实现细节、v1.6.0 Pro 版本类型增减或每个类型支持的数据约束。Pro 1.6.0 发布包是否已获取并逐项复核，需更新本记录后再判断。

散点图是可继续独立实现的类型目标，但需要显式的数值 X 字段映射、X/Y 轴范围与坐标变换、系列/图例语义、历史与持久化字段和 Office 文件映射。仅把 `scatter` 加入 `ChartType` 或把分类索引当作 X 数值，不能满足其数据契约。拟实现的第一切片应先按公开/可观察行为定义工作表列映射，再连通模型、编辑入口、Canvas 绘制和契约测试；这些工作完成前 SHEET-22 保持未验收。

## 验证边界

本轮为发布包声明、固定源码和本地 source/test 的静态对照，没有执行测试、构建、浏览器交互或 Office 文件往返验证。
