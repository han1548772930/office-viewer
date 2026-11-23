# SHEET-22 / SHEET-24 / DOC-19 图表类型与编辑契约对照

更新日期：2026-11-24。**整体状态：未验收；当前开发进度：开发中、待统一验证。** 本轮遵守用户“先实现，后统一测试”的要求，仅实现、补测试草稿并静态阅读源码/diff；没有执行测试、类型检查、前端构建、Rust/WASM 检查、基准、dev server 或浏览器。下文“已接入/已编写”只描述代码，不代表公开行为、Office 往返或性能验收。

本记录覆盖共享图表引擎与 Sheets 普通/透视图，包含 Docs 图表需要复用的契约。DOC-19 的 Docs 插入、内嵌数据、对象锚点/环绕/分页及编辑宿主尚未接入，不能由共享目录或算法推导其完成。全部 45 项、全部 Pro 配置和入口继续保留在交付范围内。

## 1. 固定来源与证据边界

OSS 只读 checkout：`E:/rust_test/univer`，tag `v1.6.0`，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 发布包研究缓存：`.research.local/univer-1.6.0/packages/`。`chart-ui`、`engine-chart`、Sheets/Docs/Pivot Chart 相关包的 `package.json` 记录版本 1.6.0；**engine-chart 已解包**。本地研究缓存不是生产依赖，没有安装或引入商业 SDK。

| 来源 | 具体文件/符号 | 本轮使用的证据及限制 |
| --- | --- | --- |
| P01 目录与面板 | `chart-ui/lib/types/components/ChartTypeOptions.d.ts::CHART_TYPE_CATALOG`；`chart-ui/lib/types/services/chart-view-state.d.ts::ChartViewStateKey`；`chart-ui/lib/types/index.d.ts` | 45 个目录选择项、46 个编辑器状态键及 section 声明；状态键不是持久化字段、命令 ID 或完整功能计数。 |
| P02 类型/转换 | `engine-chart/lib/types/chart-builder/chart-types.d.ts`；`chart-builder/conversion/chart-type-converter.d.ts`、`build-chart-preview-data.d.ts` | ChartTypeString/系列类型、bit/stack 转换与预览边界。本地使用字符串目录和独立数据角色转换，不复刻闭源内部实现。 |
| P03 数据角色 | `engine-chart/lib/types/models/chart-data/builders/` 下的 `category-series-chart-data.d.ts`、`cartesian-point-chart-data.d.ts`、`relation-chart-data.d.ts`、`sankey-chart-data.d.ts`、`chord-chart-data.d.ts`、`hierarchy-chart-data.d.ts`、`heatmap-chart-data.d.ts`、`histogram-chart-data.d.ts`、`boxplot-chart-data.d.ts`、`candlestick-chart-data.d.ts`、`gauge-chart-data.d.ts` | 图型分开消费字段与数据。声明只证明入口/类型，不证明空值、聚合、布局或最终效果；本地规则在第 4 节明确列出。 |
| P04 Builder/3D | `engine-chart/lib/types/facade/f-axis-three-d-chart-builder.d.ts`、`f-three-d-chart-builder.d.ts`、`builders/f-bubble-chart-builder.d.ts`、`builders/f-radar-chart-builder.d.ts` | 深度轴独立于相机、没有数值 min/max/format；上游轴式 3D Builder 不暴露次轴/趋势线。Bubble 公开 Builder 使用 SolidSeries API，本地通用 marker shape/size 是增强。 |
| P05 Sheets 命令 | `sheets-chart/lib/types/commands/commands/sheets-chart.command.d.ts`、`sheets-chart-update-config.command.d.ts`；`sheets-chart-ui/lib/types/commands/commands/focused-chart.command.d.ts`、`sheet-chart-menu.command.d.ts` | 插入/配置更新、可逆 mutation 准备、焦点编辑/复制/图片导出/删除命令契约。完整宿主状态与持久化仍需行为验证。 |
| P06 Sheets 菜单与宿主 | `sheets-chart-ui/lib/types/menu/schema.d.ts`、`chart-context-menu.d.ts`、`chart-sheet-schema.d.ts`、`chart-sheet-menu.d.ts`；`sheets-chart-ui/lib/es/index.js` | 静态读取发布 bundle 的 schema、命令与 permission guard；没有导入执行 SDK。图表工作表、迁移、浮动元素编辑也属于目标。 |
| P07 Docs Chart | `docs-chart/lib/types/commands/commands/` 下的 `insert-doc-chart.command.d.ts`、`remove-doc-chart.command.d.ts`、`update-doc-chart-config.command.d.ts`、`update-doc-chart-data-source.command.d.ts`、`update-doc-chart-drawing.command.d.ts`、`update-doc-chart-snapshot.command.d.ts`；`docs-chart-ui/lib/types/menu/chart.menu.d.ts`、`menu/schema.d.ts`、`lib/es/index.js` | Docs 有独立插入、数据源、配置、对象快照与宿主命令；本地尚无对应 Docs 宿主闭环。 |
| P08 Pivot Chart | `sheets-pivot-chart/lib/types/facade/f-pivot-chart-builder.d.ts`、`f-pivot-chart.d.ts`、`services/pivot-chart-source.service.d.ts`；`sheets-pivot-chart-ui/lib/types/menu/schema.d.ts`、`commands/commands/menu-insert-pivot-chart.command.d.ts`、`lib/es/index.js` | 源/cache、字段按钮/筛选/展开折叠和 `withPivotTable/chartOnly` 模式；本地透视面板接线与完整公开功能分别记录。 |
| P09 Gauge 语义与样式 | `engine-chart/lib/types/types.d.ts:185–279`；`enum.d.ts:128–140`；`facade/builders/f-gauge-chart-builder.d.ts:4–13,130–419`；`models/common/constants/default-chart-style.d.ts:4,82,160,238,294,342`；`chart-ui/lib/types/editor/sections/specific-section/gauge-style.d.ts:2–10` | 有限观测独立包裹、无效值跳过、指示器钳制而保留原始格式、量程/角度/9 个样式组、全部对应 setter/reset、5 套预设声明与 nullable 编辑 patch。声明不证明上游 UI 同时提供 5 个预设入口或运行时效果；本地差异见第 5 节。 |
| P10 可见数据默认 | `engine-chart/lib/types/types.d.ts:396–398::IChartContext.plotVisibleOnly`；`chart-builder/chart-types.d.ts:552`；`models/chart-model/chart-model.d.ts:74::getPlotVisibleOnly` | `IChartContext` 注释明确 “Absence preserves the historical true default.”；其余声明只提供兼容默认/布尔 getter，不证明运行时返回。缓存没有 engine-chart 可执行 bundle，本轮未核验上游 getter 实现。 |
| P11 Radar 专属契约与显隐 | `engine-chart/lib/types/facade/builders/f-radar-chart-builder.d.ts`；`types.d.ts::IChartRadarAxisLabelSpec`；`chart-ui/lib/es/index.js::Sl/xl` | shape/fill/axisLabel 的 6 个 setter/reset、DeepNullish 数值标签外观；shape reset 明确 Polygon，fill reset 明确仅轮廓。面板顺序为 shape → fill → 数值标签可见性，格式/字体仅在有效 visible 为真时出现；`Sl` 从 `A.radar.axisLabel.visible` 取默认，实际默认值仍未核实。 |
| P12 Cartesian 与 Boxplot 继承 | `engine-chart/lib/types/facade/f-cartesian-chart-builder.d.ts`、`facade/f-axis-chart-builder.d.ts`；`chart-builder/configuration/cartesian-series.d.ts/axes.d.ts/axis-pointer.d.ts`；`chart-builder/chart-types.d.ts:114–150,484–489::IChartAxisSpec/IChartAxesSpec/IChartAxisPointerSpec`；`enum.d.ts:242::AxisAlignEnum`；`facade/builders/f-boxplot-chart-builder.d.ts`；`chart-ui/lib/es/index.js::$c/js/Hs/zs/Bs/Vs/qi/Hi` | 系列 7 个/全系列 6 个 clear target、轴 8 个 clear target、轴标题字体/对齐、日期开关、指示器 4 个字段与恢复目标及 `number / { name }` selector。Boxplot 原始样本、标题/日期/指示器已接代码，SolidSeries fillPattern 继续待开发；上游次轴公开类型排除 reverse。标题沿轴定位和 Radar 指示器的最终效果未核实。 |
| P13 Bubble 映射与名称 selector | `engine-chart/lib/types/facade/builders/f-bubble-chart-builder.d.ts::setBubbleMapping/clearBubbleMapping`；`chart-builder/configuration/cartesian-series.d.ts::ChartPublicSeriesSelector`；`chart-builder/chart-types.d.ts::ChartSeriesSelector/IChartBubbleMappingSpec` | Bubble 的 X/Y 必为数值，size/label/series 为可选源字段；每源行一气泡，各向量按源行对齐。公开 selector 为 `number / { name }`，内部 selector 为 `{ index } / { name }`。本地保留既有数值源字段语义，以独立 namedSeries 数组表达名称覆盖；名称精确匹配、重名歧义拒绝和未活动覆盖保留是本地规则，缓存没有可执行 engine-chart bundle，尚未核实上游运行时边界。 |
| R01 本项目绘制依赖 | ECharts 6.1.0 的 `lib/chart/gauge/GaugeView.js:211–250,337–355,469,508,532`、`lib/util/shape/sausage.js`；zrender 6.1.0 的 `lib/core/PathProxy.js:34–55` | 静态读取当前锁定依赖的刻度文字定位/对齐、progress overlap/环宽/z2、锚点/文字层级与 arc normalization，作为本项目两绘制适配器的实现依据；不是 Univer 内部源码证据。 |
| R02 Radar / Boxplot 绘制依据 | ECharts 6.1.0 的 `lib/coord/radar/Radar.js/RadarModel.js`、`lib/chart/radar/RadarView.js`、`lib/chart/boxplot/boxplotLayout.js::layEndLine`、`BoxplotView.js` | Radar 原生坐标支持 clockwise，但没有径向 inverse；径向反向等配置复用 graphic。Boxplot 须线 cap 使用完整箱宽，原生 View 不绘制数据标签，启用标签时复用共享图元。ECharts 的 Radar 默认标签隐藏不是 Univer 默认值证据；实际 SSR 仍未执行。 |
| R03 真实指示器动作 | ECharts 6.1.0 `lib/component/axisPointer/modelHelper.js:93–190`、`AxisPointerModel.js`；2026-11-24 Context7 的官方 `updateAxisPointer`、`tooltip.triggerOn` 与 `axisPointer.type: cross` 文档 | tooltip 的 pointer 子模型优先于 component defaults，因此两个位置同步四个作者样式；离屏 Canvas 通过动作更新，`showContent:false/triggerOn:'none'`。文档为动态 API 依据，不是固定 Univer 运行时证据；原生和 graphic 浏览器草稿均未执行。 |
| O01 OSS 对象变换 | [固定提交的 canvas-float-dom-manager.service.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-drawing-ui/src/services/canvas-float-dom-manager.service.ts) 的 `applySheetChartTransformerConfig` | 只证明通用对象选中/变换框，不证明 Pro 图表类型、面板或 Office 行为。 |

公开产品/Office 行为对照入口：[Sheets 图表](https://docs.univer.ai/zh-CN/guides/sheets/features/charts)、[共享图表](https://docs.univer.ai/zh-CN/overview/features/charts)、[透视图](https://docs.univer.ai/zh-CN/guides/sheets/features/pivot-chart)、[Exchange](https://office.univer.ai/exchange)。动态规格、可访问演示和标准 Office 样本需分别记录观察日期、输入与结果。本轮没有浏览器或文件样本观察，不能把链接或声明当作已验证行为。

架构归属按 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的共享图表域执行；计算/内存按 [WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 执行。持久模型在 `domain/model.ts`，算法/投影/布局在 `features/charts`，宿主命令/历史在 `api/office-studio.ts`，Vue 面板只编辑草稿并调用 Facade。

## 2. 入口、布局与状态清点

以下是本轮能从固定发布 bundle 静态追踪的声明。常规 OSS 根入口清点仍见 [Sheets 入口表](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md) 与 [顶部清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)；这些文件中早期图表能力描述属于当时基线，当前图表代码状态以本记录为准。

| 分类 | 上游位置/顺序与命令 | 状态与本地接线 |
| --- | --- | --- |
| Sheets 插件扩展根声明 1 | INSERT/MEDIA，order 4；`sheet.command.menu-insert-chart`；该注册节点未声明 gridLayout，classic/grid 运行布局待验证。 | 需活动 Sheets、可编辑工作簿/工作表和有效最后选区；GRID 插入，CHART 工作表改当前主图类型，其他表类型拒绝。本地 `sheet.operation.insert-chart` → `ExcelWorkbench` 面板会话 → `FSheet.insertChart` → `sheet.command.add-chart`；共享 45 项目录。CHART 表专属分支尚未实现。 |
| Sheets 对象上下文声明 4 | DRAWING/OTHERS；order 0 编辑、1 复制、3 导出图片、4 删除；命令为 `sheet.command.chart.open-focused-panel`、`copy-focused-chart`、`export-focused-image`、`delete-focused-chart`。 | 上游按焦点 drawing、权限和对象类型确定目标。本地已有选中图表打开 Drawer、移动/缩放及 Facade 更新/删除；完整对象右键、复制、图片导出和元素浮动菜单尚未同命令闭环。 |
| Sheets 图表工作表/扩展上下文 | `createChartSheetMenuSchema/createChartSheetMenuRemovalSchema`；`sheet.menu.change-chart-type`；`sheet.operation.open-move-chart-dialog`；另有元素/轴/系列浮动菜单。 | 注册分支、元素类型、配置开关及目标表状态决定入口。只识别到声明/静态分支，未形成完整可见项计数；本地图表工作表迁移、页面适配和图内路由仍待实现。 |
| Docs 插件扩展根声明 1 | INSERT/MEDIA，order 2.5，grid row 2/column 4；`doc.operation.open-doc-chart-insert-dialog`。 | 当前文档、权限及一个有效折叠文本范围；宿主位置解析失败时禁用。该根项还在段落/其他块上下文复用，并有 below 变体；不能重复计入常驻根数。本地 Docs 根入口/宿主尚未接入。 |
| Pivot 插件扩展根声明 2 | ANALYSIS；两项 order 1，均声明 grid row 2/column 1、showLabel；`sheet.command.menu-insert-pivot-chart` 的 `withPivotTable/chartOnly` 分支。 | 两项可能由上下文选择，不能推导为同时可见两个按钮。本地 Data 入口 → `PivotChartPanel` → 透视 Facade/图表命令；已有透视可选择，完整双模式及字段按钮/筛选 popup 仍未验收。 |
| Pivot 上下文/浮动入口 | `sheet.operation.open-pivot-chart-panel/open-pivot-chart-field-list/open-pivot-chart-filter-panel`；`sheet.command.expand-pivot-chart-category-item/collapse-pivot-chart-category-item`。 | 绑定透视、源字段、分类项的展开/折叠状态；本地共享图表属性和透视配置面板已接，图内字段按钮、过滤与轴项交互未闭环。 |
| 类型目录子项 45 | P01 的 CHART_TYPE_CATALOG；Sheets 普通插入、属性和透视图复用同目录。 | 是选择项而非 Ribbon 根数；不是常驻、扩展、溢出或上下文按钮完成证明。Docs 将复用同目录，但尚无生产挂载。 |
| P12 对象属性面板叶字段 | `$c`：标题文字 → 可选对齐 → 可选字体复合控件；`js`：X 轴日期开关；`Hs`：线色 → 线型 → 标签底色 → 标签文字色。 | 本批每个可见轴接 7 个叶字段（文字 1、对齐 1、字体 5），X 轴日期开关 1，指示器字段 4。属于上下文属性字段，不增加常驻、插件根、溢出或 45 个目录计数；字体复合节点不与叶字段重复相加。运行时可见组合未验收。 |
| 溢出/运行装配 | desktop Ribbon 的 classic/grid 共享溢出机制。 | 未执行运行时菜单/权限/主题/窄屏验证；未知分支与上下文计数保留待核实，不能将静态根声明相加为完整可见入口总数。 |

本节计数仅为静态可识别根声明：Sheets 1、Docs 1、Pivot 2；Sheets 对象上下文 4；目录子项 45。图表工作表、浮动元素、Docs 复用上下文及 Pivot 字段入口另列，不冒称其完整入口总数已经清点完毕。

## 3. 全部 45 项本地映射

计数口径来自 [chart-catalog.ts](../../packages/office-studio/src/features/charts/chart-catalog.ts) 与 P01：Line 4、Column 6、Bar 6、Pie 5、Area 6，其余 18 项各 1；3D 11、2D 34。数据形状：category 33、xy 1、bubble 1、edge 3、hierarchy 2、distribution 2、ohlc 1、heatmap 1、gauge 1。**这些数字只证明目录覆盖，不证明功能完成。**

本地 legacy `bar` 表示纵向 Column，上游 Bar 映射 `bar-horizontal`；上游 `columnStacked` 的目录值为 `stacked-column`，同时保留本地历史别名 `columnStacked`；`donut` 映射 `doughnut`。`pieOfPie/barOfPie` 在上游复用 Pie bit + secondaryPlot，本地各有独立类型。历史别名不增加 45 项计数。

路径缩写：C = `chart-layout.ts/chart-draw.ts` 的分类/Cartesian；E = `echarts-option.ts` 普通 series；D = `chart-3d.ts`；P = `chart-pie-layout.ts`；S = `chart-special-layout.ts`；X = `echarts-special-option.ts`。所有路径先经过 `chart-record-projection.ts` 和 `chart-derived.ts`；共享图元进入 ECharts graphic 时复用 C/D/P/S，没有另一份统计权威。每行均为“代码已接入、未验收”。

| # | 上游目录 ID / 本地类型 | 数据契约与实际绘制路径 | 差异/待验证重点 |
| --- | --- | --- | --- |
| 1 | line / line | 分类、多数值系列；C/E 折线、空值断开/连接、标记。 | 点样式、标签、边界及 Office 往返。 |
| 2 | lineStacked / lineStacked | 同类同轴正/负分别累计；C/E 堆积折线。 | 栈/空值组合语义、合计溢出。 |
| 3 | linePercentStacked / linePercentStacked | 分类绝对值合计归一化为带符号百分比；C/E。 | Office/上游正负组合分母与百分比标签待核实。 |
| 4 | line3D / line3D | 分类和系列深度；D 实际旋转/透视投影、折线面、独立点。 | 遮挡/透明排序、单样本、公开 3D 效果。 |
| 5 | column / column | 分类、多系列；C/E 分组纵向柱形、gap/overlap。 | 分组宽度/边框/负值和类型转换。 |
| 6 | columnStacked / stacked-column | 同类同轴正/负分别累计；C/E。 | 混合次轴/组合栈、有限范围拒绝。 |
| 7 | columnPercentStacked / columnPercentStacked | 同第 3 项的百分比规则；C/E 柱形。 | 正负比例、轴边界和 Office 映射。 |
| 8 | column3D / column3D | 分类/系列深度、clustered/standard；D 六面柱体。 | grouping/gapDepth 与公开效果；局部样式。 |
| 9 | columnStacked3D / columnStacked3D | 分轴累计，D 堆积柱体。 | 遮挡及 stack/grouping 组合；本地次轴属增强。 |
| 10 | columnPercentStacked3D / columnPercentStacked3D | 百分比输入、D 堆积柱体。 | 正负比例、相机/深度轴、透明排序。 |
| 11 | bar / bar-horizontal | C/E 横向分组条形；分类 Y、主/次数值 X。 | 次轴已改为第二数值 X；首分类在上方，Office 默认顺序待核实。 |
| 12 | barStacked / barStacked | 横向正/负分别累计；C/E。 | 分类反向与主/次轴堆积。 |
| 13 | barPercentStacked / barPercentStacked | 百分比输入、横向条形；C/E。 | 正负比例、横向格式/刻度。 |
| 14 | bar3D / bar3D | D 横向柱体；分类标签左侧、数值标签下侧、次轴上侧。 | 标题/旋转投影、公开排序和布局。 |
| 15 | barStacked3D / barStacked3D | D 横向堆积柱体。 | 相机/深度轴与正负堆积。 |
| 16 | barPercentStacked3D / barPercentStacked3D | 百分比输入、D 横向堆积柱体。 | 同第 10/14 项及原生 Office 支持。 |
| 17 | pie / pie | 单一非负系列且至少一项正值；P/E 扇区、起角/爆炸/玫瑰/半饼。 | 全零/大数、标签线、系列名称和长图例。 |
| 18 | donut / doughnut | 同饼图；P/E 同一外半径与 doughnutHole 默认 0.55。 | 内外半径、爆炸留白及标签。 |
| 19 | pieOfPie / pieOfPie | P 主/副饼拆分；auto/position/value/percentage/custom、aggregate 项。 | 聚合图例、拆分索引迁移及公开连线/尺寸。 |
| 20 | barOfPie / barOfPie | P 主饼 + 副条形；副图中心与 aggregate 进入公共布局。 | aggregate 图例、拆分边界和 Office 复合图。 |
| 21 | pie3D / pie3D | D 扇面/侧壁、真实相机投影、分类配色/透明度。 | 扇区爆炸、遮挡、半饼/玫瑰等专属组合效果。 |
| 22 | area / area | C/E 面积、平滑、默认填充 0.28 与显式 opacity。 | 正负穿越、空值连接、标记/标签。 |
| 23 | areaStacked / areaStacked | C/E 同类同轴面积堆积。 | 基线/正负穿越、次轴和保存。 |
| 24 | areaPercentStacked / areaPercentStacked | 百分比输入、C/E 堆积面积。 | 正负比例、曲线/堆积交互。 |
| 25 | area3D / area3D | D 立体侧面、前后面积和正负穿越拆面。 | 透明层相交与遮挡、平滑差异。 |
| 26 | areaStacked3D / areaStacked3D | D 立体堆积面积。 | 叠层/相机/负值穿越与公开效果。 |
| 27 | areaPercentStacked3D / areaPercentStacked3D | 百分比输入、D 堆积面积。 | 同第 10/26 项、原生 Office 映射。 |
| 28 | radar / radar | 至少 3 分类、统一线性范围/半径；C/E 多边形或圆网、填充、数值刻度格式/字体、指标标签、分类/径向反向。 | 6 个专属 setter/reset、nullable patch 和条件显示已落盘；单点/值标签、分类旋转、径向反向及极端范围走共享 graphic。上游数值标签默认值、完整交互/原生/Office 效果未验收。 |
| 29 | scatter / scatter | 数值 X/Y；系列独立 X 向量、线性/对数/反向；C/E。 | 非正 log 过滤、不同 X 源、误差/趋势及 Office 往返。 |
| 30 | combination / combination | 每个系列 column/columnStacked/line/area/areaStacked；C/E、主/次轴。 | 混合栈、默认系列位置、转换/字段隐藏后的身份。 |
| 31 | wordCloud / wordCloud | 单非负权重系列；S 确定性词布局、形状/旋转/字号/重复。 | 字体实测/碰撞、满区重复、大样本预算/Worker 与公开形状。 |
| 32 | funnel / funnel | 单系列；P/X 排序保留原索引、gap、绝对值、零宽末端。 | 负值/绝对值规则、标签和公开边界。 |
| 33 | bubble / bubble | 数值 X/Y、可选 size/label/group；分组独立稀疏向量；C/E，源名独立样式与次轴。 | 同 Y 字段多组统计身份、名称 selector、2 个 mapping setter/clear 及字段/分组/单点设置已落盘，草稿未运行；重名/未活动行为、marker 增强与原生 Office 效果仍须区分和验收。 |
| 34 | relation / relation | 名义来源/目标、非负权重；S 图节点/边与布局。 | none/circular/force、环/自环、命中与大图预算。 |
| 35 | waterfall / waterfall | 系列逐项 change/total；统计结果 → C 专用基线柱/连接线。 | role/dataPoint 索引在过滤/聚合后迁移、完整专属配置。 |
| 36 | pareto / pareto | 单非负权重、同名分类聚合降序及累计比例；统计 → C/E。 | 零项/重复项、单点索引迁移、累计线样式。 |
| 37 | sankey / sankey | 名义有向边、重复边累计；S 节点流量、带宽、循环回流。 | 大图/循环碰撞、公开布局与路径交互。 |
| 38 | heatmap / heatmap | 分类为列、数值系列为行；C/X、连续/分段色标、单位/标签。 | 轴反向、显式点样式覆盖、完整色标交互。 |
| 39 | boxplot / boxplot | 每个原始数值字段一组样本；inclusive/exclusive 四分位、完整箱宽须线/异常点；统计 → C/X，共享分布样式、主/次轴、median anchor/标签。 | 使用继承的 Cartesian 系列 API，无专造 Boxplot setter；填充/边框透明度独立，log 下保留原统计并跳过非法图元。fillPattern、逐原始异常点独立覆盖与公开/Office 效果仍未完成或未验收。 |
| 40 | candlestick / candlestick | 最高/最低/收盘，开盘可选；验证 low ≤ open/close ≤ high；C/X。 | 三字段 HLC 不补造 open，走 graphic；OHLC 涨跌/hollow/颜色和 Office 往返。 |
| 41 | histogram / histogram | 单样本系列；Scott 自动分箱、箱数/箱宽、下/上溢箱、间隙；统计 → C/E。 | 分箱/边界规则、源点到箱索引和 Office 统计图。 |
| 42 | treemap / treemap | 逐层名称 + 非负权重、按完整路径累计；S/X。 | DFS 原索引/父标签已接；深层/同名路径、碰撞和大样本预算。 |
| 43 | sunburst / sunburst | 同层级契约；S/X 扇区、标签/中心文本。 | DFS 样式已接；公开标签旋转/交互和大量层级。 |
| 44 | gauge / gauge | 数值/可选 label；P/C/X 有限观测紧凑 tile、原分类索引、量程/角度/颜色段、9 个样式组、全部专属 setter/reset、5 套预设与原生分层。 | 源数值单元格显示格式元数据仍缺；极端范围 graphic、原生几何/层级、完整交互/Office/性能均待统一验证。 |
| 45 | chord / chord | 名义无向边、非负权重；S 节点扇区、ribbon/半径/起角/方向。 | 自环/权重/透明度、标签与公开有向/无向规格待核实。 |

## 4. 数据、统计与边界契约

- `ChartRecords` 保留 `fields/records/sourceCategoryFields/sourceValueFields/visibleFields/syntheticValues`。普通 Sheets 按列时首行为字段、按行时首列为字段；源表 `source.sheetId` 独立于宿主 `sheetId`，旧快照缺源表 ID 时使用宿主。原始 `rawValue: null` 不能退回格式化显示值，数值保留未舍入精度。
- `projectChartRecords` 统一普通、透视和预加载源；先校验字段/辅助引用，再处理可见性与空数据。透视数值角色由源声明，数字外观的分类 ID 不自动变为 measure；图关系端点仍是名义类别。隐藏辅助列不能被自定义误差读成零。
- `plotVisibleOnly` 缺省为 true，显式 false 包含隐藏数据；设置预览和 Switch 使用 `!== false`，与根任务 source-compute/Rust 路径统一。手动隐藏、零尺寸、filter/outline 的实际消费和 Office source helper 仍需统一验证；预览草稿覆盖双方向隐藏行列与开关后的重投影，不能由注释推导全部运行行为已验收。
- 数字/可解析数字字符串进入数值样本，非数值/非有限值作为 null；`zero/break/link` 分开处理。Gauge 无效观测不被 `zero` 策略伪造为零仪表，继续跳过。有限累加溢出显式拒绝；不删除权威源记录或降精度以获得有限场景。空/隐藏后没有可绘制数据的现存图表，全部目录均保留标题/框架和空态；创建时的必需数据约束仍校验。
- 堆积按同图型/同轴分别维护正/负栈；百分比目前以每个分类的绝对值合计为分母，保持正负符号。此为本地明确契约，尚未通过上游/Office 样本验证。普通堆积、组合堆积和次轴堆积都要求线性数值轴。
- Bubble 可省略 size，默认 1；零/null 权重隐藏，负大小拒绝。默认直径按 area/width 和 scale 缩放，显式 marker size 为绝对直径；无有效绘图区隐藏。每组保留同一分类长度的独立 X/Y/size 稀疏向量。
- 数值系列 selector 保留原有源字段语义；名称 selector 按原始系列名或 Bubble 分组字符串精确匹配，大小写、空白及空名称均保留。显示名由 style.name 单独覆盖；全系列 → 字段 → 名称继承不把显示名反写为源身份。未活动名称覆盖保留，活动源名称匹配多系列时拒绝；这是本地可测试规则，P13 的上游重名和隐藏边界仍未核实。
- 图关系跳过空端点与零/null 权重，负权重拒绝，重复边累计；Chord 将端点对归一为无向边；Relation/Sankey 保留方向和 cycle 标记。层级跳过空路径/零/null 权重，负权重拒绝，按完整路径累计；DFS 索引在布局排序前固定，根与所有子节点都可解析样式。
- HLC/OHLC 字段不重复，缺样本保留 null；最高/最低/开盘/收盘关系非法时拒绝。三字段 HLC 直接绘高低线/收盘短线，没有伪造开盘值。
- `createChart/updateChart/readChartData` 接受可选预加载 records；更新先准备候选、转换和全部校验，再修改模型。显式字段/映射 patch 优先于类型转换的默认角色。预加载能力已写契约草稿；生产命令的大源准备/复用还需与根任务计算服务接线一起验证。
- Boxplot、Histogram、Pareto、Waterfall、Trendline、ErrorBars 有共享统计契约。根任务本批实现 `crates/office-compute/src/chart/`、`chart_wasm.rs` 与 Worker 协议/planner/decoder/service；Rust/WASM、TypedArray 复制、取消与数值正确性均未执行验证。主线程统计函数仍服务显式小数据/同步公共 API，选择阈值和端到端性能必须基准后确定，不能宣称全部重操作已迁移。

### 分组统计身份

`ChartStatisticsSeries.seriesPosition` 是 `preparedData.series` 下标，必填；输出 segment 保留同一位置。job 的 `slot` 是作者配置下标，同 kind/slot 可以对不同实际系列生成多个 job：

| 结果集合 | 身份与结果字段 | 消费规则 |
| --- | --- | --- |
| trendlines | `{ configIndex, seriesPosition, statistics }` | `chart-derived` 按配置 + 位置匹配，`chart-layout/chart-draw` 按位置取轴和数据；不能只 find(field)。 |
| errorBars | `{ configIndex, seriesPosition, intervals }` | 同上；每组独立 X 与自定义误差列，保持真实 categoryIndex。 |
| boxplots | `{ field, seriesPosition, statistics? }` | 空/不可用样本仍保留项，statistics undefined 不触发同步重算。 |
| waterfalls | `{ field, seriesPosition, points }` | 按位置消费，不依赖返回数组顺序或 field 唯一。 |

显式提供集合却缺当前配置/位置对应项时拒绝，避免漏组或退回同步统计；空 points/intervals 保留合法匹配。根任务 planner 为实际 job/列/输入/输出设置预算，decoder 逐值/索引验证并可分块让出。绘制 extent 收集不再向 push 展开大批异常点、瀑布或误差/趋势数组；这只是代码修正，没有实测性能结论。

## 5. 配置、转换、放置与宿主历史

[chart-settings-draft.ts](../../packages/office-studio/src/features/charts/chart-settings-draft.ts)、[chart-settings-data.ts](../../packages/office-studio/src/features/charts/chart-settings-data.ts)、[ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) 及 option/series/statistics/collection 控件是共享面板。普通插入、选中图表 Drawer 和 PivotChartPanel 均使用同一字段/类型目录与草稿。没有增加假按钮、空操作或永久禁用图型替代实现。

| P01 状态/配置类别 | 本地持久字段与实际路径 | 差异与未完成项 |
| --- | --- | --- |
| ChartType、Mapping/MappingFields | `type/source/categoryField/valueFields/seriesToCategory/options.mapping`；projection/settings/type-conversion。 | 字段角色已区分；同源转换、跨形状点索引迁移及所有公开 mapping 细项未验收。 |
| Appearance、Title、Subtitle | `options.appearance/titleStyle/subtitleStyle`、title/subtitle；共同字体/背景/边框。 | 专属位置/富文本等公开细项待逐项追踪，不能由一组字体控件宣称完整 section。 |
| Theme、Legend | palette + 统一 theme token；`legend.visible/position/font`。 | 上游命名主题/自动渐变未完全表达；长图例、复合饼 aggregate 图例仍缺。 |
| AllSeriesStyle、Series | `allSeriesStyle/series[field]/namedSeries[{name,style}]`；name/type/color/fillOpacity/border/point/rightAxis/label/dataPoints，nullable 单字段恢复与 6/7 个 clear target。 | Bubble 名称组及 Boxplot 分布已消费同一继承样式，单点按源分类索引逐项继承；两集合可独立清除，clearAllSeriesOverrides 同时清空。SolidSeries fillPattern、过滤/排序后的单点身份迁移及逐原始异常点独立覆盖尚未完成。 |
| Axes、DepthAxis | xAxis/yAxis/rightYAxis/depthAxis；线/刻度/格式/标签/反向、标题字体及对齐，8 个 clear target、三轴标题恢复、partial 次轴与显式 linear 清除 logBase；日期序号坐标/边界、真实指示器。 | 横向 Bar 次轴是顶部数值 X；depthAxis 独立且无 min/max/format。标题字体/对齐在 Canvas、native 与 3D 接入；沿轴定位为本地决定，上游最终布局未核实。月份/年份刻度、完整公开多轴/定位和实际 UI/Office 行为待补；本地次轴 reverse 是扩展。 |
| Trendlines、ErrorBars、ErrorBarFields | 独立作者配置、字段引用、Worker 结果身份；样式/公式/R²、自定义误差。 | 支持 2D 非堆积 Cartesian/Scatter/Bubble/非堆积组合系列；3D/堆积不生成统计图元，保留配置以便换型，不冒称完成未知上游增强。 |
| ThreeD | view 的 rotX/rotY/perspective/rightAngleAxes/heightPercent/depthPercent + grouping/gapDepth。 | 模型、绘制和独立清除已接；公开投影/遮挡/透明排序仍未验收。本地 3D 次轴是增强，上游 P04 不暴露此 API。 |
| Pie、Bar、Area、Radar、Funnel | options 下同名专属配置；C/D/P/E/X。 | 配置实际影响绘制；公开完整 section、组合边界仍待逐项验收。 |
| Aggregation、InvalidValueStrategy、Pareto/ParetoDataPointOptions、Waterfall、Histogram | aggregate/topN、空值策略、Pareto/role/分箱配置。 | 统计本体已接；过滤、聚合、前 N、排序/分箱之后作者点/role 索引语义未完整迁移。 |
| Candlestick、Heatmap | 专属映射、涨跌/空心、颜色范围/分段/单位/标签。 | HLC/OHLC 区分、系列/单点覆盖已接；全部公开轴/图内编辑细项未验证。 |
| Treemap、SunburstFields/HierarchyIndexes/ValueIndex/Style、ChordFields/Mapping/Style、Relation、WordCloud | 层级/边映射、父/中心/点标签、ribbon/节点布局、词云形状/字号/旋转。 | 完整路径和 DFS 样式已接；大布局预算/Worker、公开交互/碰撞与字体实测不足。 |
| GaugeFields/LabelIndex/ValueIndex/Style | `mapping.gauge/gauge`；scale/angles 映射、ticks、ranges、axisLine/axisTick/splitLine/axisLabel/pointer/anchor/progress/title/detail、预设/清除、tile 与标签继承。 | 固定声明的 setter/reset 已表达并消费；本地扁平量程/角度与 null reset 扩展明确记录。源 numeric cell 显示格式、全部 UI/原生/Office 及性能未验收。 |
| AxisPointer | 四个作者字段、`setAxisPointer/clearAxisPointer` 与独立恢复目标、nullable patch、真实悬停线和轴标签、Canvas 位图版本更新及离开/拖拽/滚动/切表/卸载清理。 | 普通坐标/Heatmap 走 ECharts 原生动作，自绘图及 Radar 走独立图元组；缓存内容布局而不在悬停重做数据统计。上游默认颜色/完整 Radar 交互、浏览器效果及性能仍未验收。 |
| AutoGradientFill、PlotArea | 本地尚无完整对应持久/交互契约。 | 继续待实现，不能把背景色或绘图区 rectangle 标为对齐。 |

`chart-type-conversion.ts` 保留源、作者外观和非当前图型配置，按实际数值角色选 X/Y/size、edge、hierarchy、HLC/OHLC 和 gauge 字段；跨百分比变体清除不适用的固定轴边界。堆积/Radar 切换为线性轴。数据角色转换不等于所有 point overrides/作者索引已成功迁移。

`chart-placement.ts` 复用 Canvas 的 SpreadsheetGeometry 行列轴，考虑显式/跨度行高列宽、自动行高、隐藏、折叠、筛选与表格过滤。Builder 的 cell shortcut 保留在 ChartConfig，真正创建时用当前宿主轴解析；持久 ChartModel 保存绝对文档坐标。更新/重开不把绝对位置重新解释为单元格。表格筛选读取仍可能重扫源，尚无大样本/Worker 预算证据。

`ChartConfigBuilder/FChart` 提供类型/字段/映射、系列/轴/统计/options 与布局/绘制 API；`setPosition` 使用上述放置入口。`threeD` 与 `depthAxis` 可独立清除且保持对方。内容更新统一经过 `sheet.command.add-chart/update-chart/remove-chart` 和工作簿历史，内部保存/重开可携带完整模型；新接口草稿覆盖失败原子拒绝、撤销/重做和快照重建，全部未运行。命令级宿主/源表权限、复制/粘贴与所有结构变换需完整验证。

`use-sheet-chart-panels.ts` 捕获 unit/studio/sheet/会话身份；切换、关闭工作簿、更换 studio、scope dispose 时释放会话及 8 个事件监听器，旧异步完成/错误/closed/applied 回调不触碰新会话。源读取失败保留图表模型和错误状态。关闭按活动上下文、其他面板和 focus eligibility 恢复焦点；实际主题、焦点、输入法与浏览器交互未验证。预览目前仍同步调用 `FChart.getData()`，大源读取/透视预览可能阻塞，是明确剩余缺口。

### Gauge 的公开契约、本地差异与未验收项

[chart-gauge.ts](../../packages/office-studio/src/features/charts/chart-gauge.ts) 统一默认值、5 套预设、角度/比例、颜色段、刻度文字与指针 profile；[chart-gauge-api.ts](../../packages/office-studio/src/features/charts/chart-gauge-api.ts) 统一 patch/reset 的校验；[chart-gauge-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-gauge-settings-controls.ts) 编辑同一持久配置。Builder/FChart 与 UI 的更新继续经过共享 `update-chart`、文档历史和内部保存链路，测试草稿未运行。

| 参数/样式 | 已落盘的公开操作及实际消费 | 差异与验证边界 |
| --- | --- | --- |
| 字段与模式 | `setGaugeValueField/setCategoryField`、`setGaugeMode/resetGaugeMode`；公开导出及 FOffice 实例 `.Enum.ChartGaugeMode/ChartGaugePointerShape`。 | Modern/Classic 是固定公开模式；历史 `pointer/progress` 继续兼容。Builder 的 `setValueFields` 要求恰好 1 个非负安全整数，并同步已有显式 Gauge mapping；隐式单字段源仍用字段标题作名称。 |
| 量程/角度/分段 | `setGaugeScale/resetGaugeScale`、`setGaugeTicks`、`setGaugeAngles/resetGaugeAngles`、`setGaugeRanges/resetGaugeRanges`；默认 0–100、225° 到 -45°、主分段 5。 | 上游 `scale/angles` 嵌套，本项目沿用既有扁平 `min/max/splitNumber/startAngle/endAngle`，公开 setters 映射至这些字段。局部负量程更新先与当前配置合并，再校验；无序/重叠/越界颜色段拒绝。 |
| axisLine/axisTick/splitLine/axisLabel | 4 组 setter/reset，显示、独立宽度/长度、次分段、颜色、文字距离/格式及字体；布局与两绘制适配器消费。 | 线宽/刻度长度为绝对尺寸，指针长度及 title/detail offset 为 0–1 比例；小 tile 的环宽/刻度长度按半径约束，文字距离跨中心时停在中心。几何与最终定位仍待统一验证。 |
| pointer/anchor/progress | 3 组 setter/reset，指针长度/宽度/3 形状、锚点尺寸/颜色/边框、独立进度弧宽/圆端。 | 锚点在指针下方；零指针长度/宽度、零锚点和零进度宽度/值/跨度隐藏对应图元。实际 native 图形已写契约草稿，未执行。 |
| title/detail | 2 组 setter/reset，各自显示、纵向比例、字体；detail 数字格式、作者系列名标签及原值文字。 | 清除单字段恢复系列标签/图表字体继承；历史 `gauge.label` 继续兼容。当前只保留 numeric-string 原文；**源 numeric cell 的格式/显示值元数据没有进入 ChartRecords，不能宣称保留源单元格格式**。 |
| 预设与清除 | 本地 `setGaugePreset`、5 个 `antdv-next` Button、恢复默认 helper；替换全部已知 Gauge 作者样式，保留 mapping/源和其他图表配置。 | 5 套来自 P09 的样式声明，尚未核实上游 UI 的 5 个入口。颜色/字体按主题继承；“恢复默认”必须写 `options.gauge = undefined`，草稿克隆及 settings patch 保留它，省略组只表示保留。 |

所有 Gauge 样式 setter 接受本地 nullable 单字段 patch，将 null 归一为显式 undefined；这是对上游部分 `DeepNullish` setter 的一致扩展。整组 reset 恢复有效模式默认值；null/custom prototype/自有危险属性/稀疏 ranges 等非法对象拒绝，失败不写候选模型或历史。设置字段默认与真实 projection 一致：隐式单字段用 valueIndex 0，显式 mapping 缺 labelIndex 表示没有分类标签；编辑数值字段同步 `valueFields`。源范围草稿改变时面板传入实际 fieldCount，避免使用旧范围推断默认。

有限观测紧凑排列 tile，保留源 `categoryIndex` 供点颜色/命中/图例使用；无效项不参与 tile/legend，全部无效使用共享空态。原始值与文字保留，指针/进度比例及 native data value 单独钳制。225°→-45° 为顺时针 270°、反向为 90°，整圈与零弧分开处理；极端有限量程差溢出时仍由共享 graphic 路径消费，不假称 native Gauge 可计算该域。

`echarts-special-option.ts` 每个真实 tile 可生成 `:gauge:<源 categoryIndex>:axis/progress/foreground` 3 个 native series：axis 的 z=1，progress 的 z=2，刻度/锚点/指针/文字前景的 z=3。轴与进度层 silent 且关闭 tooltip，图例仅列真实观测名称。progress `overlap:false` 避免 native 随值生成高 z2；其隐藏轴宽设为实际 progress.width，使环宽独立于背景轴宽；前景保留真实 axisLine.width 供刻度几何使用。刻度文字用 native 自适应方向对齐，Canvas 消费同一阈值。无效观测不创建 native series；零宽/零跨度轴与进度层不创建。此处是绘制适配，不改变 Gauge 数据的单值语义；native layer ID/数量是本批接口变化，消费方不能将每个 series 当作新观测。

既有 `chart-data-layout.contract.test.mjs` 增加真实 ECharts GaugeChart + SVGRenderer SSR 的绘制契约草稿，读取 native display list/图形形状/文字/指针变换；SSR 只用于未来测试，产品内容仍走 Canvas。Context7 查询请求 ECharts 6.1.0 的 axisLabel/SSR API，返回官方文档的当前 master；该动态参考只用于 API 使用说明，固定几何依据仍为 R01 的本地 6.1.0 代码。**本批尚未执行 SSR、测试、主题/焦点/IME、两后端、Office 样本和性能/内存测量；Gauge 未验收，Docs 宿主及全部 Pro 范围仍保留。**

### Radar 的数值标签、轴与公开操作

[chart-radar.ts](../../packages/office-studio/src/features/charts/chart-radar.ts) 统一范围、角度、半径、刻度和数值标签继承。两绘制适配器使用同一 signed 线性 min/max 与五个径向刻度，支持固定负范围、分类反向和径向反向；缺样本保留源分类索引，有限极端范围使用稳定比例并走共享 graphic。数值标签的字体依次继承图表 appearance、历史 `yAxis.label` 和 `radar.axisLabel` 中已定义的字段，缺省格式继承 `yAxis.format`；显式 undefined 不遮掉继承字体。指标名称、旋转与字体使用 `xAxis.label`，与数值标签分开。

`ChartConfigBuilder/FChart` 已落盘 `setShape/resetShape`、`setFill/resetFill`、`setRadarAxisLabel/resetRadarAxisLabel`，以及本地 `setRadar` 配置入口；FOffice `.Enum.ChartRadarShape` 使用 `ChartRadarShapes`。resetShape/resetFill 清除对应作者覆盖，使有效默认恢复 Polygon/仅轮廓，resetRadarAxisLabel 清除专属外观后恢复有效继承。专属 patch 使用 `ChartNullablePatch`，null 单字段恢复归一为显式 undefined，并经普通对象/枚举/有限值校验与候选更新；错误图型、log 径向轴和非法 patch 拒绝，命令失败不发布模型/历史。对象隔离、历史及内部重开只有未运行草稿。

[chart-radar-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-radar-settings-controls.ts) 的 3 个基础控件按 P11 排列；有效数值标签 visible 为真时才显示格式与 5 个字体控件，隐藏时保留作者格式。指标轴和径向轴面板只呈现对应能力，径向轴不提供 log selector。控件继续由共享 `antdv-next` Form/Input/InputNumber/Select/Switch/ColorPicker 承载，没有新原生控件。**本地数值标签缺省为 false，并兼容历史 `yAxis.label.visible`；P11 未核实的 Univer 实际缺省仍保留待核实，不能据此声称默认等价。** 泛用轴样式、指标旋转及径向 reverse 是本地轴契约扩展，固定 Radar Builder 本身只暴露上述 6 个专属操作。

Native Radar 已接统一范围、分类方向、半径、圆形/多边形网和实际刻度格式/字体；分类旋转、径向反向、单点覆盖/值标签或跨度溢出的有限范围切到共享 graphic。真实 native display list 的文本/字体/形状草稿和 graphic 几何草稿均未执行，主题/焦点/IME、两后端、公开演示和 Office 往返保持未验收。

### Boxplot 的原始分布、样式和对数轴边界

[chart-boxplot.ts](../../packages/office-studio/src/features/charts/chart-boxplot.ts) 以原始 `field + seriesPosition + sourceName` 选择分布样式、主/次数值轴及几何；每组数据继续输入原始样本，由共享统计契约提供五数和异常点。boxes 保留原始 sourceName，作者 name 仅改显示标签；不能把布局下标当成作者字段，也不把五数行当成原始样本。每个分布的公共 dataPoints override 使用分布位置；异常点消费该分布的 point shape/size/填充/边框样式。SolidSeries 的完整 fillPattern 和每个原始异常点独立覆盖仍是缺口。

Canvas 箱体、median、须线、完整箱宽 cap 与异常点均消费同一分布样式；填充透明度不再衰减独立边框。矩形异常点单独绘填充及边线，以保留边框 opacity/dash。主/次轴 extent、布局的 median anchor 和数据标签使用同一比例，空分布保留位置，不使后续分布身份偏移。带数据标签时使用共享 graphic，因为 R02 的 native Boxplot View 没有标签绘制路径；native 五数及异常点仍消费相同字段、轴和透明度。

log 轴下，非正或非有限五数不能绘制完整箱体/须线/median/标签；原统计保持原值，**不从正样本重新计算四分位**。合法正异常点可独立显示并参与该轴范围，非法异常点省略。Native adapter 将不可绘制汇总映射为空数据项，保留分类位置；这与原始统计、界面源数据的保留分开。主/次轴、log 跨零汇总/正异常点、独立透明度与 native 端点的草稿均未运行。本地次轴 reverse 是对 P12 `rightY` 不含 reverse 的扩展。

### Cartesian 继承的 nullable 更新与恢复目标

[chart-cartesian-api.ts](../../packages/office-studio/src/features/charts/chart-cartesian-api.ts) 提供 `ChartSeriesClearTargets` 7 项、`ChartAllSeriesStyleTargets` 6 项和 `ChartAxisTargets` 8 项；Builder/FChart 与 FOffice Enum 接同一常量。`clearSeries/clearAllSeriesOverrides/clearAllSeriesStyle`、`clearXAxis/clearYAxis/clearRightYAxis` 及三轴 title setter/clear 已落盘；clear 单目标保留其他作者字段，全组恢复移除对应覆盖。系列 nullable patch 复制输入后将 null 字段归一为 undefined；轴 setter 保持 P12 的 partial 字段契约，通过显式 undefined 恢复字段。后续合并保留该恢复意图，不再从 getter 的 JSON 克隆丢失它。

次轴 partial patch 使用有效 linear 默认；显式改为 linear 同时清除旧 logBase，防止继承的 logBase 令候选校验拒绝。所有入口继续走共享 `update-chart`、内容历史与内部快照；新增 Radar/Cartesian 4 个与 Boxplot 1 个 Facade 草稿覆盖恢复目标、部分次轴、继承、非法输入原子拒绝、撤销/重做和重开，未执行。名称 selector 已按下一节落盘，但本地数值 selector 仍表示源字段，不能把它等同上游内部 index 的全部转换契约。轴 titleStyle/titleAlignment、useDateAxis/AxisPointer 的当前代码与边界见下一节；SolidSeries fillPattern 继续待开发。

### P12 轴标题、日期刻度与真实指示器

[chart-axis-title.ts](../../packages/office-studio/src/features/charts/chart-axis-title.ts) 合并图表字体与轴标题的已定义字段，nullable 恢复继续继承。Canvas 标题按逻辑轴起点/中心/终点定位，反向轴同时翻转位置及文字锚点；native 用 `nameLocation/nameTextStyle`，3D 投影消费同一标题字段，边距随有效字号扩展。主/副标题也继承有效字号并预留文字区域。`$c` 只证实三个选择值和字段写入，`AxisAlignEnum` 只提供枚举，`axis-title.operator.d.ts::applyAxisTitleStyle` 没有运行时实现；**不能声称上游沿轴定位已核实**。原生 DrawingML 段落对齐与沿轴起止位置并不等价，Office 来源/诊断及未核实项见 [SHEET-29](SHEET-29-native-charts.md)。

[chart-date-axis.ts](../../packages/office-studio/src/features/charts/chart-date-axis.ts) 使用现有 Sheets 日期系统：支持 1900 序号 59/60/61、1904、有限合法序号与严格 ISO 日期/时间；真实日期的 UTC offset 跨 1900 边界跳过虚构日，显式 1900-02-29 保留序号算术。缺失日期省略对应点；非法文本、非法边界、无日期/数值配对及 log 日期轴报错。`categoryValues/categoryDateValues/date1904` 为派生输入血缘，模型只保存 `mapping.useDateAxis` 和轴作者设置；时间位置按实际序号间距、min/max 和 reverse 计算，不改为序号分类间距。月份/年份刻度单位和上游重复日期行为尚未核实。

候选/API 与表单共用 `prepareChartDateSource`。Bubble 的 ISO 日期在候选校验时转成 X 序号；Boxplot 日期分类是最终系列显示名，预检按名称及有限样本配对，不在表单计算四分位；Histogram 先检查边界，最终时间坐标来自箱边界/中心，随后由 derived 校验。Pareto 保留原始序号到最终权重顺序，同名分类合并出现不同日期身份时拒绝。开启日期轴的 UI 会把已有 log X 改为 linear 并清除 logBase，这是本地转换决定；公开 API 仍对不合法组合报错。源工作簿 `date1904` 已从 `ExcelWorkbench` 接到 Insert/Drawer/Fields，不从展示字符串猜测日期系统。

[chart-axis-pointer.ts](../../packages/office-studio/src/features/charts/chart-axis-pointer.ts) 的 capability 按 `qi/Hi` 的 IndicatorLine 位判断：column/bar/line/area/combination/scatter/bubble/radar/waterfall/heatmap/boxplot/pareto 可用；3D、Candlestick、Histogram 以及无坐标图不显示该面板。四个字段及四个恢复目标与 P12 类型一致；nullable 清除保留其他作者字段，完整清除恢复主题默认。普通/日期/反向/log/主次轴、Heatmap 行列及 Radar 几何生成指示线和坐标标签，标签约束在图表 frame。Radar 径向线与值圈是本项目独立实现，**上游具体交互未核实**。

`echarts-renderer.ts::updatePointer` 在 native 路径转发 `updateAxisPointer(mousemove/leave)`；graphic 路径仅 replace/remove 独立 group，保留内容图元。两个 pointer 配置位置同步作者线色、线型、底色和文字色；专属 tooltip 同样保持离屏。`ChartOverlay.vue` 的 `antdv-next` Button 接 pointermove/leave/blur，坐标只处理 CSS、zoom、scroll 与对象原点，DPR 由 raster surface 处理；同一 RAF 合并内容/指示器，纯悬停复用内容命令和布局。指示器错误保留到成功悬停或切表，不被下一内容帧清掉；pointercancel 丢弃移动/缩放草稿。清理与缓存是代码实现，**没有真实性能/内存数据，不能声称已优化或达标**。主题、焦点、IME、native/graphic 悬停与浏览器/Office 重开均未执行。

### Bubble mapping 与名称覆盖

`ChartBubbleMappingOptions` 明确 X/Y 的必选数值字段及可选 size/label/series 字段，Builder/FChart 的 setBubbleMapping/clearBubbleMapping 更新真实 mapping；clear 保留其他 mapping 和既有 valueFields，随后恢复共享缺省角色。错误图型、缺失必选字段或越界映射通过候选校验拒绝，不写空入口。

[chart-series-selector.ts](../../packages/office-studio/src/features/charts/chart-series-selector.ts) 与 [chart-cartesian-api.ts](../../packages/office-studio/src/features/charts/chart-cartesian-api.ts) 实现 `ChartSeriesSelector = number | { name: string }` 的校验、独立读写与清除。`namedSeries` 使用 `{name, style}` 的密集数组，显式配置替换整个数组，未提供则保留；不把任意名称用作对象键。重复作者名称拒绝；活动源的同一名称匹配多个系列时拒绝名称覆盖，数值字段路径仍可使用。没有匹配的名称保留为未活动配置，以便隐藏/消失的分组重现后恢复。大小写、空白、空字符串及 prototype 字样都是源名值；style.name 只改显示，不改变目标。这些重名、未活动与空名称规则属于本地实现，P13 未核实的上游边界保持待核实。

共享样式按全系列 → 数值字段 → 名称逐层继承，border/point/label 及按分类索引的 dataPoints 保留下层未覆盖字段；nullable 恢复后重新继承。Canvas 与 native 适配器的颜色、标记、标签、图例、主/次轴和组合系列图型均消费同一源身份。Boxplot 合成显示数据另保留 sourceName，避免从作者标签反推选择器。主次轴边距与 graphic 路径按活动系列的有效样式判断，未活动次轴配置不额外挤占绘图区；Worker 统计能力判断同步传源名，preparedData/数值 bulk 协议未增加字段。

[ChartSeriesSettings.vue](../../packages/office-studio/src/features/charts/views/ChartSeriesSettings.vue) 通过 `antdv-next` Select 同时列字段目标和每个原始名称目标，保留未活动名称的编辑/恢复入口；未配置的重名目标不可选，已有重名覆盖可进入整组恢复，样式和单点编辑停用，面板同时显示绑定错误。名称级 Bubble 单点选择限定到该组的源行，数值字段目标仍覆盖该字段全部分组。系列控制、单点控制和恢复均写作者配置；尚未执行 DOM、焦点、主题或 IME 验证。趋势/误差线的作者 selector 仍为本地 seriesIndex 字段，名称仅参与类型/能力判断，不冒称已经支持 P12 内部统计名称 selector。

本批新增 10 个独立 selector/继承/投影/布局/UI/真实 native scatter SSR 契约草稿，另有 3 个 Facade 草稿覆盖 mapping setter/clear、输入隔离、7 个名称恢复目标、原子拒绝、源分组消失/重现、撤销/重做和内部保存重开；全部未执行。Native renderer 的 series.name 用于 tooltip、legend 与更新，重复显示名的 legend 映射可为多对多，参见 2026-11-24 通过 Context7 查阅的 [ECharts series name](https://github.com/apache/echarts-doc/blob/master/en/option-gl/partial/series-name.md) 与 [legend matching rules](https://github.com/apache/echarts-doc/blob/master/en/option/component/legend.md)；动态文档不是固定 Univer 运行时证据。本地独立源名称不依赖此显示名映射；完整图内 legend command/事件闭环及 Office 动态分组行为仍未验收。

## 6. Canvas 与 UI 技术路径

- `echarts-renderer.ts` 使用锁定的 ECharts 6.1.0，tree-shakeable 注册 Bar/Line/Pie/Radar/Scatter/Funnel/Gauge/Graph/Sankey/Heatmap/Boxplot/Candlestick/Treemap/Sunburst、AxisPointer 等 components 和 CanvasRenderer。领域类型、命令/历史、数据源与 Office 模型由本项目控制。
- 3D、复合饼、词云、Chord、Waterfall、Sankey、Relation 使用共享 Canvas graphic；HLC Candlestick、极端 Gauge 范围、有趋势/误差的图表同样走共享图元；Radar 的单点覆盖/值标签、分类旋转、径向反向或极端有限范围，以及 Boxplot 的 median 数据标签使用该路径。布局/命中用同源公开 ChartLayout，不把“同名 ECharts series 存在”当作图型等价证明。
- `ChartOverlay.vue` 按 sourceVersion 合成 ECharts Canvas 位图至工作表 Canvas；宿主可以 WebGPU 或 Canvas2D。**图表位图由 ECharts Canvas2D 栅格化，WebGPU 合成不表示图内已原生 WebGPU 绘制。** 直接 DrawCommand 路径可交同一 Canvas/WebGPU backend，完整两后端行为仍未验收。
- 本次涉及的 Insert/Dialog、Drawer、Pivot 面板及共享字段/系列/统计控件使用 `antdv-next` 的 Form/FormItem、Input、InputNumber、Select、Switch、Button、ColorPicker、Collapse、Alert 等；没有新增原生表单或其他 UI 库。ConfigProvider/theme token 由应用统一管理；静态控件接入不替代主题切换、弹层尺寸、禁用/悬浮/焦点和 IME 验证。
- ECharts tooltip、legend、brush/dataZoom 等图内事件没有完整路由为编辑器 command；元素选中、对象菜单、Pivot 字段按钮及 Docs 宿主仍需独立实现，不把栅格输出当作编辑闭环。

ECharts 技术参考：[6.1.0 release](https://github.com/apache/echarts/releases/tag/6.1.0)、[模块导入](https://echarts.apache.org/handbook/en/basics/import/)、[Canvas/SVG 指南](https://echarts.apache.org/handbook/en/best-practices/canvas-vs-svg/)。Apache-2.0 依赖选型属于本项目渲染实现，不能作为 Univer 内部实现证据。

## 7. Office 互操作与剩余闭环

本地内部保存 JSON、图片导出、原生可编辑 XLSX/DOCX 图表是三类独立交付。JSON/位图不会替代原生图表验收。

2026-11-24 与根任务协调的静态状态：共享 `office-io-common/src/chart_kind.rs/chart_options.rs/chart_xml.rs` 的 taxonomy/options/XML、`office-writer/src/xlsx_chart.rs/xlsx_chart_style.rs/xlsx_chart_bindings.rs` 的 traditional native writer/样式/绑定、`office-parser/src/native_chart.rs/pivot_sidecar.rs` 的 reader 与 `parser/workbookCharts.ts` 归一化适配均已落盘，仍未执行验证。原生类型/字段读写矩阵和不足项集中见 [SHEET-29 原生图表记录](SHEET-29-native-charts.md)；不能继续把新增 writer/parser 描述为尚不存在，亦不能把模型/sidecar 保留称为原生等价。

本批 source helper 已补当前手动/零尺寸/filter/outline 可见性、字段角色空图、Y zero 与 custom error 分开、默认 Bubble size 的动态公式。table filter 存储的 filteredRows 可能陈旧，export preparation 的唯一 evaluator 接线仍待补；后续在 Office 中改变可见性或新增分组等情形仍需 workbench refresh。这些静态绑定不是全部 native 动态行为等价的证明。当前没有本批编译、XLSX/DOCX 样本导出或独立 Office 重开结果。DOC-19 还缺文档图表模型/数据源宿主、DOCX drawing/chart part/关系与嵌入数据闭环。

继续保留的实质工作包括：

1. 固定版所有动态/图表工作表/浮动/上下文入口、参数、权限、历史和持久化逐入口闭环；Docs 宿主与 Pivot 字段/筛选/轴项操作。
2. 名称 selector 的公开演示/重名/隐藏边界与真实 UI/Office 验收、统计名称 selector；类型转换、过滤/聚合/Pareto/分箱后的 dataPoints 和 Waterfall role 索引迁移。
3. Radar 上游数值标签默认值、轴扩展和径向指示器效果；Boxplot 的 SolidSeries fillPattern/逐原始异常点独立覆盖；标题对齐的上游沿轴/文字盒语义、日期月份/年份单位及重复日期；Gauge 源显示格式/最终交互；这些类型与指示器的极端/native/Office 验证，以及 AutoGradientFill/PlotArea 等未映射配置。
4. 长图例/aggregate 图例、复杂图布局/碰撞/字体、完整 3D 相机/遮挡/透明排序及公开效果。
5. 源准备、预览、放置过滤和复杂布局的有界大计算/Worker 路径；实际进度、预算/取消、传输字节与内存峰值。
6. 全类型可编辑 XLSX/DOCX 写出/读回、未知信息/不等价原生图型的源/配置保留与诊断；独立 Office 重开和再编辑。
7. 相同数据集的首屏/滚动/编辑/缩放、WASM/Worker 任务、长任务、帧时间、JS/WASM/GPU 内存与前后基准。

这些是继续实施的缺口，不是最终范围缩减。

## 8. 测试草稿与验证状态

| 草稿/已有夹具 | 本轮新增或涉及的覆盖 | 状态 |
| --- | --- | --- |
| `packages/office-studio/tests/chart-data-layout.contract.test.mjs` | 45 项空态、原始 null/精度、双方向可见数据缺省/Switch 重投影、透视角色、DFS/样式、转换/原子拒绝、Bubble marker、分组统计身份、Bar 次轴、堆积/大权重、Radar 单点/标签、14 万异常点/瀑布、3D 深度/横向标签；Gauge 默认/5 预设、有效/无效 tile、源索引/原文/clamp、极端角度、零指示器、nullable/非法 patch、字体/格式继承、控件/草稿重置链路、真实 native 弧宽/层级/文字/指针变换的 SSR 草稿。 | 草稿，未执行，包括 native SSR；大数组断言不是性能基准。 |
| `chart-radar-boxplot.contract.test.mjs` | Radar nullable/字体继承、条件显隐、线性 signed/极端范围、源分类索引、分类/径向反向、值标签和 graphic；native Radar 的真实文本/字体/范围/半径/圆网 SSR；Boxplot 原始字段/位置、已知四分位、透明度/完整 cap、native path/端点 SSR、主/次轴 median anchor/标签、log 保留统计与合法正异常点；nullable patch 对象/数组隔离、危险/稀疏对象拒绝。 | 新草稿，未执行，包括 native SSR；不是浏览器交互或性能验收。 |
| `chart-facade.contract.test.mjs` | 类型/option patch、cell placement、depthAxis/view 独立清除；Gauge 公开导出与 FOffice 实例 Enum、单值绑定、全部样式/量程/角度/刻度/ranges setter/reset、5 预设/全组清除、对象隔离/非法输入原子拒绝、更新/撤销/重做/内部保存重建；另新增 Radar/Cartesian 4 个、Boxplot 1 个及 Bubble/名称 3 个草稿，覆盖 6/7/8 个 clear target、轴标题/partial 次轴/linear 清除 logBase、原始箱线样本/主次轴、名称独立覆盖/消失重现/7 个恢复目标/两集合清除、mapping/错误输入/历史/内部重开。 | 本轮修改未执行。 |
| `chart-series-selector.contract.test.mjs` | 10 个名称精确匹配/空白/大小写/空字符串/危险字样、字段与名称继承/逐点边框/nullable、集合替换/未活动保留、非法选择器/稀疏或重名配置/源歧义拒绝；Bubble Canvas 与真实 native scatter 名称/标记/标签/图例/主次轴 SSR；有效次轴边距；设置目标/所属组点范围/继承读写/重名覆盖恢复；Boxplot sourceName/标签/次轴；组合类型/统计能力及轴换型规则；专属色阶不覆盖名称颜色及 Gauge 字体的名称继承/恢复。 | 新草稿，未执行，包括 native SSR；不是浏览器交互、Office 往返或性能验收。 |
| `chart-axis-interaction.contract.test.mjs` | 15 个 P12 日期虚构闰日/1904/ISO 时区、真实间距/reverse/缺失配对/Pareto 血缘/Bubble 候选原子拒绝、标题继承/边距/旋转/native/3D、四指示器参数/缓存查询/主次轴/log/日期/横向/Heatmap/Radar/清理、控件顺序与日期预检、CSS 坐标契约草稿。 | 新草稿，未执行；不是性能或最终 UI 验收。 |
| `chart-facade.contract.test.mjs` 新增 P12 场景 | 3 个标题字体/对齐/nullable/日期保存重开、4 个指示器恢复目标/独立 reset/历史、非法日期/标题/指示器候选拒绝且不发布历史的集成草稿。 | 未执行。 |
| `echarts-renderer.browser.test.mjs`、`tests/browser/chart-workbench.spec.ts` 新增 P12 场景 | 2 个 native 动作/graphic group replace-remove、sourceVersion/缓存/作者模型保留草稿；1 个真实 Overlay 在 150% zoom/scroll 的悬停、离开及 pointercancel 草稿。缓存用重新生成的 option，避免克隆 formatter 闭包。 | 浏览器草稿均未执行；不包含截图或像素验收。 |
| `sheet-chart-panels.contract.test.mjs` | pending 插入/更新换 unit/sheet、旧 Pivot 回调隔离、关闭/更换 studio/scope 监听释放、错误保留、焦点 eligibility、源读取失败。 | 6 项草稿，未执行；不是 DOM/IME 浏览器验证。 |
| `pivot-chart-panel.contract.test.mjs` | 透视面板会话/状态与共享目录、源配置改动。 | 本轮修改未执行。 |
| `chart-statistics-compute.contract.test.mjs`、`chart-statistics-service.contract.test.mjs`、`crates/office-compute/tests/wasm_chart.rs` 及 chart 原生测试 | 根任务的协议/planner/decoder、分组身份、服务生命周期、预算/让出、数值与真实浏览器绑定草稿。 | 根任务开发中，未执行；不得由 TS 夹具推导 Rust/WASM 已通过。 |
| 既有 `chart-engine-rendering/echarts-option/echarts-renderer/chart-workbench`、Office writer 夹具 | 历史七类/有限字段切片以及已有生命周期用例。 | 旧通过结果只属于历史版本，不覆盖本批完整目录/配置/协议改动。 |

本轮唯一完成的核对是静态源码、接口、fixture 和 diff 阅读。**测试、类型、前端构建、Rust/WASM、浏览器、明暗主题、窄视口、焦点/IME、两渲染后端、真实 Office 重开、性能/内存全部未执行、未验收。** 统一验证阶段必须先生成与协议 2 匹配的 WASM 产物，再运行相应检查和真实样本/交互闭环；未验证项目持续留在本记录，不复用历史“通过”字样。
