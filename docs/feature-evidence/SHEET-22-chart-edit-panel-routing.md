# SHEET-22：图表编辑面板 route 与真实表单定位

状态：**P17 前端接线和测试草稿已落盘，未验收**。本批继续 P16 的 `sheet.command.chart.open-focused-panel`，将 `tab/section/axisId/seriesId/titleKey` 接入既有 Ant Drawer。没有新增 Ribbon 根；分组枚举、参数数目和测试草稿数均不构成功能完成证明。

P18 后续：`lineAndArea` 已接入 Line / Smooth / Step 的真实 Select、二维 / 3D 几何及公开写方法，详见 [线型独立证据](SHEET-22-line-and-area-styles.md)。本文件保留 P17 的历史范围和 33 个草稿计数；P18 新增 34 个草稿独立记录，均未运行。

## 固定版来源和分类口径

固定对照为 Univer **v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**，基础仓库在 `E:\rust_test\univer`。本批 Pro 证据使用项目内固定包 ES / 类型快照，没有安装或运行商业 SDK。

| 证据 | 已查阅的可观察契约 |
| --- | --- |
| [focused-chart.command.d.ts](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/commands/commands/focused-chart.command.d.ts)；[sheets-chart-ui ES](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/es/index.js) 的 `Ri`（约字符 20119） | open-focused 参数是 `route?: IChartEditPanelRoute`；通过 `Ii/getChartContextMenuTarget` 捕获 drawing，转发宿主 ID 和 route |
| [open-chart-panel-operation.d.ts](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/commands/operations/open-chart-panel-operation.d.ts)；[ChartEditPanel.d.ts](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/views/chart-edit-panel/ChartEditPanel.d.ts) | 内部打开操作携带 unit/subUnit/drawing/route；面板 key 是 `sheet.chart.edit.panel` |
| [float-menu/types.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/types.d.ts)；[chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 枚举（约字符 90646） | 2 个 tab、27 个 section、3 个 axisId、5 个 titleKey；11 类上下文 scope，mark panel / direct-edit 分支 |
| [chart-edit-panel-route.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/routing/chart-edit-panel-route.d.ts)；同一 ES 的 `gs`（约字符 92575）和 `vs/resolveChartEditPanelRoute`（约字符 93052） | axis/gridLine/tick 都转到对应轴分组；mark 有 direct-edit 优先级；dataPoint/dataLabel 无 seriesId 时不产生 route |
| [ChartEditorTabs.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/components/ChartEditorTabs.d.ts)；[StyleTabPanel.d.ts](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/views/chart-edit-panel/style-tab-panel/StyleTabPanel.d.ts) | controlled activeTab、tab list / panel、tab change 和 ariaLabel；样式与数据分别有真实内容区 |
| [chart-capabilities.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/sections/chart-capabilities.d.ts)；同一 ES 的 `Ui` / `Ji` | 轴、通用系列、指示线、图例能力；Line/Area/Scatter/Combination 的 lineAndArea 能力 |
| [ChartAxisSections.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/sections/ChartAxisSections.d.ts)；[ChartLineAndAreaSection.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/sections/ChartLineAndAreaSection.d.ts)；同一 ES 的 `ic` | 轴线、标签、边界、刻度和网格线分组；lineAndArea 包含 `area.lineStyle=line/smooth/step` 和 `invalidValueStrategy=break/link/zero`，Scatter 不显示 lineStyle 控件 |

分类计数：**复用公开上下文命令根 1；新增常驻 Ribbon 根 0、插件扩展根 0、溢出根 0、独立上下文菜单根 0**。现有“更多”菜单中的默认编辑仍调用同一根。route 入口分类为 **2 标签页、27 分组（通用 7、图型专属 16、网格/指示/统计 4）、3 轴参数、5 标题参数、11 上下文 scope**。这些分类互有引用，不能相加当功能数量。

`IChartPanelRouteState={chartId,route|null}`、`IChartSelectedElementState={chartId,element,route,target}` 属于上游交互状态。本地仅保存 Drawer 导航和已有图表选中状态；没有把 route、导航 key 或 UI session 写入图表作者模型，也没有由此实现所有 element 的直接命中菜单。

## 27 个分组到本地真实编辑区

以下按公开枚举声明顺序列出。它不是上游各图型实际同时显示的分组顺序；本地复用现有折叠表单，不复刻上游像素布局。标签列是本地实际标签。分组适用性结合已有图型、轴、图例、统计和 Waterfall 判断，错误目标明确失败。

| route section | 本地标签 / 真正展开位置 | 目标、能力及内容提交 |
| --- | --- | --- |
| chartStyle | 标题、图例与外观；字体与边框 → 图表字体 / 图表边框 | 聚焦真实背景控件；沿用 appearance 作者 |
| chartAndAxisTitles | 标题、图例与外观 / 标题字体 / 副标题字体；或指定坐标轴 | title/subtitle 或指定 axis title；轴标题要求该轴存在 |
| legend | 标题、图例与外观；字体与边框 → 图例字体 | 聚焦显示图例；复用图例能力判断 |
| horizontalAxis | 坐标轴 → X 轴 / 分类轴 | 推导 xAxis；聚焦其轴标题控件 |
| verticalAxis | 坐标轴 → Y 轴 / 数值轴 | 推导 yAxis；聚焦其轴标题控件 |
| rightVerticalAxis | 坐标轴 → 右侧数值轴 / 顶部数值轴 | 仅实际拥有次轴的图表；不会为了导航新增右轴 |
| series | 系列、标记与数据标签 → 所有系列 / 单独设置系列 | 有 seriesId 时选定既有字段 / 源名称作用域；没有时打开共享样式 |
| pieStyle | 图型专属设置 → 饼图实际控件 | 饼 / 环 / 复合饼家族；复用已有 pie 作者 |
| paretoBarSeries | 图型专属设置 → 帕累托柱形样式 | 展开 barStyle；不修改数据或样式 |
| paretoLineSeries | 图型专属设置 → 帕累托累计线样式 | 展开 cumulativeLineStyle；真实控件写入对应作者 |
| waterfallSeries | 图型专属设置 → 瀑布图角色样式作用系列 | 选定来源系列；保留现有角色、点和全局合计语义 |
| relation | 图型专属设置 → 关系图实际控件 | 关系图布局 / 力参数等既有作者 |
| wordCloud | 图型专属设置 → 词云实际控件 | 字体、布局、图片资源沿用原会话 / loading 处理 |
| radar | 图型专属设置 → 雷达实际控件 | 雷达网格 / 填充 / 刻度；轴设置仍复用专门轴表单 |
| funnel | 图型专属设置 → 漏斗实际控件 | 间距、排序、绝对值等既有作者 |
| heatmap | 图型专属设置 → 热力图实际控件 | 色阶与标签等既有作者 |
| lineAndArea | P18：Line/Area/Combination 同时打开标题、图例与外观、图型专属设置和所有系列；Scatter 只打开空值处理 | 11 个 Line/Area/Combination 图型含堆积、百分比及 3D，聚焦真实 `area.lineStyle` Select；Scatter 聚焦 `invalidValueStrategy`。线宽/边框沿用系列控件；运行和上游差异仍见独立线型证据 |
| gridlinesAndTicks | 坐标轴 → 指定轴 | 展开真实网格线 / 刻度表单，默认 yAxis；缺失轴明确失败 |
| indicatorLine | 指示线与坐标标签 | 聚焦真实 axisPointer 控件；复用既有能力 |
| trendline | 趋势线与误差线 → 已有趋势线 | 没有线时聚焦添加按钮；导航不添加线 |
| errorBars | 趋势线与误差线 → 已有误差线 | 没有线时聚焦添加按钮；导航不添加线 |
| candlestick | 图型专属设置 → 蜡烛 / K 线实际编辑器 | 独立蜡烛作者表单 |
| histogram | 图型专属设置 → 直方图样式 | 分箱由原数据准备服务提供；导航不进行新分箱作者提交 |
| treemap | 图型专属设置 → 树图实际编辑器 | 层级、布局和标签等既有作者 |
| sunburst | 图型专属设置 → 旭日图实际编辑器 | 层级、中心和标签等既有作者 |
| gauge | 图型专属设置 → 仪表实际控件 | 刻度、模式、范围、指针和进度等既有作者 |
| chord | 图型专属设置 → 和弦实际编辑器 | 布局、带、标签等既有作者 |

上下文 resolver 按固定 ES 重写：chartArea→chartStyle；title/axisTitle→标题组；axis/gridLine/tick→对应轴组；legend→legend；series/dataPoint/dataLabel→series。mark 先处理 seriesStyle/pieStyle/paretoStyle/waterfallStyle，再处理 7 类 panel，最后按 seriesId 回退；未提供足够目标时返回 undefined。PieSlice/PieAggregate/PieConnector 的 direct-edit 名称本身不擅自推导新 route，沿用固定版 panel / seriesId 回退行为。

## DTO、会话、作者和计算边界

- [chart-edit-panel-route.ts](../../packages/office-studio/src/features/charts/chart-edit-panel-route.ts) 提供纯 DTO 校验、既有能力复用、导航规划、上下文 resolver 和系列 ID 适配。未知字段、非法枚举、过长系列 ID、数据页附样式参数、冲突的轴 / 标题 / 分组、无效系列组合都明确失败。
- `field:2` 对应本项目已有数值字段作用域；`name:"Comparison"` 对应 canonical JSON 的精确源名称作用域，使用 `sourceName ?? name`。不是显示名称、系列顺序或任意 Univer opaque runtime ID 导入。缺失 / 重复名称拒绝，字段作用域可覆盖同字段多组。没有当前准备数据时明确报“尚未就绪”，不会猜第一系列。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的既有 edit 根依次复用目标、版本 / 作者快照、权限、route 校验、当前统计缓存和 UI host；route 不产生内容提交。默认编辑仍可省略 route，同一 Before/Executed 分发机制不变。
- [use-sheet-chart-panels.ts](../../packages/office-studio/src/features/charts/views/use-sheet-chart-panels.ts) 给同一图表保留会话和 Drawer draft，仅更新导航请求。新宿主 / 工作簿 / 工作表、关闭 / dispose 和不允许回焦的状态使旧 `navigationCurrent` 失效。
- [ChartEditorPanel.vue](../../packages/office-studio/src/components/spreadsheet/ChartEditorPanel.vue) 持有作者 draft，Ant Tabs 只切换表单内容。导航按请求应用一次；系列等待 draft source projection 真正准备就绪。输入法组合期间延后路由，保留正在输入的 DOM；新的导航 key 允许重复请求同一路由重新定位子分组，不依靠每次作者输入重新导航。
- [ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) 展开已有 Collapse 并定位实际可见、可用控件；系列须先匹配实际 selected scope。nextTick / rAF 在执行时再检查 active、disabled、composition、宿主和竞争 dialog，销毁时取消。不能因一个隐藏同名控件而回退到错误区域。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 和 [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 已为范围主题及主任务新增的 `cellShiftPending` 加回焦门禁；与单元格输入、当前宿主身份判断共同生效。该门禁仅为接线，真实弹窗 / 输入法仍未验收。
- 真正应用仍通过 `chartSettingsPatch` → 现有 `chart.update` / `sheet.command.update-chart` → 每文档历史、模型保存和重开。route、Tabs、展开状态不保存为作者内容。本批没有新增解析器、计算引擎、Worker 协议或 WASM 内存接口；原重计算 / 数据源服务和 Canvas 渲染继续使用既有路径。

## 实际文件与公开接口

本批涉及 **24 个文件：15 源码、7 测试 / fixture、2 独立证据文件**。共享文件仅修改已协调的图表接线 / 导出和指定回焦局部；中央清单由主任务更新。

源码（除新 route 模块，其余均为既有文件）：

1. [chart-edit-panel-route.ts](../../packages/office-studio/src/features/charts/chart-edit-panel-route.ts)
2. [chart-object-commands.ts](../../packages/office-studio/src/features/charts/chart-object-commands.ts)
3. [use-chart-object-actions.ts](../../packages/office-studio/src/features/charts/views/use-chart-object-actions.ts)
4. [use-sheet-chart-panels.ts](../../packages/office-studio/src/features/charts/views/use-sheet-chart-panels.ts)
5. [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue)
6. [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue)
7. [ChartEditorPanel.vue](../../packages/office-studio/src/components/spreadsheet/ChartEditorPanel.vue)
8. [ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue)
9. [ChartOptionFields.vue](../../packages/office-studio/src/features/charts/views/ChartOptionFields.vue)
10. [ChartSeriesSettings.vue](../../packages/office-studio/src/features/charts/views/ChartSeriesSettings.vue)
11. [ChartWaterfallSettings.vue](../../packages/office-studio/src/features/charts/views/ChartWaterfallSettings.vue)
12. [ChartStatisticsSettings.vue](../../packages/office-studio/src/features/charts/views/ChartStatisticsSettings.vue)
13. [ChartParetoSettings.vue](../../packages/office-studio/src/features/charts/views/ChartParetoSettings.vue)
14. [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)
15. [index.ts](../../packages/office-studio/src/index.ts)

公开变化：`ChartObjectCommandInput.route?: ChartEditPanelRoute`；UI host / Overlay 的 `openPanel/openChartEditor(target,route?)`；Drawer 的 route / routeKey / navigationCurrent，以及子表单导航 key。`src/index.ts` 追加 `CHART_EDIT_PANEL_TABS/SECTIONS/AXES/TITLES`、`normalizeChartEditPanelRoute`、`resolveChartEditPanelRoute`、`chartEditPanelSeriesId` 及 route / tab / section / axis / title / element-panel-target 类型。纯导航 plan / capabilities 保留在 feature 内。无协议版本、Office 作者模型或锁文件变化。

## 草稿、UI 与验收状态

新增 **33 个场景草稿**，按 `test` 入口手工计数，循环参数不另计：

| 文件 | P17 新增场景 | 验证意图 | 运行 |
| --- | ---: | --- | --- |
| [chart-edit-panel-route.contract.test.mjs](../../packages/office-studio/tests/chart-edit-panel-route.contract.test.mjs) | 13 | DTO / 组合拒绝、精确系列身份与歧义、5 标题 / 3 轴、27 分组适用性、实际控件 writer、统计不新增作者、11 scope / mark 优先级、重排后源名称稳定 | 未运行 |
| [chart-edit-panel-routing.contract.test.mjs](../../packages/office-studio/tests/chart-edit-panel-routing.contract.test.mjs) | 5 | 冷缓存拒绝、权限 / veto、旧版本 / 作者 / 宿主、导航不入历史、真实作者 undo/redo/JSON、关闭工作簿 | 未运行 |
| [sheet-chart-panels.contract.test.mjs](../../packages/office-studio/tests/sheet-chart-panels.contract.test.mjs) | 3 | 同图表会话保留、无效 / busy 不更改导航、旧会话 / focus gate 失效 | 未运行 |
| [chart-edit-panel-routing.browser.test.mjs](../../packages/office-studio/tests/chart-edit-panel-routing.browser.test.mjs) | 12 | 真实 Tabs / Drawer、源控件、5 标题 / 3 轴、系列 / 面积编辑、重复路由、草稿保留、统计与 Pareto / Waterfall 真表单、权限拒绝、IME、迟到导航和竞争 Modal / 主题 | 未运行 |

另改写 [P16 命令草稿](../../packages/office-studio/tests/chart-object-commands.contract.test.mjs) 的旧“合法 route 未映射”预期为真实转发 / 无效组合拒绝；纠正 [P16 浏览器草稿](../../packages/office-studio/tests/chart-object-actions.browser.test.mjs) 的实际“图表标题”控件标签。这两项沿用原测试入口，不增加 P16 的 34 场景计数。[真实工作台 fixture](../../packages/office-studio/tests/fixtures/chart-object-actions.html) 添加可选多系列 / 图型、route 参数、真实 prepare 调用和竞争 Ant Modal；未替换 chart 编辑器或计算器。

涉及的可替换 UI 使用 `antdv-next` 的 Drawer / Tabs / Collapse / Form / FormItem / Input / InputNumber / Select / Switch / Button / ColorPicker / Alert，以及现有图表子编辑器。新增 data attributes / div 仅定位已有控件，不是自绘表单或新的原生输入桥。主题由现有统一 ConfigProvider / token 管理；内容保持 Canvas。本批没有新增原生表单控件，但**主题、焦点、IME 和所有浏览器场景仍未实际验证**。

本批仅执行文件 / 固定来源读取、手工静态审阅和代码 / 文档写入。**没有运行测试、类型检查、前端构建、浏览器、dev / Vite 服务、Cargo / WASM、性能基准、服务验收或 `git diff --check` 等自动检查；没有提交、打包发布或改锁文件。** 浏览器草稿使用 Canvas2D，不证明 WebGPU。内部 JSON reopening 草稿不证明 XLSX / Office 文件保真。

## 真实剩余缺口

1. route 接线未等于完整 element 直接编辑：axis / gridLine / tick / mark / dataPoint / dataLabel 的所有 Canvas 命中、浮动菜单、selected-element 服务和直接编辑闭环尚未统一接入 UI。当前 context resolver 是有契约的适配工具，不能冒称 11 scope 全部已有直接操作。
2. 系列 ID 仅复用本项目 field / 精确源名称作用域；任意 Univer runtime opaque seriesId、drawing-series 对照导入仍未实现。
3. P18 后续已实现 `area.lineStyle=line/smooth/step` 的真实 Select、二维折线 / 面积 / 组合系列和 3D 投影几何，并新增 Scatter / 3D / 空值策略的契约与浏览器草稿。上游曲率 / 转折 / Combination 最终行为、真实浏览器和 XLSX / Office 重开仍未验收，不能将这次接线记为完整线型验收；具体差异见 [P18 线型证据](SHEET-22-line-and-area-styles.md)。
4. 完整 chart-sheet 创建 / 导入 / 独立布局、GRID / CHART paste placement、跨 workbook 源表 / pivot / 图片依赖导入、一般 / 分组多 drawing 和外部 Microsoft Office 绘图剪贴板仍是后续**前端**缺口，不能转记为后端自动补齐。
5. 本批 33 个草稿、原相关草稿以及真实浏览器 / 文件 / 服务验证均未运行；各图型真实数据、权限变更、缓存 / 大样本、图片资源、字体 / DPR、主题、IME、取消和历史仍待统一验收。XLSX / OOXML 和 Office 重开复用本项目转换链路补证，不以 JSON 或面板定位代替。
6. 性能与内存未测：首屏、滚动、编辑、缩放、统计 Worker / WASM、跨边界传输、帧时间、长任务与峰值内存保持“未验收”。本批不声称优化或达标。

以上继续属于完整 Sheets / Docs / Pro 对齐范围，没有缩减交付范围。
