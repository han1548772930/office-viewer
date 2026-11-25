# SHEET-22：Canvas 图表元素命中与直接编辑

状态：**P19 前端实现和 34 个测试声明草稿已落盘，未验收**。本批新增实际轴标题、轴线／轴标签、网格线、刻度线、图例、系列、数据点与数据标签的 Canvas 命中和 Ant 浮动编辑。标题／副标题与绘图区沿用既有实现，P17 面板 route、P18 线型也不重复计入本批。

这里只记录本批实际接入的范围；25 个上游 kind、8 类本地编辑目标及 34 个测试声明是不同统计对象，均不证明全部图型或所有 Pro 行为已完成。

## 固定来源与分类口径

基线：**Univer v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**，只读仓库 `E:\rust_test\univer`；Pro 公开 ES／类型快照位于本项目 `.research.local/univer-1.6.0/packages`。未安装、运行或依赖商业 SDK。

| 来源 | 本批核对的事实 |
| --- | --- |
| [engine-chart/types.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/types.d.ts)，约 1238–1345 行 | `IChartElementHit` 是 discriminated union；`ChartElementKind = IChartElementHit['kind']`，并无另一套 `elementType` enum。声明包括 bounds／overlay、axisId、seriesId／seriesIndex、dataIndex、semanticId、pieSourcePointIndex、Waterfall role 和 field button identity；事件有 click／dblclick／pointer 系列 |
| [float-menu/types.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/types.d.ts) | 11 个 config scope；7 个 mark direct-edit kind 和 7 个 mark panel；adapter 携带 unit／subUnit／chart／drawing／target，状态读写与关闭／More 分开 |
| [chart-element-selection.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/chart-element-selection.d.ts)；[chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 的 `hs/resolveChartElementConfigTarget`、`ql/resolveChartElementSelection` | hit → config target → selection／showFloatMenu；ES 的 resolver 分支未覆盖 d.ts 的全部声明 scope，不能以类型枚举推断每类都有运行态浮动菜单 |
| [constants.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/constants.d.ts)；[en-US.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/locale/en-US.d.ts) | 浮动菜单 label rotate 为 0／30／45／60／90，不透明度候选为 1／.9／.7／.5／.3／.2／.1／0；common／floatMenu 资源提供标题、字段、动作、控制与选项标签，不是另一套 element kind 顺序 |
| [create-chart-element-float-menu-adapter.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/create-chart-element-float-menu-adapter.d.ts)；[Sheets controller](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/controllers/chart-element-float-menu.controller.d.ts) | 共享 host adapter，订阅真实图表事件，浮动 session／popup 生命周期，双击打开完整面板，隐藏元素后清理 selection |
| [P17 面板证据](SHEET-22-chart-edit-panel-routing.md) | 复用已落盘的 `sheet.command.chart.open-focused-panel`、route 校验和真实 Drawer；不新增另一套面板作者命令 |
| 本地 ECharts **6.1.0**：`packages/office-studio/node_modules/echarts/lib/component/axis/AxisBuilder.js`、`CartesianAxisView.js`；ZRender **6.1.0**：`node_modules/.pnpm/zrender@6.1.0/node_modules/zrender/lib` | 固定 `anid`：axis `line`、`label_*`、`tick_*`、split `line_*`／`minor_line_*`；实际 model／view、display list、`findHover().topTarget`、attached text host、transform／clip path 与 `shouldBePainted` |

Microsoft Office 的公开元素选择／格式编辑行为是产品语义参照；本批没有运行 Office 或上游在线演示，不将这些交互选择冒称为已观察到的完整上游运行效果。

## 25 个 kind 清点

按 `IChartElementHit` 分类联合的声明顺序，以及各分类中字面量的声明顺序排列。分类数量：**Basic 9、Axis 9、Legend 2、Mark 2、Series 1、Field button 2，合计 25**。axis kind 在 `ChartAxisElementKind` 中先列 title、再 label、再 line；这不是浮动菜单或 Ribbon 的布局顺序。

| 序号 | 上游 kind | 本地入口／现状 |
| ---: | --- | --- |
| 1 | blank | 沿用整图选择／默认编辑；不创建虚构元素表单 |
| 2 | chartArea | 沿用对象编辑、P17 chartStyle route 与既有绘图区入口；chartArea 和 plot-area 不混作同一上游 kind |
| 3 | title | 沿用已有真实标题命中、浮动输入和拖动，不计 P19 增量 |
| 4 | subtitle | 同上，保留独立副标题作者 |
| 5 | gridLine | 新增实际主／次轴网格线命中 → 对应轴 gridLine 表单 |
| 6 | tick | 新增实际主／次轴刻度命中 → 对应轴 tick 表单 |
| 7 | dataLabel | 新增实际标签命中；普通图型写 selected series.label，Pie 家族写 pie.label，并显示作用域说明 |
| 8 | categoryHierarchyExpandControl | 多级分类展开控制未接入，继续待开发 |
| 9 | categoryHierarchyCollapseControl | 多级分类折叠控制未接入，继续待开发 |
| 10 | xAxisTitle | 新增分类／水平轴标题命中与局部编辑 |
| 11 | yAxisTitle | 新增主数值轴标题命中与局部编辑 |
| 12 | rightYAxisTitle | 新增实际次轴标题命中；不为了选择而新增次轴 |
| 13 | xAxisLabel | 新增真实分类／水平轴标签命中，局部 axis 表单 |
| 14 | yAxisLabel | 新增主数值轴标签命中 |
| 15 | rightYAxisLabel | 新增实际次轴标签命中 |
| 16 | xAxisLine | 新增真实分类／水平轴线命中；part=line |
| 17 | yAxisLine | 新增主数值轴线命中 |
| 18 | rightYAxisLine | 新增实际次轴线命中 |
| 19 | legend | 已命中的图例项选择整个图例作者范围；空白容器区域未另建估算命中 |
| 20 | legendItem | 普通点击保留原生 view toggle；双击打开全图例编辑；命中保留实际 native name，选择框覆盖当前可见图例 |
| 21 | mark | 二维 Cartesian／Radar graphic／Pie 真源点／Funnel 已标记；Pareto、Waterfall、复合 Pie aggregate／connector 等专属 mark 尚未统一接入 |
| 22 | dataPoint | 实际源点可直接设置颜色、opacity、border、marker；合成点不伪装源点；native Radar 的单 datum 只选择系列 |
| 23 | series | 实际原生／graphic 系列命中；点上 Shift 或“选择系列”升为同一绑定系列 |
| 24 | fieldButton | Pivot chart field button 真实命中与操作仍待开发 |
| 25 | fieldButtonOverflow | field button overflow 仍待开发 |

本地交互联合为 **8 类**：axisTitle、axis、gridLine、tick、legend、series、dataPoint、dataLabel。axis 带 line／label part，axisId 是语义上的 xAxis／yAxis／rightYAxis。横向 Bar 的 native x/y component 映射到同一语义轴，次轴保留独立身份。

## 本批上下文入口与写入范围

增量分类：**常驻 Ribbon 根 0、插件扩展 Ribbon 根 0、Ribbon 溢出根 0、对象右键新增根 0；新增浮动工具组件 1，动态承载上述 8 类编辑目标**。既有对象右键根、title／plot-area 工具组件与 P17 More 命令不重复计数。

浮动工具的固定布局顺序：标题／关闭 → Ant Form 的本 scope 控件 → 标签作用域说明（适用时）→ 恢复格式、选择系列（仅点／标签）、更多设置、应用修改 → 错误 Alert。共 **4 个常见动作**（关闭、恢复格式、More、Apply）和 **1 个条件动作**（选择系列）。它们不是 5 个独立作者能力，也不是 Ribbon root；classic／grid 两种 Ribbon 共享同一 Canvas 上下文入口。

| 目标／手势 | 控件／局部写入 | 显隐、禁用与命令 |
| --- | --- | --- |
| axisTitle 点击／双击 | title、titleAlignment、titleStyle 字体 | 仅真实可绘标题；标题文本保留，恢复只清格式；Apply → FChart.update → `sheet.command.update-chart` |
| axis line／label 点击 | label／line、reverse、min／max、scale／log／date、format；旋转候选 0／30／45／60／90 | numeric／date 按既有图型能力；不提交无关标题、grid、tick；Delete 只隐藏实际 part |
| gridLine 点击 | visible／color／width／opacity／dash | 仅本轴 gridLine；隐藏后关闭 selection |
| tick 点击 | 上述 line 样式＋length／position | 仅本轴 tick |
| legend 双击 | visible／position／wrap／selectMode／font | 普通点击仍由 renderer 切换显示状态，不入内容历史；More → legend route |
| series 点击或点上 Shift | 常用颜色、opacity、border、marker、label | marker 按真实图型；Pie／Funnel 不显示无效果的 series name 控件；More → 精确 series route |
| dataPoint 点击 | 本点 color／fillOpacity／border／size／shape | marker 按图型显隐；不提交 source／其他点；Delete 被元素选择消费，不删整图或数据 |
| dataLabel 点击 | 普通 selected series.label；Pie 的 options.pie.label | UI 明示系列／全扇区作用域；More 的 Pie label 到真实 pieStyle；隐藏后关闭 selection |
| Enter／F2／Escape | Enter 聚焦浮动表单，F2 打开该元素 More，Escape 放弃未提交草稿并恢复当前宿主焦点 | 输入法 composition、keyCode=229、busy 与真实权限门禁；输入框里的键盘事件不落到对象删除／剪贴板 |
| 恢复格式／Apply／More／选择系列 | 恢复先改本地 draft；Apply 一次 canonical patch；More／升系列先应用 dirty draft | no-op 不发作者命令；只读可查看禁用表单；失败显示真实 Error，不能用 toast 代替修改 |

菜单复用既有 ChartOptionFields／axis／series／label controls；不透明度使用项目已有 0–1 的 Ant InputNumber，步长 .05，允许上游候选值之外的合法精度。这是界面差异，不声称复刻上游浮动菜单像素。数据点 reset 使用项目既有 nullable patch 边界删除选中 override，保留其他点、系列 name／type／rightAxis 和源数据。

## 作者身份、几何与生命周期契约

- ECharts runtime series ID 为 `office-chart-series:${JSON.stringify([chartId, sourcePosition])}`；position 只定位当前帧原生 series，不是持久化 ID。
- 实际作者 selector 复用已有 `field:n`／`name:<canonical JSON string>`。默认唯一 field 使用 numeric selector；同字段重复时使用唯一原始 `sourceName ?? name`；field 与原始 name 双重歧义时不创建可直接修改的绑定。显示 rename 不改源身份。
- **已有精确 name 样式是最高作者层时，直接编辑选择该 named selector**，防止写入被 named style 覆盖的底层 numeric scope。binding 持有 selector／seriesId／field／sourcePosition／sourceName；原生和 graphic 走同一映射。
- 每个 native binding 明确 `raw input index → prepared semantic dataIndex`：日期重排仍保留空槽，Scatter／Bubble 与 option builder 共用有限数／log／size 过滤，原生 datum 通过 `getRawIndex()` 转换。Funnel 排序不把显示位置当 source point。native Radar 的一个 datum 代表整个系列，标 `seriesOnly`；graphic Radar 才有实际点级身份。
- `readChartCanvasElement` 在 renderer 边界重建／校验 metadata；未校验文本、显示名、native index 或随意对象不能成为作者 selector。`officeChartElements` 是 runtime 元数据，setOption 前剥离，但纳入帧 signature，防止画面相同而身份缓存过期。
- Index 来自真实 ECharts component／series view 和 ZRender display list。`findHover().topTarget` 保留真实形状、z-order 与 clip；文字沿 `__hostTarget` 再 parent 查找绑定。旋转框使用实际 computed transform；clip path 与画布边界裁剪测量 bounds；不可绘制／完全在 clip 外的元素不创建 target。
- 任意曲线、扇区等非矩形 path 的 **hit 使用 native 精确 contain／clip**，outline 仍是实际 bounds 的保守矩形。非矩形 clip 的 outline 按变换后的边界多边形保守裁剪；细线有最多约 1px 的可见选择框 padding。不能据此声称复刻上游全部 overlay circle／sector／nested-items。
- 命中先同步清除 transient pointer graphics，防止指示线遮住作者 mark。元素手势拥有独立 drag mode，消费点拖动而不移动整图；原 Sankey drag 保留。
- `useChartElementEditing` 分开 selected／commit：owner、workbook、sheet、activeKind、revision、作者 snapshot、generation、真实权限、busy 和当前 target 在提交前核对。late callback 不能把旧 selection 重绑到另一 chart／series。
- 内容提交只有 `FChart.update → sheet.command.update-chart`，沿用计算准备、权限、revision／候选事务与按文档隔离的历史。UI 不另存 undo；selection／bounds／route／runtime binding 不写作者快照。
- 每帧绘制后、legend toggle／翻页后重建 index；选中 target 消失后关闭。scroll、zoom、theme、DPR、fonts、blur、切表、owner 替换、deactivate／dispose、outside pointer 都退休草稿。author commit、对象 copy／export／paste 和拖动互斥；隐藏后仅向仍有效的原宿主恢复焦点。

本批仍属于 `features/charts` 的纯身份／planner、主线程几何适配、Vue controller 和 Ant 视图；大型数据准备复用既有 Worker／Rust/WASM，不新增第二个权威计算器。不修改 Rust／Worker 协议／lease／Office schema，也不因这个主线程命中接线宣称已经性能优化。

## 文件与接口变化

本批 **20 个文件：13 源码、6 测试／fixture、1 独立证据**。中央功能／工具栏／后端文档由主任务维护；图片与 Docs 文件未修改。

| 文件 | 本批变化 |
| --- | --- |
| [chart-element-interaction.ts](../../packages/office-studio/src/features/charts/chart-element-interaction.ts) | 新增 discriminated target／binding、runtime map、校验、clipping、More route 与 outline；bindings 可选 options 参数识别最高 named layer |
| [echarts-element-interaction.ts](../../packages/office-studio/src/features/charts/echarts-element-interaction.ts) | 新增真实 native painted index、命中／测量；图例 hit 保留实际 item name |
| [chart-element-edit.ts](../../packages/office-studio/src/features/charts/chart-element-edit.ts) | 新增 scoped controls／signature／patch／reset／Delete／hidden planner |
| [use-chart-element-editing.ts](../../packages/office-studio/src/features/charts/views/use-chart-element-editing.ts) | 新增 selection／commit 生命周期和 canonical author adapter |
| [ChartElementFloatMenu.vue](../../packages/office-studio/src/features/charts/views/ChartElementFloatMenu.vue) | 新增 Ant Form／ConfigProvider 浮动工具、IME／dirty／More／promote／error |
| [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) | native runtime IDs、point order、graphic 元素 metadata |
| [chart-axis-title.ts](../../packages/office-studio/src/features/charts/chart-axis-title.ts) | 实际轴标题 draw metadata |
| [echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts) | `getElementHit/getElementTarget/getElementTargets`；帧／图例后刷新 index |
| [chart-draw.ts](../../packages/office-studio/src/features/charts/chart-draw.ts) | Cartesian／Radar／Pie／Funnel 真实 marks／labels／axes 标签；排除 Pareto generic target |
| [chart-date-axis-draw.ts](../../packages/office-studio/src/features/charts/chart-date-axis-draw.ts) | 日期轴 labels／lines／ticks／grid metadata |
| [echarts-special-option.ts](../../packages/office-studio/src/features/charts/echarts-special-option.ts) | native Funnel runtime ID |
| [use-chart-object-actions.ts](../../packages/office-studio/src/features/charts/views/use-chart-object-actions.ts) | `execute(action,target?,route?)` 转交 edit route；optional `authorBusy()` 门禁 |
| [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) | 实际元素优先级、点／Shift 系列、浮动 UI、measured outline、键盘／互斥和生命周期接线 |
| [chart-object-actions.html](../../packages/office-studio/tests/fixtures/chart-object-actions.html) | 复用真实完整工作台；测量／真实 hit 采样、author completion gate、view 状态、zoom 用生产 WorkbenchDisplayKey 注入 |
| 以下 5 个测试文件 | 均为未运行草稿；见下表 |
| 本证据文件 | 固定来源、kind／入口分类、真实范围和缺口 |

没有 Facade／index 公共导出变更；不导入任意 Univer opaque seriesId。新 renderer 方法仅提供当前帧 runtime target，不构成作者／Office 新协议。

## 测试草稿、UI 与执行状态

按 `test(...)` 声明手工计数，参数循环不另计：**34 条，全未运行**。

| 文件 | 声明数 | 意图 |
| --- | ---: | --- |
| [chart-element-interaction.contract.test.mjs](../../packages/office-studio/tests/chart-element-interaction.contract.test.mjs) | 8 | field／name／rename／最高 named layer、歧义拒绝、日期空槽、Scatter／Bubble 过滤、metadata 校验、旋转 clip、横向／次轴、Radar／Pareto 边界、legend mask／route |
| [chart-element-edit.contract.test.mjs](../../packages/office-studio/tests/chart-element-edit.contract.test.mjs) | 7 | scope controls／不泄漏其他 draft、point reset／siblings、name selector、普通／Pie label、series reset 保数据／身份、Delete |
| [chart-element-editing.contract.test.mjs](../../packages/office-studio/tests/chart-element-editing.contract.test.mjs) | 7 | 真实 FChart 命令／history／save／reopen／no-op、外部 revision、同 ID 新 owner、busy／权限、Delete／focus、promote／canonical More、真实提交后 late completion |
| [echarts-element-interaction.contract.test.mjs](../../packages/office-studio/tests/echarts-element-interaction.contract.test.mjs) | 4 | 真实 ECharts SVG SSR model／view／ZRender：轴／标题、native raw point／attached label、graphic group clip／隐藏、scroll legend 重测量；不代替 Canvas 浏览器验收 |
| [chart-element-direct-editing.browser.test.mjs](../../packages/office-studio/tests/chart-element-direct-editing.browser.test.mjs) | 8 | 完整 ExcelWorkbench＋Canvas2D：实际命中／Ant输入、轴／点／Shift／reset／labels／Pie、view legend／dblclick、canonical Drawer More、IME／只读、zoom／theme／font event／scroll／busy／切表／新 owner |

浏览器草稿字体失效场景派发实际订阅的 FontFaceSet `loadingdone` 事件，仅验证退休机制，不证明真实字体加载、CJK／RTL shaping 或字形几何已验收。内部 JSON save／reopen 草稿不代替 XLSX／DOCX／Office 样本重开。本文没有浏览器截图、真实文件或性能测量结果。

可替换 UI 使用 `antdv-next` ConfigProvider／Form／FormItem／Input／InputNumber／Select／Switch／ColorPicker／Button／Tooltip／Alert，复用现有 ChartColorInput；正文／图内标记／选择框仍为 Canvas，未新增原生输入桥或另一组件库。统一主题 token 与局部 popup container 已接线；**明暗主题、popup、焦点和输入法尚未实际验证**。

本批只进行文件／固定源码阅读、搜索、手工静态审阅和代码／草稿／证据写入。**没有运行测试、类型检查、构建、dev／Vite、浏览器、Cargo／WASM、性能基准、服务或 `git diff --check` 等自动检查；没有提交、发布或改 package／锁文件。** 所有结论保持未验收。

## 后续缺口

1. Pareto 专属 bar／cumulative，Waterfall positive／negative／subtotal，Pie composite aggregate／connector 的真实 direct-edit target 尚未接入；Pareto generic target 主动排除，避免提供被专属 style 覆盖的无效果编辑。继续按 7 个 mark direct-edit 分支补实现。
2. Radar 轴、3D marks、heatmap／hierarchy／graph／wordCloud／chord 等专属元素尚未完整覆盖；多级分类 expand／collapse 和 Pivot field button／overflow 未实现。不能把二维通用系列命中外推到所有 45 图型。
3. 上游 semanticId／aggregate／nested overlays、任意 opaque seriesId 导入、图例空白容器、native Radar 逐点编辑仍有差异；当前 source／runtime 映射只覆盖明确可追溯的本地契约。
4. 浏览器／中文 IME／主题／缩放／DPR／真实字体、权限变更与竞争弹窗尚未实际验收；本批和此前草稿统一后置，类型与构建结果未知。
5. 内部保存不证明 Office 保真；原生图表样本写出、独立 Office 重开以及真实 source 更新／undo／redo／重排／跨工作簿复制仍待补证。
6. 首屏、滚动、编辑、缩放、native index 重建、Worker／WASM 数据准备、传输／内存和大样本均未测；不声称性能／内存已优化或达标。

上述缺口保留全部 Sheets／Docs／Pro 的最终范围，下一批继续真实专属元素闭环。
