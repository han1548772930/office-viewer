# SHEET-22 / SHEET-24 / DOC-19 图表类型与编辑契约对照

更新日期：2026-11-25。**整体状态：未验收；当前开发进度：开发中、待统一验证。** 本轮遵守用户“先实现，后统一测试”的要求，仅实现、补测试草稿并静态阅读源码/diff；没有执行测试、类型检查、前端构建、Rust/WASM 检查、基准、dev server 或浏览器。下文“已接入/已编写”只描述代码，不代表公开行为、Office 往返或性能验收。

本记录覆盖共享图表引擎与 Sheets 普通/透视图，包含 Docs 图表需要复用的契约。DOC-19 的 Docs 插入、内嵌数据、对象锚点/环绕/分页及编辑宿主尚未接入，不能由共享目录或算法推导其完成。全部 45 项、全部 Pro 配置和入口继续保留在交付范围内。

## 1. 固定来源与证据边界

OSS 只读 checkout：`E:/rust_test/univer`，tag `v1.6.0`，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 发布包研究缓存：`.research.local/univer-1.6.0/packages/`，固定发布归档在 `archives/engine-chart-1.6.0.tgz` 和 `tarballs/{chart-ui,sheets-chart,sheets-chart-ui}.tgz`。`chart-ui`、`engine-chart`、Sheets/Docs/Pivot Chart 相关包的 `package.json` 记录版本 1.6.0；**engine-chart 类型已解包，固定 ES 从归档内 `package/lib/es/index.js` 静态读取；缓存中没有 `packages/engine-chart/lib/es/index.js`**。本地研究缓存不是生产依赖，没有安装、导入执行或引入商业 SDK。

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
| P10 可见数据默认 | `engine-chart/lib/types/types.d.ts:396–398::IChartContext.plotVisibleOnly`；`chart-builder/chart-types.d.ts:552`；`models/chart-model/chart-model.d.ts:74::getPlotVisibleOnly`；固定ES `Ks.getPlotVisibleOnly/DEFAULT_PLOT_VISIBLE_ONLY`（字符offset204027/204917） | 声明注释与运行getter均已静态核实：`context.plotVisibleOnly ?? DEFAULT_PLOT_VISIBLE_ONLY`，静态常量为true；serialize省略未作者值而保留显式false。本地true缺省不需要再改；可见性实际源消费/Office仍未验收。 |
| P11 Radar 专属契约与显隐 | `engine-chart/lib/types/facade/builders/f-radar-chart-builder.d.ts`；`types.d.ts::IChartRadarAxisLabelSpec`；固定ES `R/BOe/bje/xje/Sje`（radar defaults offset13237；label converter1614386）；`chart-ui/lib/es/index.js::Sl/xl` | 6个setter/reset、nullable标签字段；R.radar.axisLabel.visible=true、splitNumber5、textStyle.fontSize12。Sl和xje均取true默认；bje最终按plot半径扣图表字体+12px，作者PlotArea不扣。面板shape → fill → 数值标签开关，visible后格式/字体。缺省已静态确认，公开视觉/交互仍未验收。 |
| P12 Cartesian 与 Boxplot 继承 | `engine-chart/lib/types/facade/f-cartesian-chart-builder.d.ts`、`facade/f-axis-chart-builder.d.ts`；`chart-builder/configuration/cartesian-series.d.ts/axes.d.ts/axis-pointer.d.ts`；`chart-builder/chart-types.d.ts:114–150,484–489::IChartAxisSpec/IChartAxesSpec/IChartAxisPointerSpec`；`enum.d.ts:242::AxisAlignEnum`；`facade/builders/f-boxplot-chart-builder.d.ts`；固定ES `wke`（titleAlign字符offset1540396）；`chart-ui/lib/es/index.js::$c/js/Hs/zs/Bs/Vs/qi/Hi` | 系列7个/全系列6个clear target、轴8个clear target、标题字体/对齐、日期开关、指示器4个字段和恢复目标及 `number / { name }` selector。wke已静态确认标题起止为物理文字盒左/右、上/下，独立于reverse；字体缺省12px，默认竖轴逐字换行。本地旋转式标题差异见P12；完整碰撞/公开效果及Radar指示器未验收，次轴公开类型仍排除reverse。 |
| P13 Bubble 映射与名称 selector | `engine-chart/lib/types/facade/builders/f-bubble-chart-builder.d.ts::setBubbleMapping/clearBubbleMapping`；`chart-builder/configuration/cartesian-series.d.ts::ChartPublicSeriesSelector`；`chart-builder/chart-types.d.ts::ChartSeriesSelector/IChartBubbleMappingSpec` | Bubble 的 X/Y 必为数值，size/label/series 为可选源字段；每源行一气泡，各向量按源行对齐。公开 selector 为 `number / { name }`，内部 selector 为 `{ index } / { name }`。本地保留既有数值源字段语义，以独立 namedSeries 数组表达名称覆盖；名称精确匹配、重名歧义拒绝和未活动覆盖保留是本地规则，尚未沿固定 ES 核实上游运行时边界。 |
| P14 SolidSeries 填充纹理 | `engine-chart/lib/types/chart-builder/chart-types.d.ts:221,234::IChartSeriesStyleSpec/IChartSeriesDataPointStyleSpec`；`types.d.ts:753,765,794::ISeriesStyle/IAllSeriesStyle`；`models/chart-render/converter/render-spec-operators/series-style-operator/series-fill-pattern.d.ts::createSeriesPatternStyle` | 系列及可填充数据点接受 `fillPattern / null`；公开转换声明包括 color/decal，本地独立生成 DrawingML 风格纹理几何。54 种纹理不是 54 个已验收功能；上游精确像素、各类型可见入口/优先级及原生 Office 往返仍须核实。 |
| P15 自动渐变与绘图区 | `engine-chart/lib/types/chart-builder/configuration/auto-gradient-fill.d.ts::setAutoGradientFill/resetAutoGradientFill`；`chart-builder/chart-types.d.ts:502,554,579,628–632::IChartPlotAreaSpec/IChartSemanticConfig/IChartDescription`；`types.d.ts:867::IChartPlotArea`；`models/chart-render/converter/render-spec-operators/auto-gradient-presets.d.ts/grid-operator.d.ts`；`chart-ui/lib/es/index.js::Oi/Qo/Zo/Yi/Wi` | 自动渐变保存布尔覆盖，reset 恢复图型默认，描述返回有效布尔值；UI 的默认表达式为 Area 开启。手动绘图区是相对完整图表外框的 left/top/width/height。声明提供 Bar/Area/Line/Pie 预设及自动 grid/编辑 seed；闭源预设色标、矩形边界限制及最终拖拽/面板行为未核实。 |
| P16 语义描述与显式线性渐变 | `engine-chart/lib/types/chart-builder/chart-color.d.ts:1–49::CHART_LINEAR_GRADIENT_TYPE/IChartLinearGradient/ChartColor/chartLinearGradient`；`chart-builder/chart-types.d.ts:153,219,232,250–258,618–650::IChartSeriesDescription/IChartCategoryDescription/IChartDescription`；`chart-builder/chart-builder.d.ts:22::describe`；`chart-live-adapter.d.ts:5`；`facade/builders/f-line-chart-builder.d.ts:52/f-area-chart-builder.d.ts:52`；`models/chart-render/converter/render-spec-operators/chart-color.d.ts::getChartColorRepresentative/toEChartsColor/resolveChartColor`；`chart-ui/lib/es/index.js::$o/sc/dc/Tc/xc/cc` | 描述按源字段索引返回无数值元数据、语义样式和轴绑定；颜色为 CSS 字符串或带稳定 discriminator 的 mark-local 渐变。公开 stop.opacity 乘 CSS alpha。声明不公开完整校验/代表色算法；共享 picker `$o` 只取作者第一个 stop.color，Pareto `sc` 则只传纯色。完整方向/色标编辑器是本地扩展，不能把当前 picker 当作完整上游渐变编辑证据。 |
| P17 Pareto 独立柱形与累计线样式 | `engine-chart/lib/types/chart-builder/chart-types.d.ts:436–483::IChartParetoBarStyleSpec/IChartParetoLineStyleSpec/IChartParetoSpec`；`facade/builders/f-pareto-chart-builder.d.ts`；`chart-builder/configuration/pareto.d.ts`；固定归档 `lib/es/index.js::R/Yi/Xi/yre/bre/xre/Sre/OOe/Jje/Yje`、`lib/es/facade.js::pi/FParetoChartBuilder`；`chart-ui/lib/es/index.js::cc/oc/Wu` | 六操作、深合并/clear null、非正累计线宽忽略、第四主题色/2px/opacity1/7px/circle 默认已静态追踪。`OOe` 按分类代数合计后保留 >0 或 >=0、稳定降序、累计 floor 两位；`Jje/Yje` 的实际样式转换可静态读取。`Wu` 总是只列正合计，本地选择器扩展为与 includeZeroValues 下的物化数据一致；逐点片段/百分比标签差异见 P17/P19。 |
| P18 Waterfall 完整累计与创建边界 | `engine-chart` 固定归档 `lib/es/index.js::ake/Nn/Rte/Xn/Mte/Pte/qn` 及 arrangement/stacked converter；`lib/types/facade/builders/f-waterfall-chart-builder.d.ts`；`chart-ui/lib/es/index.js::ji/Al/Ol`；`sheets-chart/lib/es/index.js::SheetChartLiveAdapter.commit/commitChanges/_replaceConfig`、`lib/es/facade.js::insertChart`；`sheets-chart-ui/lib/es/index.js::ChartWaterfallStyleSection` | 首有限来源的创建合计标记、own-property 默认、重复身份最后输入胜出；跨来源排列/分类堆积、每来源整图最终合计/单个全局合计；角色/连接线/标签与面板显隐。普通 patch 的 `toChartModelUpdate` 不执行创建默认；整份 config replacement 与新建分别执行 `toChartModelConfigReplacement/toChartModelConfig`。本地边缘连接、subtotal 后显示索引修正、完整样式/逐点源身份是独立实现，详见 P18 小节；没有执行 SDK 或公开演示。 |
| P19 标记、组合标签与位置 | `engine-chart/lib/types/enum.d.ts::LinePointShape/LabelContentType/SeriesLabelPosition`；固定归档 `lib/es/index.js::R6/I6/L6/POe/FOe/u5/f5/F6`；`chart-ui/lib/es/index.js::cc/oc` | 19 个形状、5 个内容 flag 与 Empty0、9 个公开位置；dot/dash 的固定 path 包含 nominal box anchors。Pareto `u5` 实际只拼 Series/Category/Value，虽然面板可选 Percentage；本地补真实百分比。insideBase/End 的固定 bar 映射按正负/reverse 定向；本地共享 anchor/空心几何与 center 扩展的差异见 P19。 |
| P20 Sankey 流量、轨迹与节点拖动 | `engine-chart` 固定归档 `lib/es/index.js::Kae/UOe/Gje/_getShowEmpty/sye`；`models/chart-data/builders/sankey-chart-data.d.ts`；`models/chart-render/converter/converters/sankey-chart-converter.d.ts`；`render-spec-operators/sankey-chart-operator.d.ts`；`facade/chart-builder-type-map.d.ts` | 三字段映射、首次端点顺序、源节点颜色、justify、name 标签和 `emphasis.focus='trajectory'` 已静态追踪。bundled native 默认 nodeWidth20/nodeGap8/iterations32/draggabletrue/line opacity.2/emphasis opacity.5；节点使用 max(in,out,value)。固定 builder 为 FChartBuilderBase；重复边累计、完整循环回流和扩展样式/标签是本地差异，见 P20。 |
| P21 Heatmap 双分类轴、标签与色标 | `engine-chart` 固定归档 `lib/es/index.js::Oae/wOe/uAe/Mje/Nje`、axis operator；`lib/es/facade.js::di`；`facade/builders/f-heatmap-chart-builder.d.ts`、`chart-builder/configuration/heatmap.d.ts`、`chart-types.d.ts:58,213,278`、`types.d.ts:454`；`chart-ui/lib/es/index.js::ml/hl/Gs/Zs` | 字段名为 X、源分类为 Y、tuple 为 `[fieldPosition,categoryIndex,value]`；实际范围/空0/0、source formatted label、inside/12px、默认首Y分类底部、splitNumber8、连续边界单位和非拖动色标已静态追踪。8个专属 setter/reset；面板/公共Canvas色标及精确区间差异见 P21。 |
| P22 日期基本单位、主刻度与重复源行 | 固定 `engine-chart` 归档 `package/lib/es/index.js::nae/Oa/hm/lOe/v8/Rke/qke/OOe`；`package/lib/types/types.d.ts:703,707–709::IAxisOptions`；`chart-builder/chart-types.d.ts:114–144::IChartAxisSpec` | 对UTF-8原文静态查找字符offset：nae89062、Oa89318、hm339951、lOe1484869、v8 1546906、Rke1551622、qke1558542、OOe1500076。普通Cartesian稳定日期排序并保留sourceDataIndices，重复日期不开启自动聚合；month/year采用连续calendar桶及首日标签，major单位独立推进。OOe专用Pareto转换器不消费calendarDateAxis。本地三字段API/UI是扩展，公开IChartAxisSpec未暴露它们；输入、自动稀疏与资源错误差异见P22小节。 |
| P23 Boxplot 固定默认与原始异常点身份 | 固定 `engine-chart` 归档 `package/lib/es/index.js::Yce/Xce/$je.boxplot`；`facade/builders/f-boxplot-chart-builder.d.ts`；`models/chart-data/builders/boxplot-chart-data.d.ts` | UTF-8原文字符offset：Yce箭头函数340896、Xce对象341159、$je对象1630106/boxplot方法1630111。Xce每原数值字段一分布，Yce上下半组中位数（奇数排除中间项）及全样本min/max；Xce不创建独立异常点，$je.boxplot不消费dataPoints。本地Rust/TS默认resolver与绘制消费已静态落盘；inclusive/exclusive、IQR和逐原始样本异常点覆盖为Office扩展，不能记为固定版同义或已验收。 |
| P24 未发布草稿的数据/统计读取 | `engine-chart/lib/types/facade/f-chart.d.ts::getInfo/update`；P03/P08数据与透视源声明；本地 `FWorksheet.getChartSettingsData/getChartSettingsStatistics` | 固定声明已核实detached配置读取和整体更新，但未找到等价的大源设置读取公开方法。本地新增Promise方法供尚未创建的图表及编辑/Pivot草稿使用；原始准备数据和物化统计采用显式绑定，不以render rows代替source，也不声称复刻私有读取协议。 |
| P25 Histogram / Candlestick 完整专属入口 | 固定 `engine-chart/lib/types/types.d.ts:74–128,480–509`；`facade/builders/f-{histogram,candlestick}-chart-builder.d.ts`、`chart-builder/configuration/{histogram,candlestick}.d.ts`；归档 `lib/es/index.js::Iae/Lae/Rae/zae/Bae/Ka/qa/Vae/xm/$De/tOe`；`chart-ui/lib/es/index.js::Vo/Uo/ha/da/fa/pa/ma` | UTF-8原文字符offset：Iae107266、Vae109566、xm342934、$De1475779；UI Vo84270、Uo84998、ha48248、da43018、fa43777、pa45849、ma48076。固定Facade Histogram setter仅接受有限作者数值，Auto省略binning；面板和算法的分组/显隐、1000箱含边箱、样本方差Scott、恒定点箱、三状态/HLC及overlay对象存在语义已静态追踪。没有执行SDK、公开演示或Office样本，差异与草稿见P25小节。 |
| P26 Treemap 层级、标签与实际原生适配 | 固定 `engine-chart/lib/types/facade/builders/f-treemap-chart-builder.d.ts`、`chart-builder/configuration/treemap.d.ts`、`models/chart-data/builders/hierarchy-chart-data.d.ts`；归档 `lib/es/index.js::Ua/Aae/eke/$Oe/H6/U6/ike/wp`；`chart-ui/lib/es/index.js::to` | UTF-8原文字符offset：Ua104244、eke1519960、$Oe1519302、UI to62258。按祖先作用域填空、strict finite number、signed selfValue/正子树剪枝、正父残差、三种父布局、独立叶标签与完整fit已静态追踪；六公开setter/reset使用共同内容命令。锁定ECharts6.1.0 Tree/TreemapView及zrender Text的本地扩展是适配依据，不是Univer SDK执行或浏览器验收。 |
| P27 Sunburst 语义角度、标签、中心及设置 | 固定 `engine-chart/lib/types/facade/builders/f-sunburst-chart-builder.d.ts`、`chart-builder/configuration/sunburst.d.ts`；归档 `package/lib/es/index.js::qOe/V6/KOe/JOe/YOe/XOe/QOe/ym/bm`；`chart-ui/lib/types/editor/sections/specific-section/{sunburst-label,SunburstLabelAtom,SunburstCenterLabelAtom,SunburstDataMappingAtom}.d.ts`、`chart-ui/lib/es/index.js::To/jo/No/Fo/Jo` | 正父残差保留semantic parent value；18%/88%环径、首圈tangential/后续radial、4%完整标签门限、category/value换行、独立中心total/format和六公开操作均已静态追踪。UI Jo字符offset87308、jo75093、No76054、Fo76954：临时节点/中心选择、各visible条件、Value格式及有序mapping，不以类型名推断完成。 |
| R01 本项目绘制依赖 | ECharts 6.1.0 的 `lib/chart/gauge/GaugeView.js:211–250,337–355,469,508,532`、`lib/util/shape/sausage.js`；zrender 6.1.0 的 `lib/core/PathProxy.js:34–55` | 静态读取当前锁定依赖的刻度文字定位/对齐、progress overlap/环宽/z2、锚点/文字层级与 arc normalization，作为本项目两绘制适配器的实现依据；不是 Univer 内部源码证据。 |
| R02 Radar / Boxplot 绘制依据 | ECharts 6.1.0 的 `lib/coord/radar/Radar.js/RadarModel.js`、`lib/chart/radar/RadarView.js`、`lib/chart/boxplot/boxplotLayout.js::layEndLine`、`BoxplotView.js` | Radar原生坐标支持clockwise但没有径向inverse；径向反向等配置复用graphic。Boxplot须线cap使用完整箱宽，原生View不绘制数据标签，启用标签时复用共享图元。P11已由固定Univer ES确认Radar标签默认显示；本地native显式覆盖ECharts隐藏缺省。实际SSR仍未执行。 |
| R03 真实指示器动作 | ECharts 6.1.0 `lib/component/axisPointer/modelHelper.js:93–190`、`AxisPointerModel.js`；2026-11-24 Context7 的官方 `updateAxisPointer`、`tooltip.triggerOn` 与 `axisPointer.type: cross` 文档 | tooltip 的 pointer 子模型优先于 component defaults，因此两个位置同步四个作者样式；离屏 Canvas 通过动作更新，`showContent:false/triggerOn:'none'`。文档为动态 API 依据，不是固定 Univer 运行时证据；原生和 graphic 浏览器草稿均未执行。 |
| R04 渐变绘制 API | 2026-11-24 Context7 查询 `/apache/echarts-doc` 返回 [option.md](https://github.com/apache/echarts-doc/blob/master/en/option/option.md) 的 serializable linear gradient | `{ type:'linear', x,y,x2,y2,colorStops,global }`；`global:false` 使用图元归一化坐标，`global:true` 使用 Canvas 像素坐标。只用于本项目 ECharts 6.1.0 适配，文档动态版本不证明固定 Univer 实现，也没有运行当前原生渲染。 |
| R05 本地渐变编辑控件 | 2026-11-24 Context7 `/antdv-next/antdv-next` 的 ColorPicker/FormItem 当前 main 文档；本地 [pnpm-lock.yaml](../../pnpm-lock.yaml) 锁定 antdv-next 1.8.2 | 使用受控 value/update:value、RGB value-format 保留 CSS alpha；FormItem、Select、InputNumber、Button 和反馈控件承载方向与色标。动态 main 只是 API 使用依据；本地 1.8.2 的实际主题、弹层、焦点和 IME 尚未运行验证。 |
| R06 Waterfall 公共图元适配 | 2026-11-24 Context7 `/apache/echarts-doc` 的 graphic elements/type/shape/style/clipPath 文档；本地锁定 ECharts 6.1.0 | Waterfall 的同一 DrawCommand 转成 graphic mark、line、text 与绘图区裁剪；无额外 ECharts 累计引擎。动态文档只用于本地适配 API，不证明 Univer 几何或浏览器像素效果。 |
| R07 标记与多选控件 API | 2026-11-25 Context7 `/apache/echarts-doc` 的 `series.symbol/path://` 文档、`/antdv-next/antdv-next` 的 Select `mode="multiple"` 文档；本地锁定 ECharts 6.1.0、antdv-next 1.8.2 | native symbol 归一化与实际 antd 多选控件的使用依据；复杂/空心/dot/dash 走公共 graphic 几何，数值 mask 由统一 helper 原子转换。动态文档不是固定 Univer 源码或本地浏览器/主题/焦点验收。 |
| R08 Heatmap native 适配与悬停 | ECharts 6.1.0 `lib/component/visualMap/ContinuousView.js::HOVER_LINK_SIZE`、`ContinuousModel.js::getValueState`、`VisualMapModel.js::resetExtent`、`visual/visualSolution.js`；2026-11-25 Context7 `/apache/echarts-doc` 的 Heatmap/visualMap/pieces 文档 | continuous 默认12px hover band及端点外数据、extent升序、datum.visualMap=false的跳过路径和显式 gt/lte区间是本项目适配依据。实际可见色标/hover由公共graphic完成，内部hidden visualMap仅保留native热力图必需模型；不冒称上游图内Canvas交互或浏览器已验收。 |
| R09 公共轴标题文字适配 | 2026-11-25 Context7 `/apache/echarts-doc` 的 graphic text `align/verticalAlign/font/width/overflow` 和轴 `nameLocation` 文档；ECharts 6.1.0 `lib/component/graphic/GraphicModel.js::optionUpdated/_flatten` | native数值轴nameLocation受轴方向影响，不能表达P12物理文字盒语义。本地关闭native name，两路径复用graphic text/宽度截断/外框clip；GraphicModel把group的children展平为elements/parentId，renderer草稿按实际展平模型读取。动态master文档只作适配API依据，不是固定Univer或当前浏览器像素验收。 |
| R10 日期/数值输入及浏览器定位 | 2026-11-25 Context7 `/antdv-next/antdv-next` InputNumber/FormItem 的当前main文档、`/microsoft/playwright` Locator.getByRole及portal作用域文档；本地锁定antdv-next1.8.2、Playwright1.65.1 | 正小数输入、清空null→undefined，以及分面板的同名字段/页面级popup option定位依据。查询传入本地版本，返回main来源，精确版本命中未证实。只写草稿并静态核对真实Drawer的“关闭”动作；没有执行浏览器、主题、焦点或IME。 |
| R11 Boxplot 原生标记与真实Canvas草稿 | ECharts 6.1.0 `lib/chart/helper/Symbol.js::_createSymbol/getSymbolPath/_updateCommon`、`lib/chart/boxplot/BoxplotView.js::render`；InputNumber/portal参考沿用R10 | 原生Scatter将symbolSize/2写入标记path缩放并保持strokeNoScale，实际path样式及datum ID可用于未来renderer草稿；Boxplot空datum保留数据位置且不创建path。getSymbolPath/display list是实现细节，仅用于本地验证适配，不作为本项目公开API或固定Univer证据。4个实际Canvas草稿未执行。 |
| R12 设置读取生命周期与类别输入 | 2026-11-25 Context7 `/vuejs/docs` 的watch清理/同步flush/onScopeDispose；`/antdv-next/antdv-next` 的InputNumber/Select/Form当前main文档；本地锁定antdv-next1.8.2 | 同步失效、microtask合并最终配置请求、scope释放和受控value/update:value的API依据。InputNumber清空null、min/max和整数precision；Select显式按label检索已生成选项。动态文档不能证明固定Univer交互或本地版本已运行；完整列表的名称搜索、主题/弹层/焦点/IME和性能保持未验收。 |
| R13 Candlestick native paint 与 Histogram Canvas 草稿 | ECharts6.1.0 `lib/chart/candlestick/{candlestickLayout,candlestickVisual,CandlestickView}.js`；控件参考沿用R05/R10/R12 | 原生布局是否提供borderColorDoji决定严格相等时的sign0；visual按sign重新写fill/stroke，故每datum同步color/color0、三种outline且设置large:false。缺失datum不创建path。display list/getItemGraphicEl仅供本地验证适配，不是公开API。Histogram实际barCategoryGap与共同binGap resolver同源，浮点百分比按数值比较；三个真实Canvas草稿未执行。 |
| R14 层级原生文字、纹理与透明度适配 | ECharts6.1.0 `lib/chart/sunburst/{sunburstLayout,SunburstPiece,SunburstView}.js`、`lib/chart/treemap/{TreemapView,treemapVisual}.js`、`model/mixin/makeStyleMapper.js`；2026-11-25 Context7 `/apache/echarts-doc` 的Sunburst rotate/formatter文档 | 锁定SunburstPiece的首圈切向、后续径向、upright flip和sector bbox用于本地共有坐标；公开labelLayout不实现hide，故marked postUpdate同步四状态ignore。TreemapView背景与内容Rect分开，shape mapper不消费itemStyle.opacity，填充alpha嵌进color；原生pattern用同步Canvas tile而非decal/异步图片解码。内部结构、alpha及真实bitmap仍未验收。 |
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
| P12 对象属性面板叶字段 | `$c`：标题文字 → 可选对齐 → 可选字体复合控件；`js`：X 轴日期开关；`Hs`：线色 → 线型 → 标签底色 → 标签文字色。 | 每个可见轴接7个标题叶字段（文字1、对齐1、字体5）及7个真实reset动作；文字reset清标题/对齐/字体整组，其余独立恢复并显示有效继承。X轴日期开关1、指示器字段4另计。属于上下文属性，不增加常驻、插件根、溢出或45目录数；字体复合节点不重复计数，运行时可见组合/主题/焦点/IME未验收。 |
| P11 Radar 上下文属性 | `Sl/xl`：shape → fill → 数值标签visible；默认visible=true，条件区顺序为format → font复合节点。仅Radar，不提供径向log。 | 基础值入口3个，visible增加format1和字体5个叶字段，合计9个值；本地9个对应reset动作，visible reset清label整组。字体复合节点不重复计数，不增加常驻、插件根、溢出或目录数；旧yAxis.label.visible兼容继承为本地差异。 |
| P15 上下文配置入口 | `Qo/Zo`：自动渐变设置支持 Column/Bar/Area/Combination/Pie/Waterfall 六族，对应目录 25 项；Line 的公开渲染预设存在但设置开关隐藏。`Yi/Wi`：PlotArea 支持 Column/Bar/Line/Area/Scatter/Bubble/Combination/Pie/Radar 九族，对应目录 31 项。 | 本地渐变为 1 个布尔值入口加 1 个默认恢复动作；绘图区为手动模式开关 1 加百分比叶字段 4，共 5 个值入口。研究 bundle 只找到 PlotArea section/状态注册，未定位到实际四坐标输入 UI；四字段控件是本地独立实现。均属于上下文属性，不增加 Ribbon 常驻/插件根或 45 项目录数，显隐/交互仍未验收。 |
| P14 上下文纹理属性 | SolidSeries 的全系列、源字段/名称系列及可填充数据点接 `fillPattern`；固定声明见 P14，完整上游 section 显隐/顺序仍未定位。 | 每个作者层为纹理 selector 1、显式纹理下前景/背景颜色各 1；本地 selector 在除 WordCloud 外的 44 个目录项显示，绘制效果取决于实际填充图元。54 个 preset 加继承/无纹理共 56 个选择值，不能计作 56 个独立功能或 Ribbon 根；上游各图型可见入口、默认纹理及运行效果仍待核实。 |
| P16 上下文颜色复合属性 | 固定 ChartColor/LineSeriesPatch/SolidSeriesPatch 区分系列填充、可填充点和纯色符号；本地颜色节点复用全系列、字段/名称系列和单点作者层。 | 目录中系列渐变可写 31 项，可填充数据点渐变可写 25 项；符号颜色仍只接受纯色。每个颜色节点含继承/纯色/渐变模式选择；渐变展开后为方向坐标 4 个、每个色标颜色/offset/opacity 3 个及真实增删动作。可变色标数不是功能数，不与外层颜色节点、常驻根或 45 项目录重复相加；完整上游渐变编辑入口未找到。 |
| P17 Pareto 上下文样式属性 | `cc` 按 target 分 barStyle/cumulativeLineStyle，顺序为颜色/opacity → bar 边框或 line 线宽/线型/点 → 标签 → bar 的 includeZeroValues → 可选逐点选择与样式。仅 Pareto，标签细节仅在可见时显示，逐点编辑另受 showDataPointEditor 和实际 options 非空控制。 | 本地一个 ChartParetoSettings 下为柱形、累计线两个 Collapse。每组顶层至多 16 个配置节点，共 32；现有零分类 Switch 另 1。选择点后柱形有 10、累计线有 12 个样式节点，另有选择器各 1；其中纹理颜色 2 个为条件项。两组整体恢复和已选点恢复各 2 个真实动作。颜色内部叶字段另按 P16 展开，字体按 5 个叶字段计，不重复计复合节点。这是本地上下文节点计数，完整固定布局/命令/公开行为未验收。 |
| P21 Heatmap 上下文属性 | `Gs/Zs`：legend/presentation内的色阶类型；`ml/hl`：同排min/max → unit → label visible；有效visible后inside/top/bottom → font复合节点。仅Heatmap；wrap/selectMode隐藏，无标签内容selector。 | 基础值入口5个（色阶类型1、范围2、单位1、标签开关1）；可见标签增加位置1和字体5个叶字段，共11个固定语义叶字段。本地格式字段另1个扩展；本地每字段真实reset，标签开关reset清整组。legend位置/显隐复用共同字段，不与本行重复计数。全部为上下文属性，不增加常驻、扩展根、溢出或45目录数；本地响应式Form网格不声明已复刻上游行列像素位置。 |
| P22 本地日期与数值刻度扩展 | P12已有X轴日期开关；IAxisOptions持有baseTimeUnit/majorTimeUnit/majorUnit，但公开IChartAxisSpec及已定位UI开关不暴露三个时间字段。 | 活动日期X轴在min/max之后增加基本单位→主刻度单位→主单位间隔，共3个值入口及3个独立reset；已有日期开关另补1个reset，不重复计开关。活动数值轴另有主刻度间隔1个值入口及1个reset。全部使用既有antdv-next FormItem/Select/InputNumber/Button，不计入常驻、插件扩展根、溢出或45目录；布局/显隐/公开交互仍未验收。 |
| P23 本地Boxplot上下文扩展 | P12继承系列/逐点设置；固定Xce/$je.boxplot没有独立outlier/dataPoints入口，本轮未核实同名上游统计属性面板。 | 专属区按四分位→须线范围→异常点开关→条件IQR倍数排列，共4个值字段（既有3个、新增须线范围1个）及4个新增独立reset。既有逐项目Select改为原始有限样本，新增原始样本序号1个值入口，可达列表外样本。合计新增值2个、reset4个；均属上下文扩展，不增加常驻、插件根、溢出或45目录数。字体/纹理子控件不在此重复计数；实际显隐与主题/焦点/IME未验收。 |
| P24 本地大源设置与项目定位 | 普通编辑、插入和Pivot面板复用同一异步读取hook；Pareto共用一次Rust物化统计。 | 非Boxplot通用项目序号1个、Pareto柱/线序号各1个、既有Waterfall角色列表项目序号1个，合计4个条件值节点；P23原始样本序号不重复计数。三个面板各有当前读取失败后的真实重试动作1个，共3个上下文动作；loading反馈不计功能入口。各类别Select只生成前16,384项及已选远端项，其余可按序号定位；完整名称搜索效果未核实。均不增加常驻、插件根、溢出或45目录数。 |
| P25 Histogram 上下文分组 | `Vo`在字段映射之后挂`Uo`：方式 → 条件Count或Width → 同排underflow/overflow；专属`ha`只显示gap，step.05。 | 样本字段1、方式1、互斥count/width各1、阈值2、gap1，共7个可复查语义值节点；Count/Width同时可见6个、Auto5个。模式/两种箱值/两阈值/gap各有本地reset，共6种恢复入口，count/width互斥时最多显示5个；已有Select清字段不重复计数。所有值/动作属对象上下文，不增加常驻、插件根、溢出或45目录数。 |
| P25 Candlestick 上下文分组 | `ma`用上游series长度3选HLC的`fa`、至少4选OHLC的`pa`。OHLC：临时state/width同排 → 非Doji body color/opacity同排 → hollow → border color/width同排 → border opacity。HLC：highLow toggle → 条件color/width、dash/opacity；close toggle → 条件color、shape/size。 | 排除共有category/open/high/low/close映射5字段：OHLC当前非Doji显示8个值节点，其中state1为临时选择、内容作者7个；Doji显示5个，其中内容作者4个。HLC有9个专属内容值，两个开关决定其余7个显隐。本地按显式mapping/candleFields判模式，并向OHLC扩展同一9个overlay字段。每个可见作者叶字段真实reset；状态、高低线、收盘标记、全部4个整组动作，HLC隐藏状态后为3个。颜色复合控件不重复计内部叶值；属于上下文属性，完整运行行列/焦点/主题/IME未验收。 |
| P26 Treemap 上下文标签与本地扩展 | `to`顺序为父布局selector → 叶visible → 条件CategoryName/Value多选 → 含Value时数字format → 字体复合控件；父布局不依赖叶visible。 | 上游5类值节点，含2条content choice；叶隐藏2类可见、默认category4类、含Value5类。本地字体展开5个值控件，标签核心最多9个值节点；mapping另按下行的n层计数。背景类型1、颜色/渐变复合1或pattern3为本地扩展，渐变内部沿P16口径。父、叶、全部3种恢复目标是本地上下文动作。新增常驻/插件根/溢出均0，不改变45目录数；不以控件数证明闭环完成。 |
| P27 Sunburst 上下文标签与中心 | `Jo`先临时选择label/centerLabel，默认label；`Fo`为visible → 条件CategoryName/Value → auto/inside/outside → 含Value时format → font；`jo`为center visible → 条件format → font。 | 上游节点5类、中心3类作者节点，另1个临时目标selector；font复合不重复计叶值。本地用中心、节点两组同时呈现：中心最多7个、节点最多9个作者叶字段，字体各5个；背景类型1、颜色/渐变1或pattern3为本地扩展，渐变内部沿P16。中心、节点、全部3个真实reset动作；节点position没有其他图型位置。没有新增常驻/插件扩展根/溢出或目录项；真实条件交互未验收。 |
| P26/P27 有序层级mapping | 固定`No`展示根→叶序列和逐项上移/下移/移除，添加Select排除已选/权重字段，权重Select排除全部层级；最后一层不可移除。Treemap复用共同层级语义。 | 本地n层为n个持久层级Select及1个权重Select（作者值n+1），另1个临时待添加Select；每层上移/下移/移除3n个动作及1个添加动作，合计显示控件4n+3，含边界disabled。逐层替换与单独添加Button为本地扩展；add/replace/move/remove/weight共用严格互斥校验和一次草稿apply，不另加Ribbon根/溢出/45目录数。 |
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
| 35 | waterfall / waterfall | 一个 Rust `waterfallLayout` job 生成完整跨来源累计、排列/堆积、自动合计和连接端点；C 公共图元按源身份消费角色/逐点样式。 | 12 个 Builder/FChart 角色/connector/stack/subtotal 操作、源 change/total 草稿及独立全局合计样式已写代码；创建/整份替换默认接线由根任务负责。role/dataPoint 在过滤/聚合/源结构变化后的迁移、公开行为、浏览器及原生 Office 往返仍未验收。 |
| 36 | pareto / pareto | 单 signed 权重字段、同名分类先合计再排除负合计、可选保留零合计、稳定降序与 floor 两位累计比例；统计 → C/E。 | 独立柱形/累计线样式、物化点索引、真实百分比标签、19种标记、组合mask、内部位置及纯色/渐变绘制已写代码；共享类型/六个 Facade 操作/Rust 合计已接。源变化索引迁移、公开行为/原生Office往返与运行效果仍未完成或未验收，详见P17/P19。 |
| 37 | sankey / sankey | 名义有向边、重复边累计、max(in,out) 流量；S 一次 SCC/justify/relax 与真实 forward/feedback 流带。 | 源颜色/标签继承、真实流带命中、有向轨迹悬浮、view-only 节点拖动/取消与宿主接口已写代码；循环与任意拖动碰撞、完整公开行为、Worker 大图、Office 和运行验收仍未完成，详见 P20。 |
| 38 | heatmap / heatmap | 数值字段为X列、源分类为Y行，默认首Y分类底部；双轴反向、8段/连续共同色标、value标签与源格式、范围/单位、8个公开操作。 | cell↔legend hover、显式系列/点样式及reset已写代码；浏览器两路径、主题/焦点/IME、Pivot格式、Office与性能未验收，见P21。 |
| 39 | boxplot / boxplot | 每个原始数值字段一组样本；固定默认半组中位数/全样本须线，显式inclusive/exclusive及IQR扩展；统计 → C/X，共享分布、主/次轴、median anchor/标签和原始异常点身份。 | 使用继承的 Cartesian 系列 API及通用options；填充/边框透明度独立，log下保留原统计并跳过非法图元。原始sample dataPoints只作用对应异常点，不套到箱体/轴/图例；有界选择/直接序号/4字段reset已写代码。默认计算主线接线与全部公开/Office效果仍未验收。 |
| 40 | candlestick / candlestick | 最高/最低/收盘，开盘可选；验证 low ≤ open/close ≤ high；C/X。 | 三字段 HLC 不补造 open，走 graphic；OHLC 涨跌/hollow/颜色和 Office 往返。 |
| 41 | histogram / histogram | 单样本系列；Scott 自动分箱、箱数/箱宽、下/上溢箱、间隙；统计 → C/E。 | 分箱/边界规则、源点到箱索引和 Office 统计图。 |
| 42 | treemap / treemap | 有序typed层级、scoped carry、strict signed数值与postorder正树；Worker/Rust hierarchy job，S/X残差面积与父/叶标签。 | 六公开API、专属设置、深层预算与native完整fit/悬停适配已落盘，详见P26；WASM产物、浏览器、源结构点身份迁移、Office和性能未验收。 |
| 43 | sunburst / sunburst | 与P26共用signed层级Worker/Rust；独立语义父角度/正residual、18%–88%环径、首圈切向/后代径向标签及中心total；共有Canvas和native适配。 | 六公开API、独立节点/中心设置、完整fit/gradient/pattern、bitmap祖先hover已落盘，见P27；WASM产物、深树、原生生命周期、Office、Docs宿主和性能仍未验收。 |
| 44 | gauge / gauge | 数值/可选 label；P/C/X 有限观测紧凑 tile、原分类索引、量程/角度/颜色段、9 个样式组、全部专属 setter/reset、5 套预设与原生分层。 | 源数值单元格显示格式元数据仍缺；极端范围 graphic、原生几何/层级、完整交互/Office/性能均待统一验证。 |
| 45 | chord / chord | 显示文字身份的无向边、strict positive finite number/可选计数；反向重复合计、空端点/自连接跳过；Worker/Rust 权重与归一化角度、共享节点扇区和弦带。 | 28 个值、11 个专属 API 操作、8 个恢复动作、真实标签/弦带命中、Tooltip 与只读宿主接线已落盘，见 P28；WASM 产物、浏览器、Office、Docs 宿主、极端角度及真实性能仍未验收。 |

## 4. 数据、统计与边界契约

- `ChartRecords` 保留 `fields/records/sourceCategoryFields/sourceValueFields/visibleFields/syntheticValues`，并可携带按源行/字段对齐的 `formattedRecords`；数值系列可携带同记录投影的 `formattedValues`。普通 Sheets 按列时首行为字段、按行时首列为字段；源表 `source.sheetId` 独立于宿主 `sheetId`，旧快照缺源表 ID 时使用宿主。原始 `rawValue: null` 不能退回格式化显示值，数值保留未舍入精度。格式文字不进入统计Worker；Pivot格式元数据仍缺，详见P21。
- `projectChartRecords` 统一普通、透视和预加载源；先校验字段/辅助引用，再处理可见性与空数据。透视数值角色由源声明，数字外观的分类 ID 不自动变为 measure；图关系端点仍是名义类别。隐藏辅助列不能被自定义误差读成零。
- `plotVisibleOnly` 缺省为 true，显式 false 包含隐藏数据；设置预览和 Switch 使用 `!== false`，与根任务 source-compute/Rust 路径统一。手动隐藏、零尺寸、filter/outline 的实际消费和 Office source helper 仍需统一验证；预览草稿覆盖双方向隐藏行列与开关后的重投影，不能由注释推导全部运行行为已验收。
- 数字/可解析数字字符串进入数值样本，非数值/非有限值作为 null；`zero/break/link` 分开处理。Gauge 无效观测不被 `zero` 策略伪造为零仪表，继续跳过。有限累加溢出显式拒绝；不删除权威源记录或降精度以获得有限场景。空/隐藏后没有可绘制数据的现存图表，全部目录均保留标题/框架和空态；创建时的必需数据约束仍校验。
- 堆积按同图型/同轴分别维护正/负栈；百分比目前以每个分类的绝对值合计为分母，保持正负符号。此为本地明确契约，尚未通过上游/Office 样本验证。普通堆积、组合堆积和次轴堆积都要求线性数值轴。
- Bubble 可省略 size，默认 1；零/null 权重隐藏，负大小拒绝。默认直径按 area/width 和 scale 缩放，显式 marker size 为绝对直径；无有效绘图区隐藏。每组保留同一分类长度的独立 X/Y/size 稀疏向量。
- 数值系列 selector 保留原有源字段语义；名称 selector 按原始系列名或 Bubble 分组字符串精确匹配，大小写、空白及空名称均保留。显示名由 style.name 单独覆盖；全系列 → 字段 → 名称继承不把显示名反写为源身份。未活动名称覆盖保留，活动源名称匹配多系列时拒绝；这是本地可测试规则，P13 的上游重名和隐藏边界仍未核实。
- 图关系跳过空端点与零/null 权重，负权重拒绝，重复边累计；Chord 将端点对归一为无向边；Relation/Sankey 保留方向和 cycle 标记。层级独立使用P26规则：仅null/undefined/空字符串为空，祖先范围内carry且先更新路径cache再检查权重；只接受finite number，零/负进入同路径signed selfValue，postorder保留正子树，最终非正节点不物化。层级typed path ID、首次创建sourceOrder、value降序与独立根paletteOrder保留；Treemap额外生成正父残差，Sunburst不借Treemap面积树改写语义模型。
- HLC/OHLC 字段不重复，缺样本保留 null；最高/最低/开盘/收盘关系非法时拒绝。三字段 HLC 直接绘高低线/收盘短线，没有伪造开盘值。
- Pareto 的 TS 小数据契约接受 signed 源行，同名分类先有限累加，再排除负合计；零合计缺省排除、includeZeroValues 时保留，等合计按首次源分类出现顺序排列。累计百分比使用缩放合计，正总量的最后点为 100，全部零合计时为 0，不改源行。Rust 的当前同步缺口见 P17；两计算路径尚不能宣称一致或验收。
- `createChart/updateChart/readChartData` 接受可选预加载 records；更新先准备候选、转换和全部校验，再修改模型。显式字段/映射 patch 优先于类型转换的默认角色。预加载能力已写契约草稿；生产命令的大源准备/复用还需与根任务计算服务接线一起验证。
- Boxplot、Histogram、Pareto、Waterfall、Trendline、ErrorBars 有共享统计契约。根任务本批实现 `crates/office-compute/src/chart/`、`chart_wasm.rs` 与 Worker 协议/planner/decoder/service；Rust/WASM、TypedArray 复制、取消与数值正确性均未执行验证。**Waterfall 非空数据只消费完整 Rust `waterfallLayout`，没有 TS 累计回退；缺计算结果时同步布局/绘制 API 抛 `COMPUTE_NOT_READY`，实时 Facade 的 `getComputedLayout/getComputedRenderCommands` 等待服务并克隆结果。** 其他统计仍有显式小数据/同步参考路径，阈值和端到端性能必须基准后确定，不能宣称全部重操作已迁移。

### 分组统计身份

`ChartStatisticsSeries.seriesPosition` 是 `preparedData.series` 下标，必填；输出 segment 保留同一位置。job 的 `slot` 是作者配置下标，同 kind/slot 可以对不同实际系列生成多个 job：

| 结果集合 | 身份与结果字段 | 消费规则 |
| --- | --- | --- |
| trendlines | `{ configIndex, seriesPosition, statistics }` | `chart-derived` 按配置 + 位置匹配，`chart-layout/chart-draw` 按位置取轴和数据；不能只 find(field)。 |
| errorBars | `{ configIndex, seriesPosition, intervals }` | 同上；每组独立 X 与自定义误差列，保持真实 categoryIndex。 |
| boxplots | `{ field, seriesPosition, statistics? }`；statistics含五数、平行`outliers/outlierSourceIndices`。 | 空/不可用样本仍保留项，statistics undefined不触发同步重算；默认单样本半组不可用。IQR异常值按(value,原prepared源索引)稳定排序，null槽位保留索引；decoder校验原值/索引/重复与同值顺序，渲染只消费平行结果，不重新排序。 |
| waterfallLayout | `{ stackType, useSubtotal, sourceCategoryCount, seriesFields, points, connectors, cumulative }`；点带 `seriesPosition/dataPointIndex/categoryIndex/synthetic`。 | 一个 job 覆盖全部当前来源；排列按真实来源位置扩展分类，堆积全局合计使用 `0xffffffff` sentinel，不放入任何来源 series。旧 `waterfalls` 分组协议只保留兼容，生产绘制不使用它替代完整 scene。 |

显式提供集合却缺当前配置/位置对应项时拒绝，避免漏组或退回同步统计；空 points/intervals 保留合法匹配。根任务 planner 为实际 job/列/输入/输出设置预算，decoder 逐值/索引验证并可分块让出。绘制 extent 收集不再向 push 展开大批异常点、瀑布或误差/趋势数组；这只是代码修正，没有实测性能结论。

## 5. 配置、转换、放置与宿主历史

[chart-settings-draft.ts](../../packages/office-studio/src/features/charts/chart-settings-draft.ts)、[chart-settings-data.ts](../../packages/office-studio/src/features/charts/chart-settings-data.ts)、[ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) 及 option/series/statistics/collection 控件是共享面板。普通插入、选中图表 Drawer 和 PivotChartPanel 均使用同一字段/类型目录与草稿。没有增加假按钮、空操作或永久禁用图型替代实现。

| P01 状态/配置类别 | 本地持久字段与实际路径 | 差异与未完成项 |
| --- | --- | --- |
| ChartType、Mapping/MappingFields | `type/source/categoryField/valueFields/seriesToCategory/options.mapping`；projection/settings/type-conversion。 | 字段角色已区分；同源转换、跨形状点索引迁移及所有公开 mapping 细项未验收。 |
| Appearance、Title、Subtitle | `options.appearance/titleStyle/subtitleStyle`、title/subtitle；共同字体/背景/边框。 | 专属位置/富文本等公开细项待逐项追踪，不能由一组字体控件宣称完整 section。 |
| Theme、Legend | palette + 统一 theme token；`legend.visible/position/font`。 | 上游命名主题未完整表达；长图例、复合饼 aggregate 图例仍缺。自动渐变另列。 |
| AllSeriesStyle、Series | `allSeriesStyle/series[field]/namedSeries[{name,style}]`；name/type/color/fillOpacity/fillPattern/border/point/rightAxis/label/dataPoints，nullable 单字段恢复与 6/7 个 clear target。 | Bubble名称组及Boxplot分布消费同一继承样式；Boxplot原始sample索引覆盖对应异常点，汇总不消费sample0。SolidSeries纹理、null清纹理和undefined恢复继承已接。两集合可独立清除，clearAllSeriesOverrides同时清空；过滤/聚合/源结构变动后的作者身份迁移、完整入口与Office仍未验收。 |
| Axes、DepthAxis | xAxis/yAxis/rightYAxis/depthAxis；线/刻度/格式/标签/反向、标题字体及对齐，8个clear target、三轴标题恢复、partial次轴与显式linear清除logBase；日期序号/calendar坐标、baseTimeUnit/majorTimeUnit/majorUnit、数值majorUnit与独立reset、真实指示器。 | 横向Bar次轴是顶部数值X；depthAxis独立且无min/max/format。普通/Heatmap/Radar标题复用共同物理边缘几何，native轴名由graphic标题替代并保留已有色标；3D投影消费共同字段/helper。物理起止已由wke静态核实，旋转式竖轴与上游逐字竖排不同。P22已静态追踪月/年桶、主单位推进与重复日期保留；本地三时间字段和次轴reverse是扩展，完整碰撞/多轴/UI/Office仍未验收。 |
| Trendlines、ErrorBars、ErrorBarFields | 独立作者配置、字段引用、Worker 结果身份；样式/公式/R²、自定义误差。 | 支持 2D 非堆积 Cartesian/Scatter/Bubble/非堆积组合系列；3D/堆积不生成统计图元，保留配置以便换型，不冒称完成未知上游增强。 |
| ThreeD | view 的 rotX/rotY/perspective/rightAngleAxes/heightPercent/depthPercent + grouping/gapDepth。 | 模型、绘制和独立清除已接；公开投影/遮挡/透明排序仍未验收。本地 3D 次轴是增强，上游 P04 不暴露此 API。 |
| Pie、Bar、Area、Radar、Funnel | options 下同名专属配置；C/D/P/E/X。 | 配置实际影响绘制；公开完整 section、组合边界仍待逐项验收。 |
| Aggregation、InvalidValueStrategy、Pareto/ParetoDataPointOptions | aggregate/topN、空值策略、pareto.includeZeroValues/barStyle/cumulativeLineStyle。 | Pareto signed合并、独立样式/组合标签/物化点、六公开操作与Rust同步已落盘，证据见P17/P19；没有运行验收。过滤、聚合、前N与源变化之后的作者点索引迁移仍未完整实现。Histogram另列P25。 |
| Histogram | 单样本字段、`histogram.binning/underflowThreshold/overflowThreshold/binGap`；11个专属Builder/FChart操作。 | 1000总箱预算、Count/Width/Scott Auto、恒定点箱、原始null保留、严格Rust decoder、同次统计设置数据/物化分类、真实reset与native/common gap已写代码，见P25。极端有限数值采用本地稳定算术；分箱后作者点索引的源变更迁移、浏览器、Office与性能仍未验收。 |
| Waterfall | pointRoles、stackType、useSubtotal、connector 与全局/来源/名称角色样式；12 个公开 Builder/FChart 操作。 | 完整跨来源 Rust scene、真实来源与 synthetic identity、草稿应用/恢复、独立全局合计和 graphic 已写代码；创建/整份替换默认与普通 patch 的边界见 P18。源结构、过滤/聚合后角色身份迁移及公开/Office 验收仍待完成。 |
| Candlestick | 显式HLC/OHLC映射、独立rising/falling/doji、candleWidth、highLowLine/closeMarker；14个专属Builder/FChart操作。 | 三状态颜色/alpha/outline/空心、绝对正宽度、全空源角色、源单点覆盖、HLC及OHLC overlay、专属Form与真实reset已落盘，见P25。OHLC overlay面板是本地扩展；shared graphic/native转换差异、完整图内交互与Office仍未验收。 |
| Heatmap | `mapping.heatmap/heatmap`；8个专属Builder/FChart操作、双category轴、范围/8段/单位/value标签、源formatted文字、公共色标与hover。 | 范围/label/type/unit的nullable/reset、作者系列/点paint、hover不写内容历史及内部JSON重开均有草稿；Pivot格式元数据、最终公开交互、原生Office与两后端性能未验收。 |
| Treemap | `mapping.hierarchyIndexes/valueFields[0]/treemap.parentLabelLayout/treemap.label`；六专属Builder/FChart操作、父叶独立设置、背景union与reset。 | strict signed层级Worker/Rust、parent-owned residual、palette与完整label fit、bitmap hover接线已落盘，见P26。原生适配依赖锁定引擎内部结构；WASM、主题/焦点/IME、Office和真实大样本未验收。 |
| SunburstFields/HierarchyIndexes/ValueIndex/Style | `mapping.hierarchyIndexes/valueFields[0]/sunburst.label/centerLabel`；六专属Builder/FChart操作、独立中心/节点组与真实reset。 | strict signed层级复用P26，父语义角度/正residual/18%–88%环径/完整fit/tooltip、native alpha/pattern与命令/history/internal reopen已写代码，见P27；大样本、WASM、主题/焦点/IME和Office未验收。 |
| ChordFields/Mapping/Style、Relation、WordCloud | 边映射、ribbon/节点布局、词云形状/字号/旋转。 | Chord/Relation/WordCloud完整公开语义、配置/交互和布局Worker/预算继续闭环；字体和性能未实测。 |
| GaugeFields/LabelIndex/ValueIndex/Style | `mapping.gauge/gauge`；scale/angles 映射、ticks、ranges、axisLine/axisTick/splitLine/axisLabel/pointer/anchor/progress/title/detail、预设/清除、tile 与标签继承。 | 固定声明的 setter/reset 已表达并消费；本地扁平量程/角度与 null reset 扩展明确记录。源 numeric cell 显示格式、全部 UI/原生/Office 及性能未验收。 |
| AxisPointer | 四个作者字段、`setAxisPointer/clearAxisPointer` 与独立恢复目标、nullable patch、真实悬停线和轴标签、Canvas 位图版本更新及离开/拖拽/滚动/切表/卸载清理。 | 普通坐标走ECharts原生动作；Heatmap、公共图元图及Radar走独立graphic组，Heatmap另可在plot外的真实色标命中。缓存内容布局而不在悬停重做数据统计。上游默认颜色/完整Radar交互、浏览器效果及性能仍未验收。 |
| AutoGradientFill | `options.autoGradientFill`；共享默认 helper、Builder/FChart set/reset、有效 description、antdv Switch/default reset、Canvas/native/3D 派生 paint。 | 25 个目录项显示开关，另 4 个 Line 目录项可经公开 API 启用；换型保留不活动覆盖。实际纹理优先，alpha 只乘一次；闭源色标/颜色空间、全部组合/最终 Office 效果未核实、未验收。 |
| PlotArea | `options.plotArea { left,top,width,height }`；编辑 seed、完整矩形校验、5 个本地值控件、公共 chartPlotFrame/布局/指示器/native grid。 | 31 个目录项使用同一归一化矩形，按原外框位置/尺寸重建；关闭清配置，换型保留不活动值。手动边界限制为本地决定，上游四字段 UI/最终拖拽和原生 Office 布局仍待核实，未验收。 |

`chart-type-conversion.ts` 保留源、作者外观和非当前图型配置，按实际数值角色选 X/Y/size、edge、hierarchy、HLC/OHLC 和 gauge 字段；跨百分比变体清除不适用的固定轴边界。堆积/Radar 切换为线性轴。数据角色转换不等于所有 point overrides/作者索引已成功迁移。

`chart-placement.ts` 复用 Canvas 的 SpreadsheetGeometry 行列轴，考虑显式/跨度行高列宽、自动行高、隐藏、折叠、筛选与表格过滤。Builder 的 cell shortcut 保留在 ChartConfig，真正创建时用当前宿主轴解析；持久 ChartModel 保存绝对文档坐标。更新/重开不把绝对位置重新解释为单元格。表格筛选读取仍可能重扫源，尚无大样本/Worker 预算证据。

`ChartConfigBuilder/FChart` 提供类型/字段/映射、系列/轴/统计/options 与布局/绘制 API；`setPosition` 使用上述放置入口。`threeD` 与 `depthAxis` 可独立清除且保持对方。内容更新统一经过 `sheet.command.add-chart/update-chart/remove-chart` 和工作簿历史，内部保存/重开可携带完整模型；新接口草稿覆盖失败原子拒绝、撤销/重做和快照重建，全部未运行。命令级宿主/源表权限、复制/粘贴与所有结构变换需完整验证。

`use-sheet-chart-panels.ts` 捕获 unit/studio/sheet/会话身份；切换、关闭工作簿、更换 studio、scope dispose 时释放会话及 8 个事件监听器，旧异步完成/错误/closed/applied 回调不触碰新会话。源读取失败保留图表模型和错误状态。关闭按活动上下文、其他面板和 focus eligibility 恢复焦点；实际主题、焦点、输入法与浏览器交互未验证。P24已把编辑/Pivot预览的同步FChart.getData替换为detached配置与异步source/statistics读取，P25使Histogram也消费同次统计；大源真实耗时、取消/复制与完整宿主生命周期仍未验收。

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

[chart-radar.ts](../../packages/office-studio/src/features/charts/chart-radar.ts) 统一范围、角度、半径、刻度和数值标签继承。P11追加固定ES静态追踪后，已修为5段/6个含边界刻度、数值标签默认显示/12px；两绘制适配器使用同一signed线性min/max，支持固定负范围、分类反向和径向反向。最终半径按bje使用plot短边/2减图表fontSize和12px，作者PlotArea使用完整半径；小plot扣减至0时不生成指示器，避免0/0。缺样本保留源分类索引，有限极端范围使用稳定比例并走共享graphic。数值标签字体依次继承12px默认、图表appearance、历史 `yAxis.label` 和 `radar.axisLabel` 的已定义字段；缺省格式继承 `yAxis.format`，undefined不遮掉继承。指标名称/旋转/字体使用 `xAxis.label`，与数值标签分开。

`ChartConfigBuilder/FChart` 已落盘 `setShape/resetShape`、`setFill/resetFill`、`setRadarAxisLabel/resetRadarAxisLabel`，以及本地 `setRadar` 配置入口；FOffice `.Enum.ChartRadarShape` 使用 `ChartRadarShapes`。resetShape/resetFill 清除对应作者覆盖，使有效默认恢复 Polygon/仅轮廓，resetRadarAxisLabel 清除专属外观后恢复有效继承。专属 patch 使用 `ChartNullablePatch`，null 单字段恢复归一为显式 undefined，并经普通对象/枚举/有限值校验与候选更新；错误图型、log 径向轴和非法 patch 拒绝，命令失败不发布模型/历史。对象隔离、历史及内部重开只有未运行草稿。

[chart-radar-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-radar-settings-controls.ts) 的3个基础控件按P11排列；有效visible为真时显示format和5个字体控件，隐藏时保留作者字段。shape/fill/label整组和6个label叶字段均有真实reset，恢复继承且不改其他作者配置。控件由共享 `antdv-next` Form/Input/InputNumber/Select/Switch/ColorPicker/Button承载；指标/径向轴只呈现对应能力，径向轴不提供log。固定默认true已核实并修正，但历史 `yAxis.label.visible=false` 仍作为本地兼容继承保留，专属visible优先；固定xje没有该legacy继承。泛用轴样式、指标旋转和径向reverse仍是本地扩展，不把默认修正宣称为全交互已验收。

Native Radar已接统一范围、分类方向、半径、圆形/多边形网和实际刻度格式/字体；数值标签显式margin4/hideOverlap true覆盖native缺省，共享Canvas也使用4px间距。分类旋转、径向反向、单点覆盖/值标签或跨度溢出的有限范围切到graphic。新增6个默认/nullable/5段/extreme/半径/零半径/9个reset/共同paint纯契约草稿及2个默认/隐藏字段/一次apply/cancel/reset/history/internal reopen工作台草稿，既有native display list的25.0刻度断言已改为5段的20.0。全部未执行；graphic尚无完整字体碰撞/hideOverlap，与native效果差异、主题/焦点/IME、两后端、公开演示及Office往返保持未验收。

### Boxplot 的原始分布、样式和对数轴边界

[chart-boxplot.ts](../../packages/office-studio/src/features/charts/chart-boxplot.ts) 以原始 `field + seriesPosition + sourceName` 选择分布样式、主/次数值轴及几何；每组继续输入原始样本，由共享统计契约提供五数和异常点。boxes保留原始sourceName，作者name仅改显示标签；不能把布局下标当作者字段，也不把五数行当原始样本。P23已静态确认固定默认为上下半组中位数/全样本须线，没有独立异常点；本地Rust/TS默认resolver及其调用点已落盘，显式inclusive/exclusive及IQR保留为Office扩展。dataPoints索引现表示prepared源样本，仅消费对应outlierSourceIndices的异常点，箱体/轴/median/汇总图例不套sample0覆盖。SolidSeries纹理/空心形状/隐藏/大小/颜色/alpha及独立边框走同一继承；完整公开/Office行为未验收。

Canvas箱体、median、须线及完整箱宽cap消费分布样式；每个异常点再合并自己的原始样本覆盖，填充透明度不衰减独立边框。矩形异常点单独绘填充及边线，以保留border opacity/dash。Native异常点ID为chartId/seriesPosition/原始sample索引，重复异常值仍分别保留；仅实际活动且可绘的异常点纹理/共享形状触发graphic，inlier/hidden/showOutliers:false覆盖原样保存。主/次轴extent、median anchor和标签使用同一比例，空/单样本半组分布保留位置。带median标签时使用graphic；本地native仍以一条Boxplot系列内的多个分类表达分布，区别于Xce每字段单独系列，最终分组视觉未验收。

log 轴下，非正或非有限五数不能绘制完整箱体/须线/median/标签；原统计保持原值，**不从正样本重新计算四分位**。合法正异常点可独立显示并参与该轴范围，非法异常点省略。Native adapter 将不可绘制汇总映射为空数据项，保留分类位置；这与原始统计、界面源数据的保留分开。主/次轴、log 跨零汇总/正异常点、独立透明度与 native 端点的草稿均未运行。本地次轴 reverse 是对 P12 `rightY` 不含 reverse 的扩展。

### Cartesian 继承的 nullable 更新与恢复目标

[chart-cartesian-api.ts](../../packages/office-studio/src/features/charts/chart-cartesian-api.ts) 提供 `ChartSeriesClearTargets` 7 项、`ChartAllSeriesStyleTargets` 6 项和 `ChartAxisTargets` 8 项；Builder/FChart 与 FOffice Enum 接同一常量。`clearSeries/clearAllSeriesOverrides/clearAllSeriesStyle`、`clearXAxis/clearYAxis/clearRightYAxis` 及三轴 title setter/clear 已落盘；clear 单目标保留其他作者字段，全组恢复移除对应覆盖。系列 nullable patch 复制输入后将 null 字段归一为 undefined；轴 setter 保持 P12 的 partial 字段契约，通过显式 undefined 恢复字段。后续合并保留该恢复意图，不再从 getter 的 JSON 克隆丢失它。

次轴 partial patch 使用有效 linear 默认；显式改为 linear 同时清除旧 logBase，防止继承的 logBase 令候选校验拒绝。所有入口继续走共享 `update-chart`、内容历史与内部快照；新增 Radar/Cartesian 4 个与 Boxplot 1 个 Facade 草稿覆盖恢复目标、部分次轴、继承、非法输入原子拒绝、撤销/重做和重开，未执行。名称 selector 已按下一节落盘，但本地数值 selector 仍表示源字段，不能把它等同上游内部 index 的全部转换契约。轴 titleStyle/titleAlignment、useDateAxis/AxisPointer 的当前代码与边界见后文；SolidSeries fillPattern 的代码/差异见 P14 小节。

### P12 轴标题、日期刻度与真实指示器

[chart-axis-title.ts](../../packages/office-studio/src/features/charts/chart-axis-title.ts) 合并图表字体与轴标题的已定义字段，nullable恢复继续继承，缺省字号12px。追加静态读取固定 `engine-chart/lib/es/index.js::wke` 后确认：horizontal start/end写D.left/D.right，vertical start/end写D.top/D.bottom；该定位分支不消费axis.reverse，center扣文字盒宽/高后居中。共同helper据此把span按物理左→右或上→下排序，标题反向轴不移动，旋转文字选择对应锚点使内容留在span内。普通坐标、横向Bar真实category/value边缘、活动次轴顶部/右侧、Heatmap和Radar复用 `chartRectangularAxisTitleCommands`；3D使用其已有投影span及共同单标题helper，3D最终语义另待公开样本验证。

两绘制路径采用同一标题DrawCommand、有效字体、plot span宽度截断及frame裁剪。`echarts-style.ts`关闭native轴name，`echarts-option.ts`追加共同graphic标题，保留Heatmap色标和已有graphic；指示器仅修改独立group，标题不进入内容历史。主/副标题仍继承有效字号并预留文字区域。这些是代码实现；没有运行Canvas/native像素比较或性能测量。

固定wke在竖轴没有显式rotate时把字符串逐字换行；本地采用旋转式文字标题，未复刻逐字竖排、作者manualLayout/rotate及完整字体碰撞。物理起止定位只有固定ES静态证据，完整公开效果、布局边界和Office行为仍未验收。原生DrawingML段落对齐与图内标题盒位置不同，Office来源/诊断及未核实项见 [SHEET-29](SHEET-29-native-charts.md)。

[chart-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-settings-controls.ts) 在每个可见轴按标题文字→对齐→字体5叶字段排列，提供7个真实reset。文字reset清标题/字体/对齐整组，字号/字体/颜色/粗体/斜体/对齐独立恢复；表单显示图表继承后的有效字体，不写回默认作者值。全部复用antdv-next Form/Input/InputNumber/Select/Switch/ColorPicker/Button。7个独立契约、2个实际renderer和2个workbench草稿覆盖物理起止/reverse/Bar次轴/字体/Heatmap色标保留及编辑→一次apply→历史→内部JSON重开；**全部未执行，主题、焦点、IME和Office均未验收**。

[chart-date-axis.ts](../../packages/office-studio/src/features/charts/chart-date-axis.ts) 使用现有 Sheets 日期系统：支持 1900 序号 0/59/60/61、1904、有限合法序号与严格 ISO 日期/时间；真实日期的 UTC offset 跨 1900 边界跳过虚构日，显式 1900-02-29 保留序号算术。缺失日期省略对应点；非法文本、非法边界、无日期/数值配对及 log 日期轴报错。`categoryValues/categoryDateValues/date1904` 为派生输入血缘，模型只保存 `mapping.useDateAxis` 和轴作者设置。基本单位日按原serial实际间距；月/年按连续calendar桶布局，major单位只改刻度。普通Cartesian稳定日期排序，保留重复日期、原categoryIndex、源样式及缺失位置；月份/年份和重复源行的固定实现已静态核对，具体来源、公开API缺字段及本地差异见P22，运行效果仍未验收。

候选/API 与表单共用 `prepareChartDateSource`。Bubble 的 ISO 日期在候选校验时转成 X 序号；Boxplot 日期分类是最终系列显示名，预检按名称及有限样本配对，不在表单计算四分位；Histogram 的空源预检检查日期字段和边界，最终时间坐标来自真实分箱中心，不能由预检代替最终layout/tick检查。Pareto 保留原始序号到最终权重顺序，同名分类合并出现不同日期身份时拒绝。主线已在隔离统计候选的结果发布前消费最终Rust统计并调用layoutChartData；命令原子发布的实际范围和缺口见P22。开启日期轴的 UI 会把已有 log X 改为 linear 并清除 logBase，这是本地转换决定；公开 API 仍对不合法组合报错。源工作簿 `date1904` 已从 `ExcelWorkbench` 接到 Insert/Drawer/Fields，不从展示字符串猜测日期系统。

[chart-axis-pointer.ts](../../packages/office-studio/src/features/charts/chart-axis-pointer.ts) 的 capability 按 `qi/Hi` 的 IndicatorLine 位判断：column/bar/line/area/combination/scatter/bubble/radar/waterfall/heatmap/boxplot/pareto 可用；3D、Candlestick、Histogram 以及无坐标图不显示该面板。四个字段及四个恢复目标与 P12 类型一致；nullable 清除保留其他作者字段，完整清除恢复主题默认。普通/日期/反向/log/主次轴、Heatmap 行列及 Radar 几何生成指示线和坐标标签，标签约束在图表 frame。Radar 径向线与值圈是本项目独立实现，**上游具体交互未核实**。

`echarts-renderer.ts::updatePointer` 在 native 路径转发 `updateAxisPointer(mousemove/leave)`；graphic 路径仅 replace/remove 独立 group，保留内容图元。两个 pointer 配置位置同步作者线色、线型、底色和文字色；专属 tooltip 同样保持离屏。`ChartOverlay.vue` 的 `antdv-next` Button 接 pointermove/leave/blur，坐标只处理 CSS、zoom、scroll 与对象原点，DPR 由 raster surface 处理；同一 RAF 合并内容/指示器，纯悬停复用内容命令和布局。指示器错误保留到成功悬停或切表，不被下一内容帧清掉；pointercancel 丢弃移动/缩放草稿。清理与缓存是代码实现，**没有真实性能/内存数据，不能声称已优化或达标**。主题、焦点、IME、native/graphic 悬停与浏览器/Office 重开均未执行。

### Bubble mapping 与名称覆盖

`ChartBubbleMappingOptions` 明确 X/Y 的必选数值字段及可选 size/label/series 字段，Builder/FChart 的 setBubbleMapping/clearBubbleMapping 更新真实 mapping；clear 保留其他 mapping 和既有 valueFields，随后恢复共享缺省角色。错误图型、缺失必选字段或越界映射通过候选校验拒绝，不写空入口。

[chart-series-selector.ts](../../packages/office-studio/src/features/charts/chart-series-selector.ts) 与 [chart-cartesian-api.ts](../../packages/office-studio/src/features/charts/chart-cartesian-api.ts) 实现 `ChartSeriesSelector = number | { name: string }` 的校验、独立读写与清除。`namedSeries` 使用 `{name, style}` 的密集数组，显式配置替换整个数组，未提供则保留；不把任意名称用作对象键。重复作者名称拒绝；活动源的同一名称匹配多个系列时拒绝名称覆盖，数值字段路径仍可使用。没有匹配的名称保留为未活动配置，以便隐藏/消失的分组重现后恢复。大小写、空白、空字符串及 prototype 字样都是源名值；style.name 只改显示，不改变目标。这些重名、未活动与空名称规则属于本地实现，P13 未核实的上游边界保持待核实。

共享样式按全系列 → 数值字段 → 名称逐层继承，border/point/label 及按分类索引的 dataPoints 保留下层未覆盖字段；nullable 恢复后重新继承。Canvas 与 native 适配器的颜色、标记、标签、图例、主/次轴和组合系列图型均消费同一源身份。Boxplot 合成显示数据另保留 sourceName，避免从作者标签反推选择器。主次轴边距与 graphic 路径按活动系列的有效样式判断，未活动次轴配置不额外挤占绘图区；Worker 统计能力判断同步传源名，preparedData/数值 bulk 协议未增加字段。

[ChartSeriesSettings.vue](../../packages/office-studio/src/features/charts/views/ChartSeriesSettings.vue) 通过 `antdv-next` Select 同时列字段目标和每个原始名称目标，保留未活动名称的编辑/恢复入口；未配置的重名目标不可选，已有重名覆盖可进入整组恢复，样式和单点编辑停用，面板同时显示绑定错误。名称级 Bubble 单点选择限定到该组的源行，数值字段目标仍覆盖该字段全部分组。系列控制、单点控制和恢复均写作者配置；尚未执行 DOM、焦点、主题或 IME 验证。趋势/误差线的作者 selector 仍为本地 seriesIndex 字段，名称仅参与类型/能力判断，不冒称已经支持 P12 内部统计名称 selector。

本批新增 10 个独立 selector/继承/投影/布局/UI/真实 native scatter SSR 契约草稿，另有 3 个 Facade 草稿覆盖 mapping setter/clear、输入隔离、7 个名称恢复目标、原子拒绝、源分组消失/重现、撤销/重做和内部保存重开；全部未执行。Native renderer 的 series.name 用于 tooltip、legend 与更新，重复显示名的 legend 映射可为多对多，参见 2026-11-24 通过 Context7 查阅的 [ECharts series name](https://github.com/apache/echarts-doc/blob/master/en/option-gl/partial/series-name.md) 与 [legend matching rules](https://github.com/apache/echarts-doc/blob/master/en/option/component/legend.md)；动态文档不是固定 Univer 运行时证据。本地独立源名称不依赖此显示名映射；完整图内 legend command/事件闭环及 Office 动态分组行为仍未验收。

### P14 SolidSeries 纹理、继承与绘制差异

[chart-fill-pattern.ts](../../packages/office-studio/src/features/charts/chart-fill-pattern.ts) 表达 `ChartSeriesFillPattern { preset, foregroundColor, backgroundColor }`，两种颜色均必填并校验；[chart-fill-pattern-geometry.ts](../../packages/office-studio/src/features/charts/chart-fill-pattern-geometry.ts) 独立生成 54 种 DrawingML 风格纹理。依据是 P14 的公开类型和 `createSeriesPatternStyle` 转换声明；本批未沿固定 ES 核验纹理实际像素，这些独立 mask 不声称复刻上游像素。

全系列 → 源字段 → 源名称 → 数据点使用既有样式继承链。`undefined` 恢复上层纹理，`null` 明确无纹理；完整纹理对象原子替换，避免新 preset 继承旧来源的额外字段。系列颜色控件同时写 `fillPattern:null`，这是本地编辑规则；颜色恢复、纹理恢复和自动渐变布尔值互相独立。界面的“无纹理”仅清除纹理，自动渐变仍按 P15 的有效配置执行。`setSeries/setAllSeriesStyle`、逐点 patch 与 clear target 复用既有命令、历史及内部 JSON 保存，未添加专造纹理 Facade 方法。

16×16 tile 的前景/背景 mask 合并成互不重叠的矩形并缓存；多边形先三角化，再与 tile 矩形相交，填充透明度只应用一次，边框独立绘制。柱体、面积/标记、Pie/环扇、Boxplot 箱体/异常标记、Scatter/Bubble、Heatmap、OHLC 实体、图节点/流带、层级和 Gauge 指针/进度等实际填充图元接入公共几何；3D 面采用同一纹理加面阴影。纹理按当前图元边界平铺，跨分段相位和闭源纹理精确缩放未核实。绘制命令保留字符串颜色，未把 ECharts decal/CanvasPattern 保存为权威模型。

活动系列或数据点含实际纹理时，大多数ECharts适配器转为共享graphic，避免原生Scatter/Heatmap的decal能力差异；Treemap/Sunburst例外，使用同一不相交mask的同步Canvas tile保留native层级布局/悬停，见P26/P27。全系列/字段/名称/逐点面板均使用 `antdv-next` FormItem、Select、ColorPicker 等；54 个选项不是完成证明，全部最终显隐、主题、焦点/IME、Canvas2D/WebGPU、公开行为与原生 XLSX/DOCX/Office 往返仍未验收。契约草稿覆盖 mask 唯一性/覆盖率、透明区域不重叠、凹面/环孔裁剪、继承和原子拒绝；没有性能测量，缓存与几何预算不得描述为已达标优化。

### P15 AutoGradientFill 与归一化 PlotArea

自动渐变的固定声明位于 `engine-chart/lib/types/chart-builder/configuration/auto-gradient-fill.d.ts:2,4`、`facade/f-chart-builder-base.d.ts:377,390` 和 `facade/f-chart.d.ts:441,453`。`chart-builder/chart-types.d.ts:554` 的 semantic 配置是可选布尔值，`:628–632::IChartDescription.autoGradientFill` 是应用图型默认后的必选布尔值。上游底层 ChartBuilder 在 `chart-builder/chart-builder.d.ts:22` 使用 `describe()`，live adapter 在 `chart-live-adapter.d.ts:5` 有同一描述契约；本项目 `getDescription()` 是由根任务接线的本地新增接口命名，不把名称说成固定版同名 API。上游完整描述另有 data-free 系列/分类 metadata 和 semantic-axis bindings，本地映射与名称 selector 差异须分别记录。上游 setter 的同步 chaining 与本项目 live facade 的异步命令返回继续区分。

[chart-auto-gradient.ts](../../packages/office-studio/src/features/charts/chart-auto-gradient.ts) 统一有效默认：Area 族开启，其余关闭；显式 false 保留，reset 删除覆盖并重新取当前图型默认。`chart-ui::Oi/Qo/Zo` 支持自动渐变开关的 25 项与实际渲染预设可支持的 29 项分别记录：额外 4 项是 Line，可通过公共 API 启用；不支持的族保留作者值但不消费该预设。Switch/default reset 编辑隔离草稿，经一次应用进入共享 `update-chart` 和内容历史。上游 Builder 注释还明确作者系列/扇区的显式渐变色对象优先；本轮 P16 已扩展 ChartColor 作者模型和绘制路径，两种配置继续独立，完整公开行为仍未验收。

自动颜色由有效系列/逐点颜色派生，实际纹理优先；Area 采用有效透明度的 6%→100%，柱/线采用 +0.28/−0.08 色调，Pie 采用 +0.28/−0.12。它们是本地独立预设，上游闭源色标及颜色空间仍待核实。负值、数值轴反向和横向柱以绘制的基线/顶端选方向；Area/Line 分片共享绘图区坐标，扇区共享外圆坐标，3D 柱六个面按投影后的基线/顶端对齐并各自应用阴影。ECharts 原生线性颜色使用可序列化对象；全局坐标为图表像素，未指定公共边界的柱颜色为图元内归一化坐标，透明度嵌入色标后原生 opacity 为 1。Line/Area 作者逐段颜色/不透明度覆盖转为共享 graphic，保持同一有效样式。

[chart-gradient-geometry.ts](../../packages/office-studio/src/features/charts/chart-gradient-geometry.ts) 对自动预设在 0–1 内使用 48 等分色带及两个端点钳制带，显式作者色标另见 P16。裁剪保留几何面积，退化方向使用中间颜色。`262_144` 指令预算检查的是**每次渐变 append 的当前目标数组加待提交输出**，失败该次 append 不提交；后续边框/图例、分次调用的扇区和单独生成后合并的 3D scene 不受这一完整场景上限保证。完整场景/Worker 预算、取消、实际传输和内存仍待完成及测量。`chart3DScene(..., emitCommands=false)` 可只产生命中/布局坐标，不生成面 paint，3D Pie 在此模式也跳过只供渐变绘制使用的 sector bounds；循环合并数组避免函数参数展开，但没有基准，不作性能结论。

[chart-plot-area.ts](../../packages/office-studio/src/features/charts/chart-plot-area.ts) 以完整图表外框为基准保存 `{ left,top,width,height }`，按外框移动/缩放重建 Canvas、原生 grid、Pie/Radar 和 3D 的共同绘图区。P15 声明包括 `IChartPlotAreaSpec`、`IChartPlotArea`、`resolveAutomaticCartesianGrid` 和 `resolveChartPlotAreaForEditing`；固定 UI 只定位到 `Yi/Wi` 的 section/状态注册，实际四坐标输入及拖拽实现未找到。手动模式开关加四个百分比 InputNumber 是本地独立界面，不声称复刻上游面板。

手动编辑种子包括当前标题/字体/图例和活动源系列的有效次轴；最小外框种子保持正宽高。所有坐标须有限、0–1、宽高正、矩形位于外框内，这是本地边界决定，未核实上游是否允许越界布局；原生 Office IO 应保留不满足本地界面边界的源信息及诊断。PlotArea patch 整体替换，关闭手动模式使用 `setOptions({ plotArea:undefined })` 恢复自动布局，换型保留不活动矩形，没有新增虚构的专属 setter。原生 grid 的 `containLabel:false` 对自动/手动路径都成立，因为共享布局已预留轴/标题空间。输入可暂时构成无效矩形，预览源系列仍保留，应用校验原子拒绝；清空单字段按有效源系列恢复其自动种子，修改不会提前进入内容历史。

新增 9 个自动渐变、10 个绘图区及 3 个 Facade 集成草稿，另有 2 个浏览器草稿覆盖隔离编辑/恢复、绘图区真实布局、位图版本、失败不发布内容/历史、纠正后应用、撤销/重做和内部保存重开。Facade 接线由根任务负责，当前测试、类型/构建、浏览器、Office 文件/独立重开、主题/焦点/IME 及大样本性能/内存全部未执行，保持“未验收”。

### P16 源字段语义描述与作者渐变

[chart-description.ts](../../packages/office-studio/src/features/charts/chart-description.ts) 显式选择作者语义、字段角色、标题/副标题、palette、legend 和 axes，返回独立的描述对象；不接受或扫描 workbook cells、records、source values、几何或绘制结果。series 使用源字段 Record，保留未活动数值覆盖，namedSeries 独立保留；按全系列 → 字段 → 已知源名称继承，style.name 只改显示。没有源名时的 `Series N` 回退不匹配作者名称覆盖；调用方可用 seriesNames 提供空字符串及真实分组身份。主题回退颜色不进入描述，显式 palette/颜色保留；自动 PlotArea 不展开为 layout seed，autoGradientFill 返回有效布尔值。

描述的类型名使用目录 canonical ID：历史 bar → column、横向 bar-horizontal → bar、doughnut → donut；getInfo 的兼容类型不变。轴绑定使用语义 category/secondaryCategory/primaryValue/secondaryValue 与 horizontal/vertical、边缘位置，Bar 交换两个维度；仅活动次轴出现 secondaryValue，Pareto 始终有该轴，Heatmap的源分类绑定vertical/left/category，数值字段名绑定horizontal/bottom/secondaryCategory，Radar不虚构矩形轴绑定。本地 `axes:{x,y,rightY}` 是公开轴配置的适配表达，不是固定版内部字段结构。

根任务已在 Builder 接 `describe(context?)` 和稳定 chart-draft ID，在实时 FChart 接本地扩展 `getDescription()`，公共描述/颜色导出也已落盘；[chart-description-host.ts](../../packages/office-studio/src/api/chart-description-host.ts) 只读取请求的普通源头部单元格。Pivot 缺源服务提供的 metadata 时保留作者字段，未扫描或重新计算数据。全部 numeric-only 角色推断、分组 Bubble/派生 Candlestick 的精确名称仍需要宿主元数据；不由描述重新执行统计。固定版底层 Builder/live adapter 的同名契约是 describe，不能把 getDescription 的本地名字说成固定版同名 API。

[chart-color.ts](../../packages/office-studio/src/features/charts/chart-color.ts) 增加 `ChartColor = string | ChartLinearGradient`、稳定 linear-gradient discriminator 和 `chartLinearGradient(input)`。作者 start/end 和 stops 深复制，保留作者 stop 顺序；坐标要求有限，可在 0–1 外，色标 offset/opacity 限 0–1，2–256 个密集色标。渐变对象只接受已知普通数据字段，拒绝稀疏、额外、symbol、隐藏或 accessor 元数据。CSS 字符串使用现有纯颜色 parser；边框、字体、marker 和 palette 继续是纯色。最大 256、offset/opacity 边界和严格普通对象规则都是本地保护边界，上游声明仅明确至少两个 stop 与 opacity 乘源 CSS alpha，完整 renderer 校验未公开。

配置深合并把颜色对象作为原子值替换，nullable 恢复也先复制整个渐变；没有从旧色标继承残余坐标。新写入按 [chart-gradient.ts](../../packages/office-studio/src/features/charts/chart-gradient.ts) 判断：系列支持 31 个目录项，数据点支持 25 个，2D Line/Area 的数据点和所有符号只接受纯色。换型及完整集合替换可保留相等的不活动渐变；在纯色契约的图型上新增渐变原子拒绝。绘制的 point resolver 会跳过不支持的点渐变覆盖，marker 使用纯色代表值，源作者配置和描述继续保留。其他专属纯色图型的历史系列渐变使用代表色回退，不能据此声称它们支持渐变填充。

实际纹理优先，显式作者渐变优先于自动预设，并在 autoGradientFill=false 时仍绘制。运行时色标稳定排序，CSS alpha × stop.opacity × fillOpacity 只乘一次；3D 的每个面再逐色标阴影，不把整条渐变压平。作者方向不会因负值、坐标轴 reverse 或横向条形改变，自动预设仍按 P15 的方向规则。runtime.authored、bounds 及 ECharts colorStops/global 不进入作者模型或描述。

Column/Bar、Pie/复合 Pie、Waterfall、Histogram/Pareto 消费 mark-local paint；Line/Area 的片段共享同一系列范围，Area 包含基线和已采样平滑 extrema。3D 柱六面共享八顶点投影边界，Pie 使用实际投影扇区边界，Line/Area 使用投影系列面范围；自动 3D 柱仍使用投影基线/顶端方向。普通 native Column/Bar/Pie 转换为可序列化 ECharts paint；明确 Line/Area/组合线型片段的作者渐变走公共 graphic，以保持同一片段坐标。固定 P16 converter 只声明 global:false 的 normalized 对象，本地公共像素 bounds/graphic 是独立适配，闭源几何与最终像素仍未核实。

共享渐变几何在自动 48 等分之外插入作者 stop 边界，重复 offset 和窄色带保留硬边，两个端点钳制，不重复覆盖同一半透明面积；退化向量只使用 0–1 单带。代表色取最低 offset 的第一个 stop 并应用其 opacity，这与固定共享 picker `$o` 直接取 stops[0].color 的可见代码不同；底层 getChartColorRepresentative 只有声明，未核实真实规则。极端有限作者坐标可能在转换到像素时溢出，当前转换/几何拒绝非有限 paint，极端边界仍未验收。P15 的每次 append 预算不等于完整场景或 Worker 内存保证。

[chart-color-editing.ts](../../packages/office-studio/src/features/charts/chart-color-editing.ts) 和 [ChartColorInput.vue](../../packages/office-studio/src/features/charts/views/ChartColorInput.vue) 提供继承/纯色/渐变、四方向坐标、逐 stop CSS alpha/offset/opacity 和真实增删。编辑保留作者顺序，自动新建渐变使用水平纯色到透明，这是本地 seed；追加复制最后 stop，至少保留两项。所有可替换控件为 antdv-next，布局 CSS 不写固定主题色；普通色仍用库 ColorPicker。已保留但当前目标不支持的渐变可恢复继承或改纯色，方向/stop 暂停编辑是当前类型契约，不是假按钮或能力占位。原生 RGB/alpha、主题、弹层、键盘、焦点恢复及 IME 均未运行验证。

本批写了 13 个 description、15 个显式渐变算法/绘制/UI 控件、3 个显式渐变 Facade 和 2 个真实颜色编辑器浏览器草稿；全部未执行。内部 JSON 重开、Canvas/native paint、完整 31 项 paint 差异等只是草稿中的待断言行为，不是产品验收或实测性能。

### P17 Pareto 独立样式、物化点与恢复

[chart-pareto.ts](../../packages/office-studio/src/features/charts/chart-pareto.ts) 实现独立 barStyle/cumulativeLineStyle 的纯 helper、nullable 复制与累计非正线宽忽略。柱形继承通用源字段/名称样式，再覆盖专属 color、opacity、border、label 和 dataPoints；清除专属组后恢复有效通用继承。累计线独立于源系列样式，缺省第四活动主题色、2px 实线、opacity1、size7、circle；固定 `R.point.shape` 的 circle、Pareto symbolSize7 和 lineThemeColorIndex3 已静态核实。point.color、边框和标签字体是纯色，线/柱填充及其可填充点支持作者渐变。

`chart-statistics.ts::chartPareto`、候选校验和 Rust `office-compute/src/chart/distribution.rs::pareto` 已改为接受 signed 源行，按第 4 节先代数合计再过滤，并按首出现稳定处理等值分类。有限 raw 总量直接累计/除法，累计百分比 floor 到两位；仅总量溢出时使用缩放，正总量最后一项为100，全零为0，同一分类自身合计溢出明确拒绝。`[6,4]→[60,100]`、`[4,4,4]→[33.33,66.66,100]` 是新增数值草稿，尚未执行。专属面板与通用系列面板的逐点选择器使用同一小数据物化函数，先复用 renderer 的 aggregation/TopN/空值和日期准备，再返回当前排序分类并按物化下标编辑；不再直接用未准备的源行列表。`chartComputeInputSteps/prepareChartComputeInputs/chartComputeInputs` 入参仅需既有 `ChartDateAxisConfig` 的 type/options/xAxis，未增加第二套聚合引擎或改变 Worker 结果 DTO。其正合计聚合降序与固定 Wu 一致，另扩展 includeZeroValues 下的零合计选择；TopN 前处理与上游最终可见效果仍未核实。源出现新分类、合计排序变化、过滤/TopN 和其他换型后如何迁移旧点覆盖仍未完成；当前按新物化位置消费，不把它说成稳定源行身份或全部上游索引契约。

本地逐点 color/opacity/border 可影响结束于该点的累计片段，marker 的颜色用纯色代表值，size/shape 消费其覆盖。固定 `Yje` 的逐点 shape/size/color 优先级与全局累计 line style 已静态读取；逐点 fillOpacity/border 影响累计片段属于本地扩展。曲线的全局 stroke 不充当 marker border；空心符号缺省以原始 marker color、1px 和有效 fillOpacity 绘轮廓，显式点 border 覆盖其色/宽/透明度。点 border.width=0 或 visible=false 只隐藏 marker 边框；曲线继续使用正线宽。边框回退用未烘入 fillOpacity 的纯色，避免两次 alpha；最终可见效果仍待统一行为验收。

Canvas 与 native 普通柱形都消费专属样式，native Pareto 对纯色填充单独烘入 fill alpha、item.opacity=1，使边框 alpha 独立；已有渐变色标仍只消费一次。请求了 border 作者字段而未写 color/width 时，两个适配器以原始填充代表色/1px 补绘制默认，这是本地决定，空 border 仍无轮廓。虚/点线矩形轮廓使用公共 line 几何，width0 不发送非法 Canvas stroke。显式累计渐变、实际活动逐点覆盖或复杂/空心 marker 走同一 graphic；未活动及 size0 marker 覆盖不强制切换。柱形百分比由已物化的非负业务值独立计算，累计线百分比消费截断累计值；相等 `[4,4,4]` 柱标签均为33.33%，不采用累计差分产生末项33.34%。`0.0%` 等 percent 数字格式先以比例值格式化，只输出一个百分号。固定 `u5` 虽有 Percentage 面板选项，formatter 实际忽略16；本地补真实百分比是明确功能差异，Office 样本仍未验证。共享 formatter 仍是既有数字格式子集，复杂 Office format 不因这一修正变成完整支持。

累计数值轴缺省 linear 0–100；author log 轴使用同一 positive extent helper，零累计点保留在统计中但不生成非法 log 图元。单个作者 min/max 在默认百分比范围外时为另一端生成有效有序边界，Canvas/native 使用相同结果。这个 log/自动边界行为是本地明确决定，公开演示和文件样本仍未观察。

[chart-pareto-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-pareto-settings-controls.ts)、[ChartParetoSettings.vue](../../packages/office-studio/src/features/charts/views/ChartParetoSettings.vue) 的两个专属组、整体/逐点恢复实际编辑独立草稿，经公共一次 apply 写内容/history。标签细节按有效 visible 展开，关闭保留作者内容/字体；bar UI 另反映通用标签继承，区别于 cc 的直接 authored visible 检查。没有预览系列时，bar 逐点颜色回退仍按 config.valueFields 的源字段取样式，和顶层柱形组一致。既有 includeZeroValues Switch 保留；全部 FormItem、Select、InputNumber、ColorPicker、Button、Collapse 和反馈用 antdv-next，无新增原生控件；共享普通颜色字段也改用 RGB value-format 保留 CSS alpha，显示格式仍为 Hex。逐点颜色/opacity/纹理/边框相较固定 cc 的逐点颜色与累计 shape/size 是本地扩展；bar.dataPoints 的公共 size/shape 字段仍保留，但当前柱形 geometry 不消费它们，面板没有写空操作。

共享 marker 已纳入全部19种公开形状与真实空心绘制，label 已支持数值组合/Empty0 和 auto/insideBase/insideEnd；旧四个字符串兼容，BubbleSize32 在 Bubble 编辑器可操作而 Pareto 专属组拒绝。完整绘制/控件/固定来源与差异见 P19。border.visible、nullable reset、center 扩展及线宽 ≤100 等本地边界仍单独记录，不能据此把 Pareto 标为全功能已验收。

共享 ChartParetoBarStyleOptions/ChartParetoLineStyleOptions/ChartParetoOptions、Builder/FChart 的 setBarStyle/clearBarStyle、setCumulativeLineStyle/clearCumulativeLineStyle、setIncludeZeroValues/resetIncludeZeroValues 六操作及 Rust signed 合计已由根任务落盘。Builder 保持同步 chaining，本地 live FChart 通过异步统一 command/history 提交，与固定 `FParetoChartBuilder` 的同步外观分开记录。结果 DTO/decoder 未改变；实际 Rust/WASM、API、历史和内部重开正确性仍待统一验证。

P17 当前17个纯算法/候选/继承/nullable/默认/透明度/边框默认/虚线/width0/百分比格式/log/native/graphic/控件草稿，包含 aggregation/TopN 选择器、源字段回退、有限 raw/极端总量与 Canvas/native 同值契约；另有3个六操作 Facade 历史/隔离/内部重开草稿和3个真实面板 edit/reset/cancel/apply/组合标签/空心标记/undo/redo/reopen 浏览器草稿。全部未执行；Pro 公开行为、Docs 宿主、原生 Office 往返、真实两后端与大样本 Worker/取消/传输/内存仍未完成或未验收。

### P18 Waterfall 全图累计、默认边界与真实源身份

固定来源是 1.6.0 发布归档的静态文本。`engine-chart/lib/es/index.js::ake` 将作者 pointRoles 构造成 `new Map(array.map(...))`，重复 `seriesIndex:dataPointIndex` 使用最后一个 role，保存时保留输入数组。`Nn/Rte/Xn` 仅在缺少 own pointRoles 且非 stacked 时，按来源顺序选择第一个有限 number 为 total；没有 own connector 时写 true。显式 `[]`、own undefined 和 connector false 均与缺省区别。`Mte/toChartModelConfig` 与 `Pte/toChartModelConfigReplacement` 执行 Nn，`qn/toChartModelUpdate` 直接执行 Jn。Sheets insertChart 调 Mte；live `commit` 调 qn，`update({dataSource,config,layout}) → commitChanges → _replaceConfig` 才调 Pte。创建/整份配置替换与普通 patch 的默认行为不能混用。

[chart-waterfall.ts](../../packages/office-studio/src/features/charts/chart-waterfall.ts) 的 `initializeChartWaterfallOptions` 克隆输入，并按本地源字段 identity 实现上述边界默认；根任务已在 [chart-engine.ts](../../packages/office-studio/src/features/charts/chart-engine.ts) 的真实 createChart 接入一次 readChartData → initialize → validate。当前没有整份 config replacement 公共 API，对应创建默认边界仍待补。普通 update/setType 继续保持 patch 语义，角色/connector clear 后不会在读取或绘制时被重新填回。运行期缺省为 arrangement、useSubtotal true、connector 隐藏；创建时写 connector true 不是 renderer 改写作者模型。三个角色默认用成对的 theme palette 正/负项和 `#CDD0D8` 合计色，neutral connector 为 `#8C8C8C`/1px/dashed/opacity1，不继承主题 axisLine 色。

来源 `[10,-3]` 与 `[4,-2]` 的 arrangement 共享累计为 `0→10→7→11→9`，各来源插入的自动合计均为整图最终值 9，显示位置 2/5；第二来源真实项目显示位置为 3/4。若第二来源首点明确为 total，则从零绘 4，并将后续累计改为 2。stacked 按分类共享同一中心：首分类两段为 `0..10`、`10..14`；第二分类的负段为 `12..9`、`14..12`，末尾只有一个全局合计 9。全局合计的 seriesPosition 为 `0xffffffff`，dataPointIndex 为真实 categoryCount，不能塞入第一个来源、采用其 subtotal 样式或以扩展分类下标寻址作者 dataPoints。

一个 `waterfallLayout` job 提供全图累计、来源位置、原始 value、start/end、角色/synthetic flags 与连接端点。[chart-waterfall-compute.ts](../../packages/office-studio/src/features/charts/chart-waterfall-compute.ts) 只解码和验证权威 tuples，检查完整来源集合、当前 stack/subtotal metadata、严格 total start=0、最终累计、真实/合成身份、有效样本数和连接关系，并按 preparation chunk 让出；它不生成第二份布局结果。[chart-waterfall-layout.ts](../../packages/office-studio/src/features/charts/chart-waterfall-layout.ts) 只展开显示标签、解析角色主/次轴和 Canvas 坐标。drawWaterfall 使用 flat scene，柱宽占完整 category band，按原始 dataPointIndex 取 role dataPoints；stacked global subtotal 只消费共享角色/单点样式。空 scene 不生成来源或自动合计空槽。旧分组 waterfall 结果没有作为生产回退，非空未计算 scene 明确报告 `COMPUTE_NOT_READY`。

连接线先绘，随后绘柱体/标签。arrangement 连接端点按两侧实际角色轴分别投影，允许显式 total 的斜入线；stacked 的 category cumulative 连接使用主轴，与固定 ES 行为一致。本地作了两处明确差异：固定 arrangement connector 没有按插入 subtotal 修正显示列，本地连接真实 `0→1→3→4`；固定连接取分类中心，本地连接柱边缘，在分类 reverse 时仍沿正确方向取边。以上是源码可见的选择，没有公开浏览器或 Office 图表样本来证明最终效果等价。log 轴下保留真实/合成分类与统计，省略 start/end 非正的 mark 及相邻非法连接。

固定入口按从上到下、同排从左到右追踪如下。计数按 section 声明的逻辑字段，选项枚举、颜色 reset、小字体按钮和折叠标题不重复算成功能；没有浏览器可见总数。

| 分类与位置 | 固定入口/条件及可复查计数 | 命令/模型与本地实现差异 |
| --- | --- | --- |
| 插件 setup 扩展 `ji` | stackType Select → useSubtotal Toggle，共 2 字段；Waterfall 类型才显示。 | `U('waterfall')` 写 view state → live commit → `sheet.command.chart-update-config`。本地 layout helper 同序；stacked 切换清旧源 pointRoles，新写非空 roles 原子拒绝。 |
| 插件 style 扩展 `Al` | 来源 Select → positive → negative → subtotal → connector → dataLabels。每角色固定 6 字段为 color/opacity、border color/width、border opacity/type；单来源或 subtotal 才默认显示 name。all 目标基础 22 字段，单来源 arrangement 24，单来源 stacked 隐藏 subtotal 后 17。 | 名称显示可由宿主 callback 覆盖；固定 Sheets callback 提供来源正/负名与合计名，没有改 showBucketName 条件。全局角色写 waterfall.styles，来源写 series.style.waterfallStyles；本地统一选择 role 后编辑，global subtotal 自动选 all 并禁用来源目标，保留不活动的来源作者覆盖。 |
| 条件 dataLabels 明细 | 有效 authored visible 时依次 content、position、format、font，共 4 个逻辑字段；font 为复合控件。基础计数分别加 4，不能将内容选项或字体子按钮另作能力完成证明。 | 固定 content 可组合 CategoryName/SeriesName/Value；position 含 auto/inside/left/right/top/bottom。本地共享编辑器已支持组合 mask 和完整公开 position，并保留 center 扩展；Waterfall 的其他类型专属百分比语义尚未按固定 ES 完成核对。 |
| 元素上下文/浮动扩展 | `kind:'waterfallStyle'` 有角色颜色、不透明度、标签 Toggle 3 个已识别 compact 控件，target 带真实 baseSeries/role。其余右键/溢出分支没有形成完整静态计数。 | 上游直接编辑目标与面板互相路由。本地共用角色模型和绘制，完整元素命中/浮动菜单同命令闭环仍未完成，不将 whole-chart hitbox 算成这些入口。 |
| 本地角色/源项目扩展 | layout 2、connector helper 5（4 个随开关显示）；role helper 21（纹理两色有条件）、point helper 10（纹理两色有条件）；另外 scope/role 2 Select、角色项目 1 Select/一基 InputNumber、源字段/项目 2 Select/一基 InputNumber、源含义 1 Select，以及对应条件 reset。 | 明确区分复用上游契约与本地扩展。16,384 行以上不构造整份 Select 列表，完整原始与合成 identity 仍可直接编辑；这只是有界实现，未测响应时间或内存。 |

Builder/FChart 的 `setPositiveStyle/clearPositiveStyle`、`setNegativeStyle/clearNegativeStyle`、`setSubtotalStyle/clearSubtotalStyle`、`setConnector/clearConnector`、`setStackType/resetStackType`、`setUseSubtotal/resetUseSubtotal` 共 12 个操作已由根任务接线；role patch 复用字段/名称 selector、深合并、nullable 恢复、候选验证及统一 chart command/history。根任务本轮也为 add/update/remove-chart 加宿主 worksheet 编辑权限判断；源读取权限、撤销/重做发布保护和其他批量变更仍须按其权限草稿验收。内部 JSON 只存作者源/配置，不存累计、flat scene、sentinel、connector 结果或 WASM lease。

[ChartWaterfallSettings.vue](../../packages/office-studio/src/features/charts/views/ChartWaterfallSettings.vue) 及共享 fields 使用 antdv-next 的 FormItem、Select、Switch、Input/InputNumber、ColorPicker、Button、Collapse 与 Alert，全部写隔离草稿后一次 apply；没有新增原生表单控件。ConfigProvider/token 沿用应用统一主题，真实明暗主题、弹层、禁用/焦点恢复及 IME 仍未验收。点/源含义 reset 和全局合计草稿不是提示消息代替的入口，但未执行测试不能宣称它们已完成行为验收。

本批新增 15 个完整 scene/布局/绘制/graphic 契约草稿、12 个默认/角色/继承/草稿/大源索引控件草稿、4 个 Builder/真实 Facade/service/decoder/history/内部重开/权限集成草稿和 3 个 workbench 浏览器草稿。共享 [chart-waterfall-fixtures.mjs](../../packages/office-studio/tests/fixtures/chart-waterfall-fixtures.mjs) 只展开固定 tuples/身份，不推导累计或堆积；已有样式、数据布局、名称 selector、自动/显式渐变草稿均改用匹配的固定 computed scene。14 万点的零变化 prefix + 固定末点只检验大数组布局契约，不是基准。全部未执行，真实 Rust/WASM/browser binding、浏览器交互/两后端、源变化身份迁移、Office 往返与大样本取消/传输/内存继续保留未验收。

### P19 全部公开标记、组合标签与真实内部位置

固定 `LinePointShape` 的19项依次为 dot/dash/circle/square/triangle/diamond/rect/roundRect/pin/arrow/none，以及 emptyCircle/emptySquare/emptyTriangle/emptyDiamond/emptyRect/emptyRoundRect/emptyPin/emptyArrow。`R6/I6/L6/POe/FOe` 将 dot 转为 `path://M0 0L0 0M10 10L10 10M3 4H7V6H3Z`、dash 转为 `path://M0 0L0 0M10 10L10 10M1 4H9V6H1Z`；零长度 anchors 保留 nominal 0..10 箱。故 scalar size 下 dot 实际宽0.4×size/高0.2×size，dash 宽0.8×size/高0.2×size；dot 不能作为 circle 别名。square/rect 在 scalar size 下共享方形边界，其余复杂形状由本地独立有界多边形生成，不声称复刻上游每一像素。

[chart-markers.ts](../../packages/office-studio/src/features/charts/chart-markers.ts) 提供唯一选择项、验证、native mapper 与纯几何；[chart-geometry.ts](../../packages/office-studio/src/features/charts/chart-geometry.ts) 由同一形状产生填充/轮廓命令。8种空心形状只画轮廓，不填中心；显式 border.visible=false/width0 保留，缺省轮廓使用 marker 原色/有效填充 opacity，曲线宽度不影响 marker 边框。坐标/size 必须有限，size 非负，生成坐标溢出拒绝，作者输入不修改。简单实心标记保留 native ECharts，实际活动 dot/dash/复杂/空心标记转公共 graphic；未活动、size0、scatter 非法数值/log、Bubble 无正权重或 Boxplot 无异常点的覆盖不会单独强制该路径。两适配器同源命令是代码契约，真实浏览器/两后端仍未验收。

固定 `LabelContentType` 为 Empty0、CategoryName2、SeriesName4、Value8、Percentage16、BubbleSize32。`u5` 的 Pareto 标签按 Series→Category→Value 逗号拼接，忽略 UI 可选的 Percentage16；普通 `F6` 也按图型去掉部分 Percentage 标志。这次本地 [chart-labels.ts](../../packages/office-studio/src/features/charts/chart-labels.ts) 按 Series→Category→Value→Percentage→BubbleSize 固定顺序消费数值 mask，undefined 保持 Value8，0明确空标签，拒绝未知bit/无效或重复选择；兼容旧 value/category/series/percentage 字符串。单独的 bubbleSize 字符串不是领域 string 契约，32由 mask 表达。Pareto 柱形使用真实业务占比、累计线使用截断累计；Bubble 传实际源 size。其他图型完整百分比语义仍须逐型核对，不能由公共 formatter 推导已完成全部公开行为。

固定公开 `SeriesLabelPosition` 有9项 auto/inside/insideBase/insideEnd/outside/top/bottom/left/right；`cc` 面板仅列 auto/inside/left/right/top/bottom 六项，累计线排除 outside。本地保留 center 扩展，Pareto bar 提供10项、line9项。`f5` 的普通 bar insideBase/End 按 value 正负、横向/纵向及 reverse 映射实际 insideTop/Bottom/Left/Right，line fallback为inside；本地共享 `chartPointLabelAnchor` 用真实投影 start/end 与短柱中点定位，auto 采用实际末端外侧。相较固定 line fallback，本地无基线点的 insideBase/End 落在实际端点，属于明确差异。新内部位置/auto 在 native 需要时使用公共 graphic，以保持命令坐标一致。

[chart-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-settings-controls.ts)、[chart-pareto-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-pareto-settings-controls.ts) 和 [ChartOptionFields.vue](../../packages/office-studio/src/features/charts/views/ChartOptionFields.vue) 新增真实 `antdv-next Select mode="multiple"`，先校验全选项再向隔离草稿写 mask，经已有一次 apply/命令/history 提交。关闭标签保留内容/字体；普通非 Bubble 编辑器保留未活动的导入 BubbleSize32，Bubble 编辑器提供5个内容选择，Pareto 专属组仅4个并拒绝32。19种 marker 选择通过公共定义复用。控件/API 枚举是选项数，不作为入口或功能完成计数；FormItem、Select、InputNumber、ColorPicker、Switch、Button 与 Collapse 仍使用统一 ConfigProvider/token，真实主题/弹层/禁用/焦点/IME 尚未运行验证。

新增 [chart-markers-labels.contract.test.mjs](../../packages/office-studio/tests/chart-markers-labels.contract.test.mjs) 的16个草稿，覆盖19项验证/有界面积/dot-dash精确箱/空心无填充/alpha/显式隐藏、独立 Pareto outline/native aliases/实际活动路由、mask/逗号顺序/原子拒绝/隔离/隐藏字体/BubbleSize保留、完整位置/正负反向短柱/真实 column 与 bar-horizontal 标签、相等业务占比和截断累计。同批扩展3个 Facade 草稿与3个 Pareto 浏览器草稿，实际编辑多选mask22、insideBase、emptyRoundRect与单点emptyPin、一次apply/history/internal reopen。所有草稿未执行；Office 原生写出与独立重开、公开演示/原生效果、Docs宿主和大样本计算/取消/复制/内存仍未验收。

### P20 Sankey 真实流带、方向轨迹与视图节点拖动

固定归档 `archives/engine-chart-1.6.0.tgz` 的 `package/lib/es/index.js` 中，`Kae`（offset112091）按 source/target/value 三字段读取、Set 保留首次端点顺序；端点按 truthy 过滤，weight 要求 number 且非 NaN，不合并重复边，只以反向边对给 isCircular 标记。`UOe`（offset1512808）是配置对象，设置 native nodes/links 和 emphasis.focus='trajectory'；`Gje`（offset1623874）是 converter arrow function，使用首次节点主题色、source 流带色、curveness.5、font family 继承/默认10px、黑色0.7与justify。`_getShowEmpty` 在 invalid source 或 isCircular 时为空，归档 bundled native Kahn 布局也拒绝全部有向环；`sye` 使用 max(sum in,sum out,raw value)。上述是固定发布源码证据，未运行商业 SDK 或公开演示。ECharts 6.1.0 和 Context7 的当前 Sankey/graphic replace API 只作本项目适配依据，不替代固定 Univer 版本证据。

| 入口分类与位置 | 固定入口及计数边界 | 本地命令/状态与差异 |
| --- | --- | --- |
| 插件目录子项 | 既有45项目录中的第37项 Sankey，SankeyIcon；不新增 Ribbon 常驻、插件根、溢出或 Docs/Pivot 根项。 | 沿 P05/P06 的插入/编辑 live commit 写统一 chart config；本地仍使用 add/update-chart、内容历史和现有属性草稿。 |
| setup 映射 | Kae 实际消费 source/target/value 3个逻辑角色；value 缺省的公开 UI 显隐/完整可见数仍未核实。 | 本地源/目标字段不可相同，缺 weight 计次数、trim/数字端点作为名称、null/zero 跳过、负权重拒绝、重复有向边累计；这些是明确的数据规则差异。 |
| chart-ui 专属 section | 固定类型中未找到 Sankey 专属 section，不把通用属性或图内拖动算成新增面板入口。 | 复用既有颜色/opacity/纹理/边框/标签作者层；通用与单点 marker size/shape/color 对此矩形节点不活动，控件隐藏并保留作者字段。 |
| Canvas 上下文交互 | native draggable=true、focus trajectory 和 .2→.5 的直接边强调已静态核实；专属浮动/右键入口未形成完整计数。 | 本地由真实节点/流带命中进入 runtime，不新增持久模型；节点优先拖动，其他区域继续原图表移动/尺寸命令。完整权限/键盘与元素浮动菜单仍依赖宿主验收。 |

[chart-data.ts](../../packages/office-studio/src/features/charts/chart-data.ts) 将 incoming/outgoing 分开有限累计，Sankey 节点取两者最大值；中转 `A→B→C` 不把 B 流量算双份，自环也只承载一次。Relation/Chord 保留其 incident 合计行为。有向重复边按首次端点身份累计，循环检测覆盖自环、双向和长环；每节点同向流量溢出明确拒绝。布局按最大边权重归一化，跨列或不同节点的有限流量不因全图合计而溢出。

[chart-sankey-layout.ts](../../packages/office-studio/src/features/charts/chart-sankey-layout.ts) 通过 iterative Kosaraju 求 SCC、condensation 最长层和终端 justify；32次带权松弛只处理 DAG，排序不改变原节点/边索引。默认20px节点和8px间距按场景空间钳制；forward band 使用两条 analytic cubic 与真实垂直端口，反馈边预留顶部 lane 和四个 quarter-circle 边界。循环回流是本地独立能力，区别于固定空态。`ChartLayout.sankey` 缓存一次几何供 draw、graphic、命中与交互复用；拖动只 clone view layout、钳制 nodeArea、重排端口/重建 bands，不重读源、重求 SCC 或松弛。水平拖动支持反馈 cap 和非正跨度 S controls，但任意节点重叠、反向拖动的碰撞、自交、标签遮挡以及 cycle 可读性仍未完成行为审查/验证，不能由有界坐标断言推导已解决。

[chart-sankey-draw.ts](../../packages/office-studio/src/features/charts/chart-sankey-draw.ts) 一次解析 shared→field→named→dataPoint，缓存节点 author paint；流带采用源节点色、0.2×有效作者 opacity，节点 opacity 独立。默认标签 visible/category/auto，终端自动朝左；组合mask、真实节点值、格式、字体和全部公开位置/center 扩展实际生成坐标、maxWidth/clip，编辑器读相同继承默认。扩展 Percentage 定义为节点值占最大 stage 总 throughput 的份额，用归一化求和；固定 Gje 只出 name，不声称这个分母是上游契约。纹理/边框与更完整作者覆盖同样是本地扩展，原生与跨后端效果仍未验收。

[chart-sankey-interaction.ts](../../packages/office-studio/src/features/charts/chart-sankey-interaction.ts) 缓存节点矩形和真实 band/AABB，节点优先，边按 point-in-polygon 命中。轨迹按 upstream/downstream 有向遍历且 cycle-safe，hover A→B 排除 A 的 sibling fork；固定 focus='trajectory' 已核实，完整浏览器可观察效果仍待对照。非轨迹 alpha 乘0.15，直接 hover 边从基础.2乘2.5到.5，离开返回精确基础 commands；不修改缓存作者 paint。`echarts-option.ts` 将内容置于唯一 `${chartId}:sankey-content` graphic group；renderer hover/drag 只 replace 该group，chrome保留。runtime scene/hitIndex/topology 不进入 option 签名、chart config、描述或内容历史。

renderer `getContentHit/beginContentDrag/updateContentDrag/endContentDrag` 使用图表局部 CSS pixels，sourceVersion 标记每次 surface 更新；相同作者 option 重绘保留已经移动的节点位置，来源/配置/尺寸/DPR/主题变化重置 runtime。cancel 恢复精确 pre-drag scene，commit 仅保留在该 Canvas owner；节点位置不跨关闭/重开，不作为持久作者字段。主任务在 ChartOverlay 接节点优先、pointerId/capture、末点逐帧消费、Escape/pointercancel/blur及源/主题/尺寸变化取消，并将尺寸句柄改为同级 antdv-next Button；图表移动/尺寸仍走统一 update-chart 内容历史。宿主接口正在统一静态核对，未运行 DOM/焦点/IME/明暗主题或真实鼠标验证。

本轮添加18个纯数据/布局/paint/命中/轨迹/拖动草稿、3个 Builder/实时 Facade/原子拒绝/历史/内部JSON重开草稿、4个 renderer 生命周期/graphic patch 草稿和3个 workbench 真实hover/拖动取消/源变更/antd标签草稿，全部未执行。100000节点/100000边/2000000几何点仅为显式分配预算；forward51/feedback114点包含中心线与band。SCC/relax和拖动 scene 当前在主线程独立小数据路径，没有另起 Worker 引擎；现有 Worker/Rust 准备协议仍唯一。大图路径、端到端性能、主线程长任务、取消/传输/复制/内存、WASM binding、公开cycle/轨迹/拖动、原生Office往返、Docs宿主继续保留未验收。

### P21 Heatmap 双分类轴、格式文字与共同色标

固定 `Oae/wOe/uAe` 将选中数值字段名作为X类别、源分类作为Y类别；数据tuple是 `[fieldPosition,categoryIndex,value]`，不是每个类别一列。axis operator按 `reverse === true` 设置inverse，因此默认首Y分类位于底部。本地 [chart-heatmap.ts](../../packages/office-studio/src/features/charts/chart-heatmap.ts) 提供共同单元格矩形、双轴反向/命中和实际finite extent，空源为0/0；常量范围使用palette中点。显式双边界要求min≤max；仅作者单边界与自动值倒置时按effective extent升序，作者字段原样保留。layout、description、原生tuple和共享Canvas已按同一方向接线，实际视觉仍未验收。

固定 `Mje` 的piecewise为splitNumber8。两条本地路径使用同一8段 `[min,b1]`、后续 `(b1,b2]`，边界属于前段；continuous范围外数据钳制颜色并可见，piecewise范围外标记不绘制且保留原数据。区间插值保留binary64，不使用上游decimal step rounding，以免极端/subnormal范围坍缩；相邻binary64值无法形成8个非空区间时仅保留实际非空段，常量为1段。区间标签使用实际边界，本地不声明已与上游显示舍入完全等价。极端范围差值非finite时共享graphic绘制，避免native visualMeta溢出。

普通Sheets source通过唯一 `formatSheetValue(value,numberFormat,date1904)` 生成 `formattedRecords`，不支持的格式保留cell.display；原始numeric不变。projection按同记录/字段生成 `series.formattedValues`，normalization改numeric时清该点旧格式文字。Heatmap标签只消费value，显式label.format优先，否则sourceformatted优先，再走共享数字formatter；固定默认hidden/inside/12px和继承family已接。保留的泛用系列/点label字体、格式和visible可继承，contentType不在Heatmap中消费、不绘第二组标签。Pivot source尚无格式元数据；所有原生Office源格式和unsupported-format实际样本继续未验收。

固定 `Nje` 不允许拖动范围，continuous两边界追加单位；本地trim单位并以空格连接，piecewise不追加单位。两渲染路径通过 [chart-heatmap-draw.ts](../../packages/office-studio/src/features/charts/chart-heatmap-draw.ts) 绘同一可见Canvas色标：continuous横向低→高/纵向高→低，128个有界色块；piecewise横向升序/纵向降序。visible/position实际控制scene和自动plot预留空间。native Heatmap保留hidden、calculable/realtime/selectedMode/hoverLink均false的内部visualMap，但每datum设置visualMap:false，保护共享颜色、8段定义和作者覆盖；内部show:false不表示可见色标缺失。

[chart-heatmap-interaction.ts](../../packages/office-studio/src/features/charts/chart-heatmap-interaction.ts) 一次建立256个continuous bins或piece bins，packed源身份为seriesPosition×categoryCount+categoryIndex；hover查询缓存矩形和相关bins，不排序/全表扫描。continuous用12 CSS px hover band，跨端点时包含范围外源值；cell→legend indicator、legend→matching cell边框、piece swatch强调走同一个graphic group replace/remove。Heatmap原生和graphic内容均使用既有 `EChartsCanvasRenderer.updatePointer`；未新增pointer union或Worker权威引擎。hover不修改min/max、源、作者模型或内容历史，缓存/sourceVersion、leave/cancel和源/主题/尺寸/DPR释放有草稿。

专属 [chart-heatmap-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-heatmap-settings-controls.ts) 按固定顺序提供范围、单位、label开关和条件位置/字体，无content selector；type在共同legend/presentation，format为本地公开API扩展。控件复用统一 `antdv-next` FormItem/Input/InputNumber/Select/Switch/ColorPicker/Button和ConfigProvider，不增加原生表单。标记控件在Heatmap中隐藏，两个轴只呈现category能力，未活动次轴不影响边距。8个固定Builder/FChart方法为 `setVisualMapType/resetVisualMapType`、`setValueRange/clearValueRange`、`setCellLabel/clearCellLabel`、`setValueUnit/clearValueUnit`，Enum.ChartVisualMapType、strict输入/nullable和同一update-chart历史已由根任务接入。

本轮18个纯extent/layout/paint/hit/format/settings草稿、3个根任务Facade/原子拒绝/history/internal reopen草稿、4个renderer草稿和4个真实workbench草稿均未执行。workbench坐标用公开layout减frame起点；native pointer几何已经是local CSS pixels。设置的关闭丢草稿、每字段/整组reset、一次apply/undo/redo和内部JSON重开仅为待跑场景，不能代替Office原生文件。bins/paint/layout当前仍在主线程，未测大源、帧时间/长任务、传输复制/内存、Worker取消或WASM绑定；明暗主题、弹层/焦点/IME、两后端、Docs宿主、公开交互、Pivot格式与原生Office均保留未验收。

### P22 日期基本单位、主刻度、源行与恢复合并

固定 `engine-chart` UTF-8 ES原文的 `nae/Oa` 使用JS Date/timestamp解析并保留每个源行，数值字符串和空串不作为日期；普通Cartesian对象 `lOe`（字符offset1484869，不是函数）稳定排序日期并保留sourceDataIndices。`hm/Rke` 把月/年转成连续ordinal，作者边界落到桶，首尾留半桶，标签使用桶首日；`v8` 以majorTimeUnit→baseTimeUnit→days选择推进单位，每一步从原anchor推进，保留Jan31→Feb29→Mar31的月末锚点。`qke` 还包含字体碰撞与自动稀疏。本轮只读取归档，没有运行商业SDK、演示或Office样本；全部字符offset和类型位置见来源表P22。

本地 [chart-date-calendar.ts](../../packages/office-studio/src/features/charts/chart-date-calendar.ts) 复用Sheets的Excel Gregorian/1900/1904算术。`baseTimeUnit` 控制源点的日/月/年坐标，`majorTimeUnit/majorUnit`独立生成刻度，month/year标签使用首日并保留原serial。日允许正小数，活动月/年majorUnit须为正安全整数；不活动数值轴可保留month/year元数据并采用0.25等数值间隔。空源也校验作者单位。作者间隔无法以binary64推进、超过10,000推进步预算时显式报错，不静默替换；上游v8耗尽预算返回undefined，这是错误契约差异。自动计划目标约5刻度，亚秒范围无自动刻度时保留可见源anchor；未实现qke的完整字体碰撞/稀疏，不能声称刻度密度等价。1900 serial0的月/年推进将day0归一到合法day1；该特殊值和1900虚构闰日保留为本地Excel契约，与JS Date归一行为不同。

普通Cartesian/Bubble保持稳定日期顺序，每个点继续携带原categoryIndex/xValue，重复日期不自动合并；连接判断使用sequenceIndex，避免日期重排后错误断线。`chart-layout.ts`在原serial域收集趋势/X误差外延，再扩展calendar extent并重生ticks。Histogram最终bin center、Boxplot最终系列名、Candlestick源日期走相同日期布局。`chart-date-axis-draw.ts`为Canvas/3D与native内容追加同一tick/首日标签，默认日期grid隐藏，显式grid/tick/label、reverse、Bar、字体和frame clip可独立配置。native日期轴保留value extent和实际点坐标，隐藏其内建日期label/tick/grid/pointer，用共享graphic绘日期外观；公开computed layout是工作表绝对坐标，事件坐标减frame起点。Pareto专用OOe不消费calendarDateAxis；本地日期轴仍保留权重累计顺序和日期血缘，是扩展，不能记为固定版等价。

同一calendar桶的pointer从左右靠近均选择该桶首个原源索引，显示原日期而非首日tick标签；缓存查询通过两次lower-bound定位，不扫描原始源值。这个tie-break是本地确定规则，公开交互尚未核实。主线includeDate现对超出当前工作簿日期系统的趋势/X误差派生外延明确抛RangeError，即使显式作者边界已设置也不跳过非法值；有效外延继续扩展未作者边界并生成最终ticks。来源编辑/import/Docs发布尚未全覆盖，不能把基础日期预检记为全部模型入口已原子拒绝。

时间字段按min/max后基本单位→主刻度单位→主单位间隔排列，3个值入口/3个独立reset；已有日期开关增加1个reset，数值majorUnit增加1值/1reset。固定IAxisOptions持有三时间字段，但公开IChartAxisSpec和本轮定位UI不暴露它们，本地API/UI属于扩展。关闭日期或恢复开关保留时间元数据，恢复主单位后继承基本单位、恢复基本单位后继承days，不把有效默认写入作者模型。全部控件使用既有antdv-next FormItem/Select/InputNumber/Button和统一主题。

静态复查修正了设置恢复的实际合并问题：原先delete草稿属性无法清除updateChart的深合并旧值；现在日期开关/三个时间字段/数值majorUnit、标题整组/对齐/字体、Heatmap type/范围/单位/label、Radar shape/fill/axisLabel均写显式undefined，canReset按值是否已定义判断。chartSettingsPatch保留恢复意图，3个新增真实merge契约草稿覆盖叶字段和整组恢复，作者无关字段保持。旧标题/Radar/Heatmap工作台草稿按真实Drawer的“关闭”动作和页面级portal定位修订；内部重开比较使用JSON规范投影，未运行。

2026-11-25与主线协调并静态查看的发布接线：`ChartStatisticsService.requestCandidate`创建隔离槽并复用有界队列，统计解码后直接复用结果调用layoutChartData，检查最终Histogram bins、Pareto日期及trend/X-error extent，候选不替换已发布缓存。`estimateChartFinalLayoutMemory`按真实bins/boxes+outliers/Pareto/waterfall/trend/error数量收费，日期原/最终双tick集合各按10,000上限计入；服务在最终日期布局前合计source/decoded/layout估算再做预算判断。该估算是分配准入，尚无设备/内存实测。`OfficeStudio`的add/update-chart先准备同一候选，await后原子写历史；insertChart/insertPivotChart等待，CommandExecuted在异步完成后通知。两命令走executeCommand，executeCommandSync拒绝；chart-final-layout-memory/chart-candidate-transaction专项草稿未运行。来源单元格编辑、通用mutate/import及Docs宿主尚未全覆盖；同步纯createChart/updateChart仍是源预检，上述代码不代表全部统计发布已验收。

本轮新增12个calendar、14个axis/data/layout、4个实际renderer、4个工作台及3个设置恢复merge草稿，共37个；新spec尚未接入统一验证入口。覆盖原锚点月末/跨年/虚构日/1904、重复源行/缺失连接、两路径坐标/标签/字体/clip、统计最终坐标/外延预算、同桶pointer、独立reset/关闭/一次apply/history/internal reopen和非法作者间隔。全部未执行；纯option/DrawCommand草稿不代表浏览器像素，内部JSON不代表Office原生往返。公开演示、来源编辑原子拒绝、明暗主题/弹层/焦点/IME、两宿主后端、真实WASM/Worker与大样本性能/内存均保留未验收。

### P23 Boxplot 固定默认、异常点样本身份与独立恢复

固定归档的Yce取上下半组中位数，奇数样本不把中间项纳入任一半组；min/max为完整有限样本范围。Xce每个数值字段创建一份分布，不产生独立outliers；`$je.boxplot`也没有消费dataPoints。按该静态证据，[chart-boxplot-options.ts](../../packages/office-studio/src/features/charts/chart-boxplot-options.ts) 成为有效统计选项的唯一resolver，缺省为`medianOfHalves/range`。显式旧inclusive/exclusive或旧whiskerMultiplier继续推导IQR（缺省倍数1.5），显式whiskerMethod优先；其数值与绘制在Rust/TS、planner/decoder和两渲染适配中共用该有效契约。这里只确认代码接线，尚未运行Rust或两后端验证。

`BoxplotData.outlierSourceIndices`为必需字段，与outliers逐项对应，表示传入prepared source.values的样本位置，包含前面的null空位；它不是五数行、排序后outlier数组位置或工作表绝对行号。对于`[100,null,-100,0,1,2,3,4,5,6,7,8,9,100,-100]`，inclusive/IQR草稿预期五数`[0,1.25,4.5,7.75,9]`、outliers`[-100,-100,100,100]`、indices`[2,14,0,13]`；固定默认预期五数`[-100,1,4.5,8,100]`且无outlier。以上均为待执行的预期样本，不是运行结果。

公共Canvas及native异常点按同一sourceIndex合并字段/名称/单样本覆盖。native datum ID为`${chart.id}:box:${seriesPosition}:outlier:${sourceIndex}`，重复数值仍独立。箱体、median、须线、汇总图例不套sample0的颜色/大小；inactive/inlier/hidden覆盖保留作者配置，只有当前实际可绘异常点需要公共形状或纹理时才切graphic。单样本shape/size/fillOpacity/border的绘制及透明度路径已写代码，真实glyph/像素、隐藏后交互和Office样本仍未验收。

专属四个字段提供显式undefined恢复，真实`chartSettingsPatch → option merge`保留清除意图；清除whiskerMethod不会同时删掉仍有效的旧quartile/multiplier，所以可能仍显示IQR。原始样本Select最多枚举前16,384个位置中的有限值，原始样本InputNumber直接定位其余位置并将已选远端项加入Select；null/infinite序号给出Form错误。P23的17,000样本草稿只有34,002源位置，覆盖的是选择器上限，**不覆盖65,536同步预览预算**；超预算闭环由P24的40,000样本草稿补充。

P23已写16个数值/身份/默认/nullable/独立绘制契约草稿、4个实际Canvas renderer草稿和4个真实工作台编辑/源激活/17,000远端样本/reset/history/internal reopen草稿，共24个；全部未执行。工作台草稿已静态修正为应用后Drawer保持打开、等待关闭按钮可用再关闭，不改变生产面板行为。没有主题、焦点、IME、真实WASM、Office或大样本性能验收。

### P24 大源设置读取、Pareto同次统计与有界项目选项

[FWorksheet](../../packages/office-studio/src/api/facade.ts) 新增本地Promise读取接口`getChartSettingsData(config, options)`和`getChartSettingsStatistics(config, options)`。它们先创建detached候选并捕获unit/宿主sheet/独立source sheet、modelRevision及user/ACL；options支持signal、timeoutMs、memoryBudgetBytes、priority。前者只返回统一projection、归一化与输入准备后的`ChartData`，不统计、不创建最终scene。后者返回同一次准备所得`ChartComputedStatistics`，含preparedData与Pareto等Rust结果，不执行最终日期scene/tick admission。两者共用有界队列和准备调度，不插入图表、不写历史、不发布既有chart cache；最终来源/派生日期和场景预算仍由真正创建/Apply候选校验。这是本地公开扩展，未找到固定版等价读取API。

[chart-settings-data.ts](../../packages/office-studio/src/features/charts/chart-settings-data.ts) 用`kind:'prepared-source'`、bindingKey和sourceRevision显式标记来源。key包括宿主与独立来源sheet、解析后的sourceText范围、方向、图型、Pivot、字段/mapping、可见性、invalid策略、聚合/Top N、辅助字段、活动组合类型及日期核心字段；Pareto额外包括有效includeZeroValues。标题/颜色/单点paint和Boxplot统计summary不进入该读取key。只有当前绑定和版本匹配的prepared source可绕过65,536同步预览预算；derived五数/Histogram bins/无标记series-only不能被误当原始样本，匹配数据直接复用arrays，不再projection或整份clone。无loader的独立小源仍有有界读取；Pareto独立消费者必须提供当前物化统计，不能在Vue computed里再算一套统计。

[use-chart-settings-source.ts](../../packages/office-studio/src/features/charts/views/use-chart-settings-source.ts) 是Editor、Insert和Pivot的共同临时读取状态。同步watch先abort并丢弃旧快照，microtask仅启动一次交互的最终绑定；请求捕获binding/version、loader/session身份，旧成功和失败都不能覆盖新状态。关闭、scope释放、源范围/字段/版本、工作簿/会话或读取器变化都失效；当前错误可见并可重试，Apply先assertReady。detached草稿的纯paint及Boxplotsummary编辑复用同一source数组；发布/内容revision变化仍重新读取，不能据此宣称已经缓存所有source变化。

工作台的`chartEditorPreview`只读detached getInfo，不再同步调用FChart.getData；Pivot原同步源读取和重复category/series构造已换成同一Facade loader。源读取收集sourceNumericFields，UI优先sourceValueFields/sourceNumericFields，不再对大records逐字段some来猜可选数值字段。prepared大源的日期预检只校验常量规模的作者配置；最终日期与派生tick合法性留给Apply。部分独立组件的小源/层级标签路径仍有有界或尚未实测的枚举，Pivot上游输出读取/复制也没有真实大样本数据，不能宣称设置管线全部重操作已满足性能标准。

Pareto只调用统计loader，不并行/串行再读取第二份完整source。`ChartSettingsFields.pointCategories`、`ChartParetoSettings`及[chartParetoSettingsCategories](../../packages/office-studio/src/features/charts/chart-pareto.ts) 都直接消费同一`statistics.pareto.categories`；helper已改成纯accessor，不再chartComputeInputs或chartPareto。这样已准备的Top N/others及invalidValueStrategy不会被面板再次应用，includeZero和聚合变化则触发新的Rust结果并立即清空旧分类。柱/累计线/通用项目索引仍指当前物化Pareto分类；源变化后的作者点索引迁移、公示交互和Office往返仍是独立缺口。

通用类别、Pareto柱/线及既有Waterfall角色列表最多生成前16,384项加已选远端项，并新增受控antdv-next InputNumber直接序号；Bubble名称组的下拉和序号使用同一membership判断，重名/缺失组不能选入其他组的点。Boxplot继续按有限原始样本定位。可搜索的是当前生成的label选项；完整类别名称搜索和与固定版交互的最终等价性未核实。这些修改只减少UI同步枚举的代码范围，**没有实测前后基准，不能称已优化或达标**；Facade少一次整份clone也不能称Worker/WASM整条链路零拷贝，复制次数/字节、lease与memory.grow仍按根任务协议验收。

本批新写28个source/statistics hook/绑定/80,002位置/null/far sample/Top N/zero/取消/旧完成/错误重试/有界类别/Bubble归属契约草稿。宿主原6个会话草稿另增4个source和3个statistics路由/取消/替换草稿，文件共13个。新增7个真实工作台草稿分别覆盖超预算Boxplot原始样本edit/apply/history/internal reopen、另一大范围与旧完成、关闭重开、真实低内存预算失败/retry、尚无FChart的插入Modal、Pareto三处共用统计及旧zero/Top N结果、实际Pivot面板大输出/关联保存/撤销保留Pivot。所有草稿均未执行，新spec尚未接统一验证入口；主题/焦点/IME、浏览器像素、两后端、Worker/WASM取消与传输/内存、首屏/滚动/编辑/缩放实测、真实Office重开全部保持未验收。

### P25 Histogram 分箱真实性、Candlestick 三状态与专属设置

本批按固定P25源码完成静态入口→字段→统计/布局→绘制→共同更新命令追踪。Histogram数据设置属于`Vo/Uo`，样式区`ha`只改gap；Candlestick `pa`默认临时选Rising，Doji不显示body/fill/hollow，`fa`的高低线与收盘标记均按visible展开。固定bundle没有为这些属性定义独立Sheets内容命令；本地仍以`sheet.command.add-chart/update-chart`发布隔离候选，撤销/重做按工作簿隔离，Vue仅编辑detached草稿。不能由共享命令或控制节点数量推导图型已验收。

[chart-histogram.ts](../../packages/office-studio/src/features/charts/chart-histogram.ts) 与Rust `chart/distribution.rs`对齐`Iae/Lae/Rae/zae/Bae/qa`：有限观测独立于原始null，`invalidValueStrategy:'zero'`不能把Histogram缺失样本变成0。Count总数最多1000，含启用的underflow/overflow箱，普通箱必须至少1个；Count不足、非整数/非正/超过cap、Width非正或超cap时，保留有限作者配置而运行时回退Scott Auto。Scott使用样本标准差的n−1分母。倒置或相等的两个阈值同时忽略；underflow消费≤阈值，overflow消费>阈值。Count/Auto普通范围为过滤后真实min/max，首箱含lower端、其余为左开右闭；Width用underflow阈值或真实min起始，以overflow或实际max结束，最后箱截断。恒定Count/Auto只生成[min,min]点箱。没有普通观测时，只留有观测的边箱；存在普通观测时，作者边箱即使count0仍存在。

[chart-histogram-compute.ts](../../packages/office-studio/src/features/charts/chart-histogram-compute.ts) 只核对保留的源样本、总数、边界/箱类型、计数及输出budget，不重算Scott或生成第二套边界；恒定点箱和阈值与首Count箱之间的实际间距合法。Rust多job流水已经按当前Histogram自己的regular样本判空，避免参考前一Pareto等job输出。极端有限输入使用稳定方差/插值及有限边界保护，是本地算术适配；未通过真实WASM或Office极端样本，不声称与上游溢出细节完全相同。gap有效范围0..0.9、默认.007，有限无效作者值保留而绘制回默认；Count/Width默认输入10/1来自固定面板。

Histogram加入P24同次statistics loader；binding key新增binning/underflow/overflow，不包含gap或纯paint。设置中的项目标签直接消费当前Rust物化bins：样本[0,1,2,3,4]、Count4、underflow1/overflow3显示`≤1、[2,2.5]、(2.5,3]、>3`，计数2/1/1/1，不能借原始行标签代替分箱。样本字段为单Select；分类、统计与日期最终scene仍由同一准备/发布链路控制。分箱后逐点作者索引当前指物化箱，源/阈值变更后的身份迁移仍是缺口。

[chart-candlestick.ts](../../packages/office-studio/src/features/charts/chart-candlestick.ts) 使用close>open、close<open、相等三个严格状态，默认green.500/red.500/gray.500为`#0DA471/#F05252/#4E5565`。新状态层与既有generic/legacy继承分开，源数据点覆盖最后生效；fillOpacity与outline opacity独立，hollow仅消掉body填充。Auto宽度为half-band且clamp8..32，有限正作者宽度为绝对值，无100px上限；固定UI min1，本地API仍可接受0.125等正小数。缺失OHLC观测保留空位置，HLC不合成Open；`ChartData.candleFields`使全空数据仍保留角色。源码`ma`按series数量判断面板，本地显式mapping优先，四个旧valueFields也不能覆盖已经省略Open的HLC。Close源字段身份与header命名复用prepared模型；创建Histogram/Candlestick默认隐藏legend，显式作者显示仍保留。

`xm`的overlay存在语义保持可观察：HLC缺省两组visible=true；OHLC缺省false，但显式overlay对象且未写visible则true。OHLC一旦highLowLine对象存在，普通wick只连open/close端点，再由overlay画完整high/low；即使对象visible=false，该缩短仍生效，整组clear才恢复完整普通wick。Y reverse已按high→较大价格端、较小价格端→low连接，不能用视觉top/bottom替代数值端。Dot收盘标记是size×.5长度水平线，Dash为size长度；None、size0或visible=false不画。其他标记复用P19几何，HLC没有蜡烛body。

[ChartCandlestickSettings.vue](../../packages/office-studio/src/features/charts/views/ChartCandlestickSettings.vue) 与Histogram controls全部使用现有统一Form、FormItem、Select、Switch、InputNumber、ColorPicker、Button及主题ConfigProvider；没有新增原生表单桥。OHLC设置两种overlay是本地扩展，固定`pa`只显示state/body/outline。普通OHLC走native Candlestick且每datum同步两种fill、两种border和borderColorDoji、large:false；HLC、overlay、日期轴、复杂样式或标签使用[chart-candlestick-draw.ts](../../packages/office-studio/src/features/charts/chart-candlestick-draw.ts)同一DrawCommand graphic Canvas。上游Candlestick custom series与本地native/common分流不是同一内部实现，实际两路径像素、裁剪、hover和性能待验收。

专属叶字段/整组reset显式传递undefined，经过`chartSettingsPatch→mergeChartOptions`才真正清除旧深层覆盖；serialized内部模型省略这些字段。Count/Width的binning为原子union替换，切换不保留另一arm；清binning恢复Scott Auto。恢复所选K线状态只清新作者层，legacy risingColor/fallingColor/hollow以及generic样式仍可继承；恢复全部清candlestick组含legacy，generic系列仍保留。通用字段series映射与namedSeries数组在命令边界原子替换，所以逐点恢复直接移除该作者点即可；本批把这一路径提取为`clearChartSeriesSettingPoint`供真实Vue和契约草稿共用，没有把专属深合并清除规则错误推广到所有集合。

本批纯草稿Histogram7个、Candlestick9个，真实Renderer Canvas3个、工作台Candlestick3个/Histogram2个；通用settings reset另加字段点/名称点两条，文件总5条。Rust补sample SD、无效有限作者、1000 cap含边箱、Width起止、恒定/极端点箱和多job无regular边箱草稿，并同步旧统计/协议fixture。Browser草稿覆盖临时选择无内容历史、三状态独立alpha/outline/宽度、HLC条件显隐/None/整组恢复/源Open变化仍HLC、全空角色、Histogram数据/样式分组/物化项目、真实Apply/history/internal reopen及gap比例。它们没有运行，内部JSON重开不能替代原生Office往返；新spec尚未接统一验证入口。测试、类型、构建、Rust/WASM、浏览器像素、明暗/窄屏主题、焦点/IME、首屏/滚动/编辑/缩放、Worker传输/取消/lease/memory.grow、内存与真实Office样本全部未验收。Facade/原生Office接线由根任务单独交接，本段不借静态类型宣称互操作完成。

### P26 Treemap 完整层级、父叶标签、背景与 bitmap 悬停

固定`Ua`由Treemap/Sunburst共用，字段必须是非空、无重复的有序根→叶列表，与唯一权重字段互斥。`Aae`的身份区分字符串、number（含−0）、boolean和Date timestamp；同显示名不同typed key不会合并。只有null/undefined/空字符串为空，空格名称仍有效。填空缓存按祖先typed key作用域维护：祖先变化清掉后代；根空白无缓存或中间空白之后再出现非空后代则该行无效，尾部空白可成为父节点自有事实。路径cache在检验权重前更新，因而非法权重行仍可建立下一行的carry作用域；权重严格只接finite number，数值外观字符串不进入树。

零和负数进入terminal selfValue，同完整typed路径进行signed累加。postorder只保留正子节点，以`selfValue + 正子节点value合计`求节点value；最终非正节点不物化。siblings按value降序，相等按首次创建顺序；根paletteIndex使用最终正根的首次创建顺序，与值排序分开。稳定ID是typed路径key的JSON。本地有限累加溢出显式报错，不从权威源删行。原始source records、formattedRecords与派生树保持分离；源显示格式的推断保留分组空格、数字后的单位空格、百分比及signed selfValue，但这仍是样本显示文字推断，完整自定义数字格式与Pivot格式元数据未验收。

[chart-hierarchy-data.ts](../../packages/office-studio/src/features/charts/chart-hierarchy-data.ts)准备typed key/label字典、column-major Float64值与Uint8状态，`hierarchy` job进入现有`chartStatistics`协议，字段为`hierarchyColumns/labelColumns/keyCount/labelCount`。本地[chart-hierarchy-compute.ts](../../packages/office-studio/src/features/charts/chart-hierarchy-compute.ts)的portable reference和Rust [hierarchy.rs](../../crates/office-compute/src/chart/hierarchy.rs)消费相同数字包；正式异步源投影通过`ChartStatisticsService`注入实际compute client。输出values按节点为`[value,selfValue]`，indices为`[nodeCount,rootCount,...(parent,key,label,sourceOrder),...rootIndex]`，parent sentinel为`0xffffffff`；preorder保证parent-before-child，末尾rootIndex按首次创建顺序。字符串字典由owner保留，不transfer Rust linear-memory buffer；通用统计decoder显式拒绝hierarchy，由专属decoder与保留字典校验节点/根预算、parent/depth、typed ID、positive totals、sourceOrder及palette tuples。

现有任务预算、权限、unit/revision、取消和transfer metrics复用于层级job；[chart-compute-memory.ts](../../packages/office-studio/src/features/charts/chart-compute-memory.ts)把所有后代metadata、共同布局/native节点与父残差纳入保守准入估算，不能仅按根类别计场景内存。普通JS输入进入WASM和有效结果区间离开WASM的必要复制仍需实测记录，typed dictionary编码或transferable不等于整条链路零拷贝。**本地WASM产物尚未重建，不认识新增hierarchy job；本段不宣称当前产物已可运行。** 统一验证阶段还须验证真实Promise/宏任务取消、lease释放、memory.grow视图重建、复制字节、深树与大样本。

Treemap`eke`额外将有后代且selfValue>0的父事实生成未命名残差叶，ID为`__univer_treemap_residual__:${semanticId}`，保留semanticParentId及原父模型。负父selfValue不生成残差；绘制容器权重改用实际render children总量，语义value/selfValue不改写。例如A1:C8草稿中North语义13/self−5、实际子面积18；South语义9/self3、Leaf A6及残差3。South虽语义排序第二，仍取首次出现palette0。Canvas [chart-treemap-layout.ts](../../packages/office-studio/src/features/charts/chart-treemap-layout.ts)使用迭代squarified，banner保留fontSize+8的标题带，overlapping不预留子面积，none不画父名称；leaf.visible不控制父标签。

`$Oe`只拼CategoryName2和Value8，组合内容用换行；残差不造类别名，Value使用self fact。`H6`仅对#rrggbb按深度向白混合12%，最多48%；`U6`显式文字色优先，否则hex亮度>150取黑、其余白。Canvas [chart-treemap-draw.ts](../../packages/office-studio/src/features/charts/chart-treemap-draw.ts)按真实字体度量完整换行，内容整体放不下即隐藏，不截断或压缩；overlapping父文字最后画在后代之上。`wp`的pattern背景消费均值色（pctN使用N%、其他50%），作者pattern仍保留；solid/linearGradient为独立union原子替换，渐变按标签矩形定位。

Native [echarts-treemap-option.ts](../../packages/office-studio/src/features/charts/echarts-treemap-option.ts)保持animationfalse/nodeClickfalse/roamfalse/breadcrumbfalse/visibleMin0、隐藏legend及ancestor emphasis；数据已稳定排序，因此sortfalse避免native desc反转等权dataIndex。锁定6.1.0的TreemapView将parent注册为background Rect，而leaf注册为Group、文字挂其content Rect，故[echarts-treemap-layout.ts](../../packages/office-studio/src/features/charts/echarts-treemap-layout.ts)显式遍历叶group。marked extension在layout后回填overlapping标题带，postUpdate移除原view会恢复width/height的text.beforeUpdate setter，清height/lineOverflow后按完整文字测fit，normal/emphasis/select/blur同步ignore并抬高parent z2。ECharts公开LabelOption仅声明string/image背景，本地marker在postUpdate给锁定zrender Text runtime设置relative gradient；这些都是需要浏览器复核的内部适配契约，不能由静态option推导实际像素通过。

P27批次补修原生纹理和alpha：leaf所有label状态均关闭时Text可不存在，因此以leaf Group内最高z2的content Rect应用[echarts-hierarchy-pattern.ts](../../packages/office-studio/src/features/charts/echarts-hierarchy-pattern.ts)，不再依赖Text host，也不把纹理刷到parent/background border Rect。锁定TreemapView的shape style mapper没有opacity，旧itemStyle.opacity不能代表实际填充alpha；datum现用已有独立fill/border helper将fillOpacity嵌进填充色，borderColor保留自身alpha。原生tile同步创建、四状态共享，DPR尺寸及最多128项缓存有界；这些只是实现和准入策略，真实透明度、内存和释放仍未验收。背景setter先clone/validate后建options/label，非法union不产生半更新。

实际Sheet overlay已转发局部pointer，[echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts)新增`officeChartPointer.mode='treemap'`，转发zrender mousemove/mouseout及公开showTip/hideTip，并同步锁定animation frame listener的pending hover状态后更新bitmap sourceVersion。leave先用范围外mousemove清掉_hovered，让同节点reentry可以再次mouseover。richText tooltip绘在Canvas，按semantic path和signed value显示，残差显示原父value/self/子合计，使用主题颜色；stateAnimation及tooltip delay为0。pointer不写作者模型/内容历史；实际hover、清理、重复hover缓存和native碰撞仍未验收。

`ChartConfigBuilder/FChart`已落盘六公开操作：`setHierarchyFields/setValueField/setParentLabelLayout/resetParentLabelLayout/setTreemapLabel/resetTreemapLabel`，共用[chart-treemap-api.ts](../../packages/office-studio/src/features/charts/chart-treemap-api.ts)校验、共同add/update命令与工作簿内容历史。reset通过显式undefined清对应组，不清其他组。shared option merge对backgroundColor对象union原子替换，null恢复或undefined重置不会留下旧pattern/gradient的另一arm。真实保存仍是内部模型快照，Office原生可编辑Treemap写出/重开不能由这些API或JSON测试证明。

[ChartTreemapSettings.vue](../../packages/office-studio/src/features/charts/views/ChartTreemapSettings.vue)使用统一FormItem、Select、Switch、Input/InputNumber、ColorPicker、Button与共同ChartColorInput。父layout独立于leaf.visible；category/value多选、Value条件format、五个font控件、solid/gradient/pattern背景、父/叶/全部恢复均修改隔离草稿后一次apply。通用series label/marker及legend控件对Treemap隐藏，避免不活动设置冒充入口。背景完整编辑和全部组reset属于本地扩展；没有新增原生输入桥。继承ConfigProvider不代表明暗/弹层/窄屏、焦点恢复或IME已验证。

| 草稿文件/样本 | 计划覆盖的行为 | 当前状态 |
| --- | --- | --- |
| `tests/fixtures/chart-hierarchy.json`、`chart-hierarchy.contract.test.mjs`、`chart/hierarchy_tests.rs` | 四份独立numeric golden覆盖signed自有事实、祖先carry、等权创建顺序、重复/zero剪枝；TS另8项typed identity/格式/映射/decoder/transfer/取消/过期，Rust另3项fixture与单owner/not-ready/consumed、预算/token、互斥字段/有限溢出草稿。 | 全部未执行；portable与Rust预期相同不等于实际WASM已验证。 |
| `chart-treemap.contract.test.mjs`、`chart-treemap-facade.contract.test.mjs` | 8项residual面积/semantic totals、banner/overlap、完整fit、palette/contrast/background/控制/预算、native填充与outline alpha；背景原子失败补到已有草稿。2项真实Builder/FChart六方法、共享独立numeric transport、命令/history/detached snapshots及internal reopen。 | 全部未执行；没有Office、真实复制/内存或UI结果。 |
| `chart-treemap-renderer.browser.test.mjs` | 5项真实native Canvas草稿：父叶独立/残差/颜色像素、overlap z2及清clamp、完整超长字符串/relative gradient、所有label状态无Text时仍绘pattern并保留outline色、残差tooltip/ancestor leave/reentry。 | 全部未执行；依赖锁定Tree/view/zrender的适配必须真实验收。 |
| `tests/browser/chart-treemap.spec.ts` | 4项工作台草稿：signed源和父叶条件/一次apply/history/reopen、category+value/format/gradient→pattern原子替换/恢复、六live API及源变化、真实bitmap sourceVersion hover/leave/reentry无历史。 | 全部未执行，尚未接统一验证入口；没有启动dev或浏览器。 |

仍待核实/验收：源结构/过滤/聚合变化后的逐点作者身份迁移、所有自定义源数字格式、Canvas/native squarified几何与完整wrap差异、native extension的内部生命周期、完整公开演示效果、Office原生Treemap往返、Docs宿主、明暗/弹层/焦点/IME、真实深树/大样本/首屏/滚动/编辑/缩放与Worker/WASM性能和内存。当前代码及草稿不将这些项标为完成。

### P27 Sunburst 语义角度、完整标签、中心与有序映射

Sunburst复用P26的typed path、scoped carry、strict signed数值、positive postorder语义树及同一`hierarchy` Worker/Rust协议；没有另一份权威聚合引擎。[chart-sunburst.ts](../../packages/office-studio/src/features/charts/chart-sunburst.ts)按固定`qOe`追加正父selfValue残差，保留原父value/selfValue。A1:C8样本North语义13/self−5，后代合计18；South9/self3，Leaf A6及残差3；中心总量22。Sunburst不套Treemap的父render value18：根North占13/22圈，而其children沿全局角度单位共占18/22圈，可超出父扇区。正残差使用`__univer_sunburst_residual__:${semanticId}`、原父path及独立own value，名称空且不绘标签；负父事实没有残差。

[chart-sunburst-layout.ts](../../packages/office-studio/src/features/charts/chart-sunburst-layout.ts)迭代生成扇区，起角−π/2、半径18%至88%、等厚depth rings；节点和残差继续取首次root palette及固定`V6`的六位hex每层18%混白、上限54%。原生和共有布局使用语义根total及全局角度比例，不钳制children到parent或悄悄改值；total有限累加溢出明确报错。扁平布局fact不重复携带完整children树；native datum只存`semanticPath/self/childValue`等必要fact，权威源模型和render data分离。

固定`KOe`只消费CategoryName2/Value8并换行；Empty0产生空字符串，残差及小于总量4%的节点隐藏标签。`JOe/YOe`令首圈tangential、后代radial；共有[chart-sunburst-draw.ts](../../packages/office-studio/src/features/charts/chart-sunburst-draw.ts)按锁定SunburstPiece的upright flip、auto/inside/outside锚点及真实字体完整旋转bbox绘制，不截断或压缩源内容。[echarts-sunburst-layout.ts](../../packages/office-studio/src/features/charts/echarts-sunburst-layout.ts)在marked native postUpdate移除clamp、测完整旋转bbox并复用同一`chartSunburstLabelFits`，normal/emphasis/select/blur同步hide。缺尺寸、非finite/负尺寸或完整文字越过sector bbox均不绘；最终碰撞/完整字体差异需要浏览器复核。

[echarts-sunburst-option.ts](../../packages/office-studio/src/features/charts/echarts-sunburst-option.ts)保持固定`QOe`的animationfalse/nodeClickfalse/sortnull、隐藏legend及ancestor emphasis；local plot适配把18%/88%公共半径转成像素，而上游写百分比，属于明确布局差异。solid、真实relative label gradient及pattern均保留作者union；pattern label背景按P26均值色，节点填充pattern使用同一同步Canvas tile与独立alpha。没有将decal、CanvasPattern或运行时gradient写回作者模型，未通过真实像素或资源验收。

中心按`bm`缺省hidden/font16、节点按`ym`缺省visible/category/font12。独立center total、数字format和五font叶字段实际生成Canvas/native中心文字，兼容旧快照`centerLabel.font`并优先新直接字体字段；公开setter只接受固定直接字段。`XOe`的value/share tooltip保留语义path，当self/child同时非零显示自身/子合计；residual显示自己的3及父South语义self3/child6。上游HTML tooltip有RTL dir，本地richText绘在Canvas，使用统一主题色；完整RTL及公开提示效果尚未核实。Treemap/Sunburst共用bitmap pointer转发、ancestor/leave/reentry状态清理和sourceVersion刷新，不写内容历史。

六公开操作`setHierarchyFields/setValueField/setSunburstLabel/resetSunburstLabel/setCenterLabel/resetCenterLabel`在[chart-sunburst-api.ts](../../packages/office-studio/src/features/charts/chart-sunburst-api.ts)集中验证、clone/nullable恢复，Builder/live Facade走共同add/update命令及工作簿内容历史。非法或inactive图型的输入不发布候选；mapping必须互斥并在源字段范围内。单组reset显式undefined，不清另组；整组background替换不串旧gradient/pattern字段。权限、异步准备取消/过期和保存仍沿共同宿主服务；内部JSON重开不证明原生Office可编辑Sunburst往返。

[ChartSunburstSettings.vue](../../packages/office-studio/src/features/charts/views/ChartSunburstSettings.vue)以统一antdv-next FormItem/Select/Switch/Input/InputNumber/ColorPicker/Button/Alert和ChartColorInput编辑隔离草稿后一次apply。固定`Jo`在节点/中心间临时切换，本地同时呈现中心、节点两组，记录为布局差异；各visible决定字体/格式/内容显隐，关闭保留作者覆盖，通用series label/marker及legend不显示不活动入口。背景完整编辑、两组font family及三reset属于本地扩展。没有新增原生表单控件；统一ConfigProvider接入不等于主题、弹层、焦点恢复及IME已经验证。

两图型的数据区共用[chart-hierarchy-settings.ts](../../packages/office-studio/src/features/charts/chart-hierarchy-settings.ts)和[ChartHierarchyMappingFields.vue](../../packages/office-studio/src/features/charts/views/ChartHierarchyMappingFields.vue)：新增、逐层替换、上移/下移、移除、单权重使用同一不空/无重复/互斥字段验证，最后一层不可移除，非法direction原子拒绝。候选及输入数组隔离，既有chart不在apply前更新；固定`No`的列表/添加即提交交互在本地改成可替换Select、临时pending与独立添加Button，分类计数见第2节。普通、插入及透视设置共享此区，不扩充Ribbon根数。

| 草稿文件/样本 | 计划覆盖的行为 | 当前状态 |
| --- | --- | --- |
| `fixtures/chart-sunburst-fixtures.mjs`、`fixtures/chart-hierarchy-facade-transport.mjs` | A1:C8 signed样本复用原始数据但生成独立Sunburst模型；两图型Facade共用独立numeric golden transport，真实structuredClone transfer并要求输入detached，不调用层级实现造期望。 | 仅草稿；没有运行Worker、WASM或复制测量。 |
| `chart-sunburst.contract.test.mjs` | 12项signed/residual/center22、全局角度与环径、默认/legacy fonts、category/value/Empty0/4%、tint/旋转/outside、完整fit、native flat fact/gradient/pattern、tooltip/center、条件控件/一次update/reset、背景原子失败、nullable/API及内存准入/total overflow。 | 全部未执行；几何草稿按容差比较，不等于实际Canvas/native已验收。 |
| `chart-sunburst-facade.contract.test.mjs` | 4项六Builder/Enum/输入隔离、live signed layout/独立reset/history/internal reopen、非法/inactive原子拒绝、worksheet保护阻止全部操作/direct command/insert且无Worker分配。历史比较实际chart内容/版本及savepoint，避免把save元数据作为功能契约。 | 全部未执行；Office、异步真实取消/权限/UI闭环未验收。 |
| `chart-sunburst-renderer.browser.test.mjs` | 6项actual native signed layout/center、residual/Empty0/4%四状态隐藏、超长完整文字/真实gradient、hidden-label pattern tile、residual tooltip/ancestor/leave/reentry/sourceVersion、主题/resize/DPR/retain/dispose。 | 全部未执行；没有启动server或浏览器，native生命周期/像素/资源仍未验收。 |
| `tests/browser/chart-sunburst.spec.ts` | 4项真实工作台defaults/center-node条件/一次apply/history/internal reopen、category/value/position/format/gradient→pattern/reset、六live API及源C3=2后的center12、精确residual bitmap hover不写history。 | 全部未执行；这些草稿不替代主题/弹层/焦点/IME、Office或大样本验收。 |
| `chart-hierarchy-settings.contract.test.mjs`、`tests/browser/chart-hierarchy-settings.spec.ts`、`tests/browser/helpers/chart-hierarchy-settings.ts` | 3项共享add/replace/move/remove/weight、非法/最后层级/互斥原子拒绝及两图型草稿隔离；2项真实工作台A1:E8顺序最终[4,3,1]/weight2、一次apply/history/internal reopen/cancel。 | 全部未执行；真实UI/Worker/IME未验收。 |

仍待核实/验收：原生递归深树及主线程layout预算/取消、真实WASM重建与浏览器binding/lease/memory.grow、输入/输出复制次数及字节、真实大样本/帧时间/长任务/内存、源结构/过滤/聚合后的作者点身份迁移、所有源数字格式/Pivot metadata、Canvas/native完整字体和outside碰撞、RTL/公开演示/原生OfficeSunburst与Treemap文件样本、Docs宿主和主题/弹层/焦点/IME。完整Pro范围不变，不由这些草稿或当前静态代码宣称已验收。

## 6. Canvas 与 UI 技术路径

- `echarts-renderer.ts` 使用锁定的 ECharts 6.1.0，tree-shakeable 注册 Bar/Line/Pie/Radar/Scatter/Funnel/Gauge/Graph/Sankey/Heatmap/Boxplot/Candlestick/Treemap/Sunburst、AxisPointer 等 components 和 CanvasRenderer。领域类型、命令/历史、数据源与 Office 模型由本项目控制。
- 3D、复合饼、词云、Chord、Waterfall、Sankey、Relation 使用共享 Canvas graphic；HLC Candlestick、极端 Gauge范围、Heatmap extent差值非finite、有趋势/误差的图表同样走共享图元；Radar的单点覆盖/值标签、分类旋转、径向反向或极端有限范围，以及Boxplot的median数据标签使用该路径。实际纹理通常切至共有graphic，Treemap/Sunburst使用同一mask的同步native Canvas tile例外；Line/Area/组合线型作者显式渐变，启用自动渐变的Line/Area逐段颜色/透明度覆盖，以及Pareto的显式累计渐变/点覆盖也切至该公共路径。普通柱/条/Pie渐变使用原生serializable paint。Heatmap可见色标和hover在两路径共用graphic，常规单元格仍为native Heatmap。布局/命中用同源公开ChartLayout，不把“同名ECharts series存在”当作图型等价证明。
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
3. Radar轴扩展、graphic字体碰撞/隐藏重叠和径向指示器效果；Boxplot逐原始异常点独立覆盖；标题旋转竖排差异、manualLayout/rotate/完整碰撞与公开/Office文字盒效果；日期自动碰撞/稀疏、同桶pointer公开效果、非法派生外延与来源编辑/import/Docs的最终发布校验覆盖；Gauge源显示格式/最终交互；所有纹理、自动/显式渐变与绘图区的上游可见效果/边界、代表色、极端/native/Office验证。Pareto共享接线/Rust signed同步、19项完整marker、组合数值mask标签/缺失位置、物化点迁移及公开/Office样本仍需继续完成。
4. 长图例/aggregate 图例、复杂图布局/碰撞/字体、完整 3D 相机/遮挡/透明排序及公开效果。
5. 源准备、预览、放置过滤和复杂布局的有界大计算/Worker 路径；实际进度、预算/取消、传输字节与内存峰值。
6. 全类型可编辑 XLSX/DOCX 写出/读回、未知信息/不等价原生图型的源/配置保留与诊断；独立 Office 重开和再编辑。
7. 相同数据集的首屏/滚动/编辑/缩放、WASM/Worker 任务、长任务、帧时间、JS/WASM/GPU 内存与前后基准。

这些是继续实施的缺口，不是最终范围缩减。

## 8. 测试草稿与验证状态

| 草稿/已有夹具 | 本轮新增或涉及的覆盖 | 状态 |
| --- | --- | --- |
| `chart-fill-pattern.contract.test.mjs` | P14 的 10 个纹理继承/null/undefined/完整对象替换、54 个 mask 唯一性/覆盖/密度、透明前景/背景不重叠、凹面/环孔裁剪、marker/边框透明度、Scatter/Heatmap/Boxplot/Area/Pie graphic 路径、候选原子拒绝与颜色/纹理控件草稿。 | 未执行；不是公开像素、Office 往返或大样本性能验收。 |
| `chart-facade.contract.test.mjs`、`tests/browser/chart-workbench.spec.ts` 新增 P14 场景 | 既有 Facade 新增 2 个全系列/字段/名称/逐点纹理保存重开、继承恢复/历史、非法纹理拒绝草稿；浏览器新增 1 个隔离草稿、无纹理实际 apply、关闭/undo/JSON reopen 场景。 | 全部未执行；原生 XLSX/DOCX 及独立 Office 重开仍待验收。 |
| `chart-auto-gradient.contract.test.mjs` | P15 的 9 个图型默认/reset、25 个 UI 与 29 个预设目录区分、有效继承/纹理优先、负值/轴反向/横向、alpha 一次消费、Canvas/native 坐标、色带面积/共享边界/端点钳制、线/环孔、退化和每次 append 预算原子拒绝、逐段作者覆盖场景。 | 新草稿，未执行；颜色预设/原生绘制/Office/性能未验收。 |
| `chart-plot-area.contract.test.mjs` | P15 的 10 个完整外框归一化/移动缩放、活动次轴/标题 seed、最小外框、编辑/清空、无效草稿保留源系列并纠正、完整矩形原子 patch/拒绝、31 个目录换型保留、Canvas/指示器/native grid、Pie/Radar/3D 几何一致及 Excel-view 快照隔离场景。 | 新草稿，未执行；浮点几何使用容差，无浏览器或 Office 样本验收。 |
| `chart-gradient-layout-facade.contract.test.mjs` | 3 个 P15 Builder/实时 Facade 有效 description/显式覆盖/reset、PlotArea 公共 setOptions/原子拒绝、内容历史、模型隔离和内部 JSON 保存重开集成草稿。 | 接口由根任务接线；草稿未执行，不是原生文件往返验收。 |
| `tests/browser/chart-workbench.spec.ts` 新增 P15 场景 | 2 个真实设置面板草稿：渐变 reset 与手动矩形在一次 apply 前不发布内容，apply 后位图版本/布局/history/reopen；无效矩形报告约束且不写内容/历史，纠正后可应用和逐次 undo。 | 未启动 dev server/浏览器，未执行；没有截图、像素、主题、焦点/IME 或 Office 结果。 |
| `chart-description.contract.test.mjs` | P16 的 13 个 data-free 字段 metadata、canonical ID、源字段 Record/未活动作者层、真实名称与空名称、继承/纯主题色省略、自动 PlotArea 省略、横向/主次/Heatmap/Radar 轴绑定、palette/渐变复制及原始值/几何/paint 排除草稿。 | 未执行；公共 getDescription 是本地接口扩展，不等于 Pivot/Docs 元数据或文件验收。 |
| `chart-explicit-gradient.contract.test.mjs` | P16 的 15 个严格纯数据构造/2–256 stop、31/25/符号 capability、原子 merge/nullable/隔离、CSS alpha 与作者 stop 顺序/方向/纹理优先、3D 阴影、硬/窄/退化色带面积/每次 append 预算、31 项实际 paint 差异、signed/reverse/native Bar/Pie bounds、Line/Area/组合公共图元、换型保留/新写拒绝以及真实编辑 helper/控件草稿。 | 未执行；option/DrawCommand 草稿不含真实 native SSR 或像素结果，极端坐标、上游代表色、两后端和性能仍未验收。 |
| `chart-explicit-gradient-facade.contract.test.mjs`、`tests/browser/chart-gradient.spec.ts` | 3 个 Builder/实时 API 草稿覆盖公开构造/描述、nullable/history/internal reopen、错误类型与作者模型隔离；2 个真实颜色编辑器草稿覆盖方向/alpha/stop 顺序、一键 apply/撤销重开以及不活动渐变恢复/取消。 | 未执行；没有 dev server、浏览器、截图、主题/焦点/IME 或原生 Office 结果。 |
| `chart-pareto.contract.test.mjs` | P17 的17个 signed 聚合/零负合计/等值顺序/有限 raw 与总量溢出百分比、单字段候选、独立默认/通用继承、深合并/nullable/宽度忽略/拒绝、原生与 Canvas 独立 fill/border/marker alpha、作者边框默认/虚线/矩形 width0、业务/累计标签及 percent format、log/单边范围、渐变 native/graphic、聚合排序索引/点样式/TopN 准备、描述隔离、源字段回退与真实控件显隐/草稿消费场景。 | 共享类型、六操作与 Rust 合计已落盘；全部草稿未执行，真实 native renderer/文件/浏览器/性能没有运行。 |
| `chart-markers-labels.contract.test.mjs` | P19 的16个19形状验证/有界几何/面积/dot-dash精确箱/空心无填充/alpha/隐藏、Pareto独立轮廓、native aliases/活动marker路径、组合mask/Empty0/逗号顺序/原子拒绝/克隆、多选隔离/隐藏字体/BubbleSize保留、完整位置/正负反向短柱、真实column/bar-horizontal anchor、相等业务占比与截断累计草稿。 | 未执行；共有命令/option 断言不是原生像素、浏览器两后端、主题/焦点/IME、Office或性能验收。 |
| `chart-pareto-facade.contract.test.mjs`、`tests/browser/chart-pareto.spec.ts` | 3个六操作 Builder/实时集成草稿覆盖深合并/非正宽度/null reset、numericmask/19形状/内部位置、内部保存重开/history/原子拒绝；3个真实专属面板草稿覆盖signed分类、bar/line/point编辑、组合多选/insideBase/emptyRoundRect/emptyPin、一次apply、整组恢复/cancel/undo/redo/internal reopen。 | 六个 Facade 操作已落盘；全部未执行，不是 API 已验收、源索引全迁移或 Office 往返证明。 |
| `chart-waterfall-layout.contract.test.mjs` | P18 的 15 个完整 computed scene 必需/陈旧 metadata 拒绝、排列/堆积真实与合成身份、完整 band 柱宽、跨 subtotal/主次轴/斜 total 连接、reverse/log、源/扩展 style 索引、独立 global role/point/legend、neutral connector、空 scene 和实际 graphic geometry/style/clip 草稿。 | 未执行；fixed tuples 不证明 Rust/WASM 正确，graphic option 断言不证明浏览器像素或 Office 效果。 |
| `chart-waterfall-settings.contract.test.mjs` | P18 的 12 个创建 own-property/首有限来源/输入隔离、layout/reset/stack 清 roles、非法新 roles 原子拒绝、重复身份最后输入胜出且保留作者数组、源 role edit/reset/边界、active role/null-as-zero/真实 auto identity、16,384 有界选择/完整直接索引、global style 隔离、有效 gap 与保留不活动 overlap 草稿。 | 未执行；不是 DOM、主题、焦点/IME 或大样本性能验收。 |
| `chart-waterfall-facade.contract.test.mjs`、`tests/browser/chart-waterfall.spec.ts` | 4 个 Builder/12 操作/真实服务与严格 decoder/transfer 固定输入、live role/duplicate/stack/subtotal/reset/history/internal reopen、错误类型和非法 patch 原子拒绝、worksheet 权限草稿；3 个真实面板 role/source-total/connector、stack/global subtotal/hide 和整组/单点/源 role/connector reset/cancel/apply/undo/reopen 草稿。 | Waterfall Facade 操作已落盘，所有草稿未执行；没有运行 dev server/浏览器/截图/像素、真实 WASM binding 或原生 Office 样本。 |
| `chart-sankey.contract.test.mjs` | P20的18个重复边/max流量/self-loop/完整cycle/非法权重、MAX有限流量/justify/身份/端口/band/budget/长链、prepared布局/源色/alpha/font/全部位置/mask/percentage、真实命中/定向轨迹/hover不变base/view拖动与源重投影草稿。 | 未执行；不是循环碰撞、公开布局/轨迹、两后端或性能验收。 |
| `chart-sankey-facade.contract.test.mjs`、`chart-sankey-renderer.browser.test.mjs`、`tests/browser/chart-sankey.spec.ts` | 3个base Builder/继承/克隆/description/live历史/源更新/内部重开/无效mapping原子拒绝；4个graphic group/retained view/cancel/来源尺寸主题DPR重置/retain-dispose；3个真实workbench hover/拖动无内容历史/取消源变更/antd标签草稿。 | 全部未执行；不包含截图或像素，内部JSON不能代替Office原生往返，主题/焦点/IME与大图Worker仍未验收。 |
| `chart-heatmap.contract.test.mjs` | P21的18个actual/empty/constant/extreme/subnormal extent、8段边界/越界、X字段/Y分类/反向/命中、formatted投影及真实Sheets数字格式、唯一value标签/位置/unit/legend、native tuple/作者paint、hover bins与源不变、控件顺序/显隐/reset/草稿隔离。 | 全部未执行；option/DrawCommand和有界数组不证明native像素、公开交互或性能，Pivot格式仍缺。 |
| `chart-heatmap-facade.contract.test.mjs`、`chart-heatmap-renderer.browser.test.mjs`、`tests/browser/chart-heatmap.spec.ts` | 根任务3个8方法/Enum/输入隔离/nullable/reset/live-history/internal reopen/非法候选草稿；4个真实native rect/text/颜色、graphic-only hover/缓存/作者范围、8段/越界、纹理graphic/source/theme/DPR/释放草稿；4个真实workbench cell/colorbar无内容历史、leave/cancel/源刷新、antd设置一次apply/history/reopen和reset/cancel草稿。 | 全部未执行；无dev server/浏览器/截图/像素、主题/焦点/IME、两后端、原生Office或性能结果。 |
| `packages/office-studio/tests/chart-data-layout.contract.test.mjs` | 45 项空态、原始 null/精度、双方向可见数据缺省/Switch 重投影、透视角色、DFS/样式、转换/原子拒绝、Bubble marker、分组统计身份、Bar 次轴、堆积/大权重、Radar 单点/标签、14 万异常点/瀑布、3D 深度/横向标签；Gauge 默认/5 预设、有效/无效 tile、源索引/原文/clamp、极端角度、零指示器、nullable/非法 patch、字体/格式继承、控件/草稿重置链路、真实 native 弧宽/层级/文字/指针变换的 SSR 草稿。 | 草稿，未执行，包括 native SSR；大数组断言不是性能基准。 |
| `chart-radar-boxplot.contract.test.mjs` | Radar nullable/字体继承、条件显隐、线性 signed/极端范围、源分类索引、分类/径向反向、值标签和 graphic；native Radar 的真实文本/字体/范围/半径/圆网 SSR；Boxplot 原始字段/位置、已知四分位、透明度/完整 cap、native path/端点 SSR、主/次轴 median anchor/标签、log 保留统计与合法正异常点；nullable patch 对象/数组隔离、危险/稀疏对象拒绝。 | 新草稿，未执行，包括 native SSR；不是浏览器交互或性能验收。 |
| `chart-radar-defaults.contract.test.mjs`、`tests/browser/chart-radar.spec.ts` | P10/P11静态追踪后的6个default true/12px、显式/legacy false、nullable继承、5段/extreme、auto/manual/零半径、9个reset/草稿隔离、native/Canvas共同数值标签草稿；2个真实workbench默认/隐藏字段保留/一次apply/reopen及shape/fill/label整组reset/cancel/history草稿。 | 全部未执行；P10 getter/default仅有固定ES静态证据，可见性真实源链/Office仍待验收；Radar无浏览器像素、完整碰撞、主题焦点IME、两后端或性能结果。 |
| `chart-facade.contract.test.mjs` | 类型/option patch、cell placement、depthAxis/view 独立清除；Gauge 公开导出与 FOffice 实例 Enum、单值绑定、全部样式/量程/角度/刻度/ranges setter/reset、5 预设/全组清除、对象隔离/非法输入原子拒绝、更新/撤销/重做/内部保存重建；另新增 Radar/Cartesian 4 个、Boxplot 1 个及 Bubble/名称 3 个草稿，覆盖 6/7/8 个 clear target、轴标题/partial 次轴/linear 清除 logBase、原始箱线样本/主次轴、名称独立覆盖/消失重现/7 个恢复目标/两集合清除、mapping/错误输入/历史/内部重开。 | 本轮修改未执行。 |
| `chart-series-selector.contract.test.mjs` | 10 个名称精确匹配/空白/大小写/空字符串/危险字样、字段与名称继承/逐点边框/nullable、集合替换/未活动保留、非法选择器/稀疏或重名配置/源歧义拒绝；Bubble Canvas 与真实 native scatter 名称/标记/标签/图例/主次轴 SSR；有效次轴边距；设置目标/所属组点范围/继承读写/重名覆盖恢复；Boxplot sourceName/标签/次轴；组合类型/统计能力及轴换型规则；专属色阶不覆盖名称颜色及 Gauge 字体的名称继承/恢复。 | 新草稿，未执行，包括 native SSR；不是浏览器交互、Office 往返或性能验收。 |
| `chart-axis-interaction.contract.test.mjs` | 15 个 P12 日期虚构闰日/1904/ISO 时区、真实间距/reverse/缺失配对/Pareto 血缘/Bubble 候选原子拒绝、标题继承/边距/旋转/native/3D、四指示器参数/缓存查询/主次轴/log/日期/横向/Heatmap/Radar/清理、控件顺序与日期预检、CSS 坐标契约草稿。 | 新草稿，未执行；不是性能或最终 UI 验收。 |
| `chart-axis-titles.contract.test.mjs`、`chart-axis-titles-renderer.browser.test.mjs`、`tests/browser/chart-axis-titles.spec.ts` | P12追加固定wke追踪后7个物理起止/输入span顺序/reverse/Bar真实主次轴/12px与继承/手动plot span/7叶reset/两路径同源/Heatmap色标草稿；2个实际native-content文字glyph bounds/字体/反向不移动及native/graphic hover保留标题和内容草稿；2个真实workbench草稿/一次apply/独立及整组reset/cancel/历史/内部JSON重开草稿。 | 全部未执行；固定ES静态追踪不是浏览器像素、两后端、主题/焦点/IME、Office或性能验收。新spec尚未接入统一验证入口。 |
| `chart-date-calendar.contract.test.mjs`、`chart-date-axis.contract.test.mjs` | P22的12个calendar原锚点月末/闰年/1900与1904/serial0/首日桶/基本与主单位/作者边界/日小数/跨年/亚秒/非法空源/10,000步预算草稿；14个稳定重复行/源style/缺失连接/native与common extent/字体grid/tick/clip/Bar/3D/Bubble/Boxplot/Candlestick/Pareto/Histogram最终中心/趋势与X误差外延/pointer/reset/关闭后numeric 0.25/设置拒绝草稿。 | 全部未执行；基础layout可生成后再测派生外延预算，仍不是Rust真实结果、浏览器或全部发布入口验收。 |
| `chart-date-axis-renderer.browser.test.mjs`、`tests/browser/chart-date-axis.spec.ts` | P22的4个实际Canvas native点/ID/共同glyph一次/日期pointer内容保留与leave/major只改密度/basic改间距/1904 Bubble与graphic Bar/3D/Boxplot草稿；4个真实antdv工作台三时间字段+numericY/一次apply/独立reset/关闭丢草稿/日期关闭后保留元数据并输入0.25/开关reset/非法小数和预算拒绝/history/internal reopen草稿。 | 全部未执行；没有启动dev server/浏览器或截图，新spec尚未接入统一验证入口，主题/焦点/IME、两后端、Office和性能未验收。 |
| `chart-settings-resets.contract.test.mjs` | 5个真实chartSettingsPatch→updateChart草稿：日期/数值单位与标题、Heatmap/Radar叶字段和整组的显式undefined恢复；P25另增字段点/名称点作者集合原子替换后的共享/字段继承、其他点/系列保留及revision。 | 全部未执行。专属深合并使用undefined，series/namedSeries替换集合使用移除作者点，各按真实边界验证；不代表异步命令/history/Office已验收。 |
| `chart-boxplot-outliers.contract.test.mjs` | P23的16个固定默认/旧IQR/奇偶样本、null位置/重复异常值、outlierSourceIndices、逐样本颜色/形状/大小/alpha/边框/隐藏/纹理、inlier不覆盖summary、字段/名称层与四个真实undefined reset草稿。 | 全部未执行；TS预期不证明Rust、浏览器、Office或资源指标已通过。 |
| `chart-boxplot-renderer.browser.test.mjs`、`tests/browser/chart-boxplot.spec.ts` | P23的4个实际Canvas native标记ID/transform/path样式及graphic/inactive/隐藏/lifecycle草稿；4个真实工作台独立重复点、null Form与源激活、17,000远端样本/edit/cancel/reset、固定默认恢复/history/internal reopen草稿。 | 共8个，全部未执行；17,000仅超过selector cap，不超过预览预算；应用后按真实行为显式关闭Drawer。没有运行浏览器/像素/主题/焦点/IME/Office，新spec未接统一验证入口。 |
| `chart-settings-source.contract.test.mjs`、`tests/browser/chart-settings-source.spec.ts` | P24的28个source/statistics身份/80,002源位置/原数组/derived拒绝/绑定和session变化/取消/失败重试/Top N/zero/有界类别/Bubble membership草稿；7个真实Editor/Insert/Pivot及Pareto旧结果隔离、最后原始样本、真实低预算错误、关联保存/history/internal reopen草稿。 | 全部未执行；source hook的零cells读取断言只覆盖hook，不宣称全部UI没有常量规模header读取。新spec未接统一验证入口，没有性能/内存、原生Office、主题/焦点/IME或两后端验收。 |
| `chart-histogram.contract.test.mjs`、`chart-candlestick.contract.test.mjs` | P25共16个纯契约草稿：样本SD Scott、Count/Width/cap含边箱、常数/空/极端、有限无效作者/null/decoder/绑定/reset；严格三状态与Close源身份、全空HLC、独立alpha/宽度/validation、overlay存在/清除、reverse wick、Dot/None/log gap、native Doji及真实merge reset/legacy。 | Histogram7个、Candlestick9个，全部未执行；固定预期与option断言不能替代真实Rust/WASM、浏览器、Office或性能。 |
| `chart-candlestick-renderer.browser.test.mjs`、`tests/browser/chart-candlestick.spec.ts`、`tests/browser/chart-histogram.spec.ts` | P25共8个真实Canvas/工作台草稿：native三状态path及gap、HLC虚线/Dot/None、Histogram物化labels/counts/gap；状态选择/独立样式、HLC条件/源变化/恢复/空角色、Histogram数据与样式分组/单字段/物化点/Width尾箱/Auto/阈值/gap reset、Apply/history/internal reopen。 | Renderer3个、工作台3+2个，全部未执行；无dev server/浏览器/截图/像素、主题/焦点/IME、两后端、Office或性能结果。新spec未接统一验证入口，内部JSON不是原生Office往返。 |
| `chart-facade.contract.test.mjs` 新增 P12 场景 | 3 个标题字体/对齐/nullable/日期保存重开、4 个指示器恢复目标/独立 reset/历史、非法日期/标题/指示器候选拒绝且不发布历史的集成草稿。 | 未执行。 |
| `echarts-renderer.browser.test.mjs`、`tests/browser/chart-workbench.spec.ts` 新增 P12 场景 | 2 个 native 动作/graphic group replace-remove、sourceVersion/缓存/作者模型保留草稿；1 个真实 Overlay 在 150% zoom/scroll 的悬停、离开及 pointercancel 草稿。缓存用重新生成的 option，避免克隆 formatter 闭包。 | 浏览器草稿均未执行；不包含截图或像素验收。 |
| `sheet-chart-panels.contract.test.mjs` | 原6个pending插入/更新换unit/sheet、旧Pivot回调、释放、错误和焦点草稿；P24新增4个source及3个statistics的捕获宿主/候选/signal、关闭替换、旧完成和abort前后路由，预览getData若被调用即失败。 | 共13项草稿，未执行；mock宿主契约不能替代真实服务权限/版本、DOM/IME或浏览器验收。 |
| `pivot-chart-panel.contract.test.mjs` | 透视面板会话/状态与共享目录、源配置改动。 | 本轮修改未执行。 |
| `chart-statistics-compute.contract.test.mjs`、`chart-statistics-service.contract.test.mjs`、`crates/office-compute/tests/wasm_chart.rs` 及 chart 原生测试 | 根任务的协议/planner/decoder、分组身份、服务生命周期、预算/让出、数值与真实浏览器绑定草稿。 | 根任务开发中，未执行；不得由 TS 夹具推导 Rust/WASM 已通过。 |
| 既有 `chart-engine-rendering/echarts-option/echarts-renderer/chart-workbench`、Office writer 夹具 | 历史七类/有限字段切片以及已有生命周期用例。 | 旧通过结果只属于历史版本，不覆盖本批完整目录/配置/协议改动。 |

本轮唯一完成的核对是静态源码、接口、fixture 和 diff 阅读。**测试、类型、前端构建、Rust/WASM、浏览器、明暗主题、窄视口、焦点/IME、两渲染后端、真实 Office 重开、性能/内存全部未执行、未验收。** 统一验证阶段必须先生成与协议 2 匹配的 WASM 产物，再运行相应检查和真实样本/交互闭环；未验证项目持续留在本记录，不复用历史“通过”字样。

## P28：Chord 完整入口与语义追踪（实现中，未验收）

在替换旧 Chord 设置前，先清点固定版 `chart-ui/lib/es/index.js` 的 `ro/uo/lo/fo/Xo` 和公开 `FChordChartBuilder`。来源版本仍为 Univer 1.6.0、OSS commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；Pro 发布包只读归档为 `.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz`，下表的 `engine-chart ES` 指归档内 `package/lib/es/index.js`，不是不存在的解包路径。

| 固定入口与次序 | 布局、显隐、范围 | 状态与公开行为 | 可复查来源 |
| --- | --- | --- | --- |
| 数据：Source → Target → Value | 3 个 Select；两端类别字段，Value 含空字符串“计数”并排除两端；Source=Target 在 Target 显示诊断 | `chordMapping` → `Sae/wae`；所有角色索引整数、非负、范围内且互不重复；省略 Value 为计数 | chart-ui ES `ro`；engine-chart ES `Sae/wae`；`models/chart-data/internal/chord-data.d.ts` |
| 样式：Layout → Label → Ribbon | `Xo` 的三个组，非 Ribbon 顶部工具栏位置；节点、强调和其余 API 字段不在这个固定页面显示 | `chordStyle` 的五个语义组；不能将未显示字段视为不在交付范围 | chart-ui ES `Xo`；`chart-builder/configuration/chord.d.ts`；`types.d.ts:283` |
| Layout：Inner/Outer 同排 → Pad Angle → radii alert | 半径各 0..1、precision2、step.01；间隔非负；非法半径诊断是条件反馈，不计值入口 | `layout.innerRadiusRatio/outerRadiusRatio/padAngle`；最终需 `0 <= inner < outer <= 1` | chart-ui ES `uo/ao/co`；engine-chart ES `Mee/vm/a6/cOe` |
| Label：Visible → 条件 Position → Font | 关闭时只保留开关；开启时 inside/outside 和四个字体子值；字体复合控件的 color/fontSize 默认恢复各为一个动作；没有 align 控件 | `label.visible/position/color/fontSize/bold/italic`；公开另有 rotation/distance/family | chart-ui ES `lo/Da/Ca/Sa/wa/Ta/Ea`；`f-chord-chart-builder.d.ts` |
| Ribbon：Color/Opacity 同排 | source/target/gradient 三选一；opacity 0..1、step.05 | `ribbon.color/opacity`；公开另有 width/type | chart-ui ES `fo/oo`；engine-chart ES `Mee/cOe` |
| 公开全部组 | Layout6、Node5、Label9、Ribbon4、Emphasis1；Mapping3；`setChordMapping` 加五组 set/reset，共11操作 | builder/FChart → 同一模型更新与候选校验、文档独立内容历史、保存/重开；默认值由渲染解析，不写回节点或弧度 | `facade/builders/f-chord-chart-builder.d.ts`；engine-chart ES `vm/cOe` |

分类计数采用“独立作者值”与“动作”分开计数：固定数据3 + 样式11 = **14 个值**；标签开启时显示14，关闭时9（条件值5）。字体默认恢复另计2动作；没有额外 Chord 专属溢出或上下文入口被这五个页面函数证实。完整公开范围为 **28 个值**（Mapping3 + 五组25）；本地专属设置已落盘另外14个公开字段，其中 start/clockwise/min3、Node5、Label rotation/distance/family3、Ribbon width/type2、Emphasis1。本地关闭标签时隐藏8个条件值，剩20个值；五组恢复、全部恢复、标签字体恢复、标签颜色恢复合计 **8 个动作**，其中前6个为本地扩展，不计功能完成证明。该批新增常驻/插件根/溢出入口均为0，沿用既有45项目录及图表设置 Drawer；不将28个控件标为28项功能验收。通用图表入口、权限、选区和历史沿用前文完整清点。

固定 `Mee` 最终默认值：startAngle90、clockwise true、padAngle2、minAngle0、inner .58、outer .78；Node opacity1、borderColor `#FFFFFF`、borderWidth1、solid；Label visible true/outside/rotation0/distance8、gray.700/12/非粗体/非斜体/Arial；Ribbon gradient/.65/width0/solid；Emphasis adjacency；legendVisible false。锁定 ECharts 6.1.0 存在原生 Chord，不能宣称它不支持。其原生默认半径、间隔与弦带透明度和固定 converter 不同，不能代替 `Mee`；原生 `ChordPiece._updateLabel` 把旋转固定为0，需要本地实际适配公开 rotation。

已落盘 API 叶字段校验与语义默认解析。专属 `setChordLayout/setChordRibbon` 和设置草稿使用共同作者迁移：先合入既有作者字段，再执行最终完整候选校验；叶字段解析不靠默认外半径提前拒绝合法局部内半径值。通用 `FChart.setOptions` 仍先完整校验其输入，不由此推导任意 partial radii 可先通过再 merge。非法最终半径在本地原子候选中拒绝；固定 converter 对非法配置返回空 series 并由 UI 诊断，属于已记录差异。旧 flat layout/ribbonOpacity 和 generic label 内容保持读取；恢复及新组 authoring 清相应 legacy fallback，避免恢复后旧值重新生效。

固定 `Sae/wae` 聚合依据：最终显示文字拥有节点身份，不 trim；空字符串端点与显示文字相同的自连接跳过；映射权重只接受 finite number 且>0，字符串、零、负、非有限跳过；计数模式每条有效非自连接计1。节点和无向链接按首次有效行创建；反向重复按两节点序号 min/max 查重，保留首次端点方向；合并链接后再按链接首次次序合计节点，格式化发生在合计之后。本地对有限合计溢出显式报错，固定源码裸 `+=` 未作该保护。

主题颜色来源为固定 OSS `packages/themes/src/default.ts/dark-blue.ts`：10色组×10 + gray12 + loop12 + highlight16，共140个 token；本地保留作者 token，绘制边界转换 CSS，主题文字/背景覆盖语义灰色，未知 token 拒绝。自定义色主题、其他固定主题和真实明暗切换仍须浏览器验收，不以字典存在声明全部主题通过。

本地实际设置入口已静态追踪：[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 捕获 stationary hitbox click → [use-sheet-chart-panels.ts](../../packages/office-studio/src/features/charts/views/use-sheet-chart-panels.ts) 的 `openChartEditor` → 既有 Drawer 草稿 → `chart.update` → 候选校验、内容历史和保存。ChartOverlay 自身的 click 只选择图表，设置由父级捕获处理。[ChartChordMappingFields.vue](../../packages/office-studio/src/features/charts/views/ChartChordMappingFields.vue) 与 [ChartChordSettings.vue](../../packages/office-studio/src/features/charts/views/ChartChordSettings.vue) 采用 `antdv-next` Form/Form.Item、Select、InputNumber、Switch、ColorPicker、Button、Alert、Divider；没有新增原生表单控件。半径可先出现临时冲突再修复，非法 Apply 不发布模型；映射临时重复在草稿反馈，权重选择排除两个端点。`Divider` 使用安装版本的 `title-placement="start"`。

[chart-chord.ts](../../packages/office-studio/src/features/charts/chart-chord.ts) 的 `chartChordAuthoredLayout/chartChordAuthoredRibbon` 和 [chart-chord-api.ts](../../packages/office-studio/src/features/charts/chart-chord-api.ts) 的 `chartChordLayoutPatch/chartChordRibbonPatch` 同供 Builder、FChart 与设置控制器。nested 已定义值优先；nested undefined 与 legacy 同存时只迁相应 legacy 作者值，旧 Canvas startAngle 仅反号一次，不作者化 resolved defaults。group/null/font/color 恢复各清相应作者字段，保留其他组、主题 token 和映射。Builder/FChart 各有 mapping1 + 五组 set/reset10，共11操作；不是新增11个 Ribbon 根项。

[chart-chord-data.ts](../../packages/office-studio/src/features/charts/chart-chord-data.ts) 在主线程准备显示字典、格式与三列 bulk，生产异步投影通过 `ChartStatisticsService/OfficeComputeClient` 进入现有 `chartStatistics` 的 `chord` job。[chord.rs](../../crates/office-compute/src/chart/chord.rs) 合计无向链接和 incident node values，计算 normalized radians；输出 `values=[N*(value,start,end),E*(value,sourceStart,sourceEnd,targetStart,targetEnd)]`、`indices=[N,E,...labelId,...sourceNode/targetNode]`。身份为 `chord-node:${index}` / `chord-link:${index}`。字符串与像素几何留在 owner；兼容 graph 引用同一节点/链接数组。全局 state3 是合法的 spreadsheet error 状态，Chord 权重明确拒绝它，未知 >3 仍由 admission 拒绝；不将所有统计任务误改为只接受0/1/2。计数模式在准备时提供单位数值，省略 Value 不保留旧 weight mapping。大权重 data/layout 草稿已改为真实 Chord 数据，不以 Sankey graph 代替。

[chart-chord-compute.ts](../../packages/office-studio/src/features/charts/chart-chord-compute.ts) 的 portable DTO 参照与专属 decoder 校验字典身份、正有限总量、节点/链接计数、方向、端口及圆周闭合。[chart-chord-layout.ts](../../packages/office-studio/src/features/charts/chart-chord-layout.ts) 只把 retained angles 映射为像素；[chart-chord-draw.ts](../../packages/office-studio/src/features/charts/chart-chord-draw.ts) 生成共同 scene、节点环扇区、真实弦带 polygon、独立 fill/border alpha、source/target/gradient 颜色、内外标签和显式 rotation/distance/family，字号0不生成文字或标签命中区。ECharts 6.1.0 原生具备 Chord；本地选择 shared scene → graphic adapter，由 Rust/Worker 保留权重与角度的权威路径，并非因引擎缺少该类型。

[chart-chord-interaction.ts](../../packages/office-studio/src/features/charts/chart-chord-interaction.ts) 提供节点环扇区、弦带 filled polygon、真实度量标签 polygon 命中和 self/adjacency/none 聚焦；adjacency 为一跳无向邻接。renderer 只替换 content group，保留标题/绘图区，按原始 paint 恢复；重复 hit 缓存，equal paint/geometry 仍刷新合计与格式 Tooltip facts。标签 polygon 由 [echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts) 读取 ZRender Text 实际 bounds/transform 回填，未用字符数估算。本地无关 mark alpha 乘 `.15`，minAngle 使用比例借角；与 native 的 borrowing/blur 有差异，最终行为未验收。

[ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 使用受控 `antdv-next Tooltip` 展示节点/无向弦带合计，不写作者模型或内容历史，不启动 Chord content drag；仅有查看权限时仍可悬停，编辑由统一保护拒绝。安装的 `antdv-next 1.8.2` 的 Tooltip 语义键为 `container`，`@v-c/tooltip 1.1.5` 把 `role="tooltip"` 放在 `.ant-tooltip-container`；本轮修正旧 `styles.body` 为 `styles.container`，保持多行文字，root pointer-events none。全局 ConfigProvider/theme token 供给弹层。leave/Escape/blur/scroll/theme/zoom/字体变化/卸载或重入清理的实现和草稿已落盘；实际主题、焦点、IME、窄视口与两后端尚未执行。工作台 hit 坐标从 Facade 的工作表绝对 layout 减去 frame origin，再应用实际 hitbox 的缩放与滚动；不能直接把绝对 center 当成局部坐标。

[chart-compute-memory.ts](../../packages/office-studio/src/features/charts/chart-compute-memory.ts) 与 `chart-tasks.ts` 纳入 decoded facts、normalized arrays、节点/弦带、hit index 与 hover staging 的保守准入估算。这不是资源峰值测量或性能达标证据。新 Worker 夹具的两行 `[2,3]` 反向边预期合计5，11个f64和6个u32；预期入站3次/62字节、出站2次/112字节复制，无额外 Rust buffer clone，transfer caller buffer detached。数字只属于待执行契约，尚无实测；serde/transferable 不代表整条链路零拷贝。

| P28 草稿与样本 | 已声明的覆盖 | 执行与证据边界 |
| --- | --- | --- |
| [chart-chord.json](../../tests/fixtures/chart-chord.json)、`packages/office-studio/tests/fixtures/chart-chord-fixtures.mjs`、[chart-chord.contract.test.mjs](../../packages/office-studio/tests/chart-chord.contract.test.mjs) | 独立解析式角度/权重/ID、显示身份/count/format、溢出、extreme angles、typed ownership/decoder、viewport映射、真实几何/gradient/alpha/label0、hit/focus/cache、28值/8条件隐藏、legacy/null/独立恢复、映射与准入。 | 草稿全部未执行；portable TS 预期不证明 Rust/WASM、浏览器或资源验收。 |
| [chart-chord-facade.contract.test.mjs](../../packages/office-studio/tests/chart-chord-facade.contract.test.mjs) | 7个 Builder/live mapping、全部11操作、legacy partial、atomic invalid radii、独立reset、source更新/count、Apply/history/internal reopen、权限与失败前 Worker allocation 草稿。 | 未执行；内部快照不等于 Office 原生往返。 |
| [chord_tests.rs](../../crates/office-compute/src/chart/chord_tests.rs)、[wasm_chart.rs](../../crates/office-compute/tests/wasm_chart.rs) | 5个Rust纯内核草稿与新增真实绑定场景：独立 golden、state3/未知state、fractional/越界字典/overlap、溢出无部分结果、both-direction closure、result owner/copy/view/memory.grow/linear-memory输入拒绝。raw view在控制DTO和其他分配完成后建立并立即同步调用；真实alias错误码为 `WASM_INPUT_ALIAS`。 | Cargo、wasm32、wasm-bindgen-test全部未执行，新增WASM尚未生成。 |
| [chart-chord-renderer.browser.test.mjs](../../packages/office-studio/tests/chart-chord-renderer.browser.test.mjs) | 6个真实ECharts Canvas草稿：只替换内容/重复命中/恢复base、真实font bounds与rotation、none/count/noDrag、equal geometry更新合计与format但保留bitmap/labels、theme/size/DPR、retain/dispose。 | 未开server或浏览器；无截图/像素/峰值或时间测量。 |
| [chart-chord.spec.ts](../../tests/browser/chart-chord.spec.ts)、[chart-chord-settings.ts](../../tests/browser/helpers/chart-chord-settings.ts) | 6个真实工作台草稿：28值/条件/radii修复、端点/weight排除与count、独立font/color/layout/all恢复、node/link Tooltip与cache/source变化、载入other-user的viewable保护快照、theme/zoom/scroll/blur/Escape/卸载重入；Apply/history/internal reopen。 | 已加入`chart-workbench.playwright.config.ts`，全部未执行；保护草稿不调用会被拒绝的other-owner创建命令。共享reopen夹具已用finally释放副本，不据此声称资源测量完成。 |
| [chart-chord-worker.spec.ts](../../tests/browser/chart-chord-worker.spec.ts)、[chart-chord-worker.ts](../../tests/browser/chart-chord-worker.ts)、[chart-chord-worker.html](../../tests/browser/chart-chord-worker.html) | 6个真实native module Worker/生产compute.worker草稿：bulk transfer+owner复用、progress0 AbortSignal、一次cancel postMessage故障强制terminate/recreate、revision0→1旧promise与旧admission拒绝、Rust state3 weight错误后的恢复、active dispose。factory只转发真实事件并记录postMessage/terminate，不供给fake结果。 | 已加入`tests/browser/playwright.config.ts`，全部未执行，默认模块URL为`/compute/office_compute.js`，统一验证前先生成协议2新产物。故障注入只验证cancel channel失败分支，不声称测量了同步Chord内核的阻塞时长。 |

剩余验收包含固定版 native 与 shared graphic 的 minAngle/blur/标签效果差异、真实WASM版本及lease/取消、普通/大样本与同数据集前后指标、主题与焦点/IME、WebGPU合成/Canvas2D回退、源结构变化后的身份/格式、Office可编辑来源/配置保存与明确诊断、独立文件重开及Docs图表宿主。ECharts位图由工作表WebGPU合成不代表图内已用WebGPU绘制，内部JSON重开不代表XLSX/DOCX可编辑图表往返。本段不将单类型代码进展推导为45图型或完整工作台完成。

**没有执行测试、类型检查、构建、Cargo/WASM、解析器检查、dev server、浏览器或基准；未提交；全部保留未验收。**
