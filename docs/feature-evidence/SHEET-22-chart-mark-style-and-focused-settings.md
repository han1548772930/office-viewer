# SHEET-22 系列、饼图标记与精确面板定位（P21，未验收）

更新：2026-11-26。承接 [P19 通用元素直接编辑](SHEET-22-chart-element-direct-editing.md)、[P20 专属标记直接编辑](SHEET-22-chart-special-mark-direct-editing.md) 和 [P17 面板路由](SHEET-22-chart-edit-panel-routing.md)。本批补剩余 `seriesStyle / pieStyle / pieSliceStyle`，并让 Waterfall 角色、Pie aggregate / connector / label / source slice 的 More 定位真实字段。实现、测试草稿和证据已落盘；按照用户要求，测试、类型检查、构建、浏览器、Office 文件样本、WASM 和性能检查统一后置，全部保持 **未验收**。

P19 / P20 的历史缺口由本批下述局部实现接续。累计 7 个 mark discriminator 只能说明这一选择协议已包含固定版枚举，不能证明 45 个图型、全部图表功能或 Sheets / Docs / Pro 已完成。最终完整范围没有缩减。

## 固定来源与可观察契约

固定基线：**Univer v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**，只读根 `E:\rust_test\univer`。对应 Pro 公开 ES / 类型快照位于 `.research.local/univer-1.6.0/packages`。本批读取固定快照，未安装或依赖商业 SDK，未启动上游演示或 Microsoft Office。

| 来源 | 核对内容与本地差异 |
| --- | --- |
| [float-menu/types.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/types.d.ts)，约 5–45 行 | 固定 7 个 mark kind 顺序为 series / Pie labels / source slice / aggregate / connector / Pareto / Waterfall。Series 由系列 ID / 索引描述，Pie slice 有 `sourcePointIndex`；Waterfall 排除 base。本批新增前三个本地 discriminator，复用 P19 / P20 浮动组件 |
| [chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 的 `Hl / Wl / Gl / Kl` | `Hl` 的系列 mark 包括 Combination、Radar、Boxplot、Line、Column、Area、Scatter；`Wl` 从 context 的来源映射定位作者，不能以 native series index 推断字段。`Gl` 区分复合来源扇区、合计、连接线和其余 Pie 标签；`Kl` 的 Relation / WordCloud 仍只走完整面板。本地 bar / bubble 沿用既有二维系列模型，属于本地图型适配，未运行上游交互对照 |
| 同一 ES 的 `Eu / Du`，及 [pie-mark-patch.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/pie-mark-patch.d.ts) | `Eu` 分全局 Pie label、`sliceStyles[sourcePointIndex]`、aggregate 与 connector 作者；`Du` 分系列、Pareto 和 Waterfall 作者。本地用纯 scoped planner 投影同一作者组，借用既有 `FChart.update` 和历史，未复刻上游内部服务 |
| [chart-types.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/chart-builder/chart-types.d.ts)，约 353–356、396–432 行 | Slice spec 有 color / fillOpacity / explosion；Pie 有稀疏 `sliceStyles`，复合图有独立 aggregate / connector。本地增加 `options.pie.sliceStyles`，继续使用原 `aggregateStyle / aggregateLabel / secondaryPlot`。独立扇区额外复用本地纹理和边框，记录为字段扩展 |
| [chart-edit-panel-route.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/routing/chart-edit-panel-route.d.ts)，约 4–10 行；同一 float-menu types 的 section 枚举 | 上游 route 仍是 `tab / section / axisId / seriesId / titleKey` 5 字段、27 个 section。本地保持这些值，增加可选第 6 字段 `focus`；这是本地瞬时导航扩展，不能称为固定版原有字段 |
| ECharts **6.1.0** / ZRender **6.1.0** 的已核对 native model、view、display list、transform / clip / contain、attached text / guide line 接口 | 继续由真实绘制元素生成命中与选区。native raw item index 先经过来源映射；graphic 命令附同一作者 metadata；没有估算 HTML hitbox |

压缩 ES 只按函数片段读取。公开类型、按钮名称和静态接线均不作为浏览器交互、像素或 Office 文件保真的完成证明。Boxplot 原始样本 / IQR / quartile 模型是本项目既有 Office 扩展，本批消费其统计结果与 `outlierSourceIndices`，没有将它改称固定上游内部统计实现。

## 入口分类与作者闭环

本批新增 **3 个 direct-edit kind**。按作者分类是 **4 类**：Series 的 numeric / exact-name 两类、Pie 全局 label 一类、Pie 来源 index 一类。Pie 的 primary / secondary 是几何上下文，仍写同一个来源 index；Boxplot summary / outlier 的增量单列如下，不另增加 kind。

本批入口根增量：**常驻 Ribbon 0、插件扩展 Ribbon 0、溢出 Ribbon 0、对象右键 0、浮动组件 0、浮动动作根 0**。复用原标题 / 关闭、Ant Form、作用域说明、恢复格式 / 选择系列 / More / Apply 和 Error Alert；“选择系列”只在原点 / 标签上下文显示，目标现在可提升为专属 seriesStyle。Shift 对通用点继续得到原 generic series。三个新 mark 使用同一权限、busy、composition、revision、宿主生命周期和焦点规则。

完整 Pie 设置区新增 **1 个组件** `ChartPieSliceSettings`。独立样式定义 **11 个控件**：color 1、纹理 3、opacity 1、explosion 1、border 5；副图 bar 段不提供 explosion，定义数为 **10**。来源选择与恢复继承样式各复用一个 Ant 控件；长来源列表使用 InputNumber 替换 Select。纹理颜色按已有样式显隐，定义数不代表同时可见入口数。其他两类 mark 复用现有完整系列 / Pie 标签 controls。

| 真实入口 / 作用域 | 作者与行为 | 显隐 / 命令 / 保存关系 |
| --- | --- | --- |
| 二维 Cartesian 线 / 面片段、Radar 多边形，以及真实点 / 标签上的“选择系列” | `seriesStyle.binding` 保存实际 field、sourcePosition、sourceName 和 numeric / exact-name selector。仅提交 color / fillOpacity / fillPattern / border / point / label | 支持 line / column / bar / area / scatter / bubble / combination / radar / boxplot；三维不走此专属 mark。不改 name / type / rightAxis、独立 dataPoints、waterfallStyles、其他来源和源数据 |
| Pie 实际文本与标签引导线 | `pieStyle` 写 `options.pie.label`，作用于全部扇区数据标签 | reset 只清这一全局 label 层；Delete 只令 label.visible=false。复合合计标签继续保留 P20 aggregate 的独立标签作用域 |
| 复合 Pie 主图来源扇区、副图扇区 / bar 段 | `pieSliceStyle` 写 `options.pie.sliceStyles[sourcePointIndex]`，有实际 sourceField / sourceName / plot | 原 `series.dataPoints` 是读取 fallback，新写入保留它及兄弟 slice / aggregate / connector。来源 index 对应当前准备数据 categories 中的项目；生成合计不会取得该身份 |
| 普通 Pie 已有独立 sliceStyles 的扇区 | native per-source map 与 graphic metadata 都选择同一独立作者层 | 没有独立层的普通扇区继续使用 P19 dataPoint。避免全表单已写独立层后，点击仍写被覆盖的旧点层而无效果；reset 留下的显式清除结构仍可再次编辑 |
| Boxplot 的真实箱体 / 须线 / 中位线 / summary label | 同一来源 `seriesStyle`，从 derived.boxes 的 field / sourceName / seriesPosition 建绑定 | 原样本 0 不会变成 summary 作者；named 层是最高作者时写该层。独立 sample dataPoints 保留 |
| Boxplot 实际异常点 | 沿用 `dataPoint`，native / graphic 都用统计的原始 `outlierSourceIndices` | null 槽、排序、重复值、log 过滤、隐藏 / 零大小点均按现有绘制顺序映射。size / shape 控件真实可用；showOutliers=false 隐藏无效 marker 控件 |

所有局部提交仍由 **`FChart.update` → `sheet.command.update-chart`** 执行；More 使用 **`sheet.command.chart.open-focused-panel`**；正文历史仍按原 workbook / 文档隔离。snapshot / save / reopen 沿用原模型链路，选择、bounds、runtime ID、`focus` 不写入内容。真实 undo / redo / save-reopen 的草稿已编写，尚未执行，不能当作文件验收。

## 精确 More 与 scoped reset

`focus` 是受校验的 plain DTO，仅接受下列 5 种 discriminator，拒绝未知字段、symbol、非 plain 对象、任意 selector 和不一致的 tab / section：

| focus | 计划与真实字段 |
| --- | --- |
| `{ kind: 'waterfallRole', role: 'positive' \| 'negative' \| 'subtotal' }` | 切换实际来源与角色后聚焦 `${series / namedSeries prefix}.waterfallStyles.${role}.color`，或全局 `waterfall.styles.${role}.color` |
| `{ kind: 'pieAggregate' }` | `pie.aggregateStyle.color` |
| `{ kind: 'pieConnector' }` | `pie.secondaryPlot.connector.visible` |
| `{ kind: 'pieLabel' }` | `pie.label.visible` |
| `{ kind: 'pieSlice', sourcePointIndex }` | 选择真实来源项目后聚焦 `pie.sliceStyles.${sourcePointIndex}.color` |

全局 stacked subtotal 没有 `seriesId`。关闭自动合计或指向不存在的来源 subtotal 会拒绝；Pie 来源必须准备就绪且 index 在 categories 内，独立 Pie focus 不适用于三维；aggregate / connector 仅用于复合 Pie。

`ChartSettingsFields.navigationControl` 先核对 Waterfall role / source 与 Pie source index，再等待真实 option key。精确 focus 尚未就绪时保持等待；原 section fallback 只服务未指定精确 focus 的路由。projection 就绪后可重新调度焦点；切 tab / owner、busy、composition 或竞争对话框仍按原生命周期约束取消。

新 mark 的 patch / reset 均只投影它负责的作者字段。独立 bar 段的 explosion 保留为 inactive 数据，不纳入 signature、patch、reset 或最大半径；重新成为主图扇区时可继续编辑。实际 sector center 使用该来源的 explosion，合计继续使用独立 explosion。

完整 Pie source reset 复用同一 planner，保留显式 undefined 清除字段，经过 `chartSettingsPatch` 和 deep merge 后可清旧格式。完整 Waterfall 全局角色 / 项目 reset 同样写显式 undefined，避免删除 draft key 后旧作者被 merge 保留。对应真实 UI 草稿已增加。

## 模块与接口变化

归属为 `features/charts` 的纯模型 / controls / scoped planner、主线程 native / graphic 几何适配，以及 Vue / `antdv-next` 视图。Boxplot 重统计继续使用项目已有计算服务与结果，本批没有增加统计权威引擎。没有修改 Rust crate、Worker 协议、WASM lease / buffer 所有权，也没有执行或声明性能 / 零拷贝优化。

- `ChartCanvasMarkDirectEditTarget` 增加上述 3 个分支。seriesStyle 使用已有 binding；pieSliceStyle 含 sourceField / sourceName / sourcePointIndex / 可选 plot。
- `ChartCanvasNativeSeriesBinding` 增加 per-item `items` 分支，Boxplot summary / outliers 用稳定 runtime ID；原来源分支增加 seriesMark / pieLabels / pieSliceMarks。raw input index 先映射到语义 / 样本 index，校验后才成为作者目标。
- `chartCanvasNativeElementRuntime` 增加可选第 5 参数 `derived`，两处生产调用已传入；Boxplot mapping 使用同一真实统计结果。
- 模型增加 `ChartPieSliceStyleOptions` 与 `options.pie.sliceStyles`；原 aggregate type 复用其字段。validation / deep clone / gradient authoring / effective style / native & graphic layout 按现有链路接线。
- `chartCompoundPieData` 入参放宽为 `Pick<ChartModel, 'options'>`，仍读取原 options，没有重写分区算法。
- 没有新增 Facade 方法、内容 command ID、历史源、Office 转换服务或商业依赖。`focus` 是上述本地导航差异。

UI 静态核对：新增视图的 FormItem、Select、InputNumber、Button、Alert 全来自 `antdv-next`；修改的 Waterfall / Pie 设置与浮动控件继续经过原统一 ConfigProvider / theme token。Canvas 命中、选区和拖拽仍为编辑器内核，没有新原生输入桥。明暗主题、弹层、实际焦点恢复和真实输入法仍待浏览器验证。

## 准确文件清单

本批 **28 个文件：19 源码、8 测试 / fixture、1 独立证据**。源码仅新增 `ChartPieSliceSettings.vue`；其余为既有或前批模块上的本批增量。共享目录其他在途改动不计入本批；中央清单由主任务同步。

| 源码文件 | 本批修改 |
| --- | --- |
| [domain/model.ts](../../packages/office-studio/src/domain/model.ts) | 独立 Pie slice 作者类型 / 稀疏 map |
| [chart-edit-panel-route.ts](../../packages/office-studio/src/features/charts/chart-edit-panel-route.ts) | 本地 focus DTO 校验 / 真实 option key / source / role plan |
| [chart-pie.ts](../../packages/office-studio/src/features/charts/chart-pie.ts) | Slice 有效层 / legacy fallback / 独立 explosion / inactive bar 判断 |
| [chart-pie-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-pie-settings-controls.ts) | 浮动与完整面板共用 11 / 10 个 slice 控件 |
| [chart-option-validation.ts](../../packages/office-studio/src/features/charts/chart-option-validation.ts) | 稀疏 index 与 slice 样式校验 |
| [chart-gradient.ts](../../packages/office-studio/src/features/charts/chart-gradient.ts) | 独立 slice 渐变作者校验 |
| [chart-pie-layout.ts](../../packages/office-studio/src/features/charts/chart-pie-layout.ts) | 真实来源 explosion 中心与半径 |
| [chart-derived.ts](../../packages/office-studio/src/features/charts/chart-derived.ts) | 分区函数仅放宽 options 入参 |
| [chart-element-interaction.ts](../../packages/office-studio/src/features/charts/chart-element-interaction.ts) | 3 kind / key / metadata / native series & item ownership / More route |
| [chart-special-mark-edit.ts](../../packages/office-studio/src/features/charts/chart-special-mark-edit.ts) | 3 类 controls / signature / scoped patch / reset / Delete / capability |
| [echarts-element-interaction.ts](../../packages/office-studio/src/features/charts/echarts-element-interaction.ts) | 实际 native root / datum / label / guide / item 映射与 painted bounds aliases |
| [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) | Slice 生效样式与 graphic fallback；runtime 的 derived 参数 |
| [echarts-special-option.ts](../../packages/office-studio/src/features/charts/echarts-special-option.ts) | Boxplot 两个稳定 runtime ID |
| [chart-draw.ts](../../packages/office-studio/src/features/charts/chart-draw.ts) | Cartesian / Radar / Boxplot / Pie 实际几何 metadata 与独立 slice 生效样式 |
| [chart-element-edit.ts](../../packages/office-studio/src/features/charts/chart-element-edit.ts) | Boxplot 通用点 / 系列 marker 控件可见性 |
| [ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) | focus 等待 / source-role 核对与完整 slice 面板装配 |
| [ChartWaterfallSettings.vue](../../packages/office-studio/src/features/charts/views/ChartWaterfallSettings.vue) | 精确角色选择；完整 reset 的显式清除 |
| [use-chart-element-editing.ts](../../packages/office-studio/src/features/charts/views/use-chart-element-editing.ts) | 原 promote 动作指向真实 dedicated series target |
| [ChartPieSliceSettings.vue](../../packages/office-studio/src/features/charts/views/ChartPieSliceSettings.vue) | 新增 Ant 来源选择 / 样式 / scoped reset 视图 |

| 测试 / fixture 文件 | 本批修改 / 声明计数 |
| --- | --- |
| [chart-mark-style-edit.contract.test.mjs](../../packages/office-studio/tests/chart-mark-style-edit.contract.test.mjs) | **新增 12**：numeric / name 系列隔离与 reset；Boxplot 原样本 point；Pie 全标签 / Delete；独立 slice / legacy / sibling / aggregate；bar inactive；full settings merge；DTO / clone / gradient；类型 / no-op / metadata；focus 精确计划 |
| [chart-mark-style-rendering.contract.test.mjs](../../packages/office-studio/tests/chart-mark-style-rendering.contract.test.mjs) | **新增 10**：真实 SVG SSR / ZRender native & graphic Cartesian、Radar、Boxplot summary / raw outliers、Pie labels / guides、普通独立 slice 作者、复合 primary / secondary / aggregate、真实 explosion / inactive bar、null / duplicate / log / hidden 过滤 |
| [chart-mark-style-direct-editing.browser.test.mjs](../../packages/office-studio/tests/chart-mark-style-direct-editing.browser.test.mjs) | **新增 8**：完整 ExcelWorkbench / Ant / Canvas2D 的 source promotion、named author、Pie label / Delete、slice 重分区、精确 More / 完整 reset、Boxplot summary / outlier、历史 / save-reopen、只读 / composition / 晚完成与切图型 / 切表 |
| [chart-element-editing.contract.test.mjs](../../packages/office-studio/tests/chart-element-editing.contract.test.mjs) | 更新 promote 断言；原 **7** 个声明不变 |
| [chart-element-direct-editing.browser.test.mjs](../../packages/office-studio/tests/chart-element-direct-editing.browser.test.mjs) | 普通 Pie 标签 selector 改为真实 pieStyle；原 **8** 个声明不变 |
| [chart-special-mark-rendering.contract.test.mjs](../../packages/office-studio/tests/chart-special-mark-rendering.contract.test.mjs) | 复合来源改为 slice mark；More 增加精确 focus；原 **8** 个声明不变 |
| [chart-special-mark-direct-editing.browser.test.mjs](../../packages/office-studio/tests/chart-special-mark-direct-editing.browser.test.mjs) | named Waterfall role / color 与 connector visible 的真实焦点断言；原 **7** 个声明不变 |
| [chart-object-actions.html](../../packages/office-studio/tests/fixtures/chart-object-actions.html) | selector 支持 sourcePointIndex / sourceField / sourceName / plot，继续询问真实绘制命中 |

按 `test(...)` 声明手工计数：**本批新增 30，全未运行**；上述 4 份旧测试合计 30 个声明不重复计入新增，涉及的 7 个测试模块合计 60 个声明。循环中的图型 / 参数不另计。P19 原 34、P20 新增 25 的阶段计数不变。本证据为第 28 个文件。

## 验证与剩余缺口

本批仅执行搜索、阅读、编辑及手工静态核对；未运行测试、类型检查、构建、Vite / dev、浏览器、Cargo / WASM、基准、协作服务或 `git diff --check`。源路径 / 固定版本、27 sections / 本地 focus 差异、文件分类和草稿声明数已手工核对；这不是运行验收。

仍需后置验证：

- Canvas2D / WebGPU 真实命中、样式效果、Drawer 字段焦点、主题弹层、实际 IME、重复样本遮挡、滚动 / 缩放 / 字体与宿主切换。
- 真实 undo / redo、内部 snapshot save / reopen、API source ownership，以及 Office / OOXML 解析写出对新增 sliceStyles 的映射与往返保真。未运行外部 Office 文件样本，不将内部 snapshot 草稿称为文件验收。
- 其余 45 图型的完整元素功能、Relation / WordCloud 等既有完整面板与全部 Pro 功能闭环；7 kind、一个表单或控件定义数不能替代逐入口的模型 / 命令 / 权限 / 持久化验收。
- 首屏、滚动、编辑、缩放、大样本、Worker / WASM 传输、峰值内存与取消的可复现测量。本批没有性能前后对照，不声明优化或达标。

上述项目全部保留 **未验收**，由后续统一检查与完整功能清单继续追踪。
