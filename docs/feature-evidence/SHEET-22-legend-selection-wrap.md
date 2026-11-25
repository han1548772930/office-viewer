# SHEET-22 / SHEET-24：通用图例选择、换行与位图宿主交互

2026-11-25。本轮是图表局部前端实现，全部保持 **未验收**。仅静态阅读源码、补丁和 `git diff --check`；没有运行测试、类型检查、构建、开发服务、浏览器、Rust/WASM 或性能基准。45 项图型目录和下述测试草稿不代表完整图表能力已经完成。

## 固定来源与入口

- 固定 Univer 源为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。商业包只作本地静态参考，没有安装商业 SDK 或调用其服务。
- [ChartLegendSection 声明](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/sections/ChartLegendSection.d.ts) 包含 position、wrap、selectMode、visualMapType 和字体字段。[chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 的 `Gs` / `Xs` / `Qs` 记录了相关显隐条件：wrap 只用于上/下位置，Heatmap 不显示通用换行/选择，组合饼图不显示选择。
- [engine-chart 归档](../../.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz) 中 `package/lib/es/index.js` 默认 `legend.selectMode = 'close'`；`xAe = {single:'single', multiple:'multiple', close:false}` 转换到渲染器选择模式。`R8` 对水平 wrap 使用 plain，其余使用 scroll；WordCloud（类型 512）直接跳过通用图例。
- [公开配置类型](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/chart-builder/chart-types.d.ts) 的 `ChartLegendSpec = boolean | IChartLegendOptions`，字体 color/fontSize/bold/italic/family 为顶层字段。[Builder](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-chart-builder-base.d.ts) 与 [FChart](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-chart.d.ts) 声明 setLegend/clearLegend；[Facade Enum](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-enum.d.ts) 使用 `ChartSelectModeEnum`、`ChartLegendPositionEnum`，位置包含 Hide。
- 本轮在现有图表上下文设置内新增 **2 个作者配置行**：图例换行、图例选择模式，各有独立恢复动作。选择模式有 single/multiple/close 三个值；值数量不计成功能数量。静态行序为现有显隐、位置、换行、选择模式；Heatmap 原有 visualMapType 行继续保留。未新增常驻 Ribbon、扩展或溢出根入口，完整顶部工作区清点仍以[图表审计](SHEET-22-chart-type-audit.md)为准；固定版完整行列布局和最终视觉顺序未验收。
- Heatmap 已有 continuous/piecewise selector 及其 API。本轮复查上游 visualMap 为 `selectedMode:false, calculable:false, realtime:false`，没有添加虚构的色阶点击过滤。
- 本地 ECharts 锁定 6.1.0（package.json 与 pnpm-lock）。Context7 的 `/apache/echarts-doc` 用于 selectedMode、legendSelect/legendToggleSelect、legendScroll、plain/scroll 和组件注册参考；最终查询指定 6.1.0，但返回 master 文档。较早 6.0.0 查询已纠正，文档查询不是本地运行验证。

## 模块归属、作者配置与视图状态

- [model.ts](../../packages/office-studio/src/domain/model.ts) 只增加 `ChartLegendSelectMode` 和可选 `legend.selectMode` / `legend.wrap`。现有 `legend.font` 继续是规范化持久结构；没有保存 renderer 对象、函数或选择名称。
- [chart-legend.ts](../../packages/office-studio/src/features/charts/chart-legend.ts) 集中类型能力、配置校验、字体归一化、图例项目、几何和派生视图过滤。公开 setLegend 接受 boolean、对象、Hide 和扁平字体字段，也保留本地 nested font。nested font 替换本地字体对象；同时给出扁平字段时，扁平字段优先。字段校验在候选替换前完成，字体解除调用者引用，顶层字体不会由该 API 写入持久 legend。
- [chart-engine.ts](../../packages/office-studio/src/features/charts/chart-engine.ts) 复用上述校验，并去除恢复操作产生的 undefined 图例字段。[facade.ts](../../packages/office-studio/src/api/facade.ts) 的 Builder/FChart setLegend 和 clearLegend 使用同一配置路径；Builder clear 删除草稿配置，live clear 明确清除字体、选择模式和换行，再恢复该图型默认显隐/位置，避免被旧配置的合并逻辑保留。Hide 规范化为 visible=false，保留原有位置。[index.ts](../../packages/office-studio/src/index.ts) 导出 `ChartLegendPositions`、`ChartLegendSelectModes`、`ChartLegendSpec`，Facade 以固定上游 Enum 名暴露这些值。
- 作者配置仍经 detached draft / `sheet.command.add-chart`、`sheet.command.update-chart` → OfficeStudio 内容事务，沿用权限、按文档历史和保存模型。本轮没有另建图例内容命令或权威计算引擎。
- [echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts) 单独拥有 selected names 与 scrollDataIndex。图例点击、分页是 view operation，不产生内容 mutation、撤销记录、文档 revision 或保存字段；只读可见图表允许这些查看操作。新名字默认可见，旧名字按同名保留，删除的名字丢弃；single 始终选中恰好一个有效名称，multiple 可以全部关闭，close 禁止选择并保持分页可用。
- 同名有效选择/当前分页保留于缓存重绘、尺寸、DPR、主题和来源刷新；改变模式重建该模式的默认选择。关闭并重开宿主会重新建立初始视图选择，仅作者配置保存。没有声称图例 view state 已进入协作、远端广播或内容历史。
- 功能归属为 charts 的模型适配/纯几何与主线程位图视图交互；重建只消费既有 derived/prepared statistics，不启动 Worker/WASM 求解。`chart-derived.ts` 的 optional legendSelection 只用于派生视图，未加入作者 DTO 或 Worker 请求。布局仍使用既有纯计算/Worker 边界，本轮未改 WASM 平台协议或线性内存所有权。

## 实际绘制与输入接线

- [chart-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-settings-controls.ts) 以既有 ChartSettingsFields 的 `antdv-next` Form.Item / Select / Switch / Button 接入新增行。显隐与恢复按类型/位置判断；没有新增原生可替换输入控件。继承统一 ConfigProvider/token，主题、焦点恢复与 IME 仍需浏览器验收。
- [chart-layout.ts](../../packages/office-studio/src/features/charts/chart-layout.ts) 与 [chart-draw.ts](../../packages/office-studio/src/features/charts/chart-draw.ts) 共用图例项目与固定槽位几何；`chartPlotFrame` 新增可选 legendCount 参数。[chart-relation-layout.ts](../../packages/office-studio/src/features/charts/chart-relation-layout.ts) 提供真实分类数，使 Worker 预排布和显示图例的预留区域一致。
- [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) 统一 native legend；共享 DrawCommand 路径用空 native pie series 提供图例名称和 action 契约，图表实际内容继续由共享 Canvas 场景绘制。重建过滤真实 marks/指针可命中对象，不用提示消息模拟功能。原生日期轴和 Radar 的指示线复用初始绘制范围、类别位置和雷达范围，只替换视图数据，防止过滤后另算范围而错位。
- 柱/线/散点等按系列过滤，饼/漏斗/仪表按分类过滤。普通柱图剔除无 points 的系列槽位；Boxplot 的样本分布和异常点可分别过滤。Waterfall 过滤实际角色 marks/相邻 connectors，保留已有累计结果、原类别和布局边界。Relation 过滤分类节点/连带边，未分类节点继续可见并保留节点索引/坐标。Sankey 过滤节点/连带边、隐藏命中带并保留拓扑索引；[chart-sankey-interaction.ts](../../packages/office-studio/src/features/charts/chart-sankey-interaction.ts) 的 optional baseScene 保存未过滤的拖动位置。空 graph scene 继续保留内容 group 和恢复所需 view metadata。
- TreeMap/Sunburst/Chord/WordCloud 不绘通用图例。Heatmap 保留自身色阶，不接通用 selection；组合饼图隐藏 selection。没有把这些固定类型条件当成未实现能力的永久禁用替代。
- [echarts-legend.ts](../../packages/office-studio/src/features/charts/echarts-legend.ts) 隔离真实 Canvas 命中适配，沿 `findHover().topTarget` 的祖先读取锁定 ECharts 6.1 的 `__legendDataIndex` 和 `pagePrev/pageNext` 名称。它们是 **内部实现依赖**，不是公开 API；升级依赖必须复核。原生 data 名称序列保留重复/空项以维持索引，selection 按有效名称去重；同名项联动。
- [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 先命中图例，再判断对象拖动。按下/释放须命中同一项目且屏幕移动不超过 4 px，才能执行；图例交互不 emit update-chart。native action/setOption/refresh 失败时保留旧选择、分页和指针场景，失效缓存并重绘旧视图，同时反馈错误。恢复时保留同一作者帧内的 Sankey 拖动位置与 Relation viewport。

## 本地差异与未核实项

- 图例文字使用有界固定槽位和 truncate，使纯 Worker 布局不依赖 DOM 字体测量。水平 wrap 超过约 40% 画布高度预算时退回 scroll，保证入口可分页到达；侧向图例使用 scroll。它与上游按 native 文本尺寸排布有差异，超长文本、字体 fallback、极小图表、分页按钮空间和实际换行效果未验收。
- 百分比堆积过滤目前保留 prepared 百分比，没有确认上游隐藏系列后是否重新归一化至 100%。3D 的真实 marks 会过滤，但各变体系列槽位/投影布局尚未逐图核验。
- Waterfall/Sankey/Relation 保留原始累计或布局是当前 view 契约；上游选择后重排/累计语义、转图型前后行为和全部 45 图型闭环仍未核实，不能标为全部通过。
- 内部图例命中适配、分页 clipping、空/重复名称、single 初始过滤、原生与共享图表外观一致性、Relation 标签命中、故障恢复、权限可见性与跨宿主 Docs 图表接线均缺运行证据。
- 没有执行任何性能对照。view 过滤和重绘仍分配 JS 数组与新命令；没有声称零拷贝、主线程长任务消除、首屏/帧时间或大样本内存达标。

## 测试草稿与验证状态

- [chart-legend.contract.test.mjs](../../packages/office-studio/tests/chart-legend.contract.test.mjs)：10 个草稿，覆盖配置原子校验/字体解除引用/扁平归一化/clear、名称选择、类型显隐、wrap 预留、native hit 索引、日期/Radar 指示线范围、真实共享 marks、分类/Candle、graph 过滤和 Waterfall 累计保留。
- [chart-legend-facade.contract.test.mjs](../../packages/office-studio/tests/chart-legend-facade.contract.test.mjs)：2 个草稿，覆盖上游 Enum 名、Builder boolean/object/Hide/flat font/clear、live 内容历史与内部 JSON 保存重开、view state 不持久化。后续需要正常 OfficeStudio/WASM/platform 环境，未执行。
- [chart-legend.browser.test.mjs](../../packages/office-studio/tests/chart-legend.browser.test.mjs) 与 [fixture](../../packages/office-studio/tests/fixtures/chart-legend.html)：4 个草稿，计划覆盖真实 ECharts Canvas hit、native/shared filtering、分页/缓存、只读 Overlay 的缩放/滚动点击与拖动取消、主题/DPR 后新位图尺寸、部分 action 故障后的旧视图重绘。fixture 明确将图表放入工作簿，并供应 preparedData 隔离 WASM 就绪；它不是完整 Workbench、WebGPU 或真实 Worker 求解验收。
- 已完成源码/diff 静态审阅、相关 tracked 改动的 `git diff --check` 和本记录 26 个本地链接的存在性检查，未报告空白错误或缺失链接；新增文件另外静态阅读。所有草稿、类型/前端/Rust/WASM、浏览器/焦点/IME、XLSX/DOCX/native Office 文件样本和大样本基准均未运行，继续标记 **未验收**。
- Sheets、Docs、Pro 完整工作台目标持续有效；本轮局部前端补充和后端后置均不构成范围缩减。
