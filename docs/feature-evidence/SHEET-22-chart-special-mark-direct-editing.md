# SHEET-22 专属图表标记直接编辑（P20，未验收）

更新：2026-11-26。承接 [P19 通用图表元素直接编辑](SHEET-22-chart-element-direct-editing.md)，本批补 Pareto 柱形／累计线、Waterfall 来源角色／全局自动合计、复合 Pie 合计扇区／主副图连接线。代码和测试草稿已经落盘；测试、类型检查、构建、浏览器、文件样本和性能检查按用户要求统一后置，状态均为 **未验收**。

P19 的专属 mark 缺口是该历史批次的状态，本批只更新下述 4 个 direct-edit kind。真源 Pie 扇区／标签、通用系列／点、标题、绘图区、对象菜单和 P17 面板路由沿用既有实现，不重复计入本批增量。所有 Sheets／Docs／Pro 的最终范围保持完整。

## 固定来源与实现差异

固定基线：**Univer v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**，只读根 `E:\rust_test\univer`。对应 Pro 公开 ES／类型快照在 `.research.local/univer-1.6.0/packages`；本批读取快照作为对照，不安装或依赖商业 SDK。

| 来源 | 本批核对的契约与差异 |
| --- | --- |
| [float-menu/types.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/types.d.ts) | 7 个 mark direct-edit kind 包含 `pieAggregateStyle`、`pieConnectorStyle`、`paretoStyle`、`waterfallStyle`；Pareto part 为 bar／line；Waterfall 不提供 base 角色编辑。本地新增这 4 个 kind，复用 P19 浮动工具 |
| [pie-mark-patch.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/pie-mark-patch.d.ts) | 独立 Pie direct color／fillOpacity patch 和 connector patch，合计与真源扇区不是同一作者目标 |
| [chart-types.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/chart-builder/chart-types.d.ts)，约 352–356、396–414 行 | Slice style 有 explosion／color／fillOpacity；composite 含 secondaryPlot.connector、aggregate.label／style。本地复用既有 `options.pie.secondaryPlot`、`aggregateLabel`，增加独立 `aggregateStyle`，不改写为上游内部 composite schema |
| [chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 的 `hs`、`ql`、`Eu`、`Du` | hit → 配置目标／特殊 mark 选择；`Eu` 分 Pie slice／aggregate 与连接线；`Du` 分 Pareto part、Waterfall source-role patch。压缩源码按函数小片段阅读，不能用按钮名或类型声明证明全部运行态行为 |
| [P17 面板证据](SHEET-22-chart-edit-panel-routing.md)、[P19 元素证据](SHEET-22-chart-element-direct-editing.md) | 沿用真实 `FChart.update`、`sheet.command.update-chart`、`sheet.command.chart.open-focused-panel`、Drawer、选区和内容历史；本批不新增 Facade／index 或另一套作者历史 |
| ECharts **6.1.0**：`packages/office-studio/node_modules/echarts`；ZRender **6.1.0**：`node_modules/.pnpm/zrender@6.1.0/node_modules/zrender` | 复用实际 native view、display list、attached text、transform／clip／contain 和 z-order。Graphic fallback 给实际绘制命令附同一身份，不用 HTML 估算 hitbox |

本批读取了固定版公开源码和类型，未运行上游演示或 Microsoft Office。Office 元素格式与复合图表的公开可观察行为作为产品语义参照；上述接线不能表述为已经核实全部上游交互、像素或 Office 文件保真。合计扇区额外复用本地边框／填充纹理能力，与上游 Slice style 的公开字段存在差异，须保留实测验收。

## 入口分类与局部写入

按作者 scope 分类，本批新增 **4 个 direct-edit kind、8 类 scope**：Pareto 2，Waterfall 来源角色 3＋全局 subtotal 1，复合 Pie 2。scope 是运行时选择类别，不是按钮总数，也不代表所有图型完成。

入口增量：**常驻 Ribbon 根 0、插件扩展 Ribbon 根 0、Ribbon 溢出根 0、对象右键根 0、浮动工具组件 0、浮动动作根 0**。复用 P19 的标题／关闭 → Ant Form → 作用域说明 → 恢复格式／More／Apply → Error Alert 顺序和 4 个已有动作；专属 mark 不显示“选择系列”。classic／grid 工作台都经同一 Canvas 上下文入口。

| scope／真实入口 | 控件与作者写入 | 显隐、权限与命令 |
| --- | --- | --- |
| Pareto bar：实际柱形及 attached 数据标签 | 复用 Pareto bar controls，写 `options.pareto.barStyle` 的 color／opacity／border／label | 不写源系列颜色或专属 point map；Apply → `FChart.update` → `sheet.command.update-chart`；More → `paretoBarSeries` |
| Pareto line：实际累计线、标记及数据标签 | 复用 cumulative controls，写 `options.pareto.cumulativeLineStyle` 的 color／opacity／width／dashType／point／label | 稳定 runtime ID 与柱形分离；More → `paretoLineSeries` |
| Waterfall positive：真实来源的正值角色及标签 | 写 numeric 或 exact-name source 的 `waterfallStyles.positive` | 按实际最高 named author layer 选择；不修改 pointRoles、其他来源或角色 |
| Waterfall negative：真实来源的负值角色及标签 | 同上，写 `waterfallStyles.negative` | 同上；More → 真实 bound `waterfallSeries` |
| Waterfall subtotal：真实来源的合计角色及自动排列合计 | 同上，写 `waterfallStyles.subtotal` | 真源／自动排列合计共享其角色作者，不发明源点 identity |
| Waterfall global subtotal：堆积模式自动合计 | 写 `options.waterfall.styles.subtotal` | 仅 stacked 且 `useSubtotal !== false`；没有 source binding，不创建 sentinel 系列。关闭自动合计、改图型或切表后旧 scope 失效 |
| Pie aggregate：主图实际合计扇区及标签 | 写独立 `options.pie.aggregateStyle` 与既有 `aggregateLabel` | 仅 pieOfPie／barOfPie；runtime aggregateIndex 只描述几何，不写 `series.dataPoints[aggregateIndex]`；More → `pieStyle` |
| Pie connector：实际两条主副图连接线 | 写 `options.pie.secondaryPlot.connector` 的 visible／color／width／opacity／dashType | 两条线共用一个作者 scope；Delete 只隐藏连接线；visible=false、width=0 或 opacity=0 不绘制／命中；More → `pieStyle` |

复合 Pie 完整面板复用同一 controls，配置控件净增 **16**：aggregateStyle 11（color 1、纹理 3、opacity 1、explosion 1、border 5），connector 5。`aggregateLabel` 是原字段，仅随合计分组移动，不计增量。纹理颜色和隐藏连接线的其他设置按状态条件显隐，不能把控件定义数当成同时可见入口数。

共同能力判断沿用真实文档／工作表权限、owner、sheet、revision、busy、当前图元和输入法状态；只读可查看禁用表单，不发作者命令。Enter／F2／Escape、dirty More、失败反馈和宿主焦点恢复沿用 P19。Delete 对其他专属 scope 被选择态消费，不删源数据、系列或整图；恢复格式保留 Pareto／Waterfall 独立 point maps、Waterfall 角色名称及 Pie aggregateLabel。

**More 定位仍有差异**：既有公开 route 只有 tab／section／seriesId／axisId／container。Pareto 能定位 part，Waterfall 能定位来源系列和 waterfall 分组，但尚不能定位具体 role；Pie 能定位 pieStyle，但尚不能精确定位 aggregate／connector 子字段。本批不扩展 Facade 或公开 route 契约，这部分保留未完成。

## 模型、渲染与生命周期

- `ChartCanvasElement` 增加 `mark` discriminator；directEdit 使用上述 4 个 kind。Waterfall 只有全局自动 subtotal 可以没有 binding。来源 binding 校验 field、sourceName、numeric／exact-name selector 和规范化 seriesId；base、无来源 positive／negative、非法 aggregateIndex 等被拒绝。
- Pareto bar／line 的 native view、实际 datum 图元及 attached text 统一映射到专属 scope；graphic 路径的柱形／累计线／标签同样带 metadata。累计线 runtime ID 为 `office-chart-pareto-line:${JSON.stringify(chartId)}`。不把累计派生值伪装为原始 dataPoint。
- Waterfall graphic 标记来自既有计算结果的真实角色／来源，堆积全局 subtotal 保留合成身份。计算仍复用本项目既有 Worker／Rust/WASM 准备路径，不增加 JavaScript 权威累计／瀑布计算器。
- 新 `ChartPieAggregateStyleOptions` 含 color／fillOpacity／fillPattern／border／explosion；connector 使用既有 `ChartLineOptions`。校验覆盖颜色／嵌套线条／分离范围，aggregate 纳入渐变作者校验。runtime theme 和缺省值不写入作者快照。
- 合计专属 style 为最高作者层；旧快照在 synthetic-index dataPoints 中的历史外观仍作为 fallback 读取。新编辑写独立 aggregateStyle，重分区／合计索引变化不会把新格式移写到源点。恢复格式移除独立层，继续保留旧快照和源点数据。
- 合计 explosion 实际改变扇区 center；布局按最大 explosion 留空间，连接线起点跟随合计扇区实际边缘。connector 缺省仍可见、主题中性颜色、1px 实线；隐藏／零宽／零透明度不留下可选 graphic。
- scoped patch 仅投影浮动表单负责的作者字段，不提交 unrelated draft、源数据、其他角色／系列／point maps。既有 `mergeChartOptions` 和 canonical update 负责模型更新、撤销重做、保存及重开；runtime binding、bounds 和 selection 不持久化。
- `chartCanvasElementCompatible`／`chartSpecialMarkCompatible` 在 pending author completion 期间也退休图型不兼容或已关闭全局自动合计的 scope。已真实提交后的晚完成只能结束其原会话，不能重新打开被关闭的标记菜单或夺走新宿主焦点。

模块归属为 `features/charts` 的纯模型／patch planner、主线程真实几何适配与 Vue／Ant 浮动视图。App、解析、文件写出、Rust crate、Worker 协议和 WASM lease 不在本批修改范围；这是对既有计算边界的接线，未进行性能优化或宣称零拷贝改进。

## 文件与接口变化

本批 **22 个文件：16 源码、5 测试／fixture、1 独立证据**。新增源码 3 个，其他源码在原有模块内接线。图片、Docs、ExcelWorkbench、Facade／index、package 和锁文件没有本批修改；共享目录其他在途改动不计入本批。

| 文件 | 本批修改 |
| --- | --- |
| [domain/model.ts](../../packages/office-studio/src/domain/model.ts) | 增加 Pie aggregate style 类型与两个嵌套作者字段；不覆盖其他模型改动 |
| [chart-pie.ts](../../packages/office-studio/src/features/charts/chart-pie.ts) | 新增合计有效样式、继承／历史 fallback、连接线缺省／隐藏与合计 explosion helper |
| [chart-pie-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-pie-settings-controls.ts) | 新增合计／connector controls，浮动和完整面板共用 |
| [chart-special-mark-edit.ts](../../packages/office-studio/src/features/charts/chart-special-mark-edit.ts) | 新增专属 scope label／controls／signature／patch／reset／Delete／兼容判断 |
| [chart-option-validation.ts](../../packages/office-studio/src/features/charts/chart-option-validation.ts) | 校验新增 Pie 作者 DTO |
| [chart-gradient.ts](../../packages/office-studio/src/features/charts/chart-gradient.ts) | aggregate 渐变作者校验 |
| [chart-element-interaction.ts](../../packages/office-studio/src/features/charts/chart-element-interaction.ts) | 专属 metadata、key、native runtime union、边界重建／校验与 More route |
| [echarts-element-interaction.ts](../../packages/office-studio/src/features/charts/echarts-element-interaction.ts) | 专属 native view／datum／attached text 身份接线 |
| [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) | Pareto 累计线稳定 runtime ID |
| [chart-element-edit.ts](../../packages/office-studio/src/features/charts/chart-element-edit.ts) | 委托专属 planner，增加图元兼容入口 |
| [chart-draw.ts](../../packages/office-studio/src/features/charts/chart-draw.ts) | 真实 Pareto／Waterfall／Pie 合计与连接线的绘制／标签 metadata 及样式 |
| [chart-pie-layout.ts](../../packages/office-studio/src/features/charts/chart-pie-layout.ts) | 独立 aggregate explosion 与跟随真实扇区的 connector endpoints |
| [chart-type-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-type-settings-controls.ts) | 复合 Pie 完整配置复用 controls，增加可选第 4 参数 `data?: ChartData` |
| [ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) | 唯一生产调用传入 projection data，继承预览用真实 source／aggregate count |
| [ChartElementFloatMenu.vue](../../packages/office-studio/src/features/charts/views/ChartElementFloatMenu.vue) | 复用原组件，显示专属作用域说明／测试身份，不显示升系列动作 |
| [use-chart-element-editing.ts](../../packages/office-studio/src/features/charts/views/use-chart-element-editing.ts) | pending 期间也退休不兼容图型／自动合计 scope |
| 下表 3 个新测试模块 | 新增 25 个未运行测试声明 |
| [chart-element-interaction.contract.test.mjs](../../packages/office-studio/tests/chart-element-interaction.contract.test.mjs) | 更新 Pareto 专属 binding 断言，原 8 个声明不变 |
| [chart-object-actions.html](../../packages/office-studio/tests/fixtures/chart-object-actions.html) | 测量 selector 支持 markKind／markPart／globalMark 与 directEdit.binding.seriesId |
| 本证据文件 | 来源、分类、作者／命令闭环、草稿和未验收缺口 |

接口变化限于本地图元 runtime union、Pie 嵌套作者类型、纯 helper 和 controls 的可选 data 参数。没有新增公开 Facade 方法、内容 command ID、undo 源、Office schema 或 Worker／WASM 协议。

## 测试草稿与执行状态

按 `test(...)` 声明手工计数，参数／图型循环不另计：**本批新增 25 条，全未运行**；本批涉及测试模块声明总数 **33**（新增 25＋原 interaction 8）。P19 的 34 条声明数未变，不将重用的 8 条再算新测试。

| 草稿 | 新增声明 | 意图 |
| --- | ---: | --- |
| [chart-special-mark-edit.contract.test.mjs](../../packages/office-studio/tests/chart-special-mark-edit.contract.test.mjs) | 10 | Pareto 两个作者层与 point 保留／恢复；numeric／name Waterfall 角色隔离、全局合计／关闭自动合计拒绝；Pie 独立 aggregate／connector patch／reset／Delete；no-op／旧图型拒绝；controls 有效样式与 DTO 校验／深克隆 |
| [chart-special-mark-rendering.contract.test.mjs](../../packages/office-studio/tests/chart-special-mark-rendering.contract.test.mjs) | 8 | 真实 ECharts SVG SSR／ZRender view／contain；native／graphic Pareto；实际 Waterfall 来源／全局合计；两种复合 Pie 真源点／独立合计／连接线；隐藏／零 paint；重分区／explosion 端点；metadata 与既有 More route |
| [chart-special-mark-direct-editing.browser.test.mjs](../../packages/office-studio/tests/chart-special-mark-direct-editing.browser.test.mjs) | 7 | 完整 ExcelWorkbench／Ant 表单／Canvas2D 的真实测量命中、独立编辑／历史／save-reopen；numeric／name Waterfall 与 dirty More；全局合计；两种 Pie aggregate；connector 样式／More／Delete；只读／模拟 composition／真实提交后晚完成与切图型／切表 |

渲染草稿为 preparedData 构造对应真实源工作表和单元格，满足 `assertChartSourcePresent`；不以空工作簿绕过源身份。浏览器草稿复用现有 canonical-completion gate：延迟的是已真实执行 `FChart.update` 后的 completion，不替换内容命令。Drawer 的 Escape 关闭使用生产 `keyboard=!busy` 契约。这些只经过静态阅读，实际运行结果仍未知。

可替换 UI 均复用 `antdv-next` 的 ConfigProvider／Form／FormItem／Input／InputNumber／Select／Switch／ColorPicker／Button／Tooltip／Alert 和现有 ChartOptionFields；绘制、命中、选择框继续 Canvas，无新增原生输入桥或其他组件库。theme token／popup container 已接线，**明暗主题、浮层、焦点和系统 IME 仍未实际验证**。

本批仅阅读／搜索文件、固定源码和手工静态审阅，写入实现、测试草稿与证据。**没有运行测试、类型检查、构建、dev／Vite、浏览器、Cargo／WASM、基准、服务或 `git diff --check` 等自动检查；没有提交、发布、安装 SDK 或修改 package／锁文件。**

## 剩余缺口

1. 25 条新草稿与 P19 草稿尚未运行，类型和构建结果未知。主题／缩放／DPR／权限变更／竞争弹窗／系统中文 IME／真实字体 shaping 尚未验收，合成 composition 事件草稿不等于系统输入法验收。
2. More 尚不能精确定位 Waterfall role 或 Pie aggregate／connector 子字段；上游不规则 overlay／semanticId／嵌套元素、opaque seriesId、3D／Radar 原生逐点等仍有差异。
3. 本批 4 个专属 direct-edit kind 不代表 7 个上游 mark kind 或 45 图型全部覆盖。P19 的多级分类 expand／collapse、Pivot field button／overflow、heatmap／hierarchy／graph／wordCloud／chord 等专属元素缺口继续保留，不添加空操作充数。
4. 内部 save／reopen 草稿不等于 XLSX／DOCX／独立 Office 重开或保真验证；真实 Office 图表样本、源更新／重分区／跨工作簿复制及完整撤销重做文件闭环仍待实测。
5. 首屏、滚动、编辑、缩放、native index 重建、Worker／WASM 准备、传输复制次数／字节、内存和大样本均未测量，保持性能未验收，不声称已优化或达标。

以上是当前实施和验收缺口，不缩减完整 Univer Sheets／Docs／Pro 工作台的交付范围。
