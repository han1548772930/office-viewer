# SHEET-22 图表类型源码对照

状态：未验收。开源参考 checkout 为 `E:\rust_test\univer` tag `v1.6.0`，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 清单及特定面板证据分来源标注。历史 1.5.2 包只作为历史对照；当前 1.6.0 Pro 发布包类型和公开声明证据来自 `.research.local/univer-1.6.0/`，但 `engine-chart` 包未在本地解包，底层配置类型与运行时语义仍未核实。

## Pro 发布包证据

- `.research.local/univer-1.6.0/packages/chart-ui/package.json` 将 `@univerjs-pro/chart-ui` 标为 1.6.0，并依赖同版本 `engine-chart`。同目录 `lib/types/components/ChartTypeOptions.d.ts` 导出 `CHART_TYPE_CATALOG`，包含 45 个目录选择项，涵盖常见图型、堆积/百分比和 3D 变体、特殊图型及 Pie-of-Pie / Bar-of-Pie。45 是目录声明项数，不代表各入口均启用；复合饼图项复用 `ChartTypeBits.Pie` 并用 `pieSecondaryPlotType` 区分。
- 同一 chart-ui 包的 `lib/types/index.d.ts` 声明系列、标题/副标题、坐标轴、图例、趋势线、误差线、3D 和类型专属设置 section；`lib/types/services/chart-view-state.d.ts` 有 46 个编辑器状态键。这些是类型与 UI 状态声明，不等于持久化字段清单或运行时实现细节。
- `packages/sheets-chart/lib/types/commands/commands/sheets-chart.command.d.ts` 声明 `InsertChartCommand` 的类型、source、theme、position、尺寸、背景、边框和 config 参数及插入时的可逆 mutation 准备；`sheets-chart-update-config.command.d.ts` 声明 config 增量/替换更新命令及对应 mutation。
- `packages/sheets-chart-ui/lib/types/commands/commands/focused-chart.command.d.ts` 声明打开焦点图表编辑面板、复制、导出图片和删除命令；`services/sheets-chart-ui.service.d.ts` 声明数据源/范围、方向、配置更新和面板路由 API。Docs Chart 另有插入、删除、配置/数据源/绘图快照更新声明。上述是发布包类型证据，不证明未公开的 `engine-chart` 内部行为。
- `.research.local/univer-1.6.0` 内的 chart manifests 记录 chart-embed、chart-ui、docs-chart、sheets-chart、pivot-chart 相关发布包。当前已复核类型目录及部分入口、命令声明；完整菜单运行时过滤、模型字段、各类型数据约束、面板组件状态转换和 Office 往返仍需补证。
- 固定版 v1.6.0 OSS checkout 未找到上述 Pro 图表 UI 源码。固定版 `packages/sheets-drawing-ui/src/services/canvas-float-dom-manager.service.ts` 的 `applySheetChartTransformerConfig` 只证明通用图表对象的选中、变换框交互，不能证明 Pro 图表类型或编辑面板语义。

## 本地实现与差异

- `domain/model.ts` / `chart-engine.ts` 当前支持 `bar | line | area | pie | doughnut | stacked-column | scatter`。Scatter 的共享 X 与按系列指定独立数值 X 字段、X/Y 线性/对数轴读数已进入模型和 `readChartData`。
- `features/charts/echarts-option.ts` 为本地模型生成 ECharts option，覆盖当前七类、类别/空值、图例、标题、轴标题/刻度、散点独立 X 与对数非正数过滤。
- `features/charts/echarts-renderer.ts` 使用 ECharts 6.1.0 tree-shakeable CanvasRenderer、注册当前七类所需的 Bar/Line/Pie/Scatter 和组件。`ChartOverlay.vue` 按数据、尺寸、DPR 和 antdv-next theme token 更新位图，再通过带 sourceVersion 的 image draw command 合成到工作表 Canvas；WebGPU backend 按版本更新纹理。图表对象自身仍由本地 ChartModel、命令/历史与 XLSX 链路控制。
- 依赖锁定 `echarts@6.1.0`，许可证为 Apache-2.0。按官方模块导入文档只注册当前适配器用到的 chart/components/CanvasRenderer；官方渲染指南说明 Canvas 适合大量图元或数据场景。该选型支持本地 Canvas 图形呈现，不是 Univer 内部架构依据。
- `components/spreadsheet/ChartEditorPanel.vue` 按固定版 Pro `ChartEditPanel` 的 Style/Data 双页能力边界建立本地 antdv-next Drawer，目前编辑标题、数据区域/方向、分类/数值字段、图例和 X/Y 轴标题；应用经 `FChart.update` 调用现有 `sheet.command.update-chart`，使用工作簿历史支持撤销/重做。其字段仅限当前 `ChartModel` 已表达内容，不代表完整复刻 Pro section。
- ECharts 当前绘制层不将图内 tooltip、legend/dataZoom/brush 等事件路由为编辑器 command。当前 Drawer 尚未对齐系列样式/类型、组合系列、标题样式、趋势线、误差线和类型专属设置。其 Canvas2D 图表位图先栅格后由工作表 WebGPU/Canvas2D 合成，并不表示 ECharts 图内绘制使用 WebGPU。
- Scatter 的模型/API/XLSX 字段接线和 ECharts 绘图 adapter 已添加；option 契约与 Edge Canvas 生命周期浏览器用例已通过，图表工作台插入/移动/缩放/数据刷新浏览器用例已通过。真实 Office/Excel 独立重开、错误恢复和完整数据源边界仍待验证，不能因此把 Scatter 标记为完整验收。
- `packages/office-studio/tests/chart-engine-rendering.contract.test.mjs` 覆盖折线图轴标签/网格和对数轴；`echarts-option.contract.test.mjs` 覆盖当前七类配置、主题、分类/空值、散点 X/Y 与对数过滤；`echarts-renderer.browser.test.mjs` 在 Edge Canvas 覆盖七类非空栅格输出、缓存、数据刷新、resize/DPR 替换、dispose/retain 和无效尺寸原子拒绝。仍未覆盖 Pro 全类型各自数据边界及编辑面板行为。
- `packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue` 提供六种本地类型选项。普通图表插入对话框提供当前七种本地模型类型；选中图表可打开受限属性 Drawer。两者均不能视为 Pro 45 项目录或完整 Pro 图表编辑器的等价实现。

## 范围与未核实项

此记录证明 Pro 1.6.0 发布包中的 45 个图表目录选项和部分 API/UI/命令类型声明，以及本地当前七类渲染和有限属性编辑闭环；不证明各菜单实际启用范围、闭源 `engine-chart` 内部字段/行为、所有 Pro 类型的数据约束或完整编辑闭环。全量命令与菜单运行时、系列/轴设置映射、类型数据契约及 Office 往返仍需逐项复核。

## 技术参考

- ECharts 6.1.0 release/license: https://github.com/apache/echarts/releases/tag/6.1.0
- Tree-shakeable package import: https://echarts.apache.org/handbook/en/basics/import/
- Canvas/SVG rendering trade-offs: https://echarts.apache.org/handbook/en/best-practices/canvas-vs-svg/

散点图第一切片已定义数值 X/Y 映射，并连通模型、API、XLSX 数据映射及 ECharts option。仍未证明上游 Pro 的全部散点设置与 Office 往返行为等价。1.6.0 Pro `CHART_TYPE_CATALOG` 明确包含组合系列、雷达、气泡、瀑布图、Pareto、直方图、箱线图、K 线、树图、旭日图、桑基图、热力图、仪表盘、和弦、关系网络、词云及多个常见类型变体。除当前七类外均未接入本地模型/编辑器；3D 是否适合本地 Canvas 目标及 ECharts 无直接同名 series 的图型映射需要逐项定义与验收。SHEET-22 整体保持未验收。

## 验证边界

本轮已执行 `pnpm.cmd check:types` 和 `pnpm.cmd build`，均通过。ECharts option 契约 5/5、图表模型/Facade/场景契约 14/14、ECharts Canvas 生命周期浏览器 5/5、图表工作台插入/移动/缩放/数据刷新浏览器测试 1/1 通过；`cargo test --manifest-path crates/office-writer/Cargo.toml` 51/51 通过，包含当前 XLSX 图表类型 writer 测试。Build 输出包含既有 Sparkline 动态导入提示及主工作台 chunk 大于 500 kB 提示。

仍未验证 WebGPU 与 Canvas2D 后端差异、真实 Office/Excel 对图表输出的独立重开、主题/长图例/窄视口专项浏览器行为、所有类型导出往返，以及 1.6.0 Pro 全图型和配置面板逐项映射。图表工作台浏览器用例已覆盖属性 Drawer 编辑、命令撤销/重做和内部保存快照后重建工作簿；这不等于 XLSX 导出后由 Microsoft Excel 独立重开验收。以上结果只证明当前七类和有限编辑字段的已测接入切片，不改变 SHEET-22 整体“未验收”状态。
