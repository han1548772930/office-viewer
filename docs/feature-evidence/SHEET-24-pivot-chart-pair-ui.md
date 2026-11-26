# SHEET-24 配对创建与透视图显示开关 UI

日期：2026-11-26。状态：生产接线与测试草稿已落盘，**未验收**。本批记录区域来源的「同时创建透视表和透视图」、报告放置目标及两个图表显示开关；不将核心事务、Canvas 按钮渲染或前批字段面板记为本批完成。

## 固定来源与入口清点

固定 OSS 为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 依据为项目内保留的同版发布包，仅读取，未安装或执行商业 SDK。

| 已读取来源 | 可观察契约与本地对应 |
| --- | --- |
| [sheets-pivot-chart-ui ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/es/index.js) | `G` 为 `sheet.command.menu-insert-pivot-chart`；创建 selector 有 `withPivotTable` / `chartOnly` 两项；`Un` 在已有透视表时直接返回引用，否则按模式创建或采用独立报告；`Hn` 是普通函数，将引用或区域变为图表来源 |
| [menu-insert-pivot-chart.command.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/commands/commands/menu-insert-pivot-chart.command.d.ts) | 图表创建命令与来源参数；本地仍由既有 Ribbon 根打开创建面板 |
| [图表 types.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts) | `controls?: Partial<IPivotChartControls>`，包括 `showFieldButtons` 与 `showExpandCollapseButtons` |
| [图表 Facade](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts)、同包 UI ES | 显示开关的作者配置与设置入口；UI ES 的图表设置区有对应两个 Switch |

| 分类 | 本批数量、位置与条件 |
| --- | --- |
| 新常驻或 Pro Ribbon 根 | 0。沿用 [前批入口证据](SHEET-24-pivot-chart-ui.md)：上游 DATA / ANALYSIS，order 1，grid row 2 / column 1；本地 DATA「分析」order 30，classic/grid 均一个既有根 |
| 创建模式选项 | 2，面板来源区。仅未创建的 range 来源显示；已有图表和 linked 来源不提供配对创建 |
| 配对报告目标控件 | 2，目标工作表 Select 与起始单元格 Input；仅配对模式显示；可选表需 worksheet read/edit 权限 |
| 图表显示设置 | 2，既有图表设置区之后；新建与已有 owned/linked 图表均可编辑；busy、过期、来源未就绪、无配置权限或 composition 期间禁用 |
| Canvas 上下文或溢出新增入口 | 本批 0。实际字段按钮、展开折叠按钮和 filter 弹层不在此批；动态字段数量不计为常驻根数 |

本地默认 `chartOnly`，创建模式在既有 Modal 内选择，未新增上游相互排斥的两个菜单 factory 或复刻其 sidebar。上游 `withPivotTable` 先执行建表、等待 collection 再建图；本地调用单一内容事务创建表与图，避免留下部分资源。这里记录实现差异，不声明最终交互已通过浏览器验收。

## 模型、命令与预览边界

归属为 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的 `SHEET-24 → SHEET-23 / SHEET-22`。Vue/antdv-next 负责输入和临时草稿；创建 helper 负责放置 DTO 与 advisory 校验；真实报告计算、注册、权限与回滚由既有内容服务负责。组件没有新建聚合或解析引擎。

| 行为 | 实际接线 |
| --- | --- |
| 新建区域来源 | 默认独立 owned 报告。改为配对时保留原字段、映射、标题和显示设置，不因模式切换重置 sourceKey |
| 报告目标 | chart drawing host 固定当前工作台 sheet；目标 Select 只改变新 PivotTable 的 anchor sheet。同源同宿主默认在来源右侧两列，例如 `D3:H7 → J3`；异表默认 `A1` |
| 起始位置 | `pivotChartCreateOptions` 解析真实单个 A1、workbook/目标 sheet、worksheet 与 anchor read/edit、anchor/source overlap；拒绝空地址、范围地址、越界和带表名地址。此处仅 advisory，不预先声称完整输出或占用已合法 |
| 最终创建 | 一次 `FWorksheet.insertPivotChart(info, context, options)`；第三参数为 `{ pivotTable: { anchor: { unitId, subUnitId, row, col } } }`。完整输出范围、占用、权限与单事务资源回滚由 [配对核心](../../packages/office-studio/src/features/pivot/services/pivot-chart-create-table.ts) 和内容命令校验 |
| 已有 linked 来源 | 继续引用当前 PivotTable；不创建第二张表，也不将已有图改为配对新建 |
| 预览 | `pivotChartPreviewConfig` 接受 prospective pair anchor；defensive clone 后移除 outputRange/generatedCellKeys/drillDownMap/axisOutput，以同一图表数据服务计算未注册的 data-only 报告；不向工作簿写报告、图表或 cells |
| 显示开关 | `draft.controls` getter/setter 的唯一存储在 `settings.config.controls`；复用 [controls 归一服务](../../packages/office-studio/src/features/pivot/services/pivot-chart-controls.ts)，缺省两项 true；相同隐式默认不增加配置 |
| 应用设置 | 两项开关与外观/fieldsConfig 一次 `configure({patch,fieldsConfig},context)`；不额外执行 update-controls 或报告 setter。仅开关变化应保留 owned/linked 报告内容与 revision；真实运行仍待验证 |
| 历史与保存 | 沿用按文档隔离的内容 history 和 workbook snapshot。草稿覆盖单 history、undo/redo、JSON 重开；真实 XLSX/Office 文件往返未验收 |

创建模式、目标 sheet 与地址仅属于创建 metadata，不塞入 chart appearance。它们进入 preview sourceRevision，改变目标会重读预览；deliberate reread 保留已选目标。用户可在目标文本错误时继续编辑，错误通过 Ant Alert 显示并禁用创建。

## 生命周期、Ant 控件与计算域

- 保留当前 draft/owner、Studio/unit/host/user、model revision、history generation、request generation 与真实内容上下文校验。预览读取前后拒绝旧 session；提交捕获创建 options 与 context；旧 finally 不清理新请求的 busy。
- 保留 composition capture 与 nextTick 接收最终输入；composition 期间禁用创建、导航、关闭和两个 Switch。目标文本仍由 Ant Input 承载，没有新原生表单桥。
- 新增 FormItem、Select、Input、Switch、Alert 均为 `antdv-next`，使用统一 ConfigProvider 与 theme token；新分隔线使用 token CSS 变量。真实主题、焦点恢复和原生 IME 尚未运行验证。
- 计算与跨边界协议沿用现有 Pivot/Chart 服务；本批没有新 Worker/WASM 协议，也没有性能优化声明。大样本继续遵守 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)，需后续 Rust/Worker、取消、传输/内存和同数据集基准验收。

## 本批文件与接口

本批 **5 个文件**：2 个既有生产文件的创建/controls 片段、2 个新测试草稿、本证据。未改 API/domain、两个 engines、chart 数据投影、ChartOverlay 或顶层功能清单；这些为主任务归属。

| 文件 | 本批职责 |
| --- | --- |
| [pivot-chart-panel-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-panel-state.ts) | 创建 mode/目标草稿、默认位置、reread 保留、options/error helper、controls 单一存储与 prospective preview anchor |
| [PivotChartPanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue) | 全 Ant 创建选项、目标输入、advisory 错误、两个显示 Switch、一次创建/配置与 owner/IME guard |
| [配对 UI 合同草稿](../../packages/office-studio/tests/pivot-chart-pair-ui.contract.test.mjs) | 默认/metadata、真实单次创建与 linked 身份、权限/占用/完整输出回滚、预览不写入、controls/no-op/history/JSON |
| [配对 UI 浏览器草稿](../../packages/office-studio/tests/pivot-chart-pair-ui.browser.test.mjs) | classic/grid 真 Ribbon、mode 切换、目标错误/权限、pending 拒绝、两种所有权 controls、真实内容/history/JSON、theme/composition/focus |
| 本证据 | 来源、入口分类、接口、计数、差异与未验收项 |

复用既有 [数值 fixture](../../packages/office-studio/tests/fixtures/pivot-chart-ui-fixture.mjs) 与 [真实工作台 fixture](../../packages/office-studio/tests/fixtures/pivot-chart-ui-workbench.html)，未新增或改动这两个 fixture，不重计为本批文件或草稿。gate 释放后调用原真实 Facade；新的第三参数由已有 `...args` 透传。

本批内部接口新增 `PivotChartCreationMode` / `PIVOT_CHART_CREATION_MODES` / `PivotChartCreationDraft`；`PivotChartDraft` 与 open input 新 createMode/pivotTableSheetId/pivotTableAddress，draft 新 controls accessor；新增 `pivotChartCreateOptions` / `pivotChartCreateOptionsError` / `setPivotChartDraftControls`；`pivotChartPreviewConfig` 新第五参数 `pairedTable`。底层 insert options 与 controls domain/Facade 由主任务实现，此处仅消费。

## 草稿计数与未验收项

以下为**人工源码清点**，不是执行结果、通过数或功能覆盖率；本轮没有运行计数脚本。

| 本批草稿 | test/test.each 声明 | 参数展开后计划场景 | 状态 |
| --- | ---: | ---: | --- |
| UI 合同 | 15 | 34 | 未运行 |
| UI 浏览器 | 13 | 31 | 未运行 |
| 合计 | 28 | 65 | 全未运行 |

合同参数覆盖 2 个目标 sheet、7 个非法 anchor、3 类目标权限、2 个预览目标、owned/linked × 3 controls 组合及 3 个非法 controls patch。浏览器覆盖 classic/grid × chartOnly/pair、3 个非法 anchor、2 创建模式 × source/history/user/host 变化、owned/linked × 3 controls 组合和 2 个 theme/composition 变体；内部断言循环不额外计场景。

复用来源 `D3:H7` 的四条记录，关闭四 totals 后真实 Region/Amount 数据预期为 `[40, 60]`。配对 `B9` 为 row 8 / col 1，标题 B9/C9、North 数值 C10=40、South 数值 C11=60。输出保护样本在 C11，起始 B9 可以通过 advisory，但完整事务应拒绝且不发布表/图；这里仅为草稿预期。

不混计前批 [创建/四区面板](SHEET-24-pivot-chart-ui.md) 的 43/70、[Show As UI](SHEET-23-24-pivot-show-data-as-ui.md) 的 55/165、共享 Show As 核心的 34/65、Facade 的 13/24，以及主任务配对核心/controls 草稿。

本轮**未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、浏览器、Office 文件样本、性能基准、git diff --check 或文档链接检查**；只读取和编辑源码/证据。来源路径为此前已读文件，不把文档落盘当作链接验收。

继续保留未验收：实际 Canvas 字段/展开折叠按钮与 filter 弹层、上游互斥 factory/sidebar、完整 member/filter 类型/格式化/空白 DTO、真实 Windows IME/Ant 焦点、用户切换响应式状态、权限与旧异步结果实际运行、XLSX pivot/cache/chart 保真往返、WebGPU 与真实大样本内存/性能。浏览器草稿计划禁用 navigator.gpu 使用 Canvas2D；合成 composition 不证明原生 IME；夹具 `Object.defineProperty` 用户变化不能证明生产用户禁用态。

SHEET-24 与 Sheets/Docs 全部 Pro 能力继续保持完整交付范围；本批代码与草稿不构成功能完成或验收证明。
