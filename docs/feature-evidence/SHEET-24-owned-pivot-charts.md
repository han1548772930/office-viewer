# SHEET-24 独立/关联透视图创建、字段命令与 Facade

日期：2026-11-26。状态：实现及合同草稿已落盘，**未验收**。本批续接 [SHEET-25 共享连接](SHEET-25-slicer-connections.md)，覆盖真实 owned/linked 创建、字段配置、一次作者事务、权限、事件和快照保存；不把入口或草稿数量计为功能通过。

## 固定依据与入口口径

固定 OSS 为 `E:\rust_test\univer`，Univer v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。只读核对项目保留的发布类型与 ES，未将商业 SDK 加入依赖：

- [sheets-pivot-chart/types.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts)：Range/PivotTable 来源 union、initialFieldsConfig、字段设置与报告编辑 DTO。
- [f-pivot-chart.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts)：来源、字段、选项、刷新、删除及 builder 的公开契约。
- [sheets-pivot-chart-ui 发布 ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/es/index.js)：`G/Hn/Un` 创建与来源选择，`W` 面板 operation，轴菜单和浮层入口。
- 完整标签/分组/order/classic/grid、条件、两条互斥根、两个创建子项、四个固定轴子项和动态层级清点见 [UI 独立记录](SHEET-24-pivot-chart-ui.md)。本批常规 Ribbon 根增量为 **0**；operation alias、字段枚举和 API 不计为新的 Ribbon 根。

模块为 `SHEET-24 → SHEET-23/SHEET-22`：Vue/Ant 管理 UI 与临时草稿，命令编排权限/历史/事件，Pivot 服务提供唯一报告聚合，Chart 服务复用源读取与作者数据投影。未向 App/store 添加聚合器。控制 DTO 属于 TypeScript 领域边界；本批尚未迁移大型 Pivot 计算到协议 2 的 retained Rust/Worker 路径，不能宣称性能或零拷贝达标。

## 已落盘行为

| 行为 | 实际路径与边界 |
| --- | --- |
| 独立区域创建 | `source.kind=range` → detached report → `chart.ownedPivot`；initialFieldsConfig 与图表在同一提交中应用，不创建 worksheet、PivotTable 或输出 cells |
| 关联创建 | `source.kind=pivotTable` 校验完整 unit/subUnit/report 身份；旧 linked ChartConfig 走同一入口，绘图 source 由真实 report 输出推导 |
| 字段编辑 | 四区、添加/移除/移动/排序、显示名/格式/11 类汇总、标签/值过滤、分组/报告选项复用 [字段引擎](SHEET-24-pivot-chart-engine.md)，不从展示值反算报告 |
| 一次应用 | `configure({patch,fieldsConfig}, context)` 在同一内容事务里校验字段与视觉候选；第二次 source-aware 作者归一化保留刚修改的 owned report 与单次 revision |
| 空报告 | 缺少 row/value 区时返回可继续编辑的空数据契约；完全配置的报告及普通图表仍验证非法字段绑定，不能借空记录绕过校验 |
| 权限与会话 | source read、图表完整 footprint、linked 输出范围的 read/edit；可选 expectedRevision/historyGeneration/userId/host。工作台绑定 active unit/sheet，绑定 Facade 可编辑非活动文档 |
| Before 与 settled | 请求先 clone，两个 Before 目标先捕获，首个观察者后重验版本；取消/异常/重入拒绝原提交。related charts/reports/slicers 的 settled payload 全部先捕获，再隔离观察者发布 |
| no-op/刷新/删除 | 无变化返回接受且不增加历史/settled；refresh 保留作者设置；删除清理未引用图表资源与连接，保留 source cells 和独立 PivotTable |
| 保存与历史 | 统一文档历史保存完整来源、owned report/linked reference、fields/options 和视觉设置；快照重开合同已编写，真实 XLSX 原生往返仍缺 |

五个内容命令为 `sheet.command.add-pivot-chart`、`edit-pivot-chart-report`、`configure-pivot-chart`、`refresh-pivot-chart`、`remove-pivot-chart`，均使用 `sheet.command` 前缀。`sheet.operation.open-pivot-chart-panel` 是面板别名，属于 UI 会话。

公开类新增 `FSheetPivotChart`、`PivotChartConfigBuilder`，FWorkbook/FWorksheet 可查询或创建 owned/linked chart。builder 的 string/FRange/coords/完整 range source、linked reference 与 initialFieldsConfig 委托真实命令；继承 FChart 的视觉 `getOptions/setOptions`，报告选项暂名 `getReportOptions/setReportOptions`，这是明确的上游 API 差异，不能宣称 Facade 全等。

## 文件与合同草稿

| 文件 | 责任 |
| --- | --- |
| [pivot-chart-api-types.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-api-types.ts) | 来源 union、创建/配置/context DTO 与五命令常量 |
| [pivot-chart-commands.ts](../../packages/office-studio/src/api/pivot-chart-commands.ts) | 内容命令、候选校验、权限/版本、Before/settled、资源清理 |
| [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)、[events.ts](../../packages/office-studio/src/api/events.ts)、[facade.ts](../../packages/office-studio/src/api/facade.ts)、[index.ts](../../packages/office-studio/src/index.ts) | 注册、公开 API/事件与统一提交 |
| [chart-statistics-service.ts](../../packages/office-studio/src/features/charts/services/chart-statistics-service.ts) | source-aware 二次准备保留新报告和 revision |
| [chart-pivot-data.ts](../../packages/office-studio/src/features/charts/chart-pivot-data.ts)、[chart-record-projection.ts](../../packages/office-studio/src/features/charts/chart-record-projection.ts) | 独立/关联同源读取与空报告绑定 |
| [pivot-chart-facade.contract.test.mjs](../../packages/office-studio/tests/pivot-chart-facade.contract.test.mjs) | 24 条声明 / 36 个计划场景 |
| [pivot-chart-empty-projection.contract.test.mjs](../../packages/office-studio/tests/pivot-chart-empty-projection.contract.test.mjs) | 5 条声明 / 9 个计划场景 |

主 API/投影草稿合计 **29 条声明 / 45 个计划场景**，使用真实 OfficeStudio/Facade 与非零偏移 source fixture，覆盖创建/配置/历史/重开、跨 sheet、旧 linked、取消/重入、source-aware、权限/过期、refresh/remove、切片器联动与空数据合法性。参数展开数只表示计划用例。

引擎批 [30/51](SHEET-24-pivot-chart-engine.md)、创建/字段 UI 批 [43/70](SHEET-24-pivot-chart-ui.md)、[值显示方式引擎](SHEET-23-24-pivot-show-data-as.md)与 [API/格式投影](SHEET-23-24-pivot-show-data-as-facade.md)各自独立计数，未执行，不混计为本文件的 29/45。

## 差异与待验收

- UI 的「区域同时创建透视表与图」、轴/member 展开折叠、categoryLevel 与 Canvas 字段/成员浮层尚待后续批接线；本地一个根和 Modal 与上游两条互斥根/360px 面板有明确差异。
- 任意 value index、完整日期条件/查询、全部 Builder/Facade/controls、source replacement/cache migration 继续保留目标，不能从现有字段方法推断已完成。
- 真实 Office PivotCache/PivotTable/PivotChart/chart/slicer parts、资源关系和重新打开、服务端权限/rebase、retained Rust 聚合/预算/取消/复制计量仍缺；详见 [后端需求](../BACKEND_REQUIREMENTS.md)。
- 仅做静态读写、接口回读和草稿计数。测试、类型检查、构建、Rust/WASM、dev/浏览器、Windows IME、主题/焦点、Canvas2D/WebGPU、Office 样本与大数据基准均未运行；状态保持未验收。
