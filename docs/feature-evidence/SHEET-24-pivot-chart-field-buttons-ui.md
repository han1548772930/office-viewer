# SHEET-24 透视图四区字段按钮与筛选条件 UI

日期：2026-11-26。状态：生产接线、fixture 和测试草稿已落盘，**未验收**。本批覆盖真实 ChartOverlay 的字段按钮、同区溢出、分组成员 checked/partial、日期和标签条件，以及定位既有字段设置面板。所有草稿均未运行，不把声明数当作通过数或功能完成证明。

## 固定来源与入口分类

固定 OSS 为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 对照来自项目内同版发布包，只读来源，未安装或执行商业 SDK。

| 来源 | 可观察契约、本地对应和差异 |
| --- | --- |
| [sheets-pivot-chart-ui ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/es/index.js) 的 `wr` | 按 filter、column、row、value 清点四区，标签为筛选、图例（系列）、轴（类别）、值；字段使用报告本地 ID，值字段重复实例分开保留。区域定义顺序已对应，不宣称 engine 精确屏幕位置已核实 |
| 同文件 `dr` | 溢出列出同区所有字段，含已显示项；224px 宽、最大 320px 高。本地每区最多显示 2 项后增加同区溢出，阈值为本地布局选择 |
| 同文件 `fr`、字段操作会话服务 `Y` | 336px 成员弹层、字段会话及取消、成员读取与提交。本地用 Vue owner/session/generation 和真实内容命令；未复刻上游 popup 服务内部实现 |
| [sheets-pivot-ui ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/es/index.js) 的 `Zo` | 分组 checkedItems 与 partialItems、all、clear、排序、日期入口。本地从 Facade 的分组状态构造草稿，不从字符串 selectedItems 猜 raw 身份；partial-only 草稿可确认 |
| 同文件 `Xo`、`Ko` | 8 个日期操作符、dateEqual 默认、useWholeDay 默认 false、原 numeric serial 与时间保留。上游初始默认当天；本地新条件日期留空，要求明确选择。`Ko` 不表示 day0/虚构 day60；本地保留已有 day60，使用 Ant Input 日历桥 |
| [共享筛选条件服务](../../packages/office-studio/src/features/pivot/services/pivot-filter-predicate.ts)、[公开筛选适配](../../packages/office-studio/src/features/pivot/services/pivot-filter-api.ts) | UI 消费共享 8 日期/14 caption 操作符与 normalizer；numeric code 由公开 API 层处理，组件没有第二份谓词或数字代码映射 |

| 分类 | 可复查的本批数量、位置与条件 |
| --- | --- |
| 常驻 Ribbon/Pro 根 | 新增 0，沿用既有透视图入口与图表选择 |
| Canvas 上下文字段组 | 4 组，filter/column/row/value；只有对应报告有字段、选中的 chart 可读取且 showFieldButtons=true 时显示 |
| 基线字段实例 | fixture 共 4 个 authored 实例：Date、Segment、Region、Amount；本地 ID 独立于源列编号 |
| 溢出样本 | 6 个 authored 实例，其中同一源列有 sum/avg/count 3 个 value 实例；显示 5 个直接字段按钮与 1 个溢出根，溢出列 3 项，不能将重叠的溢出项重复计为功能 |
| 成员与条件模式 | 成员模式；日期模式 1 个 capability 入口、标签模式 1 个入口；返回成员入口仅条件模式显示。日期显隐来自真实来源 capability，不按名称、显示标签或数值键猜测 |
| 字段管理与排序 | 设置、移除 2 项；升序、降序、清除排序 3 项。value 实例仅有字段管理，不提供成员/条件编辑 |
| 筛选操作 | 确认与清除 2 个内容动作，取消/关闭 2 个会话动作，错误时读取重试 1 个；clear 在成员模式清所有类别，条件模式只清自身类别 |
| 辅助输入 | 成员搜索与全选各 1；成员选择数量随 catalog 变化，80 项分页；每类条件 1 个 Select、1 或 2 个 operand，日期另有 whole-day Checkbox；动态成员数量不证明功能完成 |

本地四区位置为 filter 左上、value 右上、row 左下、column 右下；上游精确 plot/axis 位置尚未核实。synthetic Values 的 direct/overflow、方向/index 和拖拽已在[独立 Values UI 批](SHEET-23-24-pivot-values-position-ui.md)接线，仍未验收；这里的原计数只统计 authored 字段，新增 marker 单列，不重计本批。

## 实际行为与事务边界

归属为 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的 SHEET-24 → SHEET-23/SHEET-22。Vue/antdv-next 负责入口、分页和临时草稿；报告过滤、分组、重算与历史消费共享计算和内容命令。实际 ChartOverlay 接线由主任务负责，本批 fixture 使用该生产路由，未重复手动挂载字段按钮。

| 行为 | 已落盘的实际接线 |
| --- | --- |
| 字段与 badge | 真实 Facade field setting、报告本地 ID、authored position；`getPivotFieldFilterCategories` 覆盖 manual/grouped/record/date/label/value，不只看 selectedItems |
| 分组成员 | await `getFieldItems` 后同步读取 `getManualFilterItemState`，分别保留 checkedItems 与 partialItems；不修改的 partial key 将原 typed raw membership 交给同一事务保留；点击 partial 转 checked/unchecked 明确更换该组选择 |
| manual 确认 | 一次 `applyReportEdit({ type: 'setFilter', fieldId, items, isAll, partialItems }, ownerContext)`。全空拒绝，partial-only 合法；all 对整个 catalog 生效，包括分页/搜索隐藏项，只清 member 类别 |
| clear | 一次 `setFilter` report edit，`clearAllFilters: true`；移除 manual、record、date、caption、value 类别，而保留字段与报告身份 |
| 日期与标签 | 真实 `setDateFilter` / `setLabelConditionFilter`；日期 capability 含 date1904，原 numeric fractional serial 不因显示重新计算；区间缺值或倒序拒绝；空 caption 为有效条件 |
| 类别协调 | 默认 allowMultipleFilters=false 使用共享类别互斥语义；true 按共享模型求交集。图表按钮本批没有另建 options 权威，使用既有报告配置 |
| 排序/移除 | 真实 Facade sort/remove；重复值实例按本地 ID 移除，保留同源其他值实例与 chart appearance |
| 设置 | 调既有 settings 路由并传 initialFieldId；PivotChartPanel 切字段 Tab、定位并聚焦对应卡片；只改变 initialFieldId 不重新加载草稿 |
| 权限与异步 | owner 含 workbook/host sheet/user/revision/history generation；晚到 member 不能替换新字段，旧 finally 不能清新 loading/busy。命令提交最终再校验原 context。自己的单次提交允许 revision+1 |
| 焦点/IME | 关闭与成功提交仅向仍当前的 owner 恢复 Canvas focus；composition 期间保留输入并拒绝内容/导航动作；过期 composition 完成后关闭；Escape 尊重 child picker 的 defaultPrevented/isComposing/229 |

UI 均为 `antdv-next` 的 Dropdown、Button、Tooltip、Alert、Form/FormItem、Input、Checkbox、Pagination、Select、DatePicker，继承工作台 ConfigProvider 与 theme token。现有虚构 1900-02-29 输入桥也是 Ant Input；其局部文本在无效/未完成输入期间 pinned，避免卸载 IME。正常 DatePicker 尚不能创建新的虚构 day60，只能保留或编辑已有 day60；真实 Windows IME 与 DatePicker 的键盘输入/弹层 Escape 仍待运行。

内容仍由 Canvas 绘制。本批无新 Worker/WASM 协议、无性能优化或达标声明。真实大样本、Rust/Worker、取消、传输次数/字节、内存 lease 和两种渲染后端仍按 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 验收。

## 本批文件与草稿

本批 **12 个文件**：7 个生产文件（6 新、1 既有 Panel 的定位片段）、2 fixture、2 测试草稿、本证据。未将主任务的 Overlay、domain、Facade/API、引擎、chart 投影或顶层清单改动算入本批。

| 文件 | 本批职责 |
| --- | --- |
| [pivot-chart-field-button-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-field-button-state.ts) | 四区字段 snapshot、badge、checked/partial 草稿、搜索/全选及真实 action 适配 |
| [pivot-chart-field-filter-condition.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-field-filter-condition.ts) | typed 日期/caption 草稿、共享 normalizer、1900/1904/兼容日与日历输入适配 |
| [use-pivot-chart-field-buttons.ts](../../packages/office-studio/src/features/pivot/views/use-pivot-chart-field-buttons.ts) | owner/session/generation、异步成员/条件读取、提交、事件/history 失效、IME/焦点 |
| [PivotChartFieldButtons.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldButtons.vue) | 四区真实实例、同区全部溢出、设置与 busy 接线 |
| [PivotChartFieldFilterPopup.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldFilterPopup.vue) | 全 Ant 成员/日期/caption 模式、排序/clear/分页与 readonly 浏览 |
| [PivotChartFieldConditionEditor.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldConditionEditor.vue) | 全 Ant calendar/caption 编辑与已有 day60 Input 桥，无嵌套 Form |
| [PivotChartPanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue) | initialFieldId、既有字段卡片定位与焦点，不重计前批面板功能 |
| [数值 fixture](../../packages/office-studio/tests/fixtures/pivot-chart-field-buttons-fixture.mjs) | 真实稀疏源、owned/linked、重复值、130 成员、typed/date partial、epoch/day60、权限 |
| [工作台 fixture](../../packages/office-studio/tests/fixtures/pivot-chart-field-buttons-workbench.html) | 真 ExcelWorkbench/ChartOverlay，真实命令/data/history、可释放的原 Facade gates |
| [合同草稿](../../packages/office-studio/tests/pivot-chart-field-buttons.contract.test.mjs) | 真实算法/Facade/controller 的字段身份、过滤条件、历史/重开、权限/异步/IME/focus |
| [浏览器草稿](../../packages/office-studio/tests/pivot-chart-field-buttons.browser.test.mjs) | 真 workbench classic/grid、Ant 控件、四区/溢出、过滤条件、历史/重开、pending、主题/合成 IME/焦点 |
| 本证据 | 来源、分类、真实接线、计数、差异与未验收项 |

下列为人工源码清点，未执行计数脚本；内部断言循环不额外计场景。

| 本批草稿 | test/test.each 声明 | 参数展开后计划场景 | 状态 |
| --- | ---: | ---: | --- |
| 合同 | 34 | 93 | 未运行 |
| 浏览器 | 28 | 85 | 未运行 |
| 合计 | 62 | 178 | 全未运行 |

合同包含 9 类 badge、6 malformed state、8 date/14 caption 操作符、2 epoch/2 whole-day、5 无效日历、owned/linked、2 partial 来源、owner 变化和真实 controller 旧 finally。浏览器包含 classic/grid、2 partial 来源、owned/linked × 4 设置字段、3 可读保护/3 不可读来源、4 controls 组合、8 date/14 caption 条件、3 提交类型 × source/history/user/host 等。数值基线四条记录无 totals 合计100；Region North40/South60，typed partial numeric1 保留10，明确检查文本/数值同显示键1后总计30；year2026 partial 保留10，明确全组选择后100；130 成员合计8515，取消 Region129 后8385。

本轮**未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、浏览器、基准、Office 文件样本、git diff --check 或文档链接检查**。浏览器草稿显式禁用 navigator.gpu，只计划 Canvas2D；DatePicker fill+Tab、焦点和 Escape 仍待真实运行。fixture 的 Object.defineProperty 用户变化仅供最终 context guard 草稿，不证明生产用户状态的响应式 UI。JSON 重开草稿不等于 XLSX/pivot cache/chart 保真往返。

继续未验收：上游精确四区位置、[已接线的 Values 独立 UI](SHEET-23-24-pivot-values-position-ui.md)、新建虚构 day60、日期初值差异、原生 IME、两种渲染后端、Office 保真样本、用户切换 UI、真实大样本性能/内存。普通 PivotTablePanel 的分组/条件/allowMultipleFilters 已接入并另记[独立字段筛选批](SHEET-23-pivot-table-field-filter-ui.md)，全部草稿未运行；不重计前批创建面板43/70、Show As UI55/165、共享引擎34/65、Facade13/24或配对 UI28/65。

SHEET-24、Sheets/Docs 和全部 Pro 功能仍在完整交付范围，本批未验收项不改写为最终范围缩减。
