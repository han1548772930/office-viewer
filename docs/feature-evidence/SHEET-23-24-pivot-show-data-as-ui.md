# SHEET-23/24 值显示方式 UI

日期：2026-11-26。状态：生产接线与测试草稿已落盘，**未验收**。本批只记录普通透视表与 owned/linked 透视图的值显示方式 UI；共享计算见 [引擎证据](SHEET-23-24-pivot-show-data-as.md)，命令、Facade、关联图表与格式投影见 [API 证据](SHEET-23-24-pivot-show-data-as-facade.md)。不把前批 [创建与四字段面板](SHEET-24-pivot-chart-ui.md) 的 43 条声明 / 70 个计划场景计为本批新增。

## 固定来源与入口

固定 OSS 为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 证据来自项目内已保留的 1.6.0 发布包，未安装商业 SDK，也未执行发布包。

| 来源 | 已读取的具体契约 |
| --- | --- |
| [set-pivot-show-data-as.operation.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/types/commands/operations/set-pivot-show-data-as.operation.d.ts) | `requiresParameters`、none/item/item-or-relative、parentFieldOnly、report/local 基准字段筛选；完整成员 DTO 包含 items/itemTypes/itemLabels/format |
| [sheets-pivot-ui ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/es/index.js) | `li` 的固定 15 项；`ui` 的 `sheet.operation.set-pivot-show-data-as`；`di/hi` 字段与参数校验；`Ui/Wi/Gi/qi/Xi/Zi` 的菜单根、逐项 factory、activated 和注册顺序 |
| [PivotShowDataAsDialog.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/types/views/components/PivotShowDataAsDialog.d.ts) | 真实 report/field 身份、基准项 DTO、异步 onConfirm 和取消；发布 ES 的 400px 参数对话框确认前重新读取成员、显式 busy 与 finally |
| [PivotChartFieldView.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/PivotChartFieldView.d.ts)、[字段 helper](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/pivot-chart-field-view.d.ts) | 图表字段面板及已配置字段区身份；仅这些类型不证明固定版每张值卡片直接展示完整显示方式控件 |
| [图表字段 DTO](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts)、[图表 Facade](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts) | showDataAs 的字段设置与 get/set/reset API；本地整份字段草稿通过同一 configure 提交 |
| [Microsoft 公开值显示说明](https://support.microsoft.com/en-us/office/show-different-calculations-in-pivottable-value-fields-014d2777-baaf-480b-a32b-98431f48bfec) | 沿用前批记录的公开公式与参数行为；本批没有重新访问网页或用 Office 文件核实边缘效果 |

本批上游入口清点与本地差异：

| 分类 | 上游数量、位置与状态 | 本地对应 |
| --- | --- | --- |
| 常驻 Ribbon 新根 | 0 | 本批新增顶部根 0，classic/grid 原有顺序不增加显示方式按钮 |
| Pro 上下文根 | 1：`sheet.pivot.context-menu.show-data-as`，MAIN_AREA / DATA，order 8；`fi` 无有效值字段上下文时 hidden | 普通值字段操作菜单新增 1 个「值显示方式」子项；不声称 Canvas 右键根已接线 |
| 固定子项 | 15：`sheet.pivot.context-menu.show-data-as.${type}`，order 1–15，调用同一个 operation，并以当前规则判 activated | 统一 Ant Select 的 15 项方式；7 项无参数也通过确认提交，区别于上游无参数项立即执行 |
| 参数 operation / dialog | 各 1；参数方式有真实字段、成员和异步确认 | 普通独立「值显示方式」Modal 或已有「字段设置」组合提交；宽 460，与上游 400px 有布局差异 |
| 透视图值卡片 | 按已配置 value 实例动态出现，N，不另算常驻按钮 | 每张值卡片嵌入同一个字段控件；字段/外观通过一次 configure |
| 扩展溢出入口 | 本批新增 0 | 未运行顶部 overflow 或完整运行时入口清点 |

15 项按参数分为：7 项无参数（normal、总计/行/列百分比、父行/父列百分比、index）；5 项仅基准字段（指定父项、累计、累计百分比、升序/降序排名）；1 项具体成员百分比；2 项具体或 previous/next 差值。这里是模式/入口分类，不是功能通过数。历史 `percentOfParentGrandTotal` 只显示已保留兼容规则的说明，默认选项不提供；其语义未核实且不计入 15 项。

## 真实模型、命令与保存接线

模块归属为 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的 `SHEET-23 / SHEET-24 → SHEET-22`。Vue/antdv-next 管理草稿与输入；规则 helper 复用唯一的 `normalizePivotShowDataAs`；现有 Pivot 引擎负责真实来源成员、分组和计算。没有在组件、App 或 store 新建聚合算法。

| 本批行为 | 实际路径与边界 |
| --- | --- |
| 方式与参数编辑 | `PivotShowDataAsFields` → 纯草稿 helper；切换方式保留仍合法的 local 基准字段，清理旧成员/relative 参数；换字段也清理旧基准项 |
| 指定父项 | 只列有真实后代的行/列字段；叶字段不进入选项，不以同名标签猜身份 |
| 真实成员 | 复用 `getPivotShowDataAsBaseItems`，普通已提交结构通过 Facade，未提交结构/图表通过 detached draft report；保留原分组后成员，禁止 raw label 替代 |
| 普通已提交结构 | 单独方式 → `FPivotTable.setShowDataAs`；字段名称/汇总/格式/方式组合 → `setFieldSetting`；带 host、expectedRevision、expectedHistoryGeneration、expectedUserId |
| 普通未提交结构/目标 | 设置先留在当前 `PivotPanelDraft`，主「应用字段配置」 → `sheet.command.configure-pivot-table`；带 sourceFieldId/sourceColumn/local ID/showDataAs 和 hostSheetId，完整校验后提交 |
| 普通身份与保留 | 重读按来源列保留每个已有 local 实例、重复值字段、分组和 cache 身份；配置保留旧 filter/group 等非本批设置，不用数组下标重造字段身份 |
| 配置预检/no-op | detached report 先经同一 setter 校验真实 grouped scalar 成员和歧义，再修改输出；source/anchor/fields/options 归一等价时返回原报告，不清 cells、不增 revision |
| 透视图 | 每个 value 实例写到 defensive fieldsConfig；方式有效后与真实作者外观一次 `FSheetPivotChart.configure({patch,fieldsConfig},context)`；owned 保持 data-only，linked 沿用同一 PivotTable |
| 移出值区/重置 | 移出 value 清本实例规则；普通草稿用 null sentinel 防止移出再移回时复活旧规则；normal/null 删除权威规则，不删除显式作者格式 |
| 预览状态 | 草稿成员 report 保留 source、字段分组、options、dateGroupingMode 与 defensive collapseInfo；显式 null 清 collapseInfo；不注册报告、不写预览 cells。轴展开/折叠 UI 不在本批 |
| 历史与重开 | 内容沿用文档隔离的真实 history 与 workbook snapshot；规则持久化在字段模型。本批浏览器草稿包含 undo/redo/JSON 重开，真实 XLSX/Office 往返未验收 |

UI 基准字段始终传 report-local ID；重名标签额外显示行/列和 ID。普通配置可保留已有同源多个值实例，当前两种面板都没有增加专门的「同源新值实例」作者入口。

上游基准成员 DTO 可携带 scalar 类型、显示标签和格式；当前 helper 返回 string keys。本批没有宣称已补齐完整 DTO：空白 `''` 成员可显示但禁用，并解释协议缺口；真实 numeric/string key 碰撞由内容引擎拒绝。完整空白与格式化/分组成员协议继续在交付范围内。

## 生命周期、权限与 Ant 控件

- 普通新增 field settings session 捕获 draft、Studio、unit、host、user、revision、history generation、pivot/local field ID 和结构签名。成员读取前后检查会话；提交捕获完整上下文；旧 finally 只结束所属 request，不清新设置的 busy。
- 共享成员控件对字段/方式/context/retry 使用 generation，卸载后不发布；过期成员不会覆盖新字段，权限丢失清空可见选项。没有因值显示草稿或显示名称变化反复读取成员。
- 新设置检查 source/output read/edit 和真实宿主；已提交报告要求在 anchor 所在工作表执行字段 setter。普通 legacy 标签筛选/分组等其他异步流程仍有原有保护缺口，本记录不宣称整个普通面板生命周期完成。
- 两种面板仍复用 owner/current guard 与内容边界校验。浏览器夹具的用户变化通过 `Object.defineProperty` 模拟，计划验证提交时的真实上下文拒绝；它不能证明生产用户切换的响应式禁用态。
- 输入、选择、按钮、反馈、Modal、Form/FormItem、RadioGroup 等可替换控件均使用 antdv-next；共享控件使用 Ant Select/Alert/Button。样式通过统一 ConfigProvider、theme token 与 CSS vars，未新增其他组件库或原生表单桥。
- composition capture 期间禁用提交、导航和关闭；compositionend 等 nextTick 接收最后文字。测试草稿仅合成事件，不证明 Windows 原生 IME。
- 夹具含真实 ExcelWorkbench/Canvas 背景。关闭面板的测试目标是恢复其 Ant 启动按钮，嵌套目标是 Ant 字段操作按钮；不能将夹具结果称为 production Ribbon/Canvas 焦点已验证。
- 前批 Context7 查询已遇到 Monthly quota reached，继续静态对照已有 antdv-next 1.8.2 类型/实现；本批没有据此增加 SDK、组件库或额外原生交互控件。

## 本批文件与接口

本批 **12 个文件**（7 个生产文件、4 个新测试/夹具文件和本证据）。以下生产文件只计本批 Show As 片段；共享 API/core 与其他代理的 collapse/创建改动不归本批。

| 文件 | 本批职责 |
| --- | --- |
| [pivot-show-data-as-ui.ts](../../packages/office-studio/src/features/pivot/services/pivot-show-data-as-ui.ts) | 15 项、local 字段选项、纯参数草稿、唯一规则校验与成员结构签名 |
| [PivotShowDataAsFields.vue](../../packages/office-studio/src/features/pivot/views/PivotShowDataAsFields.vue) | 两种面板复用的全 Ant 参数/成员/reset 控件、generation 和 validity |
| [pivot-chart-field-draft.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-field-draft.ts) | showDataAs 草稿 patch、移区清理、真实 draft report 与 collapse 保留 |
| [PivotChartFieldsForm.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldsForm.vue) | 每个 value 字段接控件并传参数/有效性 |
| [PivotChartPanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue) | 成员、source/session 校验、无效草稿禁用与一次 configure |
| [pivot-panel-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-panel-state.ts) | local/source/column 身份、重复实例保留、纯草稿 report、参数预检与归一 no-op |
| [PivotTablePanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotTablePanel.vue) | 普通菜单、独立/组合设置、草稿配置、上下文与新设置生命周期、主题/IME |
| [UI 数值夹具](../../packages/office-studio/tests/fixtures/pivot-show-data-as-ui-fixture.mjs) | 复用前批真实 Region→Product / Year→Quarter 原始数据；grouped/repeated/blank/compatibility 与保护样本 |
| [UI 工作台夹具](../../packages/office-studio/tests/fixtures/pivot-show-data-as-ui-workbench.html) | 真 Ant/Canvas/面板/Facade/history；gate 释放后调用原方法，未用假成功结果 |
| [UI 合同草稿](../../packages/office-studio/tests/pivot-show-data-as-ui.contract.test.mjs) | 参数、身份、分组、真实数值、隔离、reset/no-op、invalid preflight、JSON 与 collapse 保留 |
| [UI 浏览器草稿](../../packages/office-studio/tests/pivot-show-data-as-ui.browser.test.mjs) | 控件到真实命令/数值、普通/owned/linked、相邻项、组合/配置、竞争、权限、history/JSON、主题/IME/focus |
| 本证据 | 来源、计数、差异与未验收边界 |

新增共享控件 props 为 value/report/loadBaseItems/contextKey/readable/disabled/labelPrefix，events 为 update:value 与 validity-change。`PivotChartFieldsForm` 新 show-as-validity(id,valid)；`PivotChartDraftFieldSetting` 接受 showDataAs 或显式 null。普通 configure DTO 新 sourceFieldId/sourceColumn/showDataAs 与上下文字段；底层 Facade/commands 由主任务维护，不在这里重复计文件/草稿。

## 静态草稿计数与未验收

| 本批文件 | test/test.each 声明 | 参数展开后计划场景 | 状态 |
| --- | ---: | ---: | --- |
| UI 合同 | 27 | 51 | 未运行 |
| UI 浏览器 | 28 | 114 | 未运行 |
| 新增合计 | 55 | 165 | 未运行，不能作为通过数或覆盖完成率 |

合同展开为固定 15 方式、7 无参数模式、3 item/previous/next 和 3 collapse projection 状态：`27 - 4 + 15 + 7 + 3 + 3 = 51`。浏览器 15 方式 × 3 面板/所有权 = 45；相邻比较 2 模式 × 2 相邻类型 × 3 = 12；pending 4 变化 × 3 = 12；grouped/reset/移区/history/compatibility/旧提交各 3；两类 IME 各 2；主题 3 × 2 = 6；输出保护 2；其他独立声明各 1，总计 114。内部断言循环和重复读取不额外计场景。

原始样本为 `C2:G10` 的五列、八条真实记录；North/A、2024/Q1 的数值预期来自聚合：normal=10、grand=.05、row=1/3、column=1/7、parentRow=1/4、parentColumn=1/3、index=20/21、指定 Region 父项=1/4、Product B 百分比=1/3、差值=-20、差值百分比=-2/3、Product 累计=10、累计百分比=1/4、升序排名=1、降序排名=2。previous 首项为当前实现的 #N/A，Office 边缘政策仍待样本核实。AVG 测试使用原始总平均 25，单格 10 的总计百分比为 .4，不把可见聚合值求和作为分母。

按用户当前「先实现、后统一测试」约定，本轮**未运行测试、类型检查、前端构建、Rust/Cargo/WASM、dev、浏览器、Office 文件样本、性能基准或 git diff --check**。只做静态读写、声明/参数清点及来源/链接核对。夹具计划禁用 navigator.gpu 走 Canvas2D，不能据此声称 Canvas/WebGPU 已验收。

继续保留：完整成员类型/格式化标签/空白 DTO、Canvas 右键显示方式根、上游移动端和快捷键、更多 Office 边缘样本、真实 Ant/Windows IME/焦点/用户切换验证、XLSX pivot/cache/chart 原生保真往返。区域配对创建表+图与 Canvas 轴交互进入后续独立批次，不计为本批完成。

本批未新增 Worker/WASM 协议或性能优化。大样本计算仍须按 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 完成 Rust/Worker 路径、取消/内存与同数据集基准。SHEET-23/24、Docs/Sheets 与所有 Pro 能力继续保持完整交付范围和「未验收」。
