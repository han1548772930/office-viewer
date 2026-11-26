# SHEET-23 普通透视表字段筛选与条件 UI

日期：2026-11-26。状态：生产接线、真实工作台 fixture 和测试草稿已落盘，**未验收**。本批替换普通 PivotTablePanel 的旧字符串 checklist，接入真实 grouped checked/partial、日期、标签条件、排序及 allowMultipleFilters 配置。声明数与计划场景数均不是通过数，不据此宣称功能已验收。

## 固定来源与入口清点

固定 OSS 为 `E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 来源为项目内同版本只读发布包，未安装或执行商业 SDK。

| 来源 | 可观察契约、本地对应与差异 |
| --- | --- |
| [sheets-pivot-ui ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/es/index.js) 的 `Zo` | 字段会话、checkedItems/partialItems、all/clear、排序与日期入口；本地普通表使用真实报告字段 ID 和 FPivotTable catalog，复用图表已有成员草稿与条件 UI |
| 同文件 `Xo`、`Ko` | 8 日期条件，dateEqual 默认、useWholeDay=false、typed numeric serial；本地新条件日期留空，要求明确输入。上游转换不表示 day0/虚构 day60，本地对既有 day60 使用 Ant Input 桥保留 |
| 同文件 allowMultipleFilters/setOptions 路径 | 上游开关即时提交 options；本地主配置开关先保存在草稿，应用字段配置时一次 configure 提交。关闭多筛选在同一共享事务中清理多类别，无第二条 options 内容事务 |
| [共享谓词](../../packages/office-studio/src/features/pivot/services/pivot-filter-predicate.ts)、[公开适配](../../packages/office-studio/src/features/pivot/services/pivot-filter-api.ts) | 8 日期/14 caption 条件、类别协调、typed raw membership、1900/1904 语义来自共享内核，UI 没有另建谓词或数字代码映射 |

| 分类 | 可复查的本批入口与条件 |
| --- | --- |
| 常驻 Ribbon/Pro 根 | 新增 0，沿用数据 Tab 的数据透视表入口与既有表选择 |
| 字段上下文根 | 筛选字段、值筛选、分组 3 项；字段必须来自相同真实 source/local identity。成员筛选允许只读浏览，内容工具要求编辑权限 |
| 成员与条件 | 成员模式；日期 capability 根 1、标签条件根 1，返回成员入口只在条件模式显示；真实类型决定日期显隐，重命名为 Date 不产生日期能力 |
| 排序 | row/column 的升序、降序、清除排序 3 项；filter 区不显示排序，value 字段不进入成员弹层 |
| 内容与会话动作 | 确认、清除 2 个内容动作，取消/关闭 2 个会话动作，读取失败时重试 1 个；成员 clear 清所有类别，条件 clear 只清自身类别 |
| 主配置 | allowMultipleFilters Ant Switch 1 个；与已有 source/target/fields/options 一次 configure 提交，弹层期间禁止修改主配置 |
| 动态辅助输入 | 搜索、全选各 1；catalog 80 项分页。每种条件 1 个 Select、1 或 2 个 operand，日期另有 whole-day Checkbox；动态成员数不算功能数量 |
| 既有工具保护 | 值筛选、数值/日期/manual 分组继续使用真实 Facade；新增 owner/context/generation、主题和 IME 保护，不把旧工具重复计为新引擎能力 |

归属为 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的 SHEET-23 → SHEET-22。Vue/antdv-next 管理临时配置、弹层与输入；成员 catalog、过滤、分组、重算、输出和内容历史消费共享模型与命令。内容仍由 Canvas 绘制。

## 实际接线与边界

| 行为 | 本批生产行为 |
| --- | --- |
| 字段身份与草稿 | 比较报告 ID、source sheet/range、字段 ID/sourceColumn/area，不按显示名称猜来源；未提交的 target、displayName、unused fields 与 options 可保留。完整 signature 阻止旧异步会话覆盖新草稿 |
| grouped/partial | await getFieldItems 后同步读取 getManualFilterItemState；partial-only 可确认，原 raw numeric/text/date membership 保留。显式勾选 partial 才选中整个显示组 |
| manual/all/clear | 真实 `sheet.command.set-pivot-filter`；all 覆盖搜索和分页隐藏项，仅清 member 类别；clearFilters 清 manual/record/date/caption/value 类别。全空拒绝，无 UI 直写 report/output cells |
| 日期/caption | 真实 `sheet.command.set-pivot-date-filter`、`sheet.command.set-pivot-label-condition-filter`；capability 包含 date1904，fractional serial 与 whole-day 原样提交；缺失或倒序日期拒绝 |
| sort | 真实 `sheet.command.set-pivot-label-sort`；成功后只将已提交 sort 合入现有主草稿，不重载来源/位置/字段/选项 |
| allowMultipleFilters | `sheet.command.configure-pivot-table` 的共享 setPivotOptions deferred 路径协调类别；true 保留求交集，关闭时在一次 configure 中清多类别，重复 false 配置无新增历史 |
| 其他字段工具 | value/number/date/manual grouping 使用 FPivotTable 内容 setters，透传 expected revision/history/user 与 hostSheetId；成功刷新 owner，保留主草稿 |
| 历史与关联图 | 内容命令复用文档隔离历史与关联 chart 重算。fixture 草稿断言一次内容提交、撤销/重做和 JSON save/reopen；未以 JSON 重开代替 Office 文件验收 |
| 权限/异步 | owner 捕获 workbook/host/user/revision/history generation；晚到成员不能替换新 catalog，旧 finally 不能清新 loading/busy；内容命令最后再校验 context。自己的单次提交允许 revision+1 |
| IME/焦点 | composition 期间保留 Ant 输入并阻止内容、导航、关闭；过期输入结束后使会话失效。Escape 尊重 child defaultPrevented/isComposing/229；取消/成功仅向仍当前的字段菜单按钮恢复焦点 |

涉及 UI 均为 `antdv-next` 的 Modal、Form/FormItem、Input、InputNumber、Select、DatePicker、Checkbox、Switch、Button、Pagination、Dropdown 与 Alert，继承统一 ConfigProvider/theme token。保留既有 Excel day60 的最小 Ant Input 日历桥；正常 DatePicker 不能创建新的虚构 day60。真实 DatePicker fill+Tab、弹层 Escape、Modal 关闭动画后的焦点与原生 Windows IME 尚待运行。

共享组件新增可选接口：Popup 的 subjectLabel/testId/showSort/showClearSort/showManagement，ConditionEditor 的 subjectLabel；默认图表文案和管理入口保留。FieldTools 新增 stale/readable/editable/supportsDateGrouping 及 composition-start/end，限制旧上下文动作，日期能力来自真实 catalog。无新 Worker/WASM 协议、无性能优化声明；仍按 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 验收。

## 本批文件与测试草稿

本批 **10 个文件**：6 个生产、1 个新 HTML fixture、2 个测试草稿及本证据。复用既有数据 mjs fixture，不重复计数；不将主任务的 Facade/API、panel-state、domain、engine、Overlay、chart 投影或顶层清单计入本批。

| 文件 | 职责 |
| --- | --- |
| [pivot-table-field-filter-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-table-field-filter-state.ts) | snapshot/权限、真实 source/local identity、signature、typed Facade action、multiple 草稿和 sort 合并 |
| [use-pivot-table-field-filter.ts](../../packages/office-studio/src/features/pivot/views/use-pivot-table-field-filter.ts) | 会话、异步读取/提交、事件/history 失效、旧 finally、IME/焦点 |
| [PivotTablePanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotTablePanel.vue) | 普通表生产接线、multiple Switch、主草稿保留、真实工具 context、删除旧 checklist 样式 |
| [PivotChartFieldFilterPopup.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldFilterPopup.vue) | 普通/图表共享的完整成员、条件、分页与排序表单，可配置文案/管理显隐 |
| [PivotChartFieldConditionEditor.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldConditionEditor.vue) | 普通表条件 aria、Ant calendar/caption/day60 输入 |
| [PivotFieldToolsDialog.vue](../../packages/office-studio/src/features/pivot/views/PivotFieldToolsDialog.vue) | 已触及工具的主题、readonly/stale/IME 和真实日期能力保护 |
| [工作台 fixture](../../packages/office-studio/tests/fixtures/pivot-table-field-filter-workbench.html) | 真 Ribbon → ExcelWorkbench → PivotTablePanel；未手工重复挂 Panel，gates 释放后调原 Facade |
| [合同草稿](../../packages/office-studio/tests/pivot-table-field-filter-ui.contract.test.mjs) | 真实模型/算法/Facade/controller 草稿 |
| [浏览器草稿](../../packages/office-studio/tests/pivot-table-field-filter-ui.browser.test.mjs) | 真工作台、Ant/Canvas、命令/history、草稿/权限/异步/主题/合成 IME/焦点 |
| 本证据 | 来源、入口分类、实际行为、计数、差异与未验收 |

人工源码清点，未运行计数脚本；内部断言循环不增加场景。

| 草稿 | test/test.each 声明 | 参数展开后计划场景 | 状态 |
| --- | ---: | ---: | --- |
| 合同 | 28 | 87 | 未运行 |
| 浏览器 | 31 | 88 | 未运行 |
| 合计 | 59 | 175 | 全未运行 |

合同含真实身份/来源变动、6 类 signature、2 partial 来源、manual/clear、8 日期/14 caption、2 epoch/day60、3 无效条件、3 sort、2 可读/2 不可读保护、multiple bool/单 configure/no-op、4 owner 的晚到读取、4 提交类型 × 4 owner、旧 finally、readonly/IME/focus/dispose。浏览器含 classic/grid、主草稿保留、partial/真实分组/value filter、multiple/category clear/no-op、相同条件/epoch/day60、130 成员分页、4 owner 读取/4 类型提交/2 工具提交、主题/合成 IME/焦点/Escape/dispose。

样本为 source D3:H7、primary A25，4 条记录总100，North40/South60。typed partial numeric1 保留10，明确勾整组后30；year2026 partial 保留10，明确整组后100。130 成员总8515，取消 Region129 后8385。whole-day 草稿修改 F4/F5 两条数据记录，保留 F3 表头；本次同步修正前批图表浏览器草稿的相同地址错误，前批声明数不变，不将其重复计入本批。

本轮**未运行测试、类型、前端构建、Rust/WASM、dev、浏览器、基准、Office 文件样本、git diff --check 或文档链接检查**。浏览器草稿计划 msedge headless、1440×1200，并禁用 navigator.gpu 使用 Canvas2D；不是已观测环境。Object.defineProperty 用户切换仅验证最终 context guard 的草稿，不证明生产用户状态的响应式 UI。源码回读不等于运行验证。

继续未验收：真实 DatePicker/关闭动画焦点/Windows IME、生产用户切换 UI、WebGPU/Canvas2D 两后端、Office pivot cache/chart 保真往返、大样本性能/Worker/WASM/内存。ΣValues 位置与 index 已在[独立 UI 批](SHEET-23-24-pivot-values-position-ui.md)接线、未验收；不重计本字段筛选批或图表字段按钮62/178、前批创建43/70、Show As UI55/165、共享 Show As34/65、Facade13/24或配对 UI28/65。Sheets/Docs 及全部 Pro 仍在完整交付范围，未验收不改写为范围缩减。
