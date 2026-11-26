# SHEET-23 / SHEET-24：ΣValues 位置与索引 UI

记录日期：2026-11-26。状态：**生产接线和测试草稿已落盘，未验收**。

本批将共享 position/index 契约接入普通透视表面板、owned/linked 透视图面板及真实图表覆盖层。按用户要求先实现、最后统一验证；本轮没有运行测试、类型、构建、Rust/WASM、dev、浏览器、基准、Office 文件样本、计数脚本、`git diff --check` 或文档链接检查。下面的声明及场景数量为人工源码清点，不是通过数量。

## 固定来源与入口分类

固定参考为 `E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 公开发布文本只读自项目 `.research.local/univer-1.6.0/packages/`，没有安装或执行商业 SDK。

| 来源与符号 | 已读行为、实现差异及未核实项 |
| --- | --- |
| `engine-pivot/lib/types/types/enum.d.ts` 的 `PivotTableValuePositionEnum`，`types/json-types.d.ts` 的 `valuePosition/valueIndex` | None=-1、Column=1、Row=2；位置与索引独立保存。本地领域使用字符串，UI 读取共享 raw/effective 状态，不另存一套权威轴配置 |
| `sheets-pivot/lib/types/facade/f-pivot-table.d.ts` 第 593–613 行，`sheets-pivot-chart/lib/types/types.d.ts` 第 117–119 行与 `facade/f-pivot-chart.d.ts` 第 219–226 行 | 更新同时传 position/index；本地普通和图表 Facade 复用各自真实内容命令及 captured context。图表旧第二参数 context 继续兼容 |
| `sheets-pivot-chart-ui/lib/es/index.js` 的 `br(e,t)` | 多于一个 value 且所选 Row/Column 时，按 clamp 后 valueIndex 插入特殊字段。本地独立 union marker 只占 UI/layout slot，不注册 fake source field |
| 同一 UI ES 的 `Or` | 特殊字段使用 multipleValues 标题、不可当普通字段移除。本地提供行/列/None、上下移动和插入索引；None 隐藏 marker、保留全部 measures |
| 同一 UI ES 的 `Ar` 与 special ID `c` 拖动分支 | special marker 接受 destination.index，真实维度跨 marker 移动调整 index。本地维度仍保存纯 local position，草稿调整复用共享 gap helper |
| 同一 UI ES 的 `wr`；`sheets-pivot-ui/lib/es/index.js` 对应字段列表 | 四区及特殊字段是上下文工具。上游精确四区像素、键盘拖动及菜单动画尚未核实；本地延用现有四区装配、Ant 菜单和 HTML5 drag |

本批没有新增顶部 Ribbon 根入口。完整顶部及扩展清点仍见[固定 Sheets 来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)与[透视图字段按钮清点](SHEET-24-pivot-chart-field-buttons-ui.md)。以下只统计本独立上下文 UI 的入口种类，不能当作功能完成数；classic/grid 是同一集合的两种布局。

| 分类 | 数量口径、显隐与位置 |
| --- | --- |
| 常驻面板设置 | 普通表和图表各 2 个 Form/Select：方向、插入位置；方向各 3 项，索引为自动 + 当前轴合法 gap + 已有超界 raw 值 |
| 上下文字段 marker | 每个多 measure 且非 None 的报告 1 个 ΣValues，插在所选 row/column 轴的 effective gap；0/1 measure 或 None 时隐藏 |
| marker 菜单 | 3 个方向、上移/下移/自动各 1、索引 submenu 1；索引子项按当前轴动态生成。权限、busy、composition、边界决定禁用 |
| 图表覆盖层溢出 | marker 和真实字段共用所在区域布局 slot；溢出后仍可打开同一位置菜单，外层区域菜单保持打开 |
| 拖拽 | marker 只允许 row/column 目标；filter/value 非法 drop 不提交。真实字段与 marker 的 identity 分离 |

## 模块、事务与实际效果

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，UI 投影属于 `features/pivot/services`，事件和临时会话属于 `views`；轴算法、校验、raw index 和 fallback 仍由[共享核心批](SHEET-23-24-pivot-values-position.md)负责。没有新增 Worker/WASM DTO、解析/写出逻辑或第二个权威透视引擎。

- `readPivotValuesUiState` 消费共享 getter/layout。省略或 -1 为自动末尾；非负 raw index 原样展示和保存，effective index 仅对当前轴 clamp。方向切换保留 raw index，避免将旧轴的 clamp 写回新轴。
- `PivotValuesPositionFields` 提供行、列、None 和索引 Select。None 时禁用索引，并提示汇总字段按列显示；不删除 measures。已有超界 raw 值有独立可读选项，用户仍可切回合法位置或自动末尾。
- `PivotValuesFieldCard` 使用独立 `kind:'values'` 和 `data-pivot-synthetic='values'`。真实源字段名称或 ID 为 Values 时仍是 authored 字段，不混同于 marker。
- 普通 `PivotTablePanel` 改主草稿；字段 toggle/remove/move/drop 与来源退役调用共享 gap helper。位置更改和现有字段、过滤、输出地址由一次 `sheet.command.configure-pivot-table` 提交，面板没有直接写报告。
- `PivotChartFieldsForm` 插入同一 marker，`PivotChartPanel` 的 setter 同步 top-level `valuePosition/valueIndex` aliases 与 options；两处冲突的外部草稿拒绝。其他 chart appearance options 继续走原路径。保存使用一次 `sheet.command.configure-pivot-chart`。
- 真实 `PivotChartFieldButtons` 的 direct/overflow 菜单和 drag 通过 `FSheetPivotChart.updateValuePosition(position,index,context)`，进入 `sheet.command.edit-pivot-chart-report` 的真实 `updateValuePosition` edit。普通独立 Facade 更新使用 `sheet.command.update-pivot-value-position`。
- 覆盖层 snapshot 单独返回 values 状态，不把 marker 合入 authored field IDs。`values:position`、`values:overflow:<area>` 和 `values:drag` 各自标识 session；普通 field/overflow 会话不能提交 Values。
- owner/user/source/history/host 校验、generation、publishing、自己提交后的 revision+1、旧 finally 和忙碌态沿用真实 controller。旧菜单关闭与 dragend 只清理自己的 anchor，不能取消新会话。关闭/成功仅向仍当前的 Canvas owner 恢复焦点；composition 期间拒绝内容和导航动作。

所有新增可替换控件使用 `antdv-next` 的 FormItem/Select、Dropdown/Button/Tooltip/Alert，继承工作台 ConfigProvider 与 theme token；wrapper 的原生 HTML5 drag 只承担编辑器内核的拖拽事件。内容继续 Canvas 绘制，没有新增原生输入控件。

单 measure 的旧本地 row 仍保留 Values header，但 UI marker 隐藏；这是保留的明确差异，未宣称完成上游单值迁移。None 的 column fallback、subtotal/collapse/grand-total、Show Values As、数字格式和输出 matrix 仍来自共享计算器。

## 本批文件与未运行草稿

本批 **18 个文件**：10 个生产、3 个 fixture、2 个测试、1 个新证据及 2 个旧证据更新。旧批和共享核心批不重计。

| 文件 | 本批变化 |
| --- | --- |
| [pivot-values-position-ui.ts](../../packages/office-studio/src/features/pivot/services/pivot-values-position-ui.ts) | raw/effective UI 投影、索引选项、alias 同步、普通/图表草稿 gap adapter |
| [PivotValuesFieldCard.vue](../../packages/office-studio/src/features/pivot/views/PivotValuesFieldCard.vue) | 共享 marker、方向/移动/索引菜单、dense 与拖拽事件 |
| [PivotValuesPositionFields.vue](../../packages/office-studio/src/features/pivot/views/PivotValuesPositionFields.vue) | Ant 方向与索引表单 |
| [PivotTablePanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotTablePanel.vue) | 普通主草稿 marker、gap、owner drag 与单次配置 |
| [PivotChartPanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue) | options/公开 aliases 同步 |
| [PivotChartFieldsForm.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldsForm.vue) | 图表四区 marker、表单和 captured draft drag |
| [pivot-chart-field-draft.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-field-draft.ts) | 字段增删/移动调用共享 gap adapter |
| [pivot-chart-field-button-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-field-button-state.ts) | values snapshot 与真实字段分离 |
| [use-pivot-chart-field-buttons.ts](../../packages/office-studio/src/features/pivot/views/use-pivot-chart-field-buttons.ts) | guarded Values 会话/命令/焦点 |
| [PivotChartFieldButtons.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldButtons.vue) | direct/overflow marker、原生 drag、Ant 错误反馈 |
| [pivot-values-ui-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-values-ui-fixture.mjs) | 真实 sparse source 的双 measures、两行/两列维度配置 |
| [pivot-table-field-filter-workbench.html](../../packages/office-studio/tests/fixtures/pivot-table-field-filter-workbench.html) | 真 Ribbon→工作台→面板，valuesUi 配置/状态/编辑 gate |
| [pivot-chart-field-buttons-workbench.html](../../packages/office-studio/tests/fixtures/pivot-chart-field-buttons-workbench.html) | 真 ChartOverlay、values getter 与真实计算 matrix |
| [pivot-values-position-ui.contract.test.mjs](../../packages/office-studio/tests/pivot-values-position-ui.contract.test.mjs) | 20 条声明、85 个计划场景，未运行 |
| [pivot-values-position-ui.browser.test.mjs](../../packages/office-studio/tests/pivot-values-position-ui.browser.test.mjs) | 20 条声明、78 个计划场景，未运行 |
| 本文、[图表字段按钮旧证据](SHEET-24-pivot-chart-field-buttons-ui.md)、[普通字段筛选旧证据](SHEET-23-pivot-table-field-filter-ui.md) | 本独立批与旧批的进度链接；不增加旧批计数 |

合计 **40 条测试声明、163 个计划场景，全部未运行**。共享核心的 15/83 属于独立核心证据；前批图表字段按钮 62/178、普通字段筛选 59/175 仍为原批计数。

fixture 复用真实 `D3:H7` source（Region/Amount/Date/Segment/Score）与报告 A25，SUM/AVG 两种 Amount measures，Region/Date 两行维度、Segment/Score 两列维度。不是 fake Values field，也没有手挂第二套覆盖层。JSON 重开草稿只覆盖当前快照，不是 XLSX 验收。

合同草稿计划覆盖三方向×五种 raw index、0/1/2/3 measures、真实字段同名 identity、alias 冲突/非法 index、普通一次配置与 owned/linked 一次配置、None/no-op/raw99 重开、table/chart gap/来源退役、真实 snapshot/Facade/焦点、两 owners×四类 pending 变化、旧 finally、只读/不可读、composition/dispose。

浏览器草稿计划覆盖 classic/grid 普通面板两轴×四索引、None/raw99/原生拖拽/非法 drop/字段 gap，owned/linked 图表面板与真实 overlay 的单次内容命令、轴 matrix、single measure legacy、溢出菜单、保护/不可读、pending/旧 finally、明暗主题、Ant composition、Escape→Canvas 焦点和卸载。计划 msedge headless、1440×1200、禁用 navigator.gpu 使用 Canvas2D；不是已观测环境。

## 未验收边界

真实 Ant Select/Dropdown/submenu、关闭动画、HTML5 drag、Escape、Windows IME 和焦点恢复未运行；fixture 的 Object.defineProperty 用户变化只用于最终 context guard 草稿，不能证明生产用户状态的响应式 UI。上游精确四区位置、原生 Office position/index 行为、完整 pivot cache/chart parts 与 XLSX 导出重开、两种渲染后端、真实大样本 Worker/Rust/WASM/取消/传输与内存仍待验收。

本批不声称性能优化、零拷贝或 Office 保真达标；继续遵守 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。Sheets / Docs 及全部 Pro 仍在完整交付范围，未验收项不改写为范围缩减。
