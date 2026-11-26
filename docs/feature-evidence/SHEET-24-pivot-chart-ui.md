# SHEET-24 透视图创建与字段面板

日期：2026-11-26。状态：本批实现和测试草稿已落盘，**未验收**。本记录只覆盖创建、选择、四字段区、报告选项、标签筛选以及共享图表设置的 UI 接线。后续「值显示方式」单独记录，不混入本批文件和草稿计数。

## 固定来源与清点

OSS 参考：`E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 参考为项目内已保留的发布包，未安装商业 SDK：

| 来源 | 可复查依据 |
| --- | --- |
| [package.json](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/package.json) | `@univerjs-pro/sheets-pivot-chart-ui` 版本 1.6.0 |
| [menu/schema.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/menu/schema.d.ts)、[views/menu.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/menu.d.ts) | button/selector/category factories 与注册 schema |
| [发布 ES](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/es/index.js) | `tr/nr` factories、`sr` schema、`G` 创建 command、`W/Qn/ur` 面板与浮层 operations、`or` 轴上下文菜单；只静态读取，不执行发布包 |
| [创建命令类型](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/commands/commands/menu-insert-pivot-chart.command.d.ts) | `mode: withPivotTable / chartOnly` |
| [PivotChartPanel](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/PivotChartPanel.d.ts)、[PivotChartFieldView](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/PivotChartFieldView.d.ts)、[字段 helper](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/pivot-chart-field-view.d.ts) | 字段区、来源字段、已有实例与默认区域判断 |
| [PivotChartFilterPopup](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/PivotChartFilterPopup.d.ts)、[PivotChartFieldListPopup](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/PivotChartFieldListPopup.d.ts) | 原生图表命中后打开字段列表和成员筛选的入口 |
| [PivotChartConfigView](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/PivotChartConfigView.d.ts) | 45 项图表类型目录与共享作者设置视图；目录数不能证明所有类型验收 |
| [PivotChart types](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts)、[Facade](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts) | range/linked 来源 union、initialFieldsConfig、报告字段及外观接口 |

本包上游入口按类型清点，不能与 OSS 常规根、全部 Pro 根或运行时可见总数混计：

| 分类 | 数量 | 含义 |
| --- | --- | --- |
| 常驻 OSS 根 | 0 | 本记录是 Pro 扩展包清点，不重计已存在的 OSS 入口 |
| Pro Ribbon 根声明 | 2 | ANALYSIS 中的两条互斥声明；满足条件时二者只显示一条 |
| 创建 selector 子项 | 2 | `withPivotTable`、`chartOnly`，都调用同一创建 command |
| 轴上下文菜单根声明 | 1 | `sheet.pivot-chart.expand-collapse`，属于独立 axis context menu 的 DATA 分组 |
| 轴上下文固定子项 | 4 | 展开当前成员、折叠当前成员、展开字段、折叠字段 |
| 轴层级动态子项 | N | 多于一个字段时按 fields 生成 expandTo/collapseTo，不写死运行数量 |
| 浮层 operation 类型 | 2 | 字段列表、成员筛选，按 Canvas 命中与存活宿主打开 |
| 包专属溢出根 | 0 | 顶部入口可由通用 Ribbon 进入 More；本批没有执行运行时 overflow 清点 |
| 侧面板类型 | 1 | 上游 `sheet.pivot-chart.panel`，宽度 360；字段与图表设置由它组合 |

逐入口注册及本地差异：

| 上游入口 | 标签/分组与位置 | 显隐、禁用与命令 | 本地对应及缺口 |
| --- | --- | --- | --- |
| `sheet.command.menu-insert-pivot-chart` | 透视图；DATA/ANALYSIS；order 1；grid r2c1、showLabel | 当前选区处于透视表内时显示；工作簿编辑、工作表查看/透视权限、范围查看和 exclusive range 判断；同名 command | 本地仍复用一个 `sheet.operation.open-pivot-chart` 根，在面板内选择来源；已有报告可创建 linked chart。未复刻两根互斥 factory |
| `sheet.pivot-chart.insert` | 同上，同一个 grid 槽位 | 选区不处于透视表内时显示；相同权限判断；子项传 mode 给 `sheet.command.menu-insert-pivot-chart` | `chartOnly` 的 owned range 行为已接线；「区域同时创建透视表与图」仍未实现，不能把 linked 已有报告等同 withPivotTable |
| `sheet.pivot-chart.expand-collapse` | 轴上下文 DATA 分组，order 0；selector | 需要有效轴 selection；成员 collapsed、field.canExpand/canCollapse 和层级共同决定子项禁用 | 本批没有添加 Canvas 轴命中、成员/字段展开折叠、类别层级操作 |
| `sheet.operation.open-pivot-chart-field-list` | Canvas 字段浮层，不属于常驻 Ribbon | 真实 event/hit bounds、float host、active unit/sheet 和存活字段列表；只修改 UI 会话 | 本地四字段区列表已接线；Canvas 字段按钮和此浮层未接完整闭环 |
| `sheet.operation.open-pivot-chart-filter-panel` | Canvas 成员筛选浮层 | 同类宿主/命中/session guard；筛选内容命令属于 report edit | 本地字段卡片打开 Ant 标签筛选 Modal；完整 Canvas 浮层入口和上游 filter query 仍缺 |
| `sheet.operation.open-pivot-chart-panel` | 上游侧面板打开 operation | 必须存在匹配 unit/subUnit/pivotChart 的图表；双击/通用 Chart panel 导航复用 | 本地作为既有 open operation 的别名；新增/已有透视图使用 860px Ant Modal。侧面板与双击路由效果未核实 |

本地顶部仍是 **1 条既有根声明，新增 0 条**：DATA 中「分析」分组 order 30，classic 在数据透视表之后、切片器之前；grid r2c1、showLabel。本批只解除「必须先有透视表」的 availability 条件，并更新禁用原因。operation 别名不算第二个 Ribbon 按钮。classic/grid 位置的浏览器草稿尚未运行。

## 模块与实际接线

归属为完整工作台架构 `SHEET-24 → SHEET-23/SHEET-22`。Vue/antdv-next 承担面板、草稿和输入，纯 helper 负责草稿编排与 DTO，Facade/Studio 负责内容命令，共享 pivot/chart 服务负责计算、历史和持久化。App、组件和 store 不新增聚合公式。

| 本批行为 | 命令、模型与保存路径 |
| --- | --- |
| 新建 owned chart | 真实单选区域或用户编辑区域 → `source.kind=range` + initialFieldsConfig → `FWorksheet.insertPivotChart` → `sheet.command.add-pivot-chart` → chart.ownedPivot；不插入 PivotTable、虚拟 worksheet 或输出单元格 |
| 新建 linked chart | 选择已存在 PivotTable → unit/subUnit/pivotTableRef + initialFieldsConfig → 同一创建 command → chart.pivotTableId；保留现有报告身份 |
| 选择已有图表 | 只列当前宿主中的 owned/linked chart，按真实 report 读取 config/source fields/options；已存在 chart 的来源归属不可在草稿中切换 |
| 四字段区 | 筛选/filter、轴/row、图例/column、值/value；来源字段用 canonical source ID，配置实例保留 local ID；添加、跨区移动、精确上下排序、删除都先改本地草稿 |
| 字段设置 | 显示名称、数字格式、11 类汇总；标签字段转值时清理不兼容的 filter/group/sort；全部属性由真实 configure 校验，绑定切片器等依赖仍可原子拒绝 |
| 报告选项 | repeatRowLabels、行/列总计、行/列小计共 5 个勾选项，以及 valuePosition 的 row/column 选择；与 chart author options 区分 |
| 标签筛选 | 从真实源列读取成员；multi-select 记录原始字符串键；`[]` 是全不选，`undefined` 是清除标签筛选；成员查询不读取伪造图表输出 cells |
| 字段及外观应用 | `FSheetPivotChart.configure({patch,fieldsConfig}, context)` → `sheet.command.configure-pivot-chart`；一次内容事务；patch 不携带 source/pivotTableId/ownedPivot；保留完整 ChartSettingsFields 的作者外观 |
| 刷新与删除 | `refresh/remove(context)` → 专用 PivotChart command；删除保留来源 cells 和独立 PivotTable，owned report 随所属图表移除；撤销重做由文档历史负责 |
| 保存/重开 | 复用真实 workbook snapshot 保存与 Office 链路；本批只编写内部 snapshot 重开草稿，未证明 XLSX 原生往返 |

字段四区、5 个选项和 11 个汇总属于面板控件/枚举清点，不能并入顶部入口计数，不能充当 Pro 完成率。添加已使用的来源字段沿用移动现有实例的语义；UI 不提供同一来源新增多个独立值实例的专门入口。

来源变化会清掉旧 valueFields/categoryField/seriesToCategory/options.mapping，保留其他外观。选择「新建」和重读使用草稿中的当前真实 range，避免回到最初打开时的选区。

预览读取创建一个**未注册的 data-only report**，带当前 fields/options/dateGroupingMode，去掉 outputRange/generatedCellKeys/drillDownMap，复用 `getChartSettingsData/getChartSettingsStatistics` 和共享 Pivot 计算。没有另建第二套聚合，也不向真实 worksheet 写预览 cells。

## 会话、权限、IME 与主题

- owner 捕获 Studio 实例、unit、host sheet、user、content revision、history generation；内容命令携带 expectedRevision/expectedHistoryGeneration/expectedUserId 和 host=workbench。
- 外部编辑、历史替换、用户/宿主/Studio 变化使旧草稿无法提交。面板保留草稿并提供明确重读；`owner-change` 事件仅由仍然匹配且当前的 session 接收，重读后同步宿主 owner，旧 session 不能改新面板。
- 源 worksheet、源 range、linked output 以及已有 chart 的范围查看权限在暴露 source fields 前检查；来源不可访问时不向字段卡片或成员选择器传入缓存内容，草稿保留在会话内。宿主、已有 chart 与 linked output 编辑权限用于 UI advisory。内容命令仍校验真实变更后的完整范围，包括新图表的放置范围与报告扩张。
- 异步预览和成员请求带 session/generation guard。成员请求过期时只结束自身 loading，保留筛选草稿并显示过期错误；旧 completion 不清新请求状态。
- busy generation 隔离旧内容任务。no-op 根据实际 workbook revision 判断，只显示「设置没有变化」，不发出已更新事件。
- 输入全部使用 Ant Input。composition capture 期间禁止关闭、换源、字段区域/汇总导航和内容应用；只允许显示名称/数字格式的临时文字更新。compositionend 等 Ant Input 发布最后文本后再结束 composing，避免 capture 先清状态而丢掉最后的输入字符。Tabs forceRender 且不销毁隐藏内容。仍须 Windows 真 IME 验证。
- 主 Modal 关闭的 Canvas 焦点由 session owner/current host 判断，嵌套标签 Modal 使用 Ant 的 trigger focus 恢复。已编写 Escape/重读/主题/输入草稿保留的浏览器场景，没有运行结果。
- 本批新增和修改的可替换 UI 均使用 antdv-next：Modal、Tabs、Form/FormItem、Button、Select、Input、Checkbox、Alert；字段列表由布局容器组成。主题通过已有统一 ConfigProvider 与 `theme.useToken()` 适配。Canvas 内核和现有输入桥不在本批改动范围，不能据此声称全项目已迁移验收。

Context7 已用于 Tabs items/content、forceRender、destroyOnHidden 与 Checkbox API 资料。后续精确版本查询遇到 `Monthly quota reached`，继续静态对照本地 antdv-next 1.8.2 的 [Tabs 类型](../../node_modules/.pnpm/antdv-next@1.8.2_vue@3.5.45_typescript@6.0.3_/node_modules/antdv-next/dist/tabs/index.d.ts)；没有因此引入其他组件库或原生可替换控件。

## 本批文件与接口

本批 **13 个文件**，其中 ExcelWorkbench、Ribbon、通用 chart panel hook 只修改透视图对应小片段；其他在途改动不归本记录。

| 文件 | 本批职责 |
| --- | --- |
| [PivotChartPanel.vue](../../packages/office-studio/src/components/spreadsheet/PivotChartPanel.vue) | 创建/选择来源、四字段区/外观 Tab、标签 Modal、owner/IME/busy 生命周期和真实内容调用 |
| [pivot-chart-panel-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-panel-state.ts) | 新来源 DTO、权限 advisory、草稿来源/映射、configure/create DTO 与 data-only 预览；保留原 legacy helper |
| [pivot-chart-panel-owner.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-panel-owner.ts) | owner 捕获、当前判断与命令上下文 |
| [pivot-chart-field-draft.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-field-draft.ts) | 四区/汇总目录、immutable 字段编排、设置与选项草稿 |
| [PivotChartFieldsForm.vue](../../packages/office-studio/src/features/pivot/views/PivotChartFieldsForm.vue) | 全 Ant 字段和报告选项控件 |
| [use-sheet-chart-panels.ts](../../packages/office-studio/src/features/charts/views/use-sheet-chart-panels.ts) | 捕获 owner、重读同步、旧 session 隔离与焦点恢复 |
| [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) | 既有透视图根 availability 的禁用原因 |
| [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) | 真实选区/所选 chart/所在 Pivot 输出初始化、operation 别名、Panel owner 事件接线 |
| [pivot-chart-ui-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-chart-ui-fixture.mjs) | D3:H7 的真实非零偏移源 cells、全新无报告分支、owned/linked/异宿主图表与保护样本 |
| [pivot-chart-ui-workbench.html](../../packages/office-studio/tests/fixtures/pivot-chart-ui-workbench.html) | 真实 Ant/Workbench/Canvas/Facade/history 浏览器夹具；延迟后调用原 API |
| [pivot-chart-ui-panel.contract.test.mjs](../../packages/office-studio/tests/pivot-chart-ui-panel.contract.test.mjs) | 纯草稿、源权限、真实创建/configure/聚合/历史/owner/hook 合同草稿 |
| [pivot-chart-ui.browser.test.mjs](../../packages/office-studio/tests/pivot-chart-ui.browser.test.mjs) | 实际控件、操作命令、IME capture、异步竞争、主题、焦点、保护与重开草稿 |
| 本证据 | 固定来源、清点、差异、接口、草稿和未验收项 |

本批 UI 新接口为 `owner-change(PivotChartPanelOwner)`，宿主对应 `ownerChanged(owner)`；open input 新增 sourceKind/sourceSheetId/sourceRange。内容接口依赖主任务的 [pivot-chart-api-types.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-api-types.ts)、[pivot-chart-commands.ts](../../packages/office-studio/src/api/pivot-chart-commands.ts) 和 Facade，字段引擎依赖 [pivot-chart-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts)。这些共享 API/model/core 文件不计入本 UI 批修改数；引擎单独见 [SHEET-24 字段引擎证据](SHEET-24-pivot-chart-engine.md)。

## 草稿清点与未验收项

| 文件组 | 静态 test 声明 | 参数展开后的计划用例 | 执行结果 |
| --- | ---: | ---: | --- |
| 新 UI 合同 | 22 | 42 | 未运行 |
| 新 UI 浏览器 | 21 | 28 | 未运行 |
| 本批新增合计 | 43 | 70 | 未运行 |
| 保留原 [pivot-chart-panel.contract.test.mjs](../../packages/office-studio/tests/pivot-chart-panel.contract.test.mjs) | 5 | 5 | 本轮未运行，文件未修改 |
| 新草稿加保留 legacy 的计划范围 | 48 | 75 | 不表示运行或通过数量 |

合同参数展开：2 种读取所有权、3 种隐藏范围、3 种编辑保护、2 种 configure 所有权、11 类实际汇总、4 类 owner 变化、2 种刷新/删除。浏览器参数展开：2 种 Ribbon 变体、2 种 configure 所有权、3 类延迟提交变化、3 种编辑保护、2 种刷新/删除。内部断言循环不额外计为用例。

按照用户当前「先实现、后统一测试」约定，**本轮未运行测试、类型检查、前端构建、Rust/Cargo/WASM、dev、浏览器、Office 文件样本、性能基准或 git diff --check**。静态读写、声明清点和来源/链接核对不是执行验证。浏览器夹具计划禁用 navigator.gpu 使用 Canvas2D，这不证明 WebGPU 回退已在设备上验收。

仍待实现或验收：

1. 上游互斥 Ribbon factories、withPivotTable 同时新建表+图、完整 sidebar/双击/通用图表导航和运行时 overflow 布局。
2. Canvas 轴/字段按钮、字段列表浮层、成员/字段 collapse、category level 钻进/上卷与多值虚拟字段任意位置。
3. 完整 Show Values As、值排序/手工排序、date-condition/partial/hasData/tri-state filter query、Top-N、完整分组及全部 Pro options/controls。值显示方式进入后续批，不代表从交付目标移除。
4. 来源替换、cache 与连接迁移、同源多值实例专门 UI、上游数字自动默认 value/其他自动 row 的添加交互；目前是显式四区添加。
5. 共享 45 项图表目录的每项最终效果及样本覆盖；复用 ChartSettingsFields 与保留 author options 不证明全类型完成。
6. 上述合同和浏览器真实执行；特别是 antdv-next 控件/弹层、Windows IME、焦点恢复、用户/Studio/历史变化以及 linked 输出范围扩张的跨模块行为。
7. 真 XLSX pivotChart/cache/chart relationship 往返、Microsoft Office 原生打开与保存后的独立重开。内部 snapshot 草稿不等于 Office 文件保真。
8. 大样本 retained Rust/WASM 聚合/布局、取消、内存和同数据集前后基准。本批没有新增 Worker/WASM 协议，不声称性能优化、零拷贝或性能达标。

SHEET-24 与所有 Pro 功能继续保持完整交付范围和「未验收」。本批面板收尾不等于完整透视图能力完成。
