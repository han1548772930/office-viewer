# SHEET-25 切片器报表连接与上下文入口

记录日期：2026-11-26。开发进度：已落盘，待统一验证。验收状态：**未验收**。

本批承担切片器/时间线连接面板、设置草稿、样式图库和工作台入口；共享缓存、连接事务、报告过滤及资源生命周期由同批主任务与生命周期任务承担。这里记录真实代码变化，不据此验收完整 SHEET-25。后续[正式 Timeline 设置/日期筛选 UI](SHEET-25-timeline-settings-ui.md)复用本文件的上下文根，新增草稿与未验收项独立统计，不改本批根计数。

## 1. 固定依据与差异

OSS 基线为 `E:\rust_test\univer` 的 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。整体 Sheets / Docs、插件、上下文与溢出入口清点沿用[顶部入口总表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)、[Sheets 源码入口](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)及[工具栏对照](../UNIVER_TOOLBAR_PARITY.md)。本批补充 SHEET-25 的四个上下文根，不改变常规根口径，不以根数证明功能完成。

Pro 依据是项目研究缓存中的 **v1.6.0 已发布包**，并非 OSS 仓库里的闭源内部实现。下表中的 `…/` 均代表项目内的 `.research.local/univer-1.6.0/packages/`。已静态读取以下具体来源：

| 来源 | 核对符号/契约 |
| --- | --- |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/facade/f-slicer.d.ts` | `FSheetSlicer.getCacheId/getConnections/getConnectionCandidates/setConnections`：缓存共享过滤、视图独立展示；候选按同工作簿/缓存/源字段；替换连接；Table 无报告连接；普通拒绝返回 false |
| `…/sheets-slicer/lib/types/types.d.ts` | `ISheetSlicerConnections`、`ISlicerConnectionCandidate`：PivotTable 与 owned PivotChart 两类目标；linked chart 归一为所属 PivotTable |
| `…/sheets-slicer-ui/lib/types/types.d.ts` | `ISlicerConnectionsData`，连接弹层的数据边界 |
| `…/sheets-slicer-ui/lib/types/services/slicer-ui.service.d.ts` | `SlicerUIService` 的当前对象、设置与连接面板服务 |
| `…/sheets-slicer-ui/lib/types/views/SlicerConnectionsDialog.d.ts` | `SlicerConnectionsDialog` 接收连接数据及关闭回调 |
| `…/sheets-slicer-ui/lib/types/commands/operations/open-slicer-connections-dialog.operation.d.ts` | `OpenSlicerConnectionsDialogOperation` 的 `unitId/kind/cacheId` 参数 |
| `…/sheets-slicer-ui/lib/types/menu/menu.d.ts` | 四个固定 menu ID、`getFocusedSlicerControl$` 与 `supportsConnections` |
| `…/sheets-slicer-ui/lib/es/index.js` | 菜单 factory、格式分类的样式/选项组顺序、grid 布局、候选勾选后替换两类集合 |
| `E:\rust_test\univer\packages/ui/src/views/components/ribbon/Ribbon.tsx` 与 `RibbonGrid.tsx` | classic/grid 呈现及溢出语义，来源固定于上述 OSS commit |

公开规格：[Univer Sheets 切片器](https://docs.univer.ai/zh-CN/guides/sheets/features/slicer)。这里保留发布类型与静态 ES 可核实契约；公开演示、Office 样本及闭源内部事务细节未观察，不将推断写成已经验证的上游事实。

本项目差异：Vue + `antdv-next` 替代 React UI；设置使用 Drawer，连接使用 Modal；打开面板的 UI 意图带业务 `slicerId`，控制器再解析实际 `cacheId`，不要求调用方复制上游缓存内部参数。连接提交仍走 `FSheetSlicer.setConnections` 的共享事务。设置与样式只提交实际变化的展示属性，不修改来源身份。

本批已使用 Context7 查询 `/antdv-next/antdv-next` 的 Modal、Checkbox.Group 与主题 API，并静态读取本地 Form/Alert 类型；错误与警告使用 Ant Alert 的 `title` 属性。未启动组件、浏览器或 dev。

## 2. 上下文入口清点

格式 Tab：本地 `slicer-format`；当前仅一个合法且可访问的 slicer/timeline 对象被选中时显示。进入上下文激活该 Tab，退出时恢复先前有效常规 Tab。两种布局使用同一入口集合。

| ID / 标签 | 分组与组内顺序 | grid | classic | 可见/禁用条件 | 实际命令及后续 |
| --- | --- | --- | --- | --- | --- |
| `sheet.menu.slicer-quick-style` / 快速样式 | 样式 order 0；根 order 0 | r1c1，跨两行 | 第一组第一项 | 所有 slicer/timeline；无当前对象展示编辑权限时禁用 | `sheet.command.update-slicer-presentation`；图库传业务 ID、内容版本及 style patch → 同一内容命令/历史 |
| `sheet.menu.slicer-settings` / 切片器设置 | 选项 order 1；根 order 0 | r1c1 | 第二组第一项 | 仅 slicer；按对象完整 footprint 权限禁用 | `sheet.operation.open-slicer-settings-dialog` → `openSlicerSettings` → Drawer → `FSheetSlicer.updatePresentation` |
| `sheet.menu.timeline-settings` / 时间线设置 | 选项 order 1；根 order 1 | r1c1，与切片器设置互斥 | 第二组第一项，与前项互斥 | 仅 timeline；同一对象能力判断 | 同一 settings operation → 时间线设置草稿 → presentation 内容命令 |
| `sheet.menu.slicer-connections` / 报表连接 | 选项 order 1；根 order 2 | r2c1 | 第二组设置之后 | Pivot 来源且 `supportsConnections`；Table 隐藏；无展示编辑权限时禁用 | `sheet.operation.open-slicer-connections-dialog` → 真实候选/勾选 → `sheet.command.set-slicer-connections` |

可复查分类计数：本批 **4 个上下文根声明**，新增常驻根 0、插件常驻根 0、独立 More 根 0。Pivot slicer/timeline 每次适用 3 个；Table slicer 适用 2 个。根声明中的两种设置是互斥入口。溢出时由共享 Ribbon 收纳这些相同入口，不另加根或删除语义；实际窄屏溢出尚未运行验收。

其他入口单列：Canvas 绘图右键菜单有设置/连接两个业务意图（Table 仅设置），浮动操作区同样两个；设置 Drawer 有一个“报表连接”按钮。它们不是额外 Ribbon 根。quick gallery 与设置 gallery 当前共用已有 **14 个 slicer / 12 个 timeline 内置名字**，各另含一个文档默认值；这些是图库选择项，不能计入根数或完整自定义样式验收。

新增独立 owned PivotChart 已纳入“插入切片器/时间线”的来源列表；来源键为 JSON `[kind,id]`，避免 Table、PivotTable、owned PivotChart 的相同 ID 混淆。创建 owned PivotChart 的工作台入口属于后续批次。

## 3. 已落盘行为与边界

- 连接 UI 使用 Ant Modal、Form/Form.Item、Checkbox.Group、Checkbox、Button、Alert；列出报表名、工作表名及报告类型。初始勾选来自权威共享缓存，提交去重并替换两类连接集合，允许空连接。
- 候选列表因权限或失效来源缺少**已经连接**的目标时，读取返回 `UNAVAILABLE_CONNECTED_REPORT`，不把该目标默默转成未选中项。恢复权限或来源后可重新读取。
- 读取/提交固定 Studio、unit、sheet、user、revision、history generation、业务 view ID 与 cache ID；离开工作台、替换 Studio、外部内容变化、历史替换及过期异步结果不能覆盖新面板。旧请求的 busy token 不能清掉新请求的 loading。
- 展示设置保留独立草稿。外部内容变化后保留用户输入，展示“重新读取”，禁止使用旧版本保存。读取与面板开关不写内容历史。
- `buildPanelSlicerPresentationPatch` 只包含有变化的字段；未修改布局时不会把投影 AABB 回写，保留完整作者 affine 几何、剪切/旋转和来源。锚定与纵横比独立提交。
- 排列能力来自 `createSheetDrawingScene` 与共享 `planDrawingArrange`，支持跨 chart/image/shape/control/group 的实际兄弟顺序，并检查所有受影响对象的权限；不再仅比较 slicer 数组位置。
- 快速样式实际发出 presentation 内容命令，携带目标和预期内容版本；设置、右键、浮动工具与 Ribbon 使用相同 Facade/内容命令。
- 插入字段使用 canonical `sourceFieldId`；owned report 的本地字段实例 ID 不作为源字段。Table 不能生成时间线，日期字段条件沿用现有校验。
- 关闭顶层面板由 ExcelWorkbench 校验 owner、工作台、版本、编辑/IME 与其他弹层状态后归还 Canvas 焦点；嵌套连接 Modal 归还仍有效的 Ant 设置按钮。弹层和图库颜色来自统一 ConfigProvider/theme token。

UI 编排在主线程；模型、过滤、聚合、资源与历史仍由既有服务/命令负责。本批未另建权威计算引擎，也未新增 Worker/WASM bulk 协议或声明性能优化。

## 4. 文件与接口

| 文件 | 责任 |
| --- | --- |
| [SlicerPanel.vue](../../packages/office-studio/src/components/spreadsheet/SlicerPanel.vue) | 插入/设置/连接、受控草稿、Ant 控件、反馈与关闭 owner |
| [SlicerQuickStylePicker.vue](../../packages/office-studio/src/components/spreadsheet/SlicerQuickStylePicker.vue) | 上下文快速图库，发实际展示 patch |
| [slicer-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-panel-state.ts) | owned 来源/规范字段、展示 diff、共享排列与权限 |
| [slicer-panel-owner.ts](../../packages/office-studio/src/features/slicer/slicer-panel-owner.ts) | Studio/unit/sheet/user/revision/history owner 和命令 context |
| [slicer-connections-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-connections-panel-state.ts) | 候选、JSON key、空替换、缺失目标与过期保护、Facade 提交 |
| [slicer-style-catalog.ts](../../packages/office-studio/src/features/slicer/slicer-style-catalog.ts) | 既有内置样式名字共用目录 |
| [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) | 四根格式上下文、分组、两布局和能力 |
| [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) | 必要入口接线、对象上下文/Tab 与焦点归还；仅修改本子任务片段 |
| [SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) | 绘图右键/浮动设置与连接 emit；仅修改经协调的菜单片段 |
| [slicer-connections-fixture.mjs](../../packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs) | 真实稀疏来源、同/不同 authored cache 报告、owned/linked 图表、共享视图、时间线与 Table |
| [slicer-connections-workbench.html](../../packages/office-studio/tests/fixtures/slicer-connections-workbench.html) | 真实 Ant/Canvas/Facade/history 宿主，延迟真实候选/内容方法以验证过期完成 |
| [slicer-connections-panel.contract.test.mjs](../../packages/office-studio/tests/slicer-connections-panel.contract.test.mjs) | 契约草稿 |
| [slicer-connections-ui.browser.test.mjs](../../packages/office-studio/tests/slicer-connections-ui.browser.test.mjs) | 浏览器草稿 |
| 本证据文件 | 上游来源、入口分类、差异与未验收项 |

本批局部接口变化：`SlicerPanel.initialSection: 'settings'|'styles'|'connections'`；新增 `closed(owner)`；Canvas 新增 `slicer-connections(id)` emit；Ribbon 新增 `SlicerRibbonControlState`；`SlicerPanelPivot.kind` 增加 `ownedPivotChart`；面板控制方法接受已有 `SlicerCommandContext`。

依赖主任务已落盘的公开接口：`FSheetSlicer.getCacheId/getConnections/getConnectionCandidates/setConnections`，`FWorksheet.insertSlicer(info, context)`；连接 DTO 为 `pivotTables[]` 与可选 `ownedPivotCharts[]`。`setConnections` 的 false 是业务拒绝，异常 Promise rejection 单独展示。本子任务不改第二套 DTO 或内容历史。

## 5. 测试草稿与未执行验证

| 草稿 | 静态 `test(...)` 声明 | 按现有参数循环计划展开 | 状态 |
| --- | ---: | ---: | --- |
| `slicer-connections-panel.contract.test.mjs` | 25 | 35 | 全部未运行 |
| `slicer-connections-ui.browser.test.mjs` | 17 | 18 | 全部未运行 |
| 合计 | **42** | **53** | 未执行，不是通过数 |

契约草稿涵盖：JSON 身份键、未知/非法目标、空替换、不可访问的已有连接、共享视图、Table、兼容/不兼容缓存、ordinary false、权限/内容/用户/宿主/历史过期、canonical owned 插入、affine diff、混合排列及受保护兄弟、四根与显隐。

浏览器草稿涵盖：真实 Ant 候选选择/提交、空连接重连、undo/redo/savepoint/内部重开、Table 显隐、classic/grid、上下文恢复、样式命令、过期草稿/候选/Studio/历史、旧请求 busy 隔离、主题、Escape/焦点、只读。延迟包装在释放后调用真实 Facade 方法，未用固定结果替代业务模型。

已做静态读取与编辑、声明统计，未执行测试、类型检查、构建、Cargo/WASM、浏览器/dev、`git diff --check` 或基准。两份夹具也未启动。统一验证时需先生成匹配协议的 WASM，再执行适用检查。

## 6. 仍需验收/开发

1. 所有上述草稿与真实用户闭环；Windows 中文 IME、组合输入/焦点冲突、窄屏及 More、两种渲染后端、恢复/卸载。
2. owned PivotChart 创建/字段编辑完整工作台入口，其他 Pro 上下文与扩展入口逐项闭环。
3. 文档自定义切片器/时间线样式定义与完整样式目录；当前既有内置名字不代表全量样式完成。
4. 完整交叉筛选、日期粒度/时区/日期系统及大量成员的真实数据验证，全部公共 Facade/Builder 与资源缺失占位行为。
5. XLSX 原生 slicer/timeline/cache/report/chart 关系的解析、独立写出与 Office 兼容程序重开；内部 JSON 不代替 Office 验收。
6. 同数据集的大样本时间、帧、长任务、Worker/WASM 传输与 JS/WASM/GPU 内存实测；无实测不声明优化或达标。

这些缺口保留完整 SHEET-25 与 Sheets / Docs / Pro 目标；本批按用户要求先落盘、测试后置，所有验收状态继续为“未验收”。

2026-11-26 后续[显示设置与 cache options UI 批](SHEET-25-slicer-settings-ui.md)已将普通设置接到正式 getSettings/updateSettings，混合几何保留单事务，并纳入文档自定义/未知样式、四项 cache 选项及 composition/真实 owner 边界。合并 Panel/连接 state/夹具读取内部两类控制使用 getSlicerControl(s)/getModel，正式 getInfo 与 getSlicer(s) 继续按公开分类契约维护。旧合同只迁移 raw model getter，旧浏览器 caption 等待 gate 改到真实 updateSettings；本旧批 42/53 不重计。新增批 **40 条声明 / 107 个计划场景，全未运行**，文件、来源与剩余缺口在新记录中单列。
