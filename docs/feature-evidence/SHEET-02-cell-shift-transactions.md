# SHEET-02 单元格局部移位、共享事务与确认入口

记录日期：2026-11-26。固定参考 Univer v1.6.0，`E:\rust_test\univer`，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：**实现与草稿已落盘，未验收**。本批续接[范围主题局部结构](SHEET-04-range-theme-structure.md)，将现有局部移位迁入统一候选事务、公开命令和 Facade。按用户安排，测试、类型检查、构建、dev/浏览器、Rust/WASM 和性能基准全部后置。

## 固定来源与入口清点

| 固定来源 | 已静态核对的契约 |
| --- | --- |
| [insert-range-move-right.command.ts][insert-right]、[insert-range-move-down.command.ts][insert-down] | 显式 range 优先，隐式当前选区；重叠插入判断、横纵带、插入空白的样式、尾部容量、资源 interceptor 和内容历史。 |
| [delete-range-move-left.command.ts][delete-left]、[delete-range-move-up.command.ts][delete-up] | 删除选中格并移入尾部内容；关联范围更新、撤销和选择。 |
| [右移确认][confirm-right]、[下移确认][confirm-down]、[左移确认][confirm-left]、[上移确认][confirm-up] | 先读取首选区，再沿移动轴延伸到工作表尾部检查部分合并；纵向确认还检查筛选隐藏行。允许后异步调用核心命令。 |
| [insert.menu.ts][insert-menu]、[delete.menu.ts][delete-menu]、[schema.ts][schema] | 单元格右键的插入/删除父项、组内顺序、确认命令、基础范围/exclusive-range 显隐和 workbook/worksheet/range 权限状态。 |
| [f-range.ts][f-range]、[dimension.ts][dimension] | `insertCells/deleteCells` 返回 void；`Dimension.COLUMNS=0`、`ROWS=1`，分别对应右/左和下/上。 |
| [ref-range/util.ts][ref-util]、[utils.spec.ts][ref-tests]、[util-behavior.spec.ts][ref-behavior] | 局部带范围/引用变换，以及跨垂直带时的范围处理。源码与草稿设计是静态依据，不代表上游或本地实际运行一致。 |

| 动作 | 核心命令 | 确认命令 | 本地/固定右键位置 |
| --- | --- | --- | --- |
| 插入并右移 | `sheet.command.insert-range-move-right` | `sheet.command.insert-range-move-right-confirm` | 插入父项的第 3 子项；先插入行，再插入列。 |
| 插入并下移 | `sheet.command.insert-range-move-down` | `sheet.command.insert-range-move-down-confirm` | 插入父项的第 4 子项。 |
| 删除并左移 | `sheet.command.delete-range-move-left` | `sheet.command.delete-range-move-left-confirm` | 删除父项的第 3 子项；先删除列，再删除行。 |
| 删除并上移 | `sheet.command.delete-range-move-up` | `sheet.command.delete-range-move-up-confirm` | 删除父项的第 4 子项。 |

分类：**复用右键父项 2、移位子项 4、注册核心/确认命令 8、Facade 方法 2；新增常驻 Ribbon 根 0、插件扩展根 0、溢出根 0、独立上下文根 0。** 本地 Sheets 常规 schema 仍为 69 根。classic/grid 或 More 复用同一入口；命令、子项和草稿数不可相加当功能完成数量。

[sheet-context-menu.ts](../../packages/office-studio/src/components/spreadsheet/sheet-context-menu.ts) 已按上述子项顺序接四个确认 ID。范围保护仍按实际移位尾部禁用；执行端再次检查权限。纵向确认在筛选隐藏行存在时返回实际 Ant 错误反馈，不提交内容或先创建合并弹窗。核心命令与 Facade 保留显式范围的行为，不额外套用上游 UI 确认入口的筛选限制。

## 模块、模型和候选事务

责任归属为 SHEET-02 的 sheet edit/selection，依赖 COMMON-02/03/05、SHEET-04/08 和各范围资源。领域几何、业务 planner、API 注册与 UI controller 分开；内容权威仍是现有 WorkbookModel/内容历史，Canvas 视图不是第二份编辑引擎。

- [domain/sheet-cell-shift.ts](../../packages/office-studio/src/domain/sheet-cell-shift.ts) 定义 `SheetCellShiftEdit={axis,kind,start,count,bandStart,bandEnd}`、位置映射/逆映射、物理范围拆分和 Office 上限。物理片段不静默截断。
- [sheet-cell-shift-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-shift-command.ts) 校验范围、容量及来源坐标/字典 key，准备脱离原模型的完整候选，再判定部分合并确认。显式 range 独立归一化，即使同时传 ranges；隐式选区保留 activeRangeIndex，显式 selections/ranges 默认取最后一项。
- 插入空白只保留 direct patch、目录引用、reset mask 和 cell theme；值、公式、rich runs、链接、批注等内容跟随移动格。仅由默认/行列/范围主题提供样式的空白继续稀疏，不能冻成 direct。后续通过同一继承解析器刷新有效样式。
- 完整位于移动带内的合并直接变换，完整被删的合并删除；部分重叠或跨插删边界的合并先拆分候选，必须确认。权限覆盖移位尾部及每个被拆合并的完整原范围。
- [sheet-cell-shift-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-shift-resources.ts) 读取原始坐标并只修改候选：公式/名称、迷你图源、内部普通/富文本链接、范围主题、CF/DV、保护规则、筛选、表格、图表绑定、透视和切片器。
- [formula-reference-areas.ts](../../packages/office-studio/src/features/formula/formula-reference-areas.ts) 是公式 token 扫描与局部结构变换的共用来源；带引号表名、绝对引用、跨表、外部/3D/整轴/结构化 token 按同一契约处理。CF/DV 分片重新基准相对公式，list literal/checkbox label 不当成引用。
- 连续矩形资源不能表示拆开后的范围时拒绝整次候选。表格列与筛选跟随整带；图表源字段/版本更新；透视源、字段、输出起点/矩形、generatedCellKeys 和 drill-down 失效同批处理；切片器字段依赖及 from/to 锚点跟随。删除图表完整源、透视完整源/输出起点或产生对象重叠时明确失败。
- 画布容量只为实际移后成员增长，包括无 cell 的图表源、透视源/输出/生成坐标和切片器锚点。行高列宽、手动隐藏、outline/freeze 仍为轴属性。不会因空白局部插入无条件增长整轴，也不静默丢弃超 Office 上限的内容。

[sheet-clipboard-changes.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-changes.ts) 按归一化的 per-cell predicate/真实覆盖顺序计算规则差异，分片新 ID 不扩大保留片段的权限；规则重排保护真实交集，包括非相邻重复 predicate 的几何交换。候选还通过同一范围权限与 stop 数据验证，远端受保护公式/资源依赖也会拒绝整次操作。

## API、历史和 UI 宿主

[sheet-cell-shift-commands.ts](../../packages/office-studio/src/api/sheet-cell-shift-commands.ts) 注册八个 ID；[OfficeStudio](../../packages/office-studio/src/api/office-studio.ts) 装配；[Facade](../../packages/office-studio/src/api/facade.ts) 的 `FRange.insertCells/deleteCells` 使用正确 Dimension 并走同一核心命令，支持显式非活动工作表。非法 Dimension 明确报错。

API 在准备及提交前检查 unit/model identity、baseRevision、历史 generation、用户、宿主与权限。确认请求经 [events.ts](../../packages/office-studio/src/api/events.ts) 的 `SheetCellShiftConfirmationRequested` 发布，含独立请求 ID、目标范围、拆合并范围、内容/历史版本；未确认不进入内容历史。接受后一次 `commitWorkbook` 同步候选与选择历史，发布 `SheetStructureChanged` / `SheetValueChanged`。选区之前/之后与内容在同一历史记录，保存/重开使用既有内部快照链路。

[use-sheet-cell-shift.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-cell-shift.ts) 与 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 接 API 或 standalone 共享 planner。standalone 先准备公式/自动行高视图再发布，同 ID 的宿主替换也检查原 identity/signature。Studio 在事件后再核对准确的 baseRevision+1；applyView 重入修改了未提交 view 时不回传该 view；选择、父回调和迟到回焦都检查 generation/当前发布。离开再返回相同 ID 不能授权旧候选。确认失败保留弹窗与可修复错误。

2026-11-26 用户报告工作台初始化时 `studio.events.on is not a function`。移位 controller 的订阅已改为现有 `EventHub.addEvent()`，保留 watcher 清理时的 `dispose()`；四个契约草稿文件中的八处同类旧调用也已迁移。补充真实 OfficeStudio 事件中心的初始化、替换 Studio 解绑与宿主销毁草稿。此次只做源码搜索和差异审阅，未执行测试或浏览器验证；其他控制台警告没有具体文本，尚未定位。

[SheetCellShiftConfirmationDialog.vue](../../packages/office-studio/src/components/spreadsheet/SheetCellShiftConfirmationDialog.vue) 使用 `antdv-next` Modal/Alert，复用统一 ConfigProvider/token。旧工作台的 cell-range pending、独立合并 Modal 和独立移位算法已移除；图表回焦、范围主题及 Canvas/Ant 输入的竞争弹层门禁保留。[sheet-cell-range-move.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-range-move.ts) 现在仅是旧 view 接口到共享 planner 的兼容适配，旧 move→reference helper 顺序仍只变换公式一次。

## 未运行草稿与静态核对

本批三个新文件共有 **57 条顶层 test 声明**，只按入口文本计数，循环子场景不另计。原范围主题基础 46、局部结构 32 是其他批次，不并入本批。

| 文件 | 草稿数 | 计划验证 |
| --- | ---: | --- |
| [sheet-cell-shift.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-shift.contract.test.mjs) | 33（24 纯契约 + 9 API/Facade） | 四方向、范围优先、稀疏作者/公式/规则、容量/Office 边界、资源/筛选确认、权限/stop 验证、历史、内部重开与 observer 重入。 |
| [sheet-cell-shift-host.contract.test.mjs](../../packages/office-studio/tests/sheet-cell-shift-host.contract.test.mjs) | 16（11 standalone + 5 Studio） | 确认捕获、真实事件中心初始化/解绑、相同 ID/JSON 替换、切表/用户/激活 generation、prepare/apply/selection/observer 重入、过滤尾部及迟到回焦。 |
| [sheet-cell-shift.browser.test.mjs](../../packages/office-studio/tests/sheet-cell-shift.browser.test.mjs) + [fixture](../../packages/office-studio/tests/fixtures/sheet-cell-shift.html) | 8 | 真右键/活动非最后选区、取消/接受、较新内容/切表、受保护项禁用、筛选确认拒绝、暗色/grid 与 Ant 合成 composition。 |

同时静态修正既有 [cell-range move 草稿](../../packages/office-studio/tests/sheet-cell-range-move.contract.test.mjs)、[editor-store 历史草稿](../../packages/office-studio/tests/editor-store.contract.test.mjs) 的来源/公式兼容预期及 [右键草稿](../../packages/office-studio/tests/sheet-context-menu.contract.test.mjs) 的确认 ID/次序与轴 range 参数。这些原有 test 入口不计入新增 57。

已做只读源码/在途文件审阅、入口和声明计数、补实现与文档；**没有运行任何测试或类型/构建/Rust/WASM/浏览器/性能检查**。浏览器草稿强制 Canvas2D；合成 composition 不证明真实系统 IME。草稿里的内部保存重开也不证明 Office 文件保真。

## 实现差异和未完成目标

- 固定 core handler 与 confirm 是异步实现，本地同步 boolean 配合确认事件；Facade 仍 void，但非法枚举、陈旧版本及准备失败可同步报错。固定 confirm 用首选区、core 隐式用最后选区；本地以显式 primary/实际 activeRangeIndex 捕获同一目标，避免确认和执行不同范围。差异已记录，运行效果未核实。
- 本地对跨插删边界的部分合并统一确认拆开，连续矩形资源拆分会拒绝；全部合并/复杂范围组合与上游实际结果仍待样本核验，不将此策略当成最终能力缩水。
- 普通 chart 当前只有绝对 pixel position；完整 drawing/media/组合对象锚点模型、线程评论的独立资源和所有扩展 interceptor 仍有缺口，不能宣称所有绘图/资源均已支持局部移位。
- 现行 CF/DV/主题矩形和对象重叠检查有 100,000 片段 / 5,000,000 工作量同步拒绝预算。这只是保护值，未由实测阈值证明；大样本索引、Worker/Rust/WASM 迁移、分块取消及传输/内存计量仍待实施和验收。
- UI 明暗主题、焦点恢复、真实 Windows IME、WebGPU 与 Canvas2D、大表/多表、权限动态变化、全部引用/资源组合及 Office XLSX 独立重开均未验收。跨客户端重基/权限/持久历史继续按[后端需求](../BACKEND_REQUIREMENTS.md)后置，不能由前端事件或内部 JSON 替代。

[insert-right]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/insert-range-move-right.command.ts
[insert-down]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/insert-range-move-down.command.ts
[delete-left]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/delete-range-move-left.command.ts
[delete-up]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/delete-range-move-up.command.ts
[confirm-right]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/insert-range-move-right-confirm.command.ts
[confirm-down]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/insert-range-move-down-confirm.command.ts
[confirm-left]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/delete-range-move-left-confirm.command.ts
[confirm-up]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/delete-range-move-up-confirm.command.ts
[insert-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/insert.menu.ts
[delete-menu]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/delete.menu.ts
[schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/schema.ts
[f-range]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/facade/f-range.ts
[dimension]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/core/src/types/enum/dimension.ts
[ref-util]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/services/ref-range/util.ts
[ref-tests]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/services/ref-range/__tests__/utils.spec.ts
[ref-behavior]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/services/ref-range/__tests__/util-behavior.spec.ts
