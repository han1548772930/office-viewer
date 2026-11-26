# SHEET-23 / SHEET-24 轴展开折叠与类别层级

日期：2026-11-26。状态：领域状态、共享计算、独立/关联报告编辑及契约草稿已落盘，**未验收**。本记录覆盖引擎；Facade、Studio、权限/历史、字段面板、图表投影与 Canvas 轴命中由对应主任务接入和记录，不以引擎草稿替代产品闭环。

## 固定依据与历史缓存边界

- 固定 OSS：`E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。
- v1.6.0 [sheets-pivot-chart/types.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts)：`setCollapse` 的 fieldId/collapse/item? 与 `setCategoryLevel` 的 fieldId。
- v1.6.0 [f-pivot-chart.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts)：`setCollapse` 限定 row hierarchy，`setCategoryLevel` 请求最深可见 row field；linked 图表共享报告折叠修改。
- v1.6.0 [pivot-chart.command.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/commands/commands/pivot-chart.command.d.ts)：字段/成员折叠、类别层级及 expand/collapse step 命令参数。
- v1.6.0 [sheets-pivot-chart/lib/es/index.js](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/es/index.js)：`W` 折叠命令检查 rowFields，linked 走语义 PivotTable 字段命令、owned 走同类型 report edit；`Ut` 请求 rowLevelFieldId；`Wt`/`qt` 定义 step 选择规则。只静态读取片段，未执行发布包。
- v1.6.0 [sheets-pivot-chart-ui/const.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/const.d.ts) 和 [menu/schema.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/menu/schema.d.ts)：axis context 与 `sheet.pivot-chart.expand-collapse` selector。
- v1.6.0 [sheets-pivot-chart-ui/lib/es/index.js](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/es/index.js)：`or` 的四个固定子项和 rowFields 动态层级，`Xn/Zn/Yn` 的当前成员命令，`_resolveCategoryTarget` 的原始成员身份与能力条件。
- v1.6.0 [set-pivot-collapse.command.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot/lib/types/commands/command/set-pivot-collapse.command.d.ts)、[set-pivot-collapse-by-field.command.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot/lib/types/commands/command/set-pivot-collapse-by-field.command.d.ts)：工作表单元格入口及不依赖输出单元格的 field/item 语义入口。
- v1.6.0 [sheets-pivot/const/type.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot/lib/types/const/type.d.ts)：`ISetPivotCollapseMutation` 使用 boolean/member-map/undefined；`IGetPivotCollapseDataParams` 同时声明 tableFieldId/collapse/item 与整轴 rowLevelFieldId。
- 历史 [engine-pivot/package.json](../../.research.local/univer/packages/engine-pivot/package.json) 与 [pivot-chart/package.json](../../.research.local/univer/packages/pivot-chart/package.json) 均为 **1.5.2**。其 [pivot-table.d.ts](../../.research.local/univer/packages/engine-pivot/lib/types/pivot/pivot-table.d.ts)、[model.d.ts](../../.research.local/univer/packages/engine-pivot/lib/types/pivot/model.d.ts) 和 [engine-pivot/lib/es/index.js](../../.research.local/univer/packages/engine-pivot/lib/es/index.js) 提供 `setRowCollapseLevel`、collapseInfo 及从整字段折叠展开单个缓存成员的状态物化实现。不能将该历史内核当成固定 v1.6.0 内部计算已核实。

没有安装或依赖商业 SDK。固定包装层证明可观察入口、参数、状态和命令分流，不能证明私有聚合、Office 样本或真实浏览器输出已通过。

## 入口分类清点与实际作用

本范围是 **4 个固定上下文子项 + 每个 row field 一个动态层级项**；另有 expand/collapse step 控制命令，单列而不加入四项计数。selector 在没有类别上下文选择时隐藏；当前成员只对可折叠的非叶 row field 提供目标。本批不新增顶部工作区，不修改主任务 Ribbon 分类数量。

| 固定子项 | 标签来源 | 命令与目标 | 上游禁用条件 |
| --- | --- | --- | --- |
| 展开当前成员 | `interaction.expand` | `sheet.command.expand-pivot-chart-category-item` → `set-pivot-chart-collapse`，fieldId + item，collapse=false | 当前成员未折叠 |
| 折叠当前成员 | `interaction.collapse` | `sheet.command.collapse-pivot-chart-category-item` → 同一字段命令，collapse=true | 当前成员已折叠 |
| 展开整个字段 | `interaction.expandField` | `sheet.command.set-pivot-chart-collapse`，不传 item，collapse=false | 非叶字段没有任何折叠状态 |
| 折叠整个字段 | `interaction.collapseField` | 同一字段命令，collapse=true | 叶子字段或已经整字段折叠 |

动态层级调用 `sheet.command.set-pivot-chart-category-level`，fieldId 遍历全部 rowFields，包含叶子。固定 ES 用目标/所选字段索引比较选择 `interaction.expandTo/collapseTo` 标签，禁用条件比较各非叶字段的 canExpand/canCollapse。计算语义由明确目标字段决定，不能根据标签词判断完成。

`sheet.command.step-pivot-chart-category-level` 的 expand 取从浅到深第一个含整字段或任意成员折叠的非叶 row field；collapse 取从深到浅最后一个尚未整字段折叠的非叶 field。没有可变更字段返回 false，不增加 revision。

本地 chart edit 对齐固定公开限制：`setCollapse` 只接受 row field；共享 Sheets 引擎支持 row/column，基准项查询和轴状态也能读取 column。这部分属于本地共享编辑器能力，尚不能据固定图表 Facade 声称公开列折叠 API 已对齐。

## 持久化状态与真实输出

模块归属为 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的 `SHEET-23 / SHEET-24 → SHEET-22`。`PivotTableModel.collapseInfo` 使用 report-local field ID → boolean 或 grouped member key → boolean map。类别深度通过同一状态推导，**没有第二个持久化 category depth 权威值**。

- 整字段 true 折叠该字段全部当前及后续成员；整字段展开清理其状态。
- 成员 map 只折叠指定 grouped key。相同 key 在多个祖先内出现时分别作用于各自 scope，与字段/成员 DTO 对应。
- 从整字段 true 展开单个成员时，物化现有来源分组成员，保留其他成员折叠；随后出现的新成员采用 map 的默认展开语义。此物化政策有 1.5.2 实现依据，仍待固定版运行/Office 样本核实。
- false 标记与默认展开等价，归一后不因字段/member 键顺序或显式展开标记增加内容 revision；已有展开表示可以保留而不产生一次假编辑。
- 祖先折叠保留后代折叠状态；展开祖先后，后代自己的规则继续生效。
- `setCategoryLevel` 展开目标之前的所有 row parents，折叠目标及更深的 parents；选择叶子会展开全部 row parents。column state 与 Values 方向保持原值。
- 混合成员状态允许不同分支具有不同深度。`PivotAxisState.categoryLevelIndex` 表示最浅的折叠边界，`mixed` 标记混合状态；不能将其当成所有行具有同一可见深度。

`projectPivotAxis` 只投影现有有序分组和父范围，不读取来源或实现聚合。折叠节点替换其后代，取真实父 scope 的原始汇总，标签保留原 scalar 且不追加 `Total`；其他展开分支保留叶子。显示小计时，展开的祖先小计仍可存在，折叠节点不会再生成一个相同 scope 的重复小计。

唯一来源读取、过滤、分组、汇总和 Show As 仍在原 `pivot-engine.ts`。折叠只是改变需要输出的 scope：每个 measure 使用其原 aggregation，Show As 在原聚合后变换，隐藏小计/总计不改变分母。当前 subtotal/rank/reference 的政策继续采用 [值显示方式证据](SHEET-23-24-pivot-show-data-as.md)，不以折叠绕过其未验收项。

实例：真实 source 有 32 条记录，Region→Product→Item 与 Year→Quarter→Channel 各三层，grand=1000。North 在 Q1/Retail 的四个叶值为 10、20、30、40。折叠 North 后输出一个 North=100 的父行，South 的四个叶行继续存在；展开 North 恢复其四个叶值。显示至 Product 输出 North/A=30、North/B=70、South/C=30、South/D=70。行 North 与列 Year 同时折叠时为 500；AVG 对同一原始范围为 31.25，而非相加已显示的平均值。

## 输出元数据与跨模块协议

`calculatePivotReport` 在原有矩阵、headerRows、numberFormats、drillDownMap 之外返回 `rowAxis/columnAxis`。矩阵维持完整已配置 row field 列与 column header 层数；折叠后的后代位置为空。元数据与矩阵在同一次计算中产生：

```ts
interface PivotAxisOutputEntry {
  index: number // output matrix 内含 header 的行/列索引；worksheet 坐标需加 anchor
  kind: 'detail' | 'collapsed' | 'subtotal' | 'grand-total'
  fieldId?: string // 最深可见的 report-local 字段
  level: number // 0-based；total/无维度为 -1
  labels: CellScalar[]
  fieldFilters: Record<string, CellScalar> // local IDs → 原 grouped scope
  fieldMembers: Record<string, CellScalar[]> // 原 source scalar 成员
  valueFieldId?: string // 值在行/列时对应其所在轴的真实 measure instance
}
```

worksheet render 成功后将同一元数据放在 `report.axisOutput`；空输出清空，输出冲突不会发布新 metadata。linked 图表可读取它及已生成矩阵，**无需再次聚合**。旧报告没有 metadata 时需由读取端选择同源 calculate fallback。data-only 清除 worksheet output metadata；owned 图表读取现有 calculate 的返回值，不另存第二份矩阵。

这是 JavaScript 报告的派生投影/命中缓存，不是 Rust retained buffer 或零拷贝声明。后续 Canvas 可用 matrix index、local ID、原 scope 与成员执行真实成员操作；实际命中和按钮本批未接入。

主任务图表归一化应保留 detail/collapsed，排除 subtotal/grand-total，并在 collapsed 节点后清除 descendant 的前值填充。不能继续仅用 `Total` 字符串判别，或将祖先/上一分支的子标签灌入折叠父行。本代理只提供协议，不修改 `chart-pivot-data.ts`。

## 接口、校验与生命周期

共享引擎新增：

```ts
setPivotCollapse(workbook, pivot, fieldId, collapse: boolean, item?: string,
  mode: PivotRefreshMode = 'worksheet'): void
getPivotCollapse(pivot, fieldId, item?): boolean
getPivotCollapseItems(workbook, pivot, fieldId): string[]
getPivotAxisState(workbook, pivot, area: 'row' | 'column' = 'row'): PivotAxisState
setPivotCollapseInfo(workbook, pivot, info: PivotCollapseInfo | null, mode?): void
setPivotCategoryLevel(workbook, pivot, fieldId, mode?): void
stepPivotCategoryLevel(workbook, pivot, direction: 'expand' | 'collapse', mode?): boolean
```

chart 服务提供 `getPivotChartAxisState/getPivotChartCollapseItems`；edit 增加 `setCollapse/setCategoryLevel/stepCategoryLevel`，完整 fields config 增加 `collapseInfo?: PivotCollapseInfo | null`。候选报告统一比较内容状态并最终刷新一次；owned/data-only 与 linked/worksheet 使用同一聚合。deferred 只用于既有批量编辑，再由最终调用者统一刷新。

命令目标沿用 chart canonical source 优先解析和重复来源消歧，传给共享 setter 的 ID 已是明确 local ID；存储的 collapseInfo 键、scope 和 metadata 必须精确 local ID，不在刷新时重新解释成其他来源。字段/成员键包含 prototype-like 字符串时按 own entry 处理，避免原型属性改变成员状态或聚合 scope。

未知字段、非轴字段、leaf 折叠、未知/真实 scalar 类型碰撞的具体成员、非法 boolean/map/direction/mode 拒绝；非法 refresh mode 在修改状态之前校验。计算与读取会诊断存量未知/leaf/非法 map 配置，不能静默忽略配置。具体成员使用真实来源经现有 manual/number/date 分组后的键，与 raw label-filter query 分开。

生命周期政策：

- 来源过滤或暂时改名使成员不再可见时保留 dormant key，返回原成员后恢复折叠；当前输出仅对真实分组应用状态。
- 移除/移动字段后，清理被移除、移至 filter/value 或失去后代而成为 leaf 的状态。仍在另一轴具有后代的同一字段实例保留其状态；类别深度随存活 hierarchy 推导。
- 普通分组编辑若使当前可地址的折叠成员键消失或歧义，则在修改 grouping 之前拒绝。整字段 true 不依赖具体键，可随新分组继续折叠。
- 一次完整 config 可同时重映射 grouping 与 collapseInfo。未显式提供状态时仅保留来源身份相同且仍有后代的字段状态，再校验成员；显式 null 重置。旧 dormant key 在严格完整配置中仍可能要求用户重置/重映射，此行为待 Office 样本核实。
- 既有切片器、value-filter 和 Show As 字段依赖保护继续生效。本 helper 不再次运行 selection/cache fan-out，主任务共同 candidate/历史边界负责。

## 文件与验证状态

本批文件：

- [model.ts](../../packages/office-studio/src/domain/model.ts)：仅 Pivot 状态和输出元数据片段。
- [pivot-axis-state.ts](../../packages/office-studio/src/features/pivot/services/pivot-axis-state.ts)：纯规则、状态归一、能力判断、有序轴投影和分组依赖检查。
- [pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts)：共享 setter/query、原聚合 scope、worksheet metadata、grouping guard 和资源退役。
- [pivot-chart-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts)：owned/linked edits、完整配置、语义 no-op 与候选提交。
- [pivot-axis-state-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-axis-state-fixture.mjs) 与 [pivot-axis-state.contract.test.mjs](../../packages/office-studio/tests/pivot-axis-state.contract.test.mjs)：两轴三层、非零 source column、独立/关联报告的真实数据草稿。
- 本证据文件；不编辑 API/Studio/Facade/index、Panel/field-draft、图表投影或顶层清单。

静态清点 **43 条 test/test.each 声明**。7 组 each 参数分别为 ownership 2、Values 方向 2、汇总 11、Show As/隐藏成员政策 13、类别深度 3、refresh strategy 3、非法 DTO 16；展开为 **86 个草稿用例**（`43 - 7 + 2 + 2 + 11 + 13 + 3 + 3 + 16`）。内部循环与多次断言不另计为运行用例或通过数。

草稿覆盖 field/member/nested collapse、整字段与物化成员的新来源差异、恢复展开、两轴三层、混合深度、两种 Values 方向、多 measure/汇总、Show As、隐藏/可见小计、metadata/scopes、sort/repeat labels、no-op/显式展开/step 边界、三种模式、linked revision、原始下钻、三类分组、重映射/退役、local/canonical/scalar/prototype 键、dormant source、非法存量、输出冲突回滚和内部快照重开。

按照当前“先实现、后统一测试”要求，**未运行测试、类型检查、前端构建、Rust/Cargo/WASM、dev、浏览器、Office 文件样本、性能基准或 git diff --check**。静态回读、声明清点和文档路径核对不是执行验证，不能称为测试通过或已验收。

仍需验收或实现：

1. 固定版/Office 真实样本核实混合深度、成员物化后新数据、空白/错误项、重复成员、grouping/移动生命周期及 subtotal/Show As 的边缘政策。
2. 主任务命令/Facade/权限/文档历史、linked fan-out、动态菜单、`antdv-next` UI 的主题/焦点/IME、真实 Canvas 轴按钮与语义 hit-target 的浏览器闭环。本批无 UI 改动，不能代替控件迁移验收。
3. 图表归一化消费 metadata、保留真实折叠父项、清除 descendant 填充及两种 Values 方向的集成/浏览器验证。
4. parser/writer 原生 OOXML collapse/showDetails、pivotCache 与 pivotChart relationship 的保真保存/重开；内部 JSON 草稿不能证明 XLSX/Office 往返。
5. 现有 Worker/Rust/WASM 大计算路径和同数据集性能/内存/取消基准。当前复用 TypeScript 聚合器，没有新增 Worker 协议；后续按 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 接入版本、预算、取消、lease 和复制统计。本批没有性能实测，不能声称优化、零拷贝或达标。

全部 Pro 范围继续保留；本批实现不能作为 SHEET-23 / SHEET-24 完整交付完成证明。
