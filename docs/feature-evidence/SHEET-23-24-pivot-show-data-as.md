# SHEET-23 / SHEET-24 透视值显示方式

日期：2026-11-26。状态：计算实现、领域字段与契约草稿已落盘，**未验收**。本记录覆盖普通透视表、独立透视图和关联透视图共用的 Show Values As 引擎；命令、Facade、字段面板和图表显示的跨模块验收由对应主任务补充。没有执行测试、类型检查、构建或浏览器验证。

## 来源与版本边界

- 固定 OSS：`E:\rust_test\univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。
- v1.6.0 [set-pivot-show-data-as.operation.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/types/commands/operations/set-pivot-show-data-as.operation.d.ts)：规则、参数模式、基准字段筛选和包含成员类型/标签/格式的基准项 DTO。
- v1.6.0 [sheets-pivot-ui/lib/es/index.js](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/es/index.js)：`li` 数组明确列出下表 15 项；`di` 取配置中的行/列字段，对指定父项模式排除叶子字段；`hi` 校验字段和具体项/相邻项参数。该文件以片段读取，未执行发布包。
- v1.6.0 [PivotShowDataAsDialog.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-ui/lib/types/views/components/PivotShowDataAsDialog.d.ts)：真实规则通过异步确认回调写入报告。
- v1.6.0 [sheets-pivot-chart/types.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/types.d.ts)：字段设置包含 `showDataAs`；[f-pivot-chart.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart/lib/types/facade/f-pivot-chart.d.ts) 第 59–67 行给总计百分比 DTO 示例，第 242–259 行声明 `getShowDataAs/setShowDataAs/resetShowDataAs`，读取要求防御性快照。
- v1.6.0 [pivot-chart-field-view.d.ts](../../.research.local/univer-1.6.0/packages/sheets-pivot-chart-ui/lib/types/views/components/pivot-chart-field-view.d.ts)：透视图字段面板的对应入口来源；不以存在入口证明计算完成。
- 历史 [engine-pivot/package.json](../../.research.local/univer/packages/engine-pivot/package.json) 明确为 **1.5.2**；其 [interface.d.ts](../../.research.local/univer/packages/engine-pivot/lib/types/types/interface.d.ts) 定义 `IPivotTableShowDataAsInfo` 的 `type/baseFieldId/baseItem/baseItemType?`，其 [enum.d.ts](../../.research.local/univer/packages/engine-pivot/lib/types/types/enum.d.ts) 另外包含 `percentOfParentGrandTotal`，以及 item/previous/next。不能把这些历史数值枚举或额外模式当成固定 v1.6.0 枚举已核实的证据。
- 已在此前工作中读取 [Microsoft 官方值显示方式说明](https://support.microsoft.com/en-us/office/show-different-calculations-in-pivottable-value-fields-014d2777-baaf-480b-a32b-98431f48bfec)：提供百分比、指定项、父项、差值、累计、排名和指数的公开公式。此说明不能证明本地边缘政策或真实 XLSX/Office 运行结果。

不安装或依赖商业 SDK。固定版包装层可证明规则名称和可观察参数契约；缺少固定版私有计算内核时，用上述 Office 公开公式独立实现，保留全部 Pro 范围及未核实项。

## 固定 15 项与本地计算

以下是规则清点，包含 7 项无参数、5 项仅基准字段、1 项具体基准项、2 项具体或相邻基准项。不是按钮计数或通过率。

记 `C` 为当前值字段实例的原始聚合，`G` 为筛选后的报告总聚合，`R`、`T` 分别为保留当前行/列条件的原始聚合，`B` 为替换基准字段成员后的原始聚合，`P` 为父级范围聚合。所有分母都由原聚合器按同一汇总类型重新计算；平均值等非可加汇总不取可见单元格之和。

| 固定类型 | 参数 | 实际计算路径 |
| --- | --- | --- |
| `normal` | 无 | 原始 `C`；存储规则清除 |
| `percentOfGrandTotal` | 无 | `C / G` |
| `percentOfRowTotal` | 无 | `C / R` |
| `percentOfColumnTotal` | 无 | `C / T` |
| `percentOfParentRowTotal` | 无 | 删除当前有效行维度中的最深一层，保留列条件，`C / P` |
| `percentOfParentColumnTotal` | 无 | 删除当前有效列维度中的最深一层，保留行条件，`C / P` |
| `index` | 无 | `(C × G) / (R × T)` |
| `percentOfParent` | 非叶子行/列基准字段 | 保留该轴直到所选字段的前缀，移除其后代条件，保持另一轴，`C / P` |
| `percentOf` | 行/列基准字段、具体项 | 保留其他条件并替换基准成员，`C / B` |
| `differenceFrom` | 行/列基准字段、具体项或 previous/next | `C - B` |
| `percentDifferenceFrom` | 行/列基准字段、具体项或 previous/next | `(C - B) / B` |
| `runningTotal` | 行/列基准字段 | 当前显示成员顺序中，从首项到当前项的原始 peer 聚合之和 |
| `percentRunningTotal` | 行/列基准字段 | 当前累计 / 同范围全部 peer 原始聚合之和 |
| `rankAscending` | 行/列基准字段 | `1 +` 同范围中小于当前值的数值 peer 数量 |
| `rankDescending` | 行/列基准字段 | `1 +` 同范围中大于当前值的数值 peer 数量 |

历史兼容项单列：`percentOfParentGrandTotal` 无参数，本地同时去掉行、列各自最深的有效维度后计算百分比，标记 `compatibilityOnly: true`。历史类型只证明该名字存在，**该同时去维度语义尚未用固定版内部实现或 Office 样本核实，不计入 15 项对齐数**。

本地 DTO 使用上述语义字符串；不接受历史 numeric enum 值。无参数规则要求 `baseFieldId/baseItem` 都是空字符串且无 `baseItemType`；字段模式要求真实轴字段，禁止多余基准项；具体项模式要求非空键；相邻模式只适用于两种差值，要求空 `baseItem`。本地兼容 `baseItemType: 'item'` 并归一为默认具体项。固定包装层只接受其对应参数组合，本地不宣称二进制枚举兼容。

## 共享权威计算与身份

模块归属为 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture) 的 `SHEET-23 / SHEET-24 → SHEET-22`。领域类型在 `domain/model.ts`；参数规则和聚合后变换在 `features/pivot/services/pivot-show-data-as.ts`；唯一来源读取、分组和汇总仍是 `pivot-engine.ts`。`pivot-chart-engine.ts` 只复用该引擎编排字段编辑、报告所有权和原子发布。

`buildPivotOutput` 按 value instance ID 与 scope 缓存原始聚合结果，再调用变换器。变换器只接收原引擎的 `aggregate(scope)` 和有序 `baseItems(fieldId, scope)`；不读取 source、不实现第二套聚合器。隐藏小计/总计不改变百分比分母，原 drill-down map 保留真实来源成员而非转换后的显示值。

目标 `fieldId` 与基准 `baseFieldId` 有不同的身份契约：

- 目标字段读写沿用既有 **canonical source ID 优先** 解析；一个来源有多个已配置实例时必须提供不歧义的 report-local ID。图表读写共用 `getPivotChartConfiguredField` 的解析器；未知/未配置读取返回 undefined，歧义读取拒绝，写侧未知仍拒绝。
- `baseFieldId` 按公开规则指向 **report-local 行/列实例**，精确 local ID 优先，唯一 canonical source ID 仅作 fallback；保存归一后的 local ID。否则另一轴 source ID 与已保存 local ID 碰撞时，刷新会错误改换基准。
- `getPivotShowDataAs` 返回规则拷贝。基准项查询取真实来源经现有手工/数值/日期分组后的成员；与读取 raw source 的 label-filter 查询分开。具体键发生真实 scalar 类型碰撞时拒绝，不静默选择某一种类型。

## 当前错误、顺序与格式政策

累计、相邻比较和排名沿用现有行/列成员顺序，同轴祖先条件限制 peer 范围；其他 scope 条件继续保留。升降序标签排序会改变累计和相邻关系。排名使用 competition ties，例如 `1,1,3`。

差值的具体基准单元格为空；不存在相邻项、筛选后或当前上下文中缺少基准项返回 `#N/A`。不含所选基准成员的小计/总计不参与具体项比较或排名，返回空值。累计模式在没有基准成员的总计范围累计全部可用 peers，累计百分比的分母是同范围全部 peer 聚合之和。

分母为零返回 `#DIV/0!`；非有限操作数或计算溢出返回 `#NUM!`，包括有限原始数值相减的溢出。找不到当前有序成员时返回 `#REF!`。原始空值保持为空，非数值原始聚合值保持其 scalar；累计/排名当前仅选取数值 peers，非数值 peer 的 Office 对应政策仍待样本确认。参考项被后续过滤或来源刷新移除时保留规则和错误诊断，不自动恢复普通显示。

百分比默认格式为 `0.00%`，排名为 `0`，显式 authored format 优先；重置移除派生格式但保留显式格式。图表源传播、仅错误行/系列的保留及 invalid-value strategy 由主任务在现有图表投影链路处理，不能把数值计算草稿当作图表标签已验证。

上述边缘政策已实现并有草稿，但缺少真实 Office 样本。空字符串具体成员当前被参数校验拒绝；基准项 UI 的成员类型、格式化显示标签、过滤查询 DTO 与上游完整结构仍未全部对齐。不得把这些差异改写为最终范围缩减。

## 接口与编辑闭环

共享服务新增：

```ts
setPivotShowDataAs(workbook, pivot, fieldId, showDataAs: PivotShowDataAsInfo | null,
  mode: PivotRefreshMode = 'worksheet'): void
getPivotShowDataAs(pivot, fieldId): PivotShowDataAsInfo | undefined
getPivotShowDataAsBaseItems(workbook, pivot, fieldId): string[]
```

`setPivotFieldSetting` 扩展 `showDataAs?: PivotShowDataAsInfo | null`，先校验组合编辑中的规则与聚合参数，再更新聚合/格式/规则并按所选模式刷新。等价规则按 primitive 属性比较，DTO 键顺序、canonical 基准别名、显式 item 默认值和 normal/null reset 归一后不引入额外 revision。

图表服务扩展 `PivotChartFieldSetting.showDataAs` 与真实编辑 `{ type: 'setShowDataAs', fieldId, showDataAs }`，提供 `getPivotChartShowDataAsBaseItems` 和 `getPivotChartConfiguredField`。完整字段配置也能设置或重置每个 value instance 的规则。沿用 candidate staging，最终 owned 用 data-only、linked 用 worksheet，一次编辑刷新和增加 report revision 一次。deferred 只在既有批量编辑中使用；没有另存第二份派生矩阵。

存活规则的基准字段不能单独删除或移动到 filter/value；value 转标签清理其自身规则。完整 reset 或一次完整 config 可以同时清理依赖和移除基准。非法参数和 linked output conflict 在图表候选上拒绝，保留已发布报告、cells 和来源 mirrors。共享 worksheet setter 的事务性仍由现有内容命令边界负责。

主任务接入 `FPivotTable/FSheetPivotChart` 的 get/set/reset 和基准项查询、`set-pivot-show-data-as` 内容命令及字段设置 DTO；共同候选阶段处理 selection/cache fan-out、linked source/revision、权限、历史和 `PivotChartChanged`。本批不再次运行该扇出，也不编辑上述 Facade/Studio/Panel 文件；集成效果尚未执行验证。

## 文件、草稿与未验收项

本批实际文件：

- [model.ts](../../packages/office-studio/src/domain/model.ts)：只修改 Pivot 规则类型与字段片段，保留其他在途模型改动。
- [pivot-show-data-as.ts](../../packages/office-studio/src/features/pivot/services/pivot-show-data-as.ts)：规则清点、参数归一、成员解析、格式和聚合后变换。
- [pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts)：共享 setter/query、scope 聚合、输出格式、原始下钻及依赖保护。
- [pivot-chart-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts)：owned/linked 编辑和共享目标解析。
- [pivot-show-data-as-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-show-data-as-fixture.mjs) 与 [pivot-show-data-as.contract.test.mjs](../../packages/office-studio/tests/pivot-show-data-as.contract.test.mjs)：独立真实 source 夹具和契约草稿。
- 本记录与 [上一批字段引擎记录](SHEET-24-pivot-chart-engine.md) 的后续状态链接。

夹具用 Region→Product、Year→Quarter 两个层级，来源从非零列开始；grand=200、Q1=70、North/Q1=40、North/A 行汇总=30。第一个真实 cell 的六个基础结果分别为 `.05`、`1/3`、`1/7`、`1/4`、`1/3`、`20/21`，通过原始数据而非硬编码显示矩阵表达契约。

静态清点 **34 条 test/test.each 声明**，其中三组 each 含 16、3、15 项，展开为 **65 个草稿用例**（`34 - 3 + 16 + 3 + 15`）。内部循环和多次断言不另计为运行用例。涵盖固定 15 项与历史兼容、隐藏总计/层级小计、行列累计/差值、ties、零/负值/数值溢出、AVG 原始分母、多个 measure、格式/no-op、三种 mode、owned/linked、原始下钻、三类分组、ID/scalar 碰撞、过滤后基准失效、依赖拒绝、非法 DTO、冲突原子性和快照重开。

按照当前“先实现、后统一测试”要求，**未运行测试、类型检查、前端构建、Rust/Cargo/WASM、dev、浏览器、Office 文件样本、性能基准或 git diff --check**。静态读写和声明清点不是执行验证；不能称为测试通过或功能已验收。

仍需真实闭环：

1. 真实 Office 样本核实基准缺失/空白、errors、subtotal、同轴后代上下文、非可加累计、ties 和分母政策；历史额外模式单独核实。
2. `antdv-next` 字段面板/基准选择及菜单、API、快捷键、权限、内容历史、linked fan-out、图表格式和错误策略的集成/浏览器验证，覆盖主题、焦点与 IME；本批无 UI 改动，不能代替控件迁移验收。
3. parser/writer 原生 pivot/showDataAs OOXML、pivotCache/chart relationship 的保真读写和真实 Office 保存/重开；内部 JSON 快照草稿不能证明 XLSX 往返。
4. retained Rust/Worker/WASM 大数据路径和同数据集基准。当前继续使用同一 TypeScript 聚合器，没有新增 Worker 协议；后续遵守 [WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract) 的协议版本、取消、预算、lease 和复制统计。本批没有性能实测，不能声称优化、零拷贝或达标。

SHEET-23 / SHEET-24 及所有 Pro 能力仍在完整交付范围，本批状态保持“未验收”。
