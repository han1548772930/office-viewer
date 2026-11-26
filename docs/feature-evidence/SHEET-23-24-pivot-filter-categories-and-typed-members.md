# SHEET-23 / SHEET-24：日期、标签、多类别与原始成员筛选

记录日期：2026-11-26。状态：**核心实现和合同草稿已落盘，未验收**。

本批对普通报告和 owned 图表使用同一个透视聚合器，补日期、caption 条件、多类别互斥和 typed raw 成员投影。按用户先实现、后统一验证的要求，只做静态读取与编辑；未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`。下列用例数是草稿展开数，没有通过数。

## 固定来源与观察

固定 OSS：`E:/rust_test/univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 依据为项目 `.research.local/univer-1.6.0/packages/` 中公开发布的类型及 ES 文本；没有安装或执行商业 SDK。

| 具体来源 | 可观察契约及本地实现差异 |
| --- | --- |
| `engine-pivot/lib/types/date-filter.d.ts`、`types/enum.d.ts`；固定 ES 的 `S`、`setDateFilterInfo` | 8 个日期比较操作符，serial operands 与 `useWholeDay`。本地默认按精确 serial 比较，whole-day 使用 `floor`，不把 serial 60 改成其他日期；校验有限、有序 operands。未运行上游样本 |
| `engine-pivot/lib/types/record-filter.d.ts`；固定 ES 的 `se` | 14 个 caption 条件。使用大小写敏感、JavaScript ordinal 比较，通配符按字面文本处理；本地不借用表格原生条件的 locale collator |
| `engine-pivot/lib/types/pivot/pivot-table.d.ts`、`pivot/table-field.d.ts`；固定 ES 的 `fe/de`、`setOptions` | member、condition、value 按同一 canonical source 分类；date/caption 属于一个 condition 类别。显式关闭多类别且报告存在冲突时清整个报告，而完整快照缺少显式 flag 时保留原状态 |
| `sheets-pivot/lib/types/record-filter-bridge.d.ts`；固定 ES 的 `kr/Ar/jr/Mr` | checked/partial 显示状态与原始值筛选区分，包含 `getSlicerSelectionFromPivotManualFilter(... current, isAll)`。本地按真实 records 将 grouped/partial 投影到 typed raw selection，不把组名传播为原始成员 |
| `engine-pivot/lib/types/record-value.d.ts`、`record-filter.d.ts` | 原始成员区分类型。当前本地 authority 为 `CellScalar`，可区分 number/text、boolean/text、null/empty；日期及 error 的独立 typed cache 身份尚未补齐 |

本批没有新增顶部按钮。入口分类仍见[顶部来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)与[透视图 UI 记录](SHEET-24-pivot-chart-ui.md)，操作符数、helper 数或用例数不表示功能完成。

## 模块及协议

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，领域 DTO 位于 [model.ts](../../packages/office-studio/src/domain/model.ts)，纯谓词位于 [pivot-filter-predicate.ts](../../packages/office-studio/src/features/pivot/services/pivot-filter-predicate.ts)，聚合、drill-down 与 worksheet materialization 继续使用 [pivot-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-engine.ts)。[pivot-chart-engine.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-engine.ts) 对 owned 完整字段配置应用相同的 options/filter 规范化；没有第二套公式或透视引擎。

- 日期 capability 使用真实来源值、number format、Excel 1900/1904 系统及 ISO 适配，不按字段名猜测。
- 默认 `allowMultipleFilters=false`：设置新类别清同来源其他类别；显式 true 保留交集。重复 report-local 实例使用 canonical source 身份判断。
- 日期 reset 只清日期条件；clear-all 清同来源所有实例的所有类别。显式 false 且任一来源已有多个类别时清整个报告；不同来源各自一个类别保持。
- date/caption/membership 谓词共用于聚合与 drill-down；value 筛选在已通过其他条件的 records 上按 scoped value aggregation 判断。
- `getPivotSourceRecordSelection(workbook, pivot, fieldId)` 返回防御性 `PivotRecordSelection | undefined`。直接 raw include/exclude 保留已删除但明确指定的值；grouped/partial 根据真实来源投影并保留类型。`getPivotSourceSelectedItems` 仅由该结果派生字符串 mirror。
- 完整 owned 配置缺少 `allowMultipleFilters` 时保留快照中的多类别，不把重开误作用户显式关闭。规范化条件的重复设置保持 no-op。

`PivotRecordSelection` 是 `{ mode: 'all' } | { mode: 'include' | 'exclude'; values: CellScalar[] }`。`include` 空列表过滤全部，`all` 与 `exclude` 空列表清 member filter。本批没有变更 Worker/WASM 或 Office writer 协议；API、命令、UI 投影由主任务接入，不以这些服务 exports 代替完整交付。

## 未运行草稿

[pivot-filter-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-filter-fixture.mjs) 使用真实稀疏数据范围：Data 行 2..9、列 3..7，7 条 records，总金额 280。样本包含同一天两种时间、1900/1904、无效日期、大小写不同的 North/north、number `1` 与 text `"1"`、boolean `true` 与 text `"true"`、null 与 empty string，另有 General numeric 字段 `Date count`。

[pivot-filter.contract.test.mjs](../../packages/office-studio/tests/pivot-filter.contract.test.mjs) 静态清点 **19 条 test 声明、86 个计划用例，全部未运行**：日期参数化为 2 个日期系统 × 2 种报告所有权 × 8 个操作符；caption 为 2 × 14；另含非法参数原子性、来源 capability、drill-down、canonical 重复实例、互斥/交集、日期专属 reset、全报告清理、grouped/partial、no-op 与 JSON 重开。

切片器跨报告的 typed fan-out、原生 Table 与命令历史草稿独立记录在[typed 切片器证据](SHEET-25-slicer-typed-members.md)，不重复计数。

## 未完成与未核实

1. 日期/error 的独立 typed identity、上游完整 date filter UI、全部 Office 日期/错误样本尚未验收。当前 `CellScalar` 原始成员不能表示上游全部 RecordValue 类型。
2. 原生 XLSX pivot cache、filter parts、Office 保存重开和续算保真未补齐；本批 JSON 草稿不等同于 Office 文件验收。
3. 计算当前仍在主线程 TypeScript；Worker/Rust/WASM、任务取消、typed buffers、lease、复制字节及内存契约留待真实协议接线。按[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)验收前，不声称零拷贝或性能达标。
4. 没有运行浏览器、主题、焦点、IME、大样本或上游演示。本批核心没有新增 UI 控件；UI 侧 `antdv-next` 和 ConfigProvider 仍需统一验收。

所有缺口均为完整 Sheets / Docs / Pro 目标内的后续工作，不缩减最终范围。
