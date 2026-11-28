# SHEET-18 / SHEET-29 原生表格身份、筛选与排序范围

日期：2026-11-28。开发进度：待验证；验收状态：未验收。本记录区分已提交的身份映射与本次续接的实际修改，不以 XML 字段或入口数量证明完整表格能力。

## 来源与语义

固定 Univer v1.6.0 提交为 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`：

- [table-column.ts](../../../univer/packages/sheets-table/src/models/table-column.ts)：列 ID、公式/数组标志、汇总标签/函数/公式、筛选按钮及序列化。
- [table.ts](../../../univer/packages/sheets-table/src/models/table.ts)：列 ID 和列顺序分别保存，移动/删列后的资源不能只按当前位置匹配。
- [table-filter.ts](../../../univer/packages/sheets-table/src/models/table-filter.ts)：列筛选和排序状态独立于源数据。
- [auto-fill.command.ts](../../../univer/packages/sheets/src/commands/commands/auto-fill.command.ts)：填充/清除通过值 mutation 和选区 operation 实施，不按填充范围缩减工作表尺寸。

Office 原生文件语义另外对照 Microsoft [Table](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.table?view=openxml-3.0.1)、[SortCondition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sortcondition?view=openxml-3.0.1) 和 [CT_SortCondition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/af08b870-9b06-4812-a4d6-7bc80d819e62)（查阅日期同上）。`totalsRowCount` 决定当前汇总行数，`totalsRowShown` 记录历史；排序键是排序区域内的一行或一列。上游功能合同与 OOXML 映射分别记录，不声称 Univer 内部使用本项目写出实现。

## 已有身份边界

续接基线已包含 `officeTableId` / `officeColumnId`、完整 `officeSource` XML，以及 calendar source 中的表/列稳定 ID 映射。原生 ID 分配预留全部源列 ID，避免新列复用已删除列的公式/扩展。日期计算镜像清除作者来源与原生公式，并分配独立表 ID。TypeScript 恢复表/列身份时同步更新 slicer 与 cache 字段，先准备全部计划，再修改消费者。

这些共享链路位于 [sheet_table_source.rs](../../crates/office-io-common/src/sheet_table_source.rs)、[xlsx_table.rs](../../crates/office-writer/src/xlsx_table.rs)、[xlsx_formula_calendar.rs](../../crates/office-writer/src/xlsx_formula_calendar.rs)、[table_sidecar.rs](../../crates/office-parser/src/table_sidecar.rs)、[workbookTableIdentity.ts](../../packages/office-studio/src/parser/workbookTableIdentity.ts)。本次补写了相应草稿，基线代码的存在不等于已经通过 Office 往返验收。

## 本次续接修改

原生 table 与 autoFilter 内的 `sortState.ref` / `sortCondition.ref` 现在通过 [xlsx_table_sort.rs](../../crates/office-writer/src/xlsx_table_sort.rs) 映射到当前表格位置和数据行边界。纵向排序键按原生列 ID 追踪，删除列对应的键被移除；横向排序保留行定位。大小写、排序方法、升降序、自定义序列、颜色格式 ID 和未知扩展随来源保留。无数据的标题单行表不输出对标题的排序。映射不会重新排序单元格，实际排序继续由命令修改权威模型。

替换可编辑筛选条件时，不再直接丢弃原生 filter 节点：保留根/列扩展及按钮属性，按列 ID 重映射后替换已知筛选谓词；显式显示按钮同时更新 `hiddenButton` 和 `showButton`。筛选范围包含标题和数据，排除实际汇总行。显式 `autoFilter: null` 继续表示清除，缺省表示保留未投影的原生条件。

Parser 不再把 `totalsRowShown=true` 当作当前存在汇总行。Writer 隐藏汇总行时写 `totalsRowCount=0`，同时保留原生历史标志。这样最后一行数据不会在下一次导入/筛选时被误当作汇总行。

table/parser/sort 共用有界本地 A1 范围解析，接受合法绝对引用，拒绝反向、跨表、多区域及轴越界引用。旧表格写出草稿改为语义 XML 属性检查，避免依赖属性序列。

本次另将 `sheet.command.configure-filter` 接入与异步排序相同的请求 owner 和版本检查：扩展事件执行前、候选准备与发布边界检查上下文。绑定使用 WeakMap，未向持久化 DTO 或 Worker 协议添加函数。

## 草稿与检查边界

- [xlsx_table_tests.rs](../../crates/office-writer/src/xlsx_table_tests.rs)：原生身份、删列/新增、公式 null 清除、样式标志、扩展、筛选按钮、移动/纵横排序、空数据表与独立镜像。
- [formula_calendar_roundtrip_tests.rs](../../crates/office-parser/src/formula_calendar_roundtrip_tests.rs)：作者原生表/列 ID 与计算镜像 ID 分离、原生公式不覆盖镜像单元格、manifest 与实际 part 身份一致。
- [table_sidecar.rs](../../crates/office-parser/src/table_sidecar.rs)：二进制导出/重解析草稿保留公式/数组、样式、稳定身份，并验证已隐藏汇总行的历史标志不增加实际行。
- [workbook-table-identity.contract.test.mjs](../../packages/office-studio/tests/workbook-table-identity.contract.test.mjs)：重排、外来 ID 碰撞、来源缺失、旧 manifest、切片器/cache 字段映射与失败原子性。
- [sheet-filter-command-guard.contract.test.mjs](../../packages/office-studio/tests/sheet-filter-command-guard.contract.test.mjs)：命令观察者失效 owner、嵌套修改保留、一次历史与内部快照重开。

以上均为未运行草稿。TypeScript/JavaScript 语法与本地 import 检查属于只读静态检查；Rust/WASM 检查、类型检查、测试、构建、dev、浏览器、Office 独立重开和性能基准全部后置。本次没有性能测量或零拷贝结论。

完整表格排序/计算列交互、x14 排序扩展、结构编辑后的完整原生公式/绑定变换、外来 Office 修改、差异样式资源引用与复杂样本仍需继续实现或验收。非完整表宽的排序区域经换列变成不连续范围时当前明确拒绝导出，尚需展开完整互操作行为；不输出指向错误区域的排序元数据，也不据此缩减最终范围。
