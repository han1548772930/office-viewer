# SHEET-11/29：复合筛选、通配符与共享读写边界

固定参考：Univer v1.6.0，`eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本批复用现有筛选面板、列候选、统一命令和内容历史，不新增常规 Ribbon 根入口。功能保持未验收。

## 固定源码与模型

[types.ts::ICustomFilters](../../../univer/packages/sheets-filter/src/models/types.ts) 明确只接一至两个子条件；[filter-model.ts::customFilterFnFactory](../../../univer/packages/sheets-filter/src/models/filter-model.ts) 在两条条件时按 `and` 选择 AND，否则 OR。固定 [custom-filters.ts](../../../univer/packages/sheets-filter/src/models/custom-filters.ts) 与 [escape.ts](../../../univer/packages/core/src/common/regexp/escape.ts) 是文本通配符来源；`*`、`?` 参与匹配，`~` 在该固定实现中是普通字符。

本地新增 `SheetCompoundFilter { kind: 'compound'; join: 'and' | 'or'; conditions: [SheetConditionFilter] | [SheetConditionFilter, SheetConditionFilter] }`，不允许嵌套 compound/value/color。现有 `SheetConditionFilter` 的类型化值、显式 null、第二操作数和本地大小写扩展完整保留。日期树只组织候选，选择结果仍保存原始 `selectedItems`，不新增日期树持久化模型。

Office公开语义另按2026-11-29查阅的微软[筛选说明](https://support.microsoft.com/en-gb/excel/get-started/filter-data-in-a-range-or-table-in-excel)、[文本筛选通配符](https://support.microsoft.com/en-au/excel/using-wildcard-characters-in-searches)和[Open XML CustomFilters](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customfilters?view=openxml-2.9.1)核对：文本模式可用星号/问号，波浪号用于转义，原生自定义筛选组合两条比较。本地固定Univer的字面波浪号与Office转义分别适配，不把两者说成同一规则。

## 生产闭环

- [model.ts](../../packages/office-studio/src/domain/model.ts)、[workbookAutoFilter.ts](../../packages/office-studio/src/parser/workbookAutoFilter.ts)：公开类型与解析 DTO 同步，检查 join、1–2 条数量、非嵌套条件、有限标量、操作数和布尔大小写标志，解码结果不共享输入子对象。
- [sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts)：复合条件复用现有条件谓词，以 every/some 求值；单条件和双条件使用同一通配符实现。面板、候选、权限、内容历史及日期树交互见 Sheets 子任务的专项记录。
- [sheet-clipboard-resource-validation.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-resource-validation.ts)：结构化表格资源允许颜色与复合条件，拒绝错误 join、嵌套和缺失操作数；完整合并组的过滤缓存允许越过ref末行，与IO保持相同绝对边界，并拒绝标题行和重复行。
- [chart-source-import-dates.ts](../../packages/office-studio/src/features/charts/chart-source-import-dates.ts)：跨工作簿日期依赖导入同时换算两条条件各自的操作数；非日期列与颜色条件保留，仍用统一日历转换，避免给颜色模型读取不存在的比较字段。
- [sheet_filter_import.rs](../../crates/office-parser/src/sheet_filter_import.rs)：worksheet/table 的项目快照均可恢复复合模型；标准 XLSX 两项 customFilters 可投影一般 AND/OR，既有 between/notBetween 与“空/非空同时成立”的兼容归一化保留。
- [xlsx_filter.rs](../../crates/office-writer/src/xlsx_filter.rs)、[xlsx_filter_compound.rs](../../crates/office-writer/src/xlsx_filter_compound.rs)：worksheet/table 共用验证与快照。可原生表示的两项条件写真正 `customFilters and="1|0"`，保留所有条件与当前过滤行，不使用按钮提示替代数据行为。
- native 文本写出保留 `*`、`?`，仅将本地字面 `~` 写成 Office 的 `~~`；ordering/interval 的普通文本仍按字面转义。导入支持一般 `*`/`?` 模式。标准 Office 的 `~*`/`~?` 表示字面通配符，当前固定条件模型没有等价的字面模式，遇到时保留完整 `sourceAutoFilter` 与实际隐藏行，明确留下适配缺口，不错误投影成通配匹配。
- [workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts)：外部文件的过滤/手动隐藏原因在所有 sheet、名称、日历及切片器组装后再对账，颜色条件可读取跨表引用；匹配的项目精确快照不无条件重算旧缓存。

## 原生表达限制与剩余

本地两条条件各自可含 between/notBetween 或 isError，但标准 Office 一列最多两个原生比较项，且没有同等动态错误类型条件。不能原生表示时，写出当前可见成员快照，给出 `AUTOFILTER_COMPOUND_SNAPSHOT`，完整表达式保留在项目扩展；Office 后续 reapply 不保证等价。大小写、混合标量、null 与文本 locale 的原生差异继续给出原诊断。没有将这些限制改写为本项目最终功能范围。

日期分组树、compound 与 wildcard 的运行、历史、浏览器和 Office 独立重开均未验收；标准 native dateGroupItem、任意字面 wildcard、Top10/动态筛选等适配与大型 Worker 筛选仍需继续推进。

本轮共享复核另确认既有值筛选的 `selectionMode`、`recordSelection`、`recordDate1904` 在TS仍保留，但Rust的Values变体目前只包含selectedItems，尚不能宣称这些扩展字段已完成原生Office往返；后续需同步扩展读写快照及相应原生表示/诊断。该既有缺口与本轮新增compound快照分别记录。

主线新增 [4 个共享边界契约草稿](../../packages/office-studio/tests/sheet-filter-boundaries.contract.test.mjs)，以及 [6 个 Rust 读写草稿](../../crates/office-parser/src/sheet_filter_roundtrip_tests.rs)，覆盖双条件完整重开、外部 join 失效指纹、原生表达限制、非法结构、wildcard/tilde 与无法投影的字面模式。另保留本轮两个 normalizer/cache 草稿，均未运行。WASM 构建例外与 Vite 修复验证独立记录在 [COMMON-11](COMMON-11-wasm-build-and-vite-import-2026-11-29.md)，不提升上述功能验收状态。
