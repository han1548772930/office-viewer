# SHEET-11：日期候选树与复合条件

日期：2026-11-29。固定来源 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。续接[筛选表头/颜色/合并记录](SHEET-11-filter-header-color-2026-11-28.md)，复用现有全部入口清点和筛选命令。本批状态为**已落盘、未验收**。

## 固定来源与合同

| 来源 | 已核实行为 |
| --- | --- |
| `packages/sheets-filter-ui/src/worker/generate-filter-values.service.ts::getFilterTreeByValueItems` | 日期格式且包含年/月/日部分的候选，按年→月→日分组；同日保留多个 `originValues`，按原始行记录计数；纯时间不生成日期树；月份/日期按数字排序 |
| `packages/sheets-filter-ui/src/services/util.ts::searchTree` | 搜索按空白拆关键词、忽略大小写，关键词间 OR；父节点匹配保留其全部后代，子树匹配保留祖先路径并重算当前计数 |
| `packages/sheets-filter-ui/src/services/sheets-filter-panel.service.ts::ByValuesModel` | 组节点操作影响叶节点；提交把选定日期叶的 `originValues` 展开成普通值筛选，不新增持久日期树；候选忽略本列条件、排除其他列已拒绝的行；从有实际拒绝行的颜色/条件切到值筛选时清空勾选 |
| `packages/sheets-filter/src/models/types.ts::ICustomFilters` | 一条或两条 customFilters，`and:TRUE` 表示 AND，缺省表示 OR；不嵌套 customFilters |
| `packages/sheets-filter/src/models/filter-model.ts::customFilterFnFactory` | 一条原谓词直接使用；两条短路 AND/OR，仍使用同一条件工厂和原始单元格值 |
| `packages/sheets-filter-ui/src/models/conditions.ts::FilterConditionItems.CUSTOM` | 两行独立条件及 AND/OR；两行 None 清除列条件，一行 None 则使用另一行的单条件 |
| `packages/sheets-filter-ui/src/views/components/SheetsFilterByConditionsPanel.tsx` | 自定义条件编辑器复用原字段、Radio AND/OR 与比较值输入；不是新增 Ribbon 常驻根 |
| `packages/sheets-filter/src/models/custom-filters.ts`、`packages/core/src/common/regexp/escape.ts::createFromWildcard` | 文本匹配锚定，`*` 任意数量字符、`?` 一个字符；其他正则字符按字面，缺省忽略大小写；固定版 `~` 是普通字符，不冒称它已有 Office 转义语义 |

常驻 Ribbon、扩展、溢出与上下文根增量均为 **0**。本地沿用前批 Sheets 常规根 **72** 的分类口径；本批仅增加现有筛选面板中的“自定义复合条件”模式和运行态日期树节点。日期树节点数量随数据变化，不能算新增函数或功能验收数量。

## 真实实现

[sheet-data-panel-state.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-data-panel-state.ts) 在已有合并列遍历和其他列筛选结果上派生候选，增加实际记录计数和可选日历日。日期识别使用已有 `resolveSheetCellStyle`、`classifySheetNumberFormat` 与 `excelSerialToDateTimeParts`，支持继承的日期格式和工作簿 1900/1904 日历；纯时间、持续时间、文本、无效序号和 rich body 继续作为普通值。原始 `CellScalar` 和筛选 key 不改，格式化文本仅用于搜索和显示。

[sheet-filter-value-tree.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-filter-value-tree.ts) 是独立纯服务，负责年/月/日树、同日多时间集合、组计数、月名 locale、搜索路径、完整/部分勾选和整组开关。键使用独立 `date:`/`value:` 命名空间，普通文本“2024”和日期年节点不冲突。日节点可处于部分勾选，保留同日不同原始时间的独立身份。

[SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue) 用 `antdv-next` Tree 呈现含日期的候选列，复用 Form、Input、Button 和主题；普通列保留原 Checkbox 分页。Tree 使用真实原值集合控制 checked/halfChecked，搜索、全选/全不选和手动展开仅改变草稿；确认时把当前选中叶值展开成 `SheetValueFilter.selectedItems`，不保存组件树。已有 include/exclude 值规则按同一已注册谓词回显勾选，不能把排除日期显示为已选。

共享模型由主线增加 `SheetCompoundFilter { kind:'compound'; join:'and'|'or'; conditions:[SheetConditionFilter] | [SheetConditionFilter,SheetConditionFilter] }`。本子任务的 [sort-filter-engine.ts](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts) 在唯一现有筛选引擎中校验 1–2 条、合法 join、非嵌套 condition、有限标量和布尔 caseSensitive，再用原 `matchesCondition` 短路合成。数据、样式、手动隐藏和范围资源保持原模型，未另建权威引擎。

[sheet-filter-condition-draft.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-filter-condition-draft.ts) 负责显式类型/第二边界/大小写草稿与领域条件转换。[SheetFilterConditionFields.vue](../../packages/office-studio/src/components/spreadsheet/SheetFilterConditionFields.vue) 用 Ant Form.Item、Select、Input、Checkbox 呈现每条条件；父面板用 RadioGroup 连接两行。无效数值在提交前失败并保留两行草稿。None 的降阶/清除与固定来源相同。

[sheet-filter-wildcard.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-filter-wildcard.ts) 为现有字符串 equals/doesNotEqual、contains/doesNotContain、beginsWith/endsWith 补上固定 `*`/`?` 匹配，同一规则可作为复合子条件。保留本地已公开的每条 `caseSensitive` 扩展。正则元字符作为字面，不将任意输入当正则表达式。

面板继续调用已有 `sheet.command.configure-filter`，两条条件、日期集合及其他列条件一次发布。打开、搜索、分组开关、取消和失败均不写内容历史；权限/内容版本/owner/取消复用前批 `useSheetSortFilterWorkbench` 与主线统一 filter command context。组字期间保留 Ant Input，禁用改变规则、列、连接方式和提交，Enter 不重新载入并清空字段。共享 parser、native 快照、Rust customFilters/Office 来源由主线接入，不把本子任务的前端草稿当作 Office 验收。

## 草稿与静态检查

本子任务新增 **19 个 test 声明**：

- [sheet-filter-value-tree.contract.test.mjs](../../packages/office-studio/tests/sheet-filter-value-tree.contract.test.mjs)：7 条，覆盖同日多时间、部分选择、父/子树搜索、继承日期格式、时间/持续时间排除、1900 伪闰日/1904 纪元、合并计数及 include/exclude 回显。
- [sheet-filter-compound.contract.test.mjs](../../packages/office-studio/tests/sheet-filter-compound.contract.test.mjs)：6 条，覆盖 AND/OR/空白/错误、不可变捕获、非法嵌套和数目、None 降阶、类型/二边界恢复及通配符。
- [sheet-data-panel.contract.test.mjs](../../packages/office-studio/tests/sheet-data-panel.contract.test.mjs)：追加 3 条，覆盖复合条件一次历史/撤销重做/native 重开，日叶展开原时间值、重开重建日期树，以及公开值筛选拒绝对象/非有限值/undefined/稀疏成员且保持原 filter、完整模型和 undo/redo。有效的 null/string/boolean/finite number 及 selectionMode 保持。
- [sheet-data-toolbar.spec.ts](../../tests/browser/sheet-data-toolbar.spec.ts)：追加 3 条，覆盖真实双条件 Form 失败草稿/AND/OR/回显，日期树取消/搜索/原值提交，以及 composition/Enter 保留草稿。

上述草稿**全部未运行**。已对 11 份代码/草稿进行 TypeScript `createSourceFile` 与 Vue parse/compileTemplate 的只读语法解析，11 段脚本、2 份模板、47 条本地引用路径无错误；其中发现并修复新字段组件一处动态 label 的反引号缺失。本批没有执行测试、类型检查、前端构建、浏览器、Office 或性能。主线应在交接后记录最终全局静态结果；主线针对用户报告执行的 Rust/WASM 打包不由本记录代替或重复。

主线复核后，`normalizeCriteria` 的 values 分支补充逐成员有限标量验证，使用迭代器同时拒绝稀疏数组中的 undefined，不只验证数组外壳。原 selectionMode、recordSelection 与日期上下文保留；拒绝发生在候选发布和内容历史之前。该追加的生产/草稿两文件再做只读语法检查，运行验证仍为 0。

## 差异与继续项

- 固定日期树代码以原始值 truthy 判断排除 serial 0；本地按 Office 日期系统和项目已有序号合同，保留 1904 serial 0→1904-01-01，1900 serial 0 的日 0 保持普通数值。该差异单独保留，不称完全复刻该上游边界。
- 日期树是候选的展示分组。本地保留数字/文本/逻辑值身份，故底层规则仍按项目已有 typed 值语义过滤；不能用分组完成推断所有 Univer 字符串/空值/数值转换完全相同。
- 本地复合子条件允许已公开的 between/notBetween/isError 等既有操作符；固定自定义编辑器二级列表排除需要两个参数的 between/notBetween。复杂组合无法对应两条原生 Office customFilter 时，完整规则依赖项目精确快照及明确来源诊断，不能声称所有组合都原生等价。
- 固定 `~` 字面与 Office `~*`/`~?` 转义有差异；主线负责保留无法投影的原生来源/实际隐藏行。完整 literal 通配符模型和真实外部编辑往返仍需后续实现/验收。
- 固定无勾选时不允许提交值筛选，而本项目原有 API/面板支持空集合隐藏全部数据；该既有行为本批保留，不以日期树改写兼容契约。颜色/条件→值的拒绝行初始状态已按来源清空勾选，原显式全不选仍可提交。
- 当前仍用现有 Modal 承载面板；固定列旁浮层、完整条件转换边界、结构化表专属入口、复杂 Pro 交互与真实 Windows IME/主题/权限/Office 场景保持未验收。
- 候选和组合谓词复用现有 TypeScript 路径；大型数据的 Worker/Rust 热点、取消与真实性能基准继续保留，不宣称已优化或已达标。

该批不提升 SHEET-11 或其他 65 组的验收状态，完整 Univer/Pro 目标继续推进。
