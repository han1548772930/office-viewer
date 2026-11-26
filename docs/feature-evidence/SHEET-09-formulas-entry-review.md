# SHEET-09 FORMULAS 顶部入口续核

状态：2026-11-27 完成固定来源与本地链路静态核对；参数函数菜单接回既有表单，补文本/日期图标、无工作表时计算模式禁用及公式选区权限接线。随后仅将已有真实求值的 `NA()` 接入信息分类，再修复 `RIGHT(text,0)` 返回整串的问题。契约及浏览器草稿未运行，功能未验收。固定来源为 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。参数和 NA 入口两批沿用既有输入链路；RIGHT 小修只改现有权威引擎的零长度分支，未变更 Worker/WASM 协议或数据模型。

## 固定入口与计数

依据为 [formula schema](../../../univer/packages/sheets-formula-ui/src/menu/schema.ts)、[menu factories](../../../univer/packages/sheets-formula-ui/src/menu/menu.ts) 和 [中文文案](../../../univer/packages/sheets-formula-ui/src/locale/zh-CN.ts)。各函数分类均派发 `formula-ui.operation.insert-function`，菜单标识为该 ID 加下表后缀，参数为具体函数名称。

| 分组 | order | 固定标签 | ID 后缀或完整 ID | Grid 位置 |
| --- | --- | --- | --- | --- |
| BASIC | 0 | 常用函数 | `.common` | r1c1，跨两行，不显示标签 |
| BASIC | 1 | 财务 | `.financial` | schema 未指定，自动排布 |
| BASIC | 2 | 逻辑 | `.logical` | schema 未指定，自动排布 |
| BASIC | 3 | 文本 | `.text` | schema 未指定，自动排布 |
| BASIC | 4 | 日期与时间 | `.date` | schema 未指定，自动排布 |
| BASIC | 5 | 查找与引用 | `.lookup` | schema 未指定，自动排布 |
| BASIC | 6 | 数学与三角函数 | `.math` | schema 未指定，自动排布 |
| BASIC | 7 | 统计 | `.statistical` | schema 未指定，自动排布 |
| BASIC | 8 | 工程 | `.engineering` | schema 未指定，自动排布 |
| BASIC | 9 | 信息 | `.information` | schema 未指定，自动排布 |
| BASIC | 10 | 数据库 | `.database` | schema 未指定，自动排布 |
| OTHERS | 0 | 计算选项 | `sheets-formula-ui.operation.calculation-options` | r1c1，跨两行两列，显示标签 |

classic 按相同分组和 order 排列。常用下拉的静态列表是 `SUMIF / SUM / AVERAGE / IF / COUNT / SIN / MAX`；description service 可按已注册函数过滤。其他分类读取 description service 的类型目录，不应把静态根数视作完整函数数。common 内另有 `formula-ui.operation.more-functions` 的“全部函数”子入口，order 0，不能另算常驻根。

函数根及“全部函数”复用 `getInsertFunctionDisabled$`：工作簿可编辑、工作表可编辑/可设置单元格值、范围编辑权限共同控制禁用；`getCurrentRangeDisable$` 也提供当前上下文状态。所有根按当前 Sheets 实例控制显隐。计算选项根仅配置 Sheets 显隐，表单内部另检查工作簿编辑权限。

本次统计范围限定为上述 formula-ui FORMULAS schema：固定常驻根 **12**（11 函数分类 + 1 计算选项），源码另注册的“全部函数”子入口 **1**；函数列表按实际服务数据生成。窄屏溢出是这些根的复用。formula-ui 的复制/粘贴扩展属于 START 下的特殊菜单，不计入这 12 根；本记录未重新清点其他插件、上下文或 Pro 公式入口。

## 本地入口与已完成小批

[sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的 `formula-edit` 组在 NA 批之前提供 **8** 根；接入 NA 后依序提供常用与统计、财务、逻辑、文本、日期与时间、查找与引用、数学、统计、信息共 **9** 根。common 仍为 r1c1 跨两行。分类按目录支持集合过滤，信息目前只有 NA；工程/数据库仍没有目录项。

[现有函数目录](../../packages/office-studio/src/features/sheet-workbench/sheet-function-catalog.ts) 从 **117** 增至 **118** 项：常用与统计 6、财务 17、逻辑 11、文本 21、日期与时间 26、查找与引用 11、数学 9、统计 16、信息 1。它们形成 118 个函数子项，common 底部“插入函数”再形成 1 个子项；这是代码数量，未逐函数验收。本地常用列表与固定版不同，SIN 未列入。

`calculation-options` 组提供 **3** 根：计算模式（r1c1 跨两行两列）、编辑公式（r1c3）、重新计算（r2c3）。本地 FORMULAS 常驻根从 **11** 增至 **12**。参数入口批的根增量为 0；NA 批常驻根 +1、函数子项 +1，扩展、溢出和上下文新增根均为 0，窄屏继续复用现有根。本地 12 是 9 分类 + 3 计算入口，固定 12 是 11 分类 + 1 计算选项，根数相等不代表分类或功能齐全。

已完成的小批修复为：

- 文本和日期分类分别复用已导入的 `TextCursor`、`CalendarDays`，对应固定版 `TextFunctionIcon`、`DateFunctionIcon` 的分类语义；未复刻图标像素。
- 计算模式 Select 补 `disabled: !state.hasSheet`，保持现有 automatic/manual 值和命令；其他根及子项已有相同无工作表禁用。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 `sheet.operation.insert-function` 先读取目录项。需要参数的函数预选实际名称并打开既有分类表单；原来向构造器传空参数会抛错且面板未打开。TRUE、FALSE、NOW、TODAY 保持直接进入 `=NAME()` 公式编辑的行为。
- `SheetRibbonState.formulaEditable` 接收工作台 `canInsertFormula`：工作台活动、当前整选区的工作表/范围编辑权限、非 IME 组合输入且无绘图选区。分类根/函数子项、打开函数面板、编辑公式共用禁用状态，并保留无工作表及绘图焦点判断；计算模式和重算不使用这一值。工作台三种 operation 派发及现有表单/确认按钮/`insertFunction` 也复用相同判断。
- NA 批仅新增 `{ name: 'NA', category: '信息', example: '' }` 目录项并加入既有零参数判定。现有分类排序、Info 图标、禁用状态和 `sheet.operation.insert-function` 自动复用，不修改 Ribbon schema，也不创建新表单或命令。选择 NA 直接进入 `=NA()` 编辑，正常求值结果就是 `#N/A`，不是未知函数错误。

本地入口按**整个当前选区**判断权限，接近上游 `getCurrentRangeDisable$` 的范围语义；函数确认后的现有单格命令只写当前 anchor，并按其实际合并范围检查权限。二者承担不同阶段的职责：若 C1 可写但选区 C1:D1 的 D1 受保护，顶部入口仍禁用；缩到 C1 后可恢复。`formulaEditable` 复用的是现有保护判断，不取 `fontEditable` 的样式状态。上游独立 WorkbookEditable/WorksheetSetCellValue 权限点、全部图表工作表与上下文状态尚未逐项验证，不能因此标为全部权限对齐。

## 函数目录与求值续核

固定 [FUNCTION_LIST 汇总](../../../univer/packages/engine-formula/src/services/formula/function-list/function-list.ts) 提供描述目录；[DescriptionService](../../../univer/packages/engine-formula/src/services/formula/description.service.ts) 先与 [ALL_IMPLEMENTED_FUNCTIONS_SET](../../../univer/packages/engine-formula/src/functions/index.ts) 取交集，再按类型生成分类菜单。以下为静态分类计数，不计动态注册及其他插件扩展；common 为重复引用函数的静态快捷列表，不能把各行相加当成唯一函数总数。

| 左到右分类 | 描述项 | 描述与全局实现集交集 | 本地目录项 |
| --- | ---: | ---: | ---: |
| 常用与统计 | 7 个快捷项 | 7 | 6 |
| [财务](../../../univer/packages/engine-formula/src/services/formula/function-list/financial.ts) | 55 | 55 | 17 |
| [逻辑](../../../univer/packages/engine-formula/src/services/formula/function-list/logical.ts) | 19 | 19 | 11 |
| [文本](../../../univer/packages/engine-formula/src/services/formula/function-list/text.ts) | 51 | 48 | 21 |
| [日期与时间](../../../univer/packages/engine-formula/src/services/formula/function-list/date.ts) | 27 | 27 | 26 |
| [查找与引用](../../../univer/packages/engine-formula/src/services/formula/function-list/lookup.ts) | 37 | 36 | 11 |
| [数学](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts) | 82 | 82 | 9 |
| [统计](../../../univer/packages/engine-formula/src/services/formula/function-list/statistical.ts) | 113 | 109 | 16 |
| [工程](../../../univer/packages/engine-formula/src/services/formula/function-list/engineering.ts) | 57 | 57 | 0 |
| [信息](../../../univer/packages/engine-formula/src/services/formula/function-list/information.ts) | 25 | 23 | 1（NA） |
| [数据库](../../../univer/packages/engine-formula/src/services/formula/function-list/database.ts) | 12 | 12 | 0 |

固定描述中未实际注册的 `PHONETIC / CALL / REGISTER.ID / RTD / FORECAST.ETS / FORECAST.ETS.CONFINT / FORECAST.ETS.SEASONALITY / FORECAST.ETS.STAT / INFO / ISOMITTED` 不计入交集。该计数来自固定文件中的真实注册项，排除了注释掉的 map 项，不以枚举名称数推断可用函数。

本地实际求值入口是 [formula-engine.ts](../../packages/office-studio/src/features/formula/formula-engine.ts)，由 [OfficeStudio](../../packages/office-studio/src/api/office-studio.ts) 的工作簿事务准备阶段调用。现有 [ComputeTask](../../packages/office-studio/src/compute/protocol.ts) 及 [office-compute](../../crates/office-compute/src/lib.rs) 只有 aggregate、sort、filter、chartStatistics、slicerMembers 等任务，没有另一套工程或数据库公式求值。未知函数的旧缓存保留行为不能当成计算实现。NA 沿用现有 TS 权威求值，本批没有声称迁移到 Worker/WASM。

候选遗漏按左到右核实结果如下，除 NA 外均未接入本批目录：

- 文本：`LEFT / MID` 对长度及位置要求整数，固定实现会取整；`RIGHT("abc",0)` 的本地旧分支会返回整串，固定 [RIGHT](../../../univer/packages/engine-formula/src/functions/text/right/index.ts) 返回空串。随后的小修仅让零长度返回空串，保留默认长度、正长度及负数拒绝行为；RIGHT 的小数长度、字符串参数转换、数组和 Unicode 语义仍未完整对齐，未加入目录。`VALUE` 当前只有 `Number(text)`，没有完整 Office 数字、日期和时间文本转换。`CONCATENATE` 本地直接复用 CONCAT 的范围扁平拼接，固定 [CONCATENATE](../../../univer/packages/engine-formula/src/functions/text/concatenate/index.ts) 按数组位置广播。这些有求值分支，但仍有真实语义缺口。
- 日期：`WEEKDAY` 仅支持 return_type 1/2/3，固定 [WEEKDAY](../../../univer/packages/engine-formula/src/functions/date/weekday/index.ts) 还支持 11–17，并取整类型值。其完整语义未完成。
- 工程：固定 57 个已注册描述函数，本地没有对应求值分支或目录，仍未完成。
- 信息：`ISNUMBER / ISTEXT / ISBLANK / ISERROR` 虽有分支，但统一错误传播先于分支执行，`ISERROR(1/0)` 会直接返回错误；ISBLANK 还把空文本视为真正空白，与固定 [ISBLANK](../../../univer/packages/engine-formula/src/functions/information/isblank/index.ts) 的引用/公式空白判断不同。其余信息能力仍未完成。仅 [NA](../../../univer/packages/engine-formula/src/functions/information/na/index.ts) 的零参数返回 #N/A 已有真实实现并接入；额外参数的错误类型及信息函数数组行为仍未对齐。
- 数据库：固定 12 个已注册描述函数，本地没有对应求值分支或目录，仍未完成。工程、信息、数据库完整能力仍在交付范围内，以上是当前缺口记录。

## 命令、输入、历史与保存

[固定插入 operation](../../../univer/packages/sheets-formula-ui/src/commands/operations/insert-function.operation.ts) 能推断上方/左方连续数值区域；多选时还可寻找结果空格并批量落公式，无引用时也能进入编辑。参数入口批只恢复函数名称与参数表单接线，NA 批复用同一编辑链路；自动范围与批量插入交互另行核对。

[固定“全部函数” operation](../../../univer/packages/sheets-formula-ui/src/commands/operations/more-functions.operation.ts) 打开 sidebar。本地复用已有 Ant Modal/Form/Select/Input；菜单、表单和公式输入保持既有 `antdv-next` 控件。表单校验仍由 `buildSheetFunctionFormula` 处理函数、参数与解析语法，成功后调用 `beginEdit(..., 'formula')`，此时尚未提交内容。取消表单不会创建编辑会话；公式栏取消丢弃会话草稿。

公式栏确认经 `finishEdit` → `setCellValue` → `sheet.command.commit-cell-input`。现有 [单格命令注册](../../packages/office-studio/src/api/sheet-cell-input-commands.ts) 调 [planSheetCellInput](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-input-command.ts)，按合并区 anchor 解析目标，检查工作表查看/编辑与目标范围编辑权限及数据验证，最后经 `commitWorkbook` 发布内容历史。入口及打开中的表单使用整选区能力；最终命令继续以真实写入范围处理提交时的权限变化，不新增第二套写入逻辑。

撤销/重做使用当前工作簿内容历史。`FWorkbook.save()` 克隆领域模型并标记已保存，重新 `createWorkbook(saved)` 保留公式与缓存值；本批浏览器草稿覆盖此 native 模型保存重开，不代表独立 XLSX 文件往返已验收。

计算模式通过 `FWorkbook.setCalculationMode` → `sheet.command.set-calculation-mode` → `commitWorkbook` 保存。重算通过 `FWorkbook.recalculate` → `sheet.command.recalculate-workbook` → 强制公式求值提交。既有 [计算模式契约](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs) 已包含 manual 缓存、显式重算、切回 automatic 与保存重开场景，本批未运行。

固定 [CalculationOptions operation](../../../univer/packages/sheets-formula-ui/src/commands/operations/calculation-options.operation.ts) 打开 420px dialog；[表单实现](../../../univer/packages/sheets-formula-ui/src/views/calculation-options/CalculationOptions.tsx) 配置 `iterativeCalculation / maxIterations / maxChange` 并派发 `SetWorkbookCalculationSettingsCommand`。本地顶部 automatic/manual 与重新计算不是这三个设置的替代实现；本批未接入对应迭代配置表单，底层迭代求值、资源与 Office 映射仍未核实。

## 草稿与未验收项

[既有 Ribbon 契约](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 补文本/日期图标、确切函数名称派发、所有公式根/子项的无工作表禁用、公式能力/绘图焦点禁用但计算模式/重算仍可用，以及 calculation mode 默认/手动/禁用状态。

[既有浏览器用例](../../tests/browser/sheet-toolbar.spec.ts) 改用当前分类菜单和 Ant Menu 的 menuitem：SUM 参数错误/填写/编辑确认/撤销重做/native 保存重开；ABS 的数学名称预选与取消无编辑；跨到逻辑 IF 的名称预选、参数填写、确认和撤销；TRUE、FALSE、NOW、TODAY 直接编辑并取消。另补一条真实 Facade 选区变化：打开参数表单后从可写 C1 扩到含受保护 D1 的 C1:D1，表单输入/确认与顶部入口同步禁用，缩回 C1 后恢复。三条用例均未运行。

参数入口批仅执行限定静态检查：Ribbon 契约 `node --check` 无语法错误；schema 与浏览器草稿的 TypeScript AST 解析共 2 个文件、0 条语法诊断；当时本证据 15 个链接均存在，限定变更及新证据无空白错误。

NA 批补 [既有公式契约](../../packages/office-studio/tests/formula-engine.contract.test.mjs) 中的构造、#N/A、IFNA、IFERROR 和错误传播草稿；全目录契约明确 NA 应返回 #N/A，其他示例仍要求正常求值。Ribbon 契约明确信息根只有 NA、沿用 Info 图标、分类顺序和函数派发，并继续覆盖权限禁用；浏览器草稿将 NA 加入既有零参数编辑/取消循环。均未运行。

NA 批静态检查通过：3 个既有契约 `node --check`；catalog/schema 两个 TypeScript AST 解析 0 条语法诊断；本证据 38 个链接均存在；26 个固定 map/description 文件中的枚举引用无缺失；按已有 CRLF 规则执行限定空白检查通过。RIGHT 小修在同一公式契约追加零长度、默认长度、正长度、超出文本长度及负数参数的真实实现回归草稿，未运行。追加后该契约 `node --check`、引擎 TypeScript AST（0 条语法诊断）、38 个证据链接和限定空白检查均通过。

入口两批及 RIGHT 小修均未运行测试、类型检查、构建、Rust/WASM、dev、浏览器或性能基准。实际菜单 hover/选中/禁用、明暗主题、焦点恢复、系统 IME、宽窄屏和 Canvas 两后端仍未验收。完整目录、全部函数/定义名称交互、自动范围与批量插入、迭代计算、SHEET-10 Pro、真实 Office 文件与大样本性能继续保持未完成或未验收，不能由这些小修推断能力闭环。
