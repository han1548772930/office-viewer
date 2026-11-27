# SHEET-09 FORMULAS 顶部入口续核

状态：2026-11-27 继续按固定来源核对现有公式入口和求值链路。本批续接 PERCENTOF、ISDATE、AVEDEV / DEVSQ / GEOMEAN / HARMEAN / MODE.SNGL，并修正既有逻辑函数的参数数目错误；数学名称累计新增48个。当前唯一目录215项、8个真实类别，Ribbon 为8类加常用投影共9个函数根；函数菜单出现项222、另有插入函数子项1，根入口不变。固定来源为 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；公式批未变更模型、命令或 Worker/WASM 协议。全部运行验证后置，完整函数功能未验收。

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

[sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的 `formula-edit` 组在 NA 批之前提供 **8** 根；接入 NA 后依序提供常用与统计、财务、逻辑、文本、日期与时间、查找与引用、数学、统计、信息共 **9** 根。common 仍为 r1c1 跨两行。本批 common 读取独立投影，其余按真实类别过滤；信息续接四项后共 13 项，工程/数据库仍没有目录项。

[现有函数目录](../../packages/office-studio/src/features/sheet-workbench/sheet-function-catalog.ts) 当前为 **215** 项、**8** 类：财务17、逻辑12、文本25、日期与时间27、查找与引用11、数学75、统计26、信息22，无重复名称。common 投影固定七项 SUMIF / SUM / AVERAGE / IF / COUNT / SIN / MAX，复用同一目录对象；菜单函数出现项 **222**，common 底部“插入函数”另计 **1** 子项，全部子项223，唯一函数仍为 **215**。NA 批117→118、126/131、三角批137/9类、常用投影批137/8类、角度转换140、反三角148、双曲/对数160、奇偶/整商164、阶乘/组合168、ROUND系列172、CEILING/FLOOR179、整数聚合183、SUMSQ184、数组求和189、罗马/进制转换193、随机函数195、MDETERM 196、信息类型200、七个信息函数207及 ISBETWEEN 批208均为历史快照。MDETERM 后累计新增48个数学名称；原函数修正不增加唯一名称，数量不等于逐函数验收。

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
| 常用函数（本地“常用与统计”） | 7 个快捷项 | 7 | 7，重复引用真实类别 |
| [财务](../../../univer/packages/engine-formula/src/services/formula/function-list/financial.ts) | 55 | 55 | 17 |
| [逻辑](../../../univer/packages/engine-formula/src/services/formula/function-list/logical.ts) | 19 | 19 | 12 |
| [文本](../../../univer/packages/engine-formula/src/services/formula/function-list/text.ts) | 51 | 48 | 25 |
| [日期与时间](../../../univer/packages/engine-formula/src/services/formula/function-list/date.ts) | 27 | 27 | 27 |
| [查找与引用](../../../univer/packages/engine-formula/src/services/formula/function-list/lookup.ts) | 37 | 36 | 11 |
| [数学](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts) | 82 | 82 | 75 |
| [统计](../../../univer/packages/engine-formula/src/services/formula/function-list/statistical.ts) | 113 | 109 | 26 |
| [工程](../../../univer/packages/engine-formula/src/services/formula/function-list/engineering.ts) | 57 | 57 | 0 |
| [信息](../../../univer/packages/engine-formula/src/services/formula/function-list/information.ts) | 25 | 23 | 22 |
| [数据库](../../../univer/packages/engine-formula/src/services/formula/function-list/database.ts) | 12 | 12 | 0 |

固定描述中未实际注册的 `PHONETIC / CALL / REGISTER.ID / RTD / FORECAST.ETS / FORECAST.ETS.CONFINT / FORECAST.ETS.SEASONALITY / FORECAST.ETS.STAT / INFO / ISOMITTED` 不计入交集。该计数来自固定文件中的真实注册项，排除了注释掉的 map 项，不以枚举名称数推断可用函数。

[逻辑 map](../../../univer/packages/engine-formula/src/functions/logical/function-map.ts) 实际注册21项；GROUPBY / PERCENTOF 未列入19项描述目录，仍属完整目标。LET 已在该逻辑 map 注册，不能因数学 map 未注册就记为全局未注册。本地逻辑12项仍缺 BYCOL / BYROW / GROUPBY / LAMBDA / LET / MAKEARRAY / MAP / REDUCE / SCAN；信息23注册项中尚缺 CELL，统计109注册项仅26项进入本地目录。名称进入目录也不表示数组、引用和文件行为已验收。

本地实际求值入口是 [formula-engine.ts](../../packages/office-studio/src/features/formula/formula-engine.ts)，由 [OfficeStudio](../../packages/office-studio/src/api/office-studio.ts) 的工作簿事务准备阶段调用。现有 [ComputeTask](../../packages/office-studio/src/compute/protocol.ts) 及 [office-compute](../../crates/office-compute/src/lib.rs) 只有 aggregate、sort、filter、chartStatistics、slicerMembers 等任务，没有另一套工程或数据库公式求值。未知函数的旧缓存保留行为不能当成计算实现。NA 沿用现有 TS 权威求值，本批没有声称迁移到 Worker/WASM。

候选遗漏及现有分支按左到右续核如下；本次五个目录缺项已接回既有入口：

- 文本：此前已修复 RIGHT 零长度与三个函数的数值取整；本次按固定 [LEFT](../../../univer/packages/engine-formula/src/functions/text/left/index.ts)、[RIGHT](../../../univer/packages/engine-formula/src/functions/text/right/index.ts) 和 [MID](../../../univer/packages/engine-formula/src/functions/text/mid/index.ts) 补数值文本、单格范围和 UTF-16 截取。位置/长度复用现有数字、日期与时间解析器，空文本参数返回 #VALUE!，空单格仍转为 0；文本下标改用字符串 slice，保留代理对的代码单元边界。数组广播、带格式数值转文字、完整转换样本与错误优先级仍需继续核对/验收；三个函数已加入目录。VALUE 已改为复用格式化数值/日期/时间转换并加入目录，空白文本、逻辑值拒绝及单格引用详见[VALUE 续核](#value-input-continuation)。`CONCATENATE` 本地直接复用 CONCAT 的范围扁平拼接，固定 [CONCATENATE](../../../univer/packages/engine-formula/src/functions/text/concatenate/index.ts) 按数组位置广播，仍有真实语义缺口。
- 日期：WEEKDAY 已按固定 [实现](../../../univer/packages/engine-formula/src/functions/date/weekday/index.ts) 支持 return_type 1/2/3/11–17、类型取整及双日期系统偏移；本次补数值/日期/时间文本、单格范围和 serial 上限。日期参数只按上游接受日期、时间或纯数值文本，类型参数沿用格式化数值转换；WEEKDAY 已加入目录，日期分类名称数达到固定 27 项，数组广播、全部日期文本规格与运行行为仍未验收。
- 工程：固定 57 个已注册描述函数，本地没有对应求值分支或目录，仍未完成。
- 信息：续作前已有 `ISBLANK / ISERR / ISERROR / ISLOGICAL / ISNA / ISNONTEXT / ISNUMBER / ISTEXT` 八个谓词及目录项。谓词分支先于统一错误传播，`ISERROR(1/0)` 返回 true，`ISERR(NA())` 返回 false；不把数值/逻辑文本强制转换成对应类型。ISBLANK 区分空文本常量与单格引用，并保留公式、富文本、图片及 INDEX 的引用身份，依据固定 [ISBLANK](../../../univer/packages/engine-formula/src/functions/information/isblank/index.ts) 和 [引用空白判断](../../../univer/packages/engine-formula/src/engine/reference-object/base-reference-object.ts)。多格范围目前返回 #VALUE!，固定数组广播及其余信息函数仍未齐；NA 的额外参数错误已按[本批续修](#na-argument-count)接回 #N/A。原有八个谓词没有重写或重复计数。
- 数据库：固定 12 个已注册描述函数，本地没有对应求值分支或目录，仍未完成。工程、信息、数据库完整能力仍在交付范围内，以上是当前缺口记录。

<a id="scalar-function-continuation"></a>

## 2026-11-27 文本取整与 WEEKDAY 续核

本批只改现有 TypeScript 权威求值器。LEFT/RIGHT 的默认长度仍为 1，MID 保留三个参数；数字位置及长度先 `Math.floor`，零长度/超出文本范围返回空串，负小数取整后仍按固定版返回 #VALUE!。参数已有错误继续传播。`LEFT("abcdef",2.9)` 为 `"ab"`，`RIGHT("abcdef",2.9)` 为 `"ef"`，`MID("abcdef",2.9,3.9)` 为 `"bcd"`。

WEEKDAY 使用固定 [日期基础函数](../../../univer/packages/engine-formula/src/basics/date.ts) 的 1900 offset 6、1904 offset 5；复用本地 `serialWeekday`，删除不再调用的旧 `excelWeekday`。类型 2/11 以周一为首日，12–17 分别以周二至周日为首日，类型 3 返回周一 0；非法类型返回 #NUM!，类型转换产生的 #VALUE! 或已有 #N/A 保持原错误。原日期有效范围判断继续复用。示例 `WEEKDAY(DATE(2024,1,7),12)` 为 6，`WEEKDAY(DATE(2024,1,9),12)` 为 1。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增两条声明，覆盖小数、零长度、负小数、越界、参数错误、全部返回类型及 serial 0/59/60/61 的双日期系统。八个信息谓词的类型/错误和 ISBLANK 引用草稿此前已在同一文件中存在。[计算模式契约草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增一条真实 Facade 集成声明，覆盖源格编辑后四个函数重算、一次范围编辑的撤销重做、公式文本与缓存保存重开及重开后再次重算。全部未运行；内部快照重开不替代 XLSX 独立重开。

本批常驻、插件扩展、上下文、溢出根与函数子项增量均为 0。控制 DTO、模型、命令 ID 及 Worker/WASM 协议未变；现有 TS 计算路径的修正不表示 Rust 公式引擎迁移或性能目标已完成。

<a id="scalar-input-continuation"></a>

## 2026-11-27 单格范围、文本转换与 UTF-16 续核

固定文本函数先用 [checkVariantsErrorIsStringToNumber](../../../univer/packages/engine-formula/src/engine/utils/check-variant-error.ts) 转换位置/长度，再取整；[StringValueObject](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts) 拒绝空白文本，使用 [getNumfmtParseValueFilter](../../../univer/packages/core/src/shared/numfmt/utils.ts) 解析日期、时间、百分比、货币和数值，其后才尝试纯数字。原本本地 `numericCoercion` 拒绝所有非空文本。现在四个目标函数先复用 `singleRangeValue` 读取单格引用或 INDEX，再由私有 `numericTextCoercion` 复用已有 [parse-value.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/parse-value.ts)、[number-input.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/number-input.ts) 与 [time-input.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/time-input.ts)。其他函数的通用数值转换保持原语义。

转换使用固定默认英文解析规则和工作簿日期系统；缺年份日期使用本次求值时钟的年份。数值文本、科学记数法、常规分组、货币、百分比及会计负数复用同一解析器，未解析文本再按固定版纯数值规则处理；空/空白文本拒绝，引用空格仍按 null→0。固定 StringValueObject 的纯数值分支可产生无限大的长度，文本函数继续按字符串下标规则截取；WEEKDAY 对同样的无界序列返回 #NUM!。所有转换结果与特殊格式仍需样本验收，不以新增草稿推断全部等价。

WEEKDAY 的日期文本转换只使用日期/时间解析和纯数值，保留上游不接受货币或百分比日期参数的差异；return_type 才使用完整格式化转换。其上限复用本地 [serial-date.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/serial-date.ts)，1900 为 2958465，1904 为 2957003；超过上限或负值返回 #NUM!。两种系统的 `"2024-01-07"` 都按同一周日求值。

LEFT/RIGHT/MID 从 `Array.from(...).slice(...).join('')` 改为字符串 slice，与固定版 substring 的正下标语义一致。`LEFT("A😀B",2)` 返回 A 和高代理代码单元，`MID("A😀B",2,2)` 保留整个 Emoji；组合标记按 UTF-16 下标截取，不按字素簇扩展。多格范围仍返回 #VALUE!，未把标量修正标记为数组/spill 完成。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)追加三条声明，直接覆盖格式化/空白/无界数值文本、真实单格及错误引用、代理对/组合标记、日期时间文本和双系统序列上限；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)把四个公式扩展为数值文本及单格范围，继续验证源编辑、重算、一次撤销重做、公式/缓存保存与重开后重算。均未运行；该标量小修阶段没有新增 UI、目录、命令、模型字段或计算引擎，随后目录补项见下文。Worker/Rust/WASM 迁移与 Office/性能仍未完成验收。

<a id="scalar-catalog-continuation"></a>

## 2026-11-27 四个函数目录缺项续接

固定 [文本描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/text.ts) 注册 LEFT/MID/RIGHT，LEFT/RIGHT 为必填文本和可选长度，MID 为必填文本、起始位置及长度；[日期描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/date.ts) 注册 WEEKDAY 的日期与可选返回类型。四个名称均在固定实现集内，本地求值分支此前已存在。本次只在既有 catalog 增加四条目录，示例沿用固定 `"Univer"` 与 `"2008-2-14"` 参数，不另建命令或表单。

现有 Ribbon 分类按 catalog 生成子项，全部函数表单也读取同一目录。四项都派发 `sheet.operation.insert-function` 并预选真实名称，通过原 Ant Modal/Form/Input 进入公式栏，确认后走现有单格输入命令、权限、内容历史和重算。新增目录子项 4，文本 21→24、日期 26→27，唯一名称 126→130；9 个分类、12 个公式常驻根及其他入口口径不变。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增一条从真实目录构造公式再求值的声明；[Ribbon 契约](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)在现有分类用例中明确四项的名称派发；[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)在现有分类用例中追加菜单选择、参数填写、确认结果和一次撤销。随后三处草稿再加入 VALUE，见下文。全部未运行。函数数量记录目录补项，不计作四个新算法或行为验收；完整目录、数组、定义名称、Office 重开和性能仍有缺口。

<a id="value-input-continuation"></a>

## 2026-11-27 VALUE 文本转换续核

固定 [VALUE](../../../univer/packages/engine-formula/src/functions/text/value/index.ts) 保留已有错误，拒绝逻辑值；字符串先经 StringValueObject 转数值，其余标量直接转数值，单元素数组可解包。本地此前把参数转成文本后只调用 Number：货币、百分比、日期时间无法求值，空文本反而成为 0，单格范围也没有解包。本次只将 VALUE 接入已有单格范围解包和 `numericTextCoercion`，在转文本之前处理，空单格仍为 0、空/空白文本与逻辑值返回 #VALUE!，已有 #N/A 保留。格式化转换、日期系统、求值年份与上述四函数复用同一解析器；没有添加第二份引擎或输入校验层。

VALUE 的固定字符串纯数值路径可返回 Infinity；本地保留该结果，未把它作为有限数值的 Office 保存证明。完整格式化转换、紧凑英文月年及数组/spill 仍需核对。catalog 追加固定文本分类 VALUE（示例 `"123"`），文本 24→25、唯一目录 130→131；当前相对续作前共新增 **5** 个函数子项，9 分类和所有常驻/扩展/上下文/溢出根不变。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)追加真实转换声明，覆盖分组/货币/百分比/会计负数、双系统日期、时间、空白、逻辑拒绝、单格/INDEX/空引用、多格范围与已有错误；目录构造、Ribbon 派发和既有浏览器分类用例追加 VALUE。浏览器草稿覆盖 `VALUE("25%")` 的确认结果与撤销，公式源/缓存保存重开继续由已有 Facade 草稿和统一 Office 样本验证。均未运行，不能由目录 131 项推断全部函数或跨模块行为验收。

<a id="scalar-text-reference-continuation"></a>

## 2026-11-27 LEN/UPPER/LOWER/TRIM 单格引用与空白续核

固定 [LEN](../../../univer/packages/engine-formula/src/functions/text/len/index.ts)、[UPPER](../../../univer/packages/engine-formula/src/functions/text/upper/index.ts)、[LOWER](../../../univer/packages/engine-formula/src/functions/text/lower/index.ts) 和 [TRIM](../../../univer/packages/engine-formula/src/functions/text/trim/index.ts) 都解包单元素数组并保留错误。本地原先把单格 RangeValue 经 scalarValue 得到的错误再转文本，LEN 可能返回错误标签长度，其余函数可能返回错误标签文本。本次将四个既有函数接入同一 `singleRangeValue`，原有统一错误传播先执行；多格范围仍为 #VALUE!，不把它计为上游数组广播。

TRIM 固定版只处理 U+0020 与 U+3000：删除首尾连续空格，内部连续空格保留第一个代码单元，Tab、换行和 NBSP 保留。本地原先 trim 与 `\s+` 同时清掉这些字符；现在按相同两步规则处理。UPPER/LOWER 的 toLocaleUpperCase/toLocaleLowerCase 已与固定版相同，本次只接单格/INDEX；LEN 仍按 UTF-16 长度，带格式数值转文字与全数组/spill 未验收。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)追加两条真实实现声明：普通/全角/混合空格和 Tab/换行/NBSP，以及四函数的单格/INDEX、已有错误、空引用与多格范围。[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)追加一条源格编辑、四个缓存重算、撤销重做、公式/缓存快照重开及重开后再次输入 Tab/换行的集成声明。全部未运行；该文本批无目录或入口增量，当时为131项/9类，真实 Office、主题/IME 和性能继续未验收。

<a id="trigonometric-scalar-continuation"></a>

## 2026-11-27 六个三角函数标量续接

固定 [数学 function map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts) 和 [描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts) 均注册 [SIN](../../../univer/packages/engine-formula/src/functions/math/sin/index.ts)、[COS](../../../univer/packages/engine-formula/src/functions/math/cos/index.ts)、[TAN](../../../univer/packages/engine-formula/src/functions/math/tan/index.ts)、[COT](../../../univer/packages/engine-formula/src/functions/math/cot/index.ts)、[SEC](../../../univer/packages/engine-formula/src/functions/math/sec/index.ts)、[CSC](../../../univer/packages/engine-formula/src/functions/math/csc/index.ts)，六者均只接受一个弧度参数。字符串经 [StringValueObject](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts) 转数值，已有错误原样保留；空引用为 0，逻辑值为 0/1。

本地在现有数学分支旁新增六者的标量求值，复用 singleRangeValue、numericTextCoercion 和 finiteResult，不另建解析器或引擎。单格/INDEX、数字/百分比/货币文本使用既有转换；空白文本拒绝。SIN/COS/TAN 使用 native Math，与固定 NumberValueObject 相同，不附加 2^27 上限；COT/SEC/CSC 在绝对值达到 134217728 时返回 #NUM!，余弦/正弦/正切倒数复用有限结果判断，分母为零返回 #DIV/0!。SIN(0)=0、COS(0)=1、TAN(0)=0、SEC(0)=1；COT/CSC 的 0、false 或空引用为 #DIV/0!。

三角批 catalog 增加六个数学名称，当时数学 9→15、唯一名称 131→137、9 类；所有常驻/扩展/上下文/溢出根不变。六项沿现有分类菜单派发 sheet.operation.insert-function，通过已有 Ant Modal/Form/Input、公式输入确认、权限与 commitWorkbook。common 当时仍为独占六项分类，随后按[常用函数投影](#common-function-projection)修正；MIN/COUNTA 仍保留在真实统计分类和完整目录中。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)补真实标量转换、单格/INDEX/错误引用、空值、参数数目、上下限及目录构造求值；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)补源格编辑、六个缓存重算、一次撤销重做、公式/缓存快照重开及重开后零值错误缓存。[Ribbon 草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)核对六项的实际名称派发，[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)扩展既有分类循环，覆盖名称预选、填写、确认结果与撤销，浮点结果按容差判断。全部未运行。

数组/spill、紧凑英文月年、全部格式化转换、Office 文件往返及 Worker/Rust/WASM 路径仍未完成或未验收。common 投影随后已接，但实际交互与完整分类目录仍未验收。六个新标量算法、六个目录子项与完整功能验收分别记录；该批没有新增 UI 控件、命令 ID、模型字段或跨计算域协议，也没有性能优化结论。

<a id="common-function-projection"></a>

## 2026-11-27 常用七项与真实分类续接

固定 [InsertCommonFunctionMenuItemFactory](../../../univer/packages/sheets-formula-ui/src/menu/menu.ts) 的顺序为 SUMIF / SUM / AVERAGE / IF / COUNT / SIN / MAX。common 是现有描述目录的快捷投影；[数学目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts) 的 SUM（1396 行）属于 Math，[统计目录](../../../univer/packages/engine-formula/src/services/formula/function-list/statistical.ts) 的 AVERAGE / COUNT / COUNTA / MAX / MIN 属于 Statistical。本地此前把 SUM 及五个统计函数放入独占“常用与统计”，导致它们在真实分类中缺失，common 还缺 SUMIF / IF / SIN。

本批在既有 catalog 导出 SHEET_COMMON_FUNCTION_CATALOG，以固定七名称查找同一 canonical 对象；没有复制函数实现或创建第二份权威目录。SUM 归数学，五个统计函数归统计；Ribbon 的 common 用投影，其他根继续按真实类别过滤。ExcelWorkbench 既有函数表单使用同一投影/分类 computed 列表，保留 open-function 的常用上下文、名称预选与首项选择；从 common 选择 SUM 后，参数表单显示真实“数学”类别。

常用投影批静态读取确认唯一函数 **137**、真实类别 **8**，数学 **16**、统计 **21**；加 common 后函数根 **9**。当时菜单函数出现项 **144 = 137 + 7**，另外“插入函数”子项 **1**，总子项 **145**；相对三角批函数菜单出现项 +7，唯一函数、算法、常驻/插件扩展/上下文/溢出根增量均为 0，Sheets 71 / Docs 54 不变。common 的重复出现不能计作新增功能。

三份生产文件仅修改目录和既有 UI 接线，Form / Select / Input / Modal 继续使用 antdv-next；公式确认、权限、数据验证、内容历史及 native 保存链路沿用上文记录。模型、公共命令/Facade 和 Worker/WASM 协议未变，本批没有迁移计算域或性能结论。

[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增一条七项顺序、同一目录对象、真实类别与构造后实际求值的声明；[Ribbon 草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)扩展既有分类用例，检查精确顺序、真实分类入口、重复出现计数和原名称派发；[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)在既有 SUM 编辑/历史/native 重开闭环检查七项菜单和参数表单的数学分类。全部未运行。常用根仍用本地“常用与统计”标签，与固定“常用函数”文案有差异；全目录/数组、自动范围/批量插入、迭代计算、主题/焦点/IME/Office 与性能继续未完成或未验收。

<a id="na-argument-count"></a>

## 2026-11-27 NA 参数数目错误续修

固定 [Na](../../../univer/packages/engine-formula/src/functions/information/na/index.ts) 的 minParams/maxParams 均为0，正常结果为 #N/A；[FunctionNode._calculate / _calculateAsync](../../../univer/packages/engine-formula/src/engine/ast-node/function-node.ts) 在参数数目不匹配时同样返回 #N/A，先于函数计算。因此 NA(1)、NA(1/0)、NA(#REF!)、NA(A1:B1) 都应保持 #N/A 身份。本地原先只在末尾识别 NA()，多余参数可能成为 #NAME? 或被统一错误传播改成其他错误。

本批把既有 NA 分支移到参数求值前，统一返回 #N/A；目录、零参数菜单、命令及持久模型不变。[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)扩展既有用例，覆盖六种多余参数及 IFNA/IFERROR；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增一次三格公式编辑、错误处理、撤销重做和 native 保存重开。均未运行，没有新增算法名称、UI 入口或 Worker/WASM 协议，独立 Office 与完整公式功能仍未验收。

<a id="pi-angle-conversion-continuation"></a>

## PI 与角度转换标量续接

本批按固定实现、[数学描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)和对应上游测试新增三个分支。参数数目不匹配依照 [FunctionNode](../../../univer/packages/engine-formula/src/engine/ast-node/function-node.ts) 返回 #N/A，先于参数错误传播。

| 函数与固定实现 | 参数数目 | 标量行为 |
| --- | ---: | --- |
| [Pi](../../../univer/packages/engine-formula/src/functions/math/pi/index.ts) | 0 | 返回 Math.PI；多余参数，包括 1/0 或 #REF!，返回 #N/A |
| [Degrees](../../../univer/packages/engine-formula/src/functions/math/degrees/index.ts) | 1 | 字符串转数值，保留输入错误；有限输入乘以 (180 / Math.PI) |
| [Radians](../../../univer/packages/engine-formula/src/functions/math/radians/index.ts) | 1 | 同一标量转换，有限输入乘以 (Math.PI / 180) |

角度转换的非有限输入返回 #VALUE!。固定 [NumberValueObject.create](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts) 直接持有数值，Degrees 对有限输入的溢出结果保留 Infinity；本地保持该结果及乘法次序。原输入错误、布尔、空单元格、数值/百分比/货币字符串和单格 range 复用既有 singleRangeValue / numericTextCoercion；没有新增转换 helper 或第二套权威求值器。固定 Degrees / Radians 还支持数组 map，本地统一数组/spill 引擎仍未实现，不能由标量分支计作完整函数验收。

计算归属为 SHEET-09 既有纯同步 TS 求值器；UI 目录、命令编排与模型历史沿用原接口，Worker/Rust/WASM 批量公式路径仍待实现和基准。目录数学由16增至19，唯一函数137→140；函数菜单出现项144→147，加插入函数子项共148。函数根9、Sheets71 / Docs54以及常驻/扩展/上下文/溢出根均不变。新增 DEGREES / RADIANS 的示例分别为 PI() / 270，使用原 antdv-next 参数表单；PI 的空示例和零参数构造进入既有公式输入与确认链路。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增两个用例，覆盖零参数、严格参数数目、转换、原错误、非有限输入及溢出；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增真实公式编辑、依赖重算、撤销重做、native 保存重开及错误处理。既有 [Ribbon 草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)补三个名称派发，[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)补两个参数表单提交及 PI 直接编辑/提交/撤销。全部未运行；主题、焦点/IME、两种 Canvas 后端、真实 Office 文件、数组与大样本性能继续未验收。

<a id="inverse-trigonometric-scalar-continuation"></a>

## 反三角函数标量续接

紧接角度转换批，按固定实现增加八个单参数数学分支；全部使用同一 singleRangeValue / numericTextCoercion、严格单参数 #N/A 判断和既有编辑事务。

| 函数与固定入口 | 定义域与可观察结果 |
| --- | --- |
| [Acos](../../../univer/packages/engine-formula/src/functions/math/acos/index.ts) | 有限 x 且 -1≤x≤1，Math.acos；越界或非有限输入为 #NUM! |
| [Acosh](../../../univer/packages/engine-formula/src/functions/math/acosh/index.ts) | 有限 x≥1，Math.acosh；越界或非有限输入为 #NUM! |
| [Acot](../../../univer/packages/engine-formula/src/functions/math/acot/index.ts) | atan(1/x)，x<0 时加 π；非有限输入为 #VALUE!，保留固定 -0 结果 |
| [Acoth](../../../univer/packages/engine-formula/src/functions/math/acoth/index.ts) | 0.5×log((x+1)/(x-1))；abs(x)≤1 为 #NUM!，非有限输入为 #VALUE! |
| [Asin](../../../univer/packages/engine-formula/src/functions/math/asin/index.ts) | 有限 x 且 -1≤x≤1，Math.asin；越界或非有限输入为 #NUM! |
| [Asinh](../../../univer/packages/engine-formula/src/functions/math/asinh/index.ts) | 有限输入，Math.asinh；非有限输入为 #NUM! |
| [Atan](../../../univer/packages/engine-formula/src/functions/math/atan/index.ts) | 有限输入，Math.atan；非有限输入为 #NUM! |
| [Atanh](../../../univer/packages/engine-formula/src/functions/math/atanh/index.ts) | 有限 x 且 -1<x<1，Math.atanh；端点、越界或非有限输入为 #NUM! |

六个委托函数的具体数值规则来自 [NumberValueObject、NullValueObject、BooleanValueObject](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts) 的对应方法；ACOT/ACOTH 使用固定独立实现。原输入错误保持身份，非法/空字符串为 #VALUE!，空单元格和布尔走原数值转换；不使用 atan2 改写 ACOT 的负数/负零约定。目录示例按固定 [数学描述](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts) 取 0/2/0/6/0/10/1/0.1，原 antdv-next 菜单和参数表单自动复用，无新增 UI、命令、模型或 Worker/WASM 协议。

本小批唯一函数140→148、数学19→27，函数菜单出现项147→155，加插入函数子项共156；与续作起点137项相比共新增11个标量分支和数学子项。函数根9、Sheets71 / Docs54及扩展/上下文/溢出根保持。[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)补八个名称的转换/引用、参数数目、原错误与定义域；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)补负数、定义域变化、撤销重做与 native 重开；[Ribbon](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)与[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)补八个名称的派发和实际参数表单提交/撤销，全部未运行。完整数组/spill、Office 文件、大样本与全部功能验收仍有缺口。

同批复查既有 [SIN/COS/TAN/COT/SEC/CSC](#trigonometric-scalar-continuation) 的 minParams/maxParams 均为1。将参数数目判断移到通用参数求值前，错误从本地 #VALUE! 改为固定 FunctionNode 的 #N/A，并移除原分支中的重复长度判断；数值与倒数计算未变。既有公式草稿分开记录参数数目与参数内容错误，Facade 草稿追加六项 IFNA 错误处理。均未运行，没有新增算法名称、目录或入口。

[反三角批静态收尾](../UNIVER_TOOLBAR_PARITY.md#math-numbering-values-static-close)当时确认目录148项/8类、数学27、常用七项、函数根9与155个函数菜单出现项；引擎、目录及四份既有草稿的语法/引用和本证据链接通过。检查未执行求值器、Facade 或浏览器，草稿与完整功能仍未验收。

<a id="hyperbolic-log-atan2-scalar-continuation"></a>

## 双曲、指数对数与 ATAN2 标量续接

本批在既有同步 TS 权威求值器和数学目录补12项，继续复用 singleRangeValue、numericTextCoercion 和原输入事务。全部参数数目错误先于参数求值返回 #N/A；已有错误保留身份。空引用、布尔和数字/格式化字符串沿用既有转换，非法或空白文本为 #VALUE!。固定 [数学 function map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)、[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)、[primitive value objects](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts) 和 [BaseValueObject](../../../univer/packages/engine-formula/src/engine/value-object/base-value-object.ts) 是本批依据。

| 固定入口 | 参数与标量契约 |
| --- | --- |
| [Sinh](../../../univer/packages/engine-formula/src/functions/math/sinh/index.ts)、[Cosh](../../../univer/packages/engine-formula/src/functions/math/cosh/index.ts)、[Tanh](../../../univer/packages/engine-formula/src/functions/math/tanh/index.ts) | 单参数；Math.sinh/cosh/tanh，非有限输入或结果为 #NUM!；有限 ±1000 的 TANH 为 ±1 |
| [Coth](../../../univer/packages/engine-formula/src/functions/math/coth/index.ts) | 单参数；0 为 #DIV/0!，其他有限值为 1/tanh(x)；非有限输入按固定 tanh 错误后的 getReciprocal 返回 #VALUE! |
| [Sech](../../../univer/packages/engine-formula/src/functions/math/sech/index.ts) | 单参数；cosh 非有限先返回0，包括原始 NaN；固定溢出分支先于 2^27 判断，有限分母取倒数 |
| [Csch](../../../univer/packages/engine-formula/src/functions/math/csch/index.ts) | 单参数；0 为 #DIV/0!，sinh 溢出为0，原始 NaN 为 #VALUE!，有限分母取倒数 |
| [Exp](../../../univer/packages/engine-formula/src/functions/math/exp/index.ts)、[Ln](../../../univer/packages/engine-formula/src/functions/math/ln/index.ts)、[Log10](../../../univer/packages/engine-formula/src/functions/math/log10/index.ts) | 单参数；对应 Math 方法，非有限输入或结果为 #NUM!，LN/LOG10 的非正输入为 #NUM! |
| [Sign](../../../univer/packages/engine-formula/src/functions/math/sign/index.ts) | 单参数；正/负值为 ±1，其余为正0；保留 ±Infinity 的 ±1、NaN 与负零的正0 |
| [Log](../../../univer/packages/engine-formula/src/functions/math/log/index.ts) | 1–2参数，默认底数10；非正 number/base 为 #NUM!，底数1为 #DIV/0!；保留计算产生的 Infinity/NaN，原 base 错误先于 number 文本转换 |
| [Atan2](../../../univer/packages/engine-formula/src/functions/math/atan2/index.ts) | 两参数 `(x,y)`，使用 Math.atan2(y,x)；x 转换错误先于 y 错误；数值原点为 #DIV/0!，固定 x 布尔分支保留 FALSE/FALSE 为0 |

本地没有复刻上游 value-object 类或数组广播，使用既有标量类型和分支。SECH 的有限 cosh 必然位于固定 2^27 检查的有效范围内，直接保留溢出优先和倒数结果；SIGN 使用比较，保持固定 NaN/负零结果。LOG 未追加统一有限结果判断；非有限缓存的独立 Office/JSON 往返仍未核实。数组/spill 与批量 Worker/Rust/WASM 公式路径继续属于未完成项，不能以这些标量分支标记完整函数验收。

目录示例按固定来源接回既有 antdv-next 菜单和参数 Form：ATAN2 `4,3`、COSH `4`、COTH `2`、CSCH `1.5`、EXP `2`、LN `EXP(3)`、LOG `8,2`、LOG10 `100000`、SECH `30`、SIGN `10`、SINH `0.0342*1.03`、TANH `1`。该小批唯一函数148→160、数学27→39，函数菜单出现项155→167，含 picker 子项156→168；后续四项见下文。函数根9、Sheets71/Docs54及常驻/扩展/上下文/溢出根不变。模型、公开命令、历史和跨边界协议不变，没有性能优化结论。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)补三条真实实现声明，覆盖12名称、双曲转换/溢出、指数对数定义域与错误、SIGN 和 ATAN2 顺序/象限；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)补两条依赖编辑、错误缓存、撤销重做、native 保存重开及再次改值；[Ribbon 草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)和[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)补12名称派发与原参数表单提交/撤销。全部未运行，Office 文件、主题/焦点/IME、完整公式与大样本仍未验收。

<a id="even-odd-quotient-sqrtpi-scalar-continuation"></a>

## 奇偶取整、整商与 SQRTPI 标量续接

紧接双曲/对数批，在同一引擎/目录两文件新增四个标量分支，没有新 UI 或计算域。

| 固定来源 | 本地保留的可观察契约 |
| --- | --- |
| [Even](../../../univer/packages/engine-formula/src/functions/math/even/index.ts)、[Odd](../../../univer/packages/engine-formula/src/functions/math/odd/index.ts)、[math-kit.ceil](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts) | 单参数，向外取偶/奇整数；保留 ceil 在 precision=0 时的 epsilon 修正。非有限输入为 #VALUE!，EVEN 的零归正0，ODD(0/-0)为1；空引用/布尔和数值文本沿用既有转换 |
| [Quotient](../../../univer/packages/engine-formula/src/functions/math/quotient/index.ts)、[checkVariantsErrorIsArrayOrBoolean](../../../univer/packages/engine-formula/src/engine/utils/check-variant-error.ts) | 两参数，按顺序解包单格并拒绝布尔/已有错误，再对标量取 Number；空/空白文本为0，货币/百分比/日期文本为 #VALUE!。NaN 输入先于零分母返回 #VALUE!，零分母为 #DIV/0!，Math.trunc 保留负零和非有限结果 |
| [Sqrtpi](../../../univer/packages/engine-formula/src/functions/math/sqrtpi/index.ts) | 单参数，拒绝多格；负值为 #NUM!，其他值为 sqrt(x×π)，按固定源码保留负零、Infinity 和 NaN，不追加统一有限结果限制 |

四者的参数数目错误均先于参数求值返回 #N/A，singleRangeValue 复用现有单格/INDEX 解包。EVEN/ODD 直接实现固定零精度 ceil 算式，未抽取单次 helper；QUOTIENT 特有的 Number 文本规则不套用格式化解析，已有 denominator 错误仍先于 numerator 的字符串转换错误。

固定目录示例分别为 EVEN/ODD `1.5`、QUOTIENT `5,2`、SQRTPI `1`。唯一函数160→164、数学39→43；函数菜单出现项167→171，含 picker 子项168→172。函数根9、Sheets71/Docs54及扩展/上下文/溢出根保持，16个新增标量名称不等于16个已验收完整函数。

既有[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)补三条精度边界、零/符号、布尔/特殊文本、单格/错误/参数数目和非有限结果声明；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)补一条四依赖的编辑、错误缓存、撤销重做与 native 保存重开；[Ribbon](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)和[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)补四名称派发和原表单提交/撤销。全部未运行；完整数组、非有限缓存的 Office/JSON 往返、主题/焦点/IME 和批量性能仍未验收。

<a id="zero-argument-function-arity"></a>

## 既有零参数函数的数目错误续修

固定 [True](../../../univer/packages/engine-formula/src/functions/logical/true/index.ts)、[False](../../../univer/packages/engine-formula/src/functions/logical/false/index.ts)、[Now](../../../univer/packages/engine-formula/src/functions/date/now/index.ts)、[Today](../../../univer/packages/engine-formula/src/functions/date/today/index.ts) 的 minParams/maxParams 均为0，[FunctionNode](../../../univer/packages/engine-formula/src/engine/ast-node/function-node.ts) 对多余参数返回 #N/A。本地 TRUE/FALSE 原为 #VALUE!，NOW/TODAY 还会先传播参数错误；本批恢复 #N/A 身份及数目错误优先，删掉 NOW/TODAY 在通用求值后的重复长度判断。零参数正常结果、原时钟/日期系统、菜单和输入行为不变，没有新增函数名称或入口。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)修正两个旧错误断言并补一条四函数的普通/错误/范围多余参数、IFNA 和空参数构造声明；[既有 NA Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)把同一编辑扩为七格，补四个数目错误处理、一次撤销重做与 native 保存重开。均未运行；该零参数小修时目录为164项/8类，随后四个阶乘/组合名称另计，完整功能未验收状态保持。

<a id="factorial-combination-scalar-continuation"></a>

## 阶乘与组合数标量续接

在同一引擎/目录两文件追加四项，沿用 singleRangeValue、原格式化数字文本转换和严格参数数目 #N/A 判断。已有错误先于定义域判断，原参数表单和提交事务复用，没有新增模型、公开命令或 Worker/WASM 协议。

| 固定来源 | 保留的标量契约 |
| --- | --- |
| [Fact](../../../univer/packages/engine-formula/src/functions/math/fact/index.ts)、[Factdouble](../../../univer/packages/engine-formula/src/functions/math/factdouble/index.ts) | 单参数先 floor；负数为 #NUM!，0/1 为1，分别按步长1/2自高向低乘。结果非有限为 #NUM!；固定 calculateFactorial 对原始 NaN 不进入循环、返回1，不能擅自添加输入 finite 限制 |
| [Combin](../../../univer/packages/engine-formula/src/functions/math/combin/index.ts) | 两参数 floor，负值或 chosen>number 为 #NUM!，按固定乘除顺序计算组合数；数值/字符串原错误先于定义域判断 |
| [Combina](../../../univer/packages/engine-formula/src/functions/math/combina/index.ts) | 两参数 floor，允许重复选择；number=0 且 chosen>0 为 #NUM!，使用 C(number+chosen−1,number−1)，保留 (0,0) 为1 |

数值依据为固定 [calculateFactorial / calculateCombin](../../../univer/packages/engine-formula/src/basics/math.ts)、[数学 map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)与[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)。本地直接写入既有分支，按循环中的有限结果判断停止溢出计算；未引入一次性 helper 或对数近似。COMBIN/COMBINA 的部分 NaN/Infinity 组合按固定零次循环返回1，结果仍由原 finiteResult 判断。本地只处理标量/单格；FACT 和两种组合数的固定数组映射/广播尚未完成，FACTDOUBLE 固定也拒绝多格。完整数组、非有限缓存的文件往返和性能继续未验收。

目录示例沿固定来源：FACT `5`、FACTDOUBLE `6`、COMBIN/COMBINA `8,2`。唯一函数164→168、数学43→47，函数菜单出现项171→175，含 picker 子项172→176；函数根9、Sheets71/Docs54及扩展/上下文/溢出根保持。与反三角批148项相比，本批合计20个新增数学标量名称和子项，不等于20个已验收完整函数。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增两条，覆盖 floor/步长、零/布尔/单格/格式化文本、参数数目、定义域、错误顺序、固定 NaN/Infinity 特例和溢出；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增一条四依赖编辑、撤销重做、native保存重开与错误恢复；[Ribbon](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)和[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)补四名称派发及原 Ant 参数表单提交/撤销。均未运行，没有产品、类型、构建或性能通过结论。

<a id="rounding-scalar-continuation"></a>

## 2026-11-27 ROUND 系列与 TRUNC 标量续接

本批继续使用固定数学分类根与命令链路，在原同步 [formula-engine.ts](../../packages/office-studio/src/features/formula/formula-engine.ts) 补四个分支，修正既有 ROUND；目录、Ant 参数表单、公式输入、权限、内容历史和 native 保存沿用原实现。没有新增模型、公开命令、Facade、Worker/WASM 协议或计算域。参数数目错误先于求值返回 #N/A；ROUND 系列、TRUNC 按 number 转换在前处理错误，ROUNDBANK 先传播两原参数错误，再转换文字。

| 固定实现 | 本地标量契约与差异 |
| --- | --- |
| [Round](../../../univer/packages/engine-formula/src/functions/math/round/index.ts) | 将小数位先 trunc，按固定 epsilon 修正半值；数值小数位对负数先取绝对值再还原符号。删除原直接指数与 ±308 参数限制，输入/结果的有限性沿固定 value object 判断 |
| [Rounddown](../../../univer/packages/engine-formula/src/functions/math/rounddown/index.ts) | 使用 floor(base×factor+epsilon)，数值小数位时朝零舍入；保持整数及小数位取整 |
| [Roundup](../../../univer/packages/engine-formula/src/functions/math/roundup/index.ts) | 使用 ceil(base×factor−epsilon)，数值小数位时朝外舍入；避免已有整数倍被浮点误差推进一格 |
| [Trunc](../../../univer/packages/engine-formula/src/functions/math/trunc/index.ts) | 可省略小数位，默认0；直接对带符号乘积加 epsilon 后 trunc，不额外限制非有限结果。固定表达式的负整数边界，例如 -8 得到 -7、-1.2 保留一位得到 -1.1，按源码保留，实际 Univer/Office 样本未核实 |
| [Roundbank](../../../univer/packages/engine-formula/src/functions/math/roundbank/index.ts) | 精度 trunc 后，>16 原样返回 number，<-16 返回0；其他精度按乘积 toFixed(8) 和 1e-8 半值容差舍入到最近偶数，不增加非有限结果拒绝 |

数值依据为 [primitive-object.ts 的 round/floor/ceil](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)、[math-kit.ts 的 round/floor/ceil/baseEpsilon/roundHalfEven](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts)和[原参数字符串转换](../../../univer/packages/engine-formula/src/engine/utils/check-variant-error.ts)。空引用对应 BaseValueObject 的数值0。前三个函数的小数位为 Boolean 时保留固定有符号路径和原非有限结果；例如负值与 TRUE 小数位的结果可与数值1不同，不把 Boolean 提前丢失为纯数值。TRUNC / ROUNDBANK 的固定数组映射，以及 ROUND 系列 value-object 数组分派/广播仍未实现；本批只实现标量/单格，不能据目录声称完整函数。

已静态读取固定 [ROUND 测试](../../../univer/packages/engine-formula/src/functions/math/round/__tests__/index.spec.ts)、[ROUNDDOWN 测试](../../../univer/packages/engine-formula/src/functions/math/rounddown/__tests__/index.spec.ts)、[ROUNDUP 测试](../../../univer/packages/engine-formula/src/functions/math/roundup/__tests__/index.spec.ts)、[TRUNC 测试](../../../univer/packages/engine-formula/src/functions/math/trunc/__tests__/index.spec.ts)和[ROUNDBANK 测试](../../../univer/packages/engine-formula/src/functions/math/roundbank/__tests__/index.spec.ts)，没有运行上游或本地求值器。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增三条真实求值声明，覆盖截断精度、半值/整数 epsilon、负数/布尔小数位、默认/空引用、单格/INDEX、参数数目、格式化文本、错误优先、非有限和负零；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增一条五依赖输入、精度修改、撤销重做、native 保存重开与错误恢复。[Ribbon](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)派发和[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)的原参数表单提交/撤销循环补五项，包含既有 ROUND。全部未运行，真实 Office/JSON 非有限缓存往返、主题/焦点/IME、完整数组和性能未验收。

新增目录示例为 ROUNDBANK `2.345,2`、ROUNDDOWN/ROUNDUP `3.2,0`、TRUNC `0.45,1`，沿固定描述目录。唯一名称168→172、数学47→51，菜单出现项175→179，含 picker 子项176→180；ROUND 只修原名称，不重复计数。函数根9、Sheets71/Docs54以及扩展/上下文/溢出根保持。

<a id="ceiling-floor-scalar-continuation"></a>

## 2026-11-27 CEILING / FLOOR 七个变体续接

在原引擎与目录两份生产文件追加七个数学名称和标量分支，复用 singleRangeValue、原格式化数字文本转换及同步权威求值器。点号函数名称由已有 parser 和参数表单构造路径处理；没有新增解析 DTO、领域字段、公开命令、Facade 或 Worker/WASM 协议。

| 固定实现 | 保留的标量契约 |
| --- | --- |
| [Ceiling](../../../univer/packages/engine-formula/src/functions/math/ceiling/index.ts)、[Floor](../../../univer/packages/engine-formula/src/functions/math/floor/index.ts) | 两个必需参数；正 number 与负 significance 为 #NUM!。保留带符号步长：负 number / 正步长时 CEILING 朝零、FLOOR 朝外；负步长时方向相反。number=0 先返回正0；CEILING 的零步长返回0，FLOOR 的非零 number / 零步长为 #DIV/0! |
| [CeilingMath](../../../univer/packages/engine-formula/src/functions/math/ceiling-math/index.ts)、[FloorMath](../../../univer/packages/engine-formula/src/functions/math/floor-math/index.ts) | 1–3 参数，省略步长/模式分别为1/0；步长按绝对值处理。负 number 的非零 mode 切换方向，mode 不截断，负值、布尔或非零小数均按固定 !==0 判断；零 number 或步长为0 |
| [CeilingPrecise](../../../univer/packages/engine-formula/src/functions/math/ceiling-precise/index.ts)、[FloorPrecise](../../../univer/packages/engine-formula/src/functions/math/floor-precise/index.ts) | 1–2 参数，省略步长为1、忽略步长符号；分别朝正/负无穷方向舍入，number 或步长为0返回正0 |
| [IsoCeiling](../../../univer/packages/engine-formula/src/functions/math/iso-ceiling/index.ts) | 固定实现继承 CeilingPrecise；ISO.CEILING 独立注册与目录名称保留，同一标量分支产生相同结果 |

七者按固定 min/max 在求值前返回参数数目 #N/A。除 FLOOR.MATH 外，原参数错误先于文本转换错误传播；FLOOR.MATH 按 number、步长、mode 逐个转换，非法 number/步长文本可以先于后面的原错误返回 #VALUE!。空引用为0、Boolean 数值转换和货币/百分比文本沿原转换，空/非法文字为 #VALUE!。模式转换和错误处理先于零值返回，不能以 number/步长为0跳过错误参数。

数值来源是 [math-kit 的 floor/ceil/baseEpsilon](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts)及 [NumberValueObject.create / 字符串转换](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)。本地把现代变体的负步长写成等价绝对值算式，保留 precision=0 的 epsilon、负零与大整数边界；不新增输入/结果有限性过滤，固定 NumberValueObject.create 可以保留 Infinity / NaN。未运行 Univer 或 Office 样本，非有限缓存与独立文件往返仍未核实。固定七者的数组扩展/广播与 spill 尚待实现，本批标量/单格候选不代表完整函数已验收。

固定 [数学 map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)和[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)决定分类和示例：CEILING `2.5,1`、CEILING.MATH/FLOOR.MATH `-5.5,2,-1`、CEILING.PRECISE `4.3,2`、FLOOR `3.7,2`、FLOOR.PRECISE `-3.2,-1`、ISO.CEILING `-4.3,2`。唯一名称172→179、数学51→58；函数菜单出现项179→186，含 picker 子项180→187。常用七项仍为原目录投影，函数根9、Sheets71/Docs54及扩展/上下文/溢出新增根为0。

已读取固定 [CEILING](../../../univer/packages/engine-formula/src/functions/math/ceiling/__tests__/index.spec.ts)、[CEILING.MATH](../../../univer/packages/engine-formula/src/functions/math/ceiling-math/__tests__/index.spec.ts)、[CEILING.PRECISE](../../../univer/packages/engine-formula/src/functions/math/ceiling-precise/__tests__/index.spec.ts)、[ISO.CEILING](../../../univer/packages/engine-formula/src/functions/math/iso-ceiling/__tests__/index.spec.ts)、[FLOOR](../../../univer/packages/engine-formula/src/functions/math/floor/__tests__/index.spec.ts)、[FLOOR.MATH](../../../univer/packages/engine-formula/src/functions/math/floor-math/__tests__/index.spec.ts)和[FLOOR.PRECISE](../../../univer/packages/engine-formula/src/functions/math/floor-precise/__tests__/index.spec.ts)测试源码，未执行。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增三条真实求值声明，覆盖带符号步长、模式/default、点号名称、目录示例、文本/空引用/单格/INDEX、零步长、错误优先、参数数目、epsilon、负零与非有限结果；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增一条七依赖的真实输入、步长/模式修改、撤销重做、native保存重开和错误恢复。[Ribbon 草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)原名称派发列表与[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)原 Ant 参数表单提交/撤销循环补七项。全部运行验证继续后置，完整数组、自动范围/批量插入、Office/主题/焦点/IME/性能仍未验收。

<a id="integer-aggregate-mround-continuation"></a>

## 2026-11-27 GCD / LCM / MULTINOMIAL / MROUND 续接

在原引擎和目录两份生产文件补四个数学名称与真实求值。GCD / LCM / MULTINOMIAL 复用稀疏区域读取、既有数字文本转换及错误值，区域内实际单元格按行再列处理，缺失或 null 单元格略过；数字文本会转换，布尔值会返回 #VALUE!，不能套用 SUM 对区域文字/布尔的忽略规则。参数按输入顺序处理，区域内按坐标顺序传播第一个错误；得到 GCD=1 或 LCM=0 后仍处理后面的值。

| 固定实现 | 保留的契约 |
| --- | --- |
| [Gcd](../../../univer/packages/engine-formula/src/functions/math/gcd/index.ts) | 1–255 参数，Math.floor 后的负数或 ≥2^53 输入为 #NUM!；从0聚合最大公约数，全空为0 |
| [Lcm](../../../univer/packages/engine-formula/src/functions/math/lcm/index.ts) | 1–255 参数，负数为 #NUM!，输入没有 GCD 的单值上限；全空为0，正常零值参加计算。最终结果非有限或 ≥2^53 为 #VALUE!，最终判断发生在所有输入处理后，后面的原错误仍可先返回，有限大结果也能被后面的0归零 |
| [Multinomial](../../../univer/packages/engine-formula/src/functions/math/multinomial/index.ts) | 1–255 参数，负数为 #NUM!；每个 floor 后的数累加，一旦和>170立即 #NUM!，先于更晚的错误参数；输出 sum! / 各数阶乘乘积，全空/全0为1 |
| [Mround](../../../univer/packages/engine-formula/src/functions/math/mround/index.ts) | 两个必需参数，允许单格区域，拒绝较大区域；原 number/multiple 错误先于布尔拒绝和 Number 转换。空文本、空格文本和空引用可转0，货币/百分比文本不按格式化数值转换；NaN 为 #VALUE!。multiple=0 返回正0，非零 number/multiple 异号为 #NUM!，其余按 round(number/multiple,0)×multiple |

数值来源为 [calculateGcd / calculateLcm / calculateFactorial](../../../univer/packages/engine-formula/src/basics/math.ts)、[round/baseEpsilon](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts)及 [StringValueObject 转换](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)。本地把现有 FACT/FACTDOUBLE 的同一阶乘循环移入实际复用的 factorial 方法，供 MULTINOMIAL 的分子/分母调用；原阶乘取整、负数错误和有限结果策略保持。MROUND 保留固定零精度 epsilon、大整数偏移、负零及 Infinity/NaN 结果，不另加有限性过滤。

可复查差异：固定 calculateGcd 以余数 !==0 循环，NaN 余数没有终止条件；本地按正余数继续迭代，余数成为 NaN 时结束并传递 NaN，LCM 再按固定末尾规则返回 #VALUE!。合法非负有限整数路径采用相同欧几里得算法；这条非有限输入路径没有固定运行对照，继续记未核实。三个聚合函数接回矩形/单格区域，不代表显式数组常量、完整动态数组及 spill 已实现；MROUND 的较大数组拒绝来自固定实现。

已读取固定 [GCD](../../../univer/packages/engine-formula/src/functions/math/gcd/__tests__/index.spec.ts)、[LCM](../../../univer/packages/engine-formula/src/functions/math/lcm/__tests__/index.spec.ts)、[MULTINOMIAL](../../../univer/packages/engine-formula/src/functions/math/multinomial/__tests__/index.spec.ts)和 [MROUND](../../../univer/packages/engine-formula/src/functions/math/mround/__tests__/index.spec.ts)测试源码，未执行。固定[数学 map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)及[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)提供分类与示例：GCD/LCM/MULTINOMIAL `5,2`，MROUND `10,3`；已有分类菜单及 Ant 参数表单直接复用。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增三条声明，覆盖空值、文本/布尔转换、255/256参数、负数与2^53/170边界、反序插入的稀疏区域错误顺序、MROUND 单格约束/错误优先/epsilon/非有限结果及目录示例；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增一条四依赖的范围重算、文本小数、布尔/原错误恢复、撤销重做与 native 保存重开。[Ribbon 草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)原名称派发列表和[浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)原 Ant 参数提交/撤销循环各补四项，全部未运行。

唯一目录179→183、数学58→62、菜单出现项186→190、含 picker 子项187→191；常用七项仍引用原目录对象，函数根9、Sheets71/Docs54及扩展/上下文/溢出新增根为0。模型、公开命令、权限、单格输入、内容历史、native 保存及 Worker/WASM 协议沿原链路。完整函数/数组、自动范围与批量插入、Office 样本、真实 IME/主题/焦点/两后端和性能仍未验收。

<a id="abs-sqrt-int-mod-power-sumsq-continuation"></a>

## 2026-11-27 ABS/SQRT/INT/MOD/POWER 与 SUMSQ 续核

本批沿用[原权威求值器](../../packages/office-studio/src/features/formula/formula-engine.ts)，将五个原数学分支移到全局参数求值/错误扫描之前，按固定顺序处理参数数目、单格区域、文本转换和错误。没有新增引擎、领域模型、公开命令或计算协议。

| 名称与固定源码 | 本批行为 |
| --- | --- |
| [ABS](../../../univer/packages/engine-formula/src/functions/math/abs/index.ts)、[SQRT](../../../univer/packages/engine-formula/src/functions/math/sqrt/index.ts) | 各1参数；数字文本、日期系统、Boolean/null和单格区域沿现有转换；非有限值为 NUM，SQRT 负数为 NUM并保留 sqrt(-0)。 |
| [INT](../../../univer/packages/engine-formula/src/functions/math/int/index.ts) | 各1参数；先按固定 stripErrorMargin 的12/15位、整数长度及1e-10误差规则修正再 floor，5.999999999999999 得6；原非有限数不新增 finite 限制。 |
| [MOD](../../../univer/packages/engine-formula/src/functions/math/mod/index.ts) | 2参数依次转换；零/null/false 除数先返回 DIV/0，数值路径保留 abs(divisor)×1125900000000 的界限和 finite 规则；true 除数沿固定 Boolean 路径保留未过滤结果。 |
| [POWER](../../../univer/packages/engine-formula/src/functions/math/power/index.ts) | 2参数依次转换；指数 NaN 为 VALUE，其他非有限数为 NUM；0 的负/零/正指数分别为 DIV/0、NUM、正0，其他非有限结果为 NUM。 |
| [SUMSQ](../../../univer/packages/engine-formula/src/functions/math/sumsq/index.ts) | 新增1–255参数平方和；标量数字文本/Boolean参与，null为0；cell/range/冒号与INDEX引用按行/列遍历，忽略文本/Boolean/空值，保留原错误与固定算术结果。 |

参数数目按既有数学契约返回 N/A。前参数文本转换错误先于后参数原错误；不再由旧 evaluated.map 的原错误扫描抢先返回。转换、MOD/POWER、ABS/SQRT 数值语义来自 [primitive-object.ts](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)，INT精度来自 [math-kit.ts](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts)，全部仅阅读固定源码。

SUMSQ 区域文本语义已追到 [BaseReferenceObject.getCellValueObject/toArrayValueObject](../../../univer/packages/engine-formula/src/engine/reference-object/base-reference-object.ts) 和 [ValueObjectFactory/transformToValueObject](../../../univer/packages/engine-formula/src/engine/value-object/array-value-object.ts)：固定测试中的未加类型字符串100/2.34会先转Number，实际声明为文本的单元格仍是String，应忽略；不能据测试数组推断文本引用也参与计算。本地复用 [isSheetTextNumberFormat](../../packages/office-studio/src/features/sheet-workbench/sheet-number-format.ts)，将声明为数值且格式为 @/@@@ 的引用按文本处理，与固定 [isTextFormat](../../../univer/packages/core/src/shared/numfmt/utils.ts)一致。

固定 SUMSQ 在平方结果后还调用 NumberValueObject.plus；[BaseValueObject.convertToNumberObjectValue/plus](../../../univer/packages/engine-formula/src/engine/value-object/base-value-object.ts)与 plusBy 将平方计算错误转换为 VALUE。有限平方累加溢出为 NUM；同一区域后续数值会将累积错误转为 VALUE，后续原单元格错误则覆盖并结束遍历；不同参数在当前参数结束后立即返回错误。本批保留这些源码路径，未运行固定对照，不将它们泛化为其他聚合函数语义。

已读固定 [ABS](../../../univer/packages/engine-formula/src/functions/math/abs/__tests__/index.spec.ts)、[SQRT](../../../univer/packages/engine-formula/src/functions/math/sqrt/__tests__/index.spec.ts)、[INT](../../../univer/packages/engine-formula/src/functions/math/int/__tests__/index.spec.ts)、[MOD](../../../univer/packages/engine-formula/src/functions/math/mod/__tests__/index.spec.ts)、[POWER](../../../univer/packages/engine-formula/src/functions/math/power/__tests__/index.spec.ts)与 [SUMSQ](../../../univer/packages/engine-formula/src/functions/math/sumsq/__tests__/index.spec.ts)测试源码，未执行。[目录](../../packages/office-studio/src/features/sheet-workbench/sheet-function-catalog.ts)只新增 SUMSQ，示例3,4来自固定[数学描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)；常用七项、分类顺序、Ant参数表单和 sheet.operation.insert-function 共用原链路。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增4声明，覆盖转换/精度/参数数目/错误顺序、Boolean divisor、零底数、typed text/@/@@@、INDEX与稀疏区域算术错误顺序；[Facade草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增2声明，覆盖六依赖的重算/错误恢复及文本引用与标量区别、撤销重做/native保存重开。原 [Ribbon名称派发](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)和 [Ant参数提交/撤销循环](../../tests/browser/sheet-toolbar.spec.ts)各补6项，全部运行0。

唯一目录183→184、数学62→63、函数菜单出现项190→191、含picker子项191→192；五个修正名称不重复计数。函数根9、Sheets71/Docs54及常驻/扩展/上下文/溢出新增根0。完整数组/broadcast/spill、其他引用返回函数和类型细节、自动范围/批量插入、全部目录/迭代计算/Pro、Office、真实IME/主题/焦点/两后端和性能仍未完成或未验收；本批不提升功能组验收状态。

<a id="pairwise-sumproduct-seriessum-continuation"></a>

## 2026-11-27 SUMX / SUMPRODUCT / SERIESSUM 续接

原[公式引擎](../../packages/office-studio/src/features/formula/formula-engine.ts)新增五个名称的求值分支，沿用同步权威引擎。引用参数复用 evaluateReferenceOperand：cell、矩形/冒号及 INDEX 保留来源和形状；共享矩阵描述只保存行列数量与读取方法，按行再列访问真实模型，没有建立另一份二维数据或第二个求值器。SUMSQ 同时复用原引用读取和 @/@@@ 类型转换。

| 固定来源 | 本批保留的行为 |
| --- | --- |
| [SUMX2MY2](../../../univer/packages/engine-formula/src/functions/math/sumx2my2/index.ts)、[SUMX2PY2](../../../univer/packages/engine-formula/src/functions/math/sumx2py2/index.ts)、[SUMXMY2](../../../univer/packages/engine-formula/src/functions/math/sumxmy2/index.ts) | 各2参数，顶层原错误优先，然后比较元素总数；行列形状可以不同，数量不同时为 N/A。单元素先解包/传播错误，Null 为 VALUE，Boolean 或无法经原始 Number 转换的文本为 DIV/0；空/空格文本为0。多元素先略过左侧文本/Boolean/Null，右侧对应错误可被遮蔽；左错误先于右侧略过规则，无有效数值对为 DIV/0。 |
| [SUMPRODUCT](../../../univer/packages/engine-formula/src/functions/math/sumproduct/index.ts) | 1–255参数；一个参数沿数组 sum 规则忽略文本/Boolean/Null，保留累积错误顺序。多个参数先扫描首区域原错误，再按参数顺序检查后续顶层错误与行列一致；后续标量可广播，后续单格引用仍是区域，不能替代标量。每个位置按参数顺序计算，非数值将当前乘积重置0，但继续检查后面的原错误；末尾直接返回数值，不新增统一有限性过滤。 |
| [SERIESSUM](../../../univer/packages/engine-formula/src/functions/math/seriessum/index.ts) | 严格4参数；任一顶层 Null 先为 N/A。x/n/m 先完成数组形状、Boolean 和原错误检查，再依次转换字符串；单格区域中的 Null 可转0。coefficients 按行再列读取，文本转换、Boolean 拒绝、原错误传播；缺失/Null 系数占用原索引。计算 Σ coefficient[i]×x^(n+i×m)，n/m 不取整，0系数仍参与乘法；所有系数处理后才将非有限结果转为 NUM。 |

SUMX 的字符串有限性依据是固定 [isRealNum](../../../univer/packages/core/src/shared/generate.ts)：Number 路径只拒绝 NaN，因此字符串1e999可以成为 Infinity，不能套用 numericTextCoercion 或 Number.isFinite。原 [flatten/iterator/sum](../../../univer/packages/engine-formula/src/engine/value-object/array-value-object.ts)和 [NumberValueObject.plus/plusBy](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)决定 SUMPRODUCT 单参数的错误顺序：有限加法溢出为 NUM，同一区域后续数值将累积错误转 VALUE，后续原单元格错误覆盖并结束；多参数结果可以保留 Infinity/NaN。

引用类型追到固定 [getCellValueObject](../../../univer/packages/engine-formula/src/engine/reference-object/base-reference-object.ts)与 ValueObjectFactory：数值类型的 @/@@@ 格式先转文本；其余非有限数值引用经 createNumberValueObjectByRawValue 转 NUM。共享读取方法在这一次转换中保留该顺序；SUMSQ 不再重复检查已转换的引用。SERIESSUM 的分阶段校验来自 [check-variant-error.ts](../../../univer/packages/engine-formula/src/engine/utils/check-variant-error.ts)，本地未将两个阶段合并为逐参数即时转换。

目录分类依据固定[数学 map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)和[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)。SUMPRODUCT 示例 C2:C5,D2:D5、SERIESSUM 示例0.785398163,0,2,A1:A4沿用描述参数；SUMX 三者将上游 A2:A8/B2:B8 示例改为可直接验证单对数值的3,4，公式仍支持这两个区域参数。中文短标签用于本地参数表单，不增加函数或命令。Context7 没有1.6.0索引，检索返回当前注册/精度文档，本批行为和样本均以固定本地源码为准。

已读取固定 [SUMX2MY2](../../../univer/packages/engine-formula/src/functions/math/sumx2my2/__tests__/index.spec.ts)、[SUMX2PY2](../../../univer/packages/engine-formula/src/functions/math/sumx2py2/__tests__/index.spec.ts)、[SUMXMY2](../../../univer/packages/engine-formula/src/functions/math/sumxmy2/__tests__/index.spec.ts)、[SUMPRODUCT](../../../univer/packages/engine-formula/src/functions/math/sumproduct/__tests__/index.spec.ts)和 [SERIESSUM](../../../univer/packages/engine-formula/src/functions/math/seriessum/__tests__/index.spec.ts)测试源码，未执行。固定不同形状的六元素样本对应 SUMX 结果28/56/8；幂级数样本对应0.707103215205。

[原公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增5声明：标量/空格/原始 Number 转换、参数数目、不同形状配对、反序区域/INDEX/冒号、左侧略过与错误顺序、@/@@@、单/多参数算术、255/256参数、分阶段检查及稀疏系数索引。[原 Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增2声明，覆盖五依赖的源编辑、typed text/Boolean 差异、撤销重做、native保存重开和错误恢复。原 [Ribbon 名称派发](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)与 [Ant参数提交/撤销循环](../../tests/browser/sheet-toolbar.spec.ts)各补5个样本；全部运行0。

唯一目录184→189、数学63→68、函数菜单出现项191→196、含picker子项192→197；常用七项仍是原对象投影，函数根9、Sheets71/Docs54及常驻/扩展/上下文/溢出新增根0。数组常量、区域算术产生的派生数组、完整广播/spill、显式数组缺失项与引用Null区别、其他引用返回函数、自动范围/批量插入、完整目录/迭代与 Pro 继续未完成或未验收。此次矩阵读取仍同步遍历逻辑范围，大样本 Worker/Rust/WASM、预算/取消、内存和真实性能没有本批实测；没有声称优化或完整函数已验收。真实主题/焦点/IME、两渲染后端、Office文件与独立重开仍待统一验证。

<a id="roman-radix-conversion-continuation"></a>

## 2026-11-27 罗马数字与进制转换续接

[原公式引擎](../../packages/office-studio/src/features/formula/formula-engine.ts)新增 ARABIC/ROMAN/BASE/DECIMAL 四个名称；单格引用沿用 singleRangeValue，字符串转换复用现有日期、时间和数字格式解析，DECIMAL 的数位文本保持原值。同步计算仍在 formula 层，不向组件、store、Docs 编号服务或另一份引擎添加计算。模型、统一命令、内容历史和 Worker/WASM 协议均无变化。

| 固定来源 | 保留的可观察行为 |
| --- | --- |
| [ARABIC](../../../univer/packages/engine-formula/src/functions/math/arabic/index.ts) | 严格1参数，原错误直接传播；Null 为0，数字/Boolean为VALUE。文本大小写不敏感、不trim、长度上限255，允许首负号及负零。逐字符保留原四项判定与减法，不改为经典罗马正则；MXMVII=1997、IL=49、ID=499，四连字符和VX为VALUE。 |
| [ROMAN](../../../univer/packages/engine-formula/src/functions/math/roman/index.ts)与[罗马形式表](../../../univer/packages/engine-formula/src/basics/math.ts) | 1–2参数，默认form0；先传播number/form原错误，再依次字符串转换。number/form取floor后范围0–3999/0–4，Boolean form的true=0、false=4。499五种结果为CDXCIX/LDVLIV/XDIX/VDIV/ID。原数值NaN配有效form返回空串。 |
| [BASE](../../../univer/packages/engine-formula/src/functions/math/base/index.ts) | 2–3参数，默认minLength0；所有顶层原错误先于字符串转换，再按number/radix/minLength顺序转换并取floor。number范围[0,2^53)，radix2–36，minLength≥0；Boolean/Null参与数字转换，大写输出并补0。固定实现没有255长度上限，原数值NaN保留NAN文本。 |
| [DECIMAL](../../../univer/packages/engine-formula/src/functions/math/decimal/index.ts)与[isRealNum](../../../univer/packages/core/src/shared/generate.ts) | 严格2参数，原错误优先，radix字符串转换/floor；text直接转字符串，Null按固定getValue为0。数字形式文本检查非负、小于2^53及整数；true/false为NUM，空白为0。逐字符检查进制，结果只将NaN转NUM；文本1e3在16进制为483，数值1e3为4096。原数值NaN radix跳过字符分支并沿parseInt默认进制，保留并记录此固定行为。 |

ROMAN 沿同一[固定词元及形式表](../../../univer/packages/engine-formula/src/basics/math.ts)从大到小贪心取词元，替换内部二分逐次减法；4000是上游查找哨兵，本地不加入可输出词元。BASE 使用原生padStart替代临时数组补零；没有端到端基准，不宣称性能优化。字符串转换追到 [StringValueObject](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)，未将DECIMAL数位文本按金额/百分比先转换。

分类、参数数目和示例核对固定[数学map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)及[数学描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)；四份固定测试源码 [ARABIC](../../../univer/packages/engine-formula/src/functions/math/arabic/__tests__/index.spec.ts)、[ROMAN](../../../univer/packages/engine-formula/src/functions/math/roman/__tests__/index.spec.ts)、[BASE](../../../univer/packages/engine-formula/src/functions/math/base/__tests__/index.spec.ts)、[DECIMAL](../../../univer/packages/engine-formula/src/functions/math/decimal/__tests__/index.spec.ts)只读未执行。

[既有公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增4声明，覆盖五种罗马形式、字符规则/255边界、Boolean/Null、负零、进制/2^53/填充长度、原错误顺序、单格/矩形区别和固定非有限路径。[Facade草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增1声明，四条相互依赖公式覆盖文本/形式编辑、重算/错误恢复、撤销重做及native保存重开。原 [Ribbon名称派发](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)和 [Ant参数表单提交/撤销循环](../../tests/browser/sheet-toolbar.spec.ts)各补4项，全部运行0。

唯一目录189→193、数学68→72、函数菜单出现项196→200、含picker子项197→201；常用七项投影与函数根9不变，各类新增根0。当前RuntimeValue未承载完整数组输出；固定ARABIC映射及其他三函数的数组广播/spill、完整目录/Pro、自动引用/批量插入、非有限值Office/JSON往返及大样本Worker/Rust/WASM仍有未完成或未核实项。主题、焦点、原生IME、两后端、Office样本与性能全部留待统一验证，SHEET-09/SHEET-10继续未验收。

<a id="rand-randbetween-continuation"></a>

## 2026-11-27 RAND / RANDBETWEEN 续接

原同步引擎新增两个随机函数，复用同一求值、重算和单格引用读取。固定 [RAND](../../../univer/packages/engine-formula/src/functions/math/rand/index.ts) 严格零参数，直接返回 Math.random()，错误参数数目为 #N/A。[RANDBETWEEN](../../../univer/packages/engine-formula/src/functions/math/randbetween/index.ts) 严格两个参数，先解 bottom 的单格形状和原错误，再解 top，随后拒绝 Boolean。上下界使用原始 Number 转换，空文本、空格和 Null 为0，百分比/金额文本不按格式化数字转换；NaN 为 #VALUE!，原始 bottom>top 为 #NUM!，再 ceil/floor 并按固定公式生成含边界整数。

保留固定空整数区间行为：0.1,0.2 返回1，-1.9,-1.1 返回-1；没有添加统一 finite 拒绝，两个 Infinity 上界/下界可返回 NaN。固定 [RANDBETWEEN 测试](../../../univer/packages/engine-formula/src/functions/math/randbetween/__tests__/index.spec.ts) 已只读；分类和示例来自 [数学 map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)及 [数学描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)。RAND 示例为空，RANDBETWEEN 使用固定 1,100；两项进入现有数学菜单和 Ant 参数表单，RAND 直接进入零参数公式编辑。

[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增2声明，检查真实随机区间、整数/单值上下界、空白/Null、单格/矩形、Boolean/原错误顺序、非有限值及严格参数数目，不断言相邻随机结果必然不同。[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增1声明，覆盖 manual 缓存冻结、上下界修改、撤销重做、显式重算和 native 保存重开。原 [Ribbon 名称派发](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)循环补两个名称；原 [Ant 表单浏览器草稿](../../tests/browser/sheet-toolbar.spec.ts)补固定7,7的提交/撤销及 RAND 的直接编辑/提交/撤销，全部运行0。完整 volatile 依赖调度、重算时钟/种子、动态随机数组、Office 缓存和真实运行仍待核实或验收。

唯一目录193→195、数学72→74、菜单出现项200→202、含 picker 子项201→203；常用七项、函数根9和各类新增根0保持。该批固定数学 map 注册82项，对应74，仍缺 AGGREGATE、MDETERM、MINVERSE、MMULT、MUNIT、RANDARRAY、SEQUENCE、SUBTOTAL；[数学名称枚举](../../../univer/packages/engine-formula/src/functions/math/function-names.ts)另有未在数学 map 注册的 LET，但它已在[逻辑 map](../../../univer/packages/engine-formula/src/functions/logical/function-map.ts)注册，应归入逻辑缺口。当前 RuntimeValue 未承载完整数组输出/spill，不能只补矩阵/数组菜单冒充功能。既有数组/引擎方案、自动范围/批量插入、迭代设置、完整 Pro 继续保留；模型/公开命令/历史/Worker/WASM 协议无新增，SHEET-09/10仍未验收。

<a id="mdeterm-matrix-continuation"></a>

## 2026-11-27 MDETERM 续接

[固定实现](../../../univer/packages/engine-formula/src/functions/math/mdeterm/index.ts)严格一个参数，先按行读取内容、保留原错误并拒绝 Null/Boolean，再检查方阵；数值文本走既有数字转换。[calculateMdeterm / performLUDecomposition](../../../univer/packages/engine-formula/src/basics/math.ts)对一阶直接返回原值、二阶直接计算，三阶起先转置再按 LU 顺序求值、交换主元，非有限或绝对值不大于 1e-11 的主元返回0，最终零规范为正零。本地复用引用矩阵读取，读取时填转置矩阵；不改源模型，不新建求值或历史引擎。

已读取[固定测试](../../../univer/packages/engine-formula/src/functions/math/mdeterm/__tests__/index.spec.ts)、[数学 map](../../../univer/packages/engine-formula/src/functions/math/function-map.ts)及[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/math.ts)，目录示例为 A1:C3。原[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增2声明，覆盖正常/奇异方阵、主元交换、两阶与三阶阈值差异、30阶单位阵、数值文本、行序错误、形状、参数数目和源模型不变；[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增1声明，覆盖矩阵修改、错误恢复、manual 缓存、撤销重做和 native 保存重开。原[Ribbon 名称派发](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)与[Ant 参数表单](../../tests/browser/sheet-toolbar.spec.ts)补 MDETERM，均未运行。

目录195→196、数学74→75，函数根不变。固定数学 map 注册82项，本地对应75，仍缺 AGGREGATE、MINVERSE、MMULT、MUNIT、RANDARRAY、SEQUENCE、SUBTOTAL；LET 已注册于逻辑 map，继续保留本地求值缺口。标量非有限值沿既有路径保留，引用非有限值沿公共读取返回 #NUM!；数组常量、派生数组、spill、真实 Office、性能和大型 Worker/Rust/WASM 路径仍未完成或未验收，不据算法改写宣称优化。

<a id="information-types-reference-continuation"></a>

## 2026-11-27 信息类型与引用续接

固定 [ERROR.TYPE](../../../univer/packages/engine-formula/src/functions/information/error-type/index.ts)按原值映射 NULL1、DIV/0 2、VALUE3、REF4、NAME5、NUM6、NA7、GETTING_DATA8、CALC14；文本 "#NULL!" 同样得到1，其余值返回 #N/A。[N](../../../univer/packages/engine-formula/src/functions/information/n/index.ts)对矩形取首格、原错误传播、所有文本为0，Boolean/Null 用 Number 转换。[TYPE](../../../univer/packages/engine-formula/src/functions/information/type/index.ts)对多格引用返回64且不读取内容；单格数值/Null1、文本2、Boolean4、错误16。[ISREF](../../../univer/packages/engine-formula/src/functions/information/isref/index.ts)消费引用身份，对含错误的有效引用仍返回 true；文本/数值/Boolean/错误值返回 false。四项严格一个参数，错误数目为 #N/A。

本地使用 evaluateArrayArgument 保留单格、矩形、INDEX 和冒号表达式的引用对象；复用 formulaArgumentMatrix/valueAt 读取首格。[固定引用 getCellValueObject](../../../univer/packages/engine-formula/src/engine/reference-object/base-reference-object.ts)会把 number 类型且 @ 格式的值作为文本，本地沿既有 formulaReferenceCellValue 保留该行为。ERROR.TYPE 多格数组广播尚无输出协议，本地仍返回 #VALUE!，保留为明确缺口；TYPE 的128对象类型、数组常量及 ISREF(XYZ1) 等在解析阶段失败的名称路径也未对齐，不把四个名称计作完整功能验收。

已读取四项固定测试：[ERROR.TYPE](../../../univer/packages/engine-formula/src/functions/information/error-type/__tests__/index.spec.ts)、[N](../../../univer/packages/engine-formula/src/functions/information/n/__tests__/index.spec.ts)、[TYPE](../../../univer/packages/engine-formula/src/functions/information/type/__tests__/index.spec.ts)、[ISREF](../../../univer/packages/engine-formula/src/functions/information/isref/__tests__/index.spec.ts)。示例取[信息描述](../../../univer/packages/engine-formula/src/services/formula/function-list/information.ts)的 "#NULL!"、7、A2、A1；原公式草稿新增2声明，Facade 新增1声明，原名称/权限循环和 Ant 表单循环补四项提交/撤销，运行0。RAND 构造器本轮补入零参数白名单；Workbench 按 entry.example 为空直接调用原 beginEdit，没有另外的名称白名单。

该批 AST 核对[信息 map](../../../univer/packages/engine-formula/src/functions/information/function-map.ts)为23注册项，当时本地13，缺 CELL、ISBETWEEN、ISDATE、ISEMAIL、ISEVEN、ISFORMULA、ISODD、ISURL、SHEET、SHEETS，INFO/ISOMITTED为未注册枚举项。目录196→200、信息9→13、菜单出现项203→207、含 picker 子项204→208；常用七项、函数根9和各类新增根0保持。模型、公开命令/API、内容历史与 Worker/WASM 协议无增量；真实 IME/主题/焦点、两后端、文件和性能继续未验收。

<a id="information-parity-sheet-continuation"></a>

## 2026-11-27 奇偶、公式身份与工作表信息续接

固定 [ISEVEN](../../../univer/packages/engine-formula/src/functions/information/iseven/iseven.ts) / [ISODD](../../../univer/packages/engine-formula/src/functions/information/isodd/isodd.ts) 严格一个参数：先解单格形状和原错误，Boolean 返回 #VALUE!，其余值按原始 Number 转换并 trunc；空白与 Null 为0，金额/百分比文本不按格式化数字转换。仅拒绝 NaN，标量 Infinity 对 ISEVEN 为 false、ISODD 为 true，保留固定行为。[ISFORMULA](../../../univer/packages/engine-formula/src/functions/information/isformula/index.ts) 需要引用身份；本地读取真实 cell.formula，不求值，因此有效错误公式和自引用均为 true，文本/数值/Boolean 返回 #N/A，原错误继续传播。多格数组输出和共享公式 si 身份仍缺，不能伪造 metadata 作为已实现依据。

[SHEET](../../../univer/packages/engine-formula/src/functions/information/sheet/index.ts) 接受0–1参数；当前表或引用返回包含隐藏表的1-based位置，标量用 String 转换后忽略大小写匹配名称，缺失名称为 #N/A。[SHEETS](../../../univer/packages/engine-formula/src/functions/information/sheets/index.ts) 严格零参数，返回全部表数量。两项复用原参考对象和工作簿顺序；构造器补零参数接线，SHEET 的 A1 示例仍打开参数表单，SHEETS 直接进入原公式编辑。改名契约已回读 [rewriteWorksheetFormulaReferences](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-references.ts) 及 [formatArea](../../packages/office-studio/src/features/formula/formula-reference-areas.ts)，草稿使用实际保存的 `=SHEET('Renamed'!A1)`，字符串参数 "Other" 保持原文。

已读取五项固定测试：[ISEVEN](../../../univer/packages/engine-formula/src/functions/information/iseven/__tests__/iseven.spec.ts)、[ISODD](../../../univer/packages/engine-formula/src/functions/information/isodd/__tests__/isodd.spec.ts)、[ISFORMULA](../../../univer/packages/engine-formula/src/functions/information/isformula/__tests__/index.spec.ts)、[SHEET](../../../univer/packages/engine-formula/src/functions/information/sheet/__tests__/index.spec.ts)、[SHEETS](../../../univer/packages/engine-formula/src/functions/information/sheets/__tests__/index.spec.ts)。原公式草稿新增3声明，Facade 新增2声明，覆盖类型/错误、源公式变化、重排/隐藏/改名/插表、manual 缓存、历史和内部保存重开；原 Ribbon / Ant 表单循环补名称、SHEETS 提交/撤销。均未运行。

<a id="information-email-url-continuation"></a>

## 2026-11-27 邮件与网址判断续接

固定 [ISEMAIL](../../../univer/packages/engine-formula/src/functions/information/isemail/index.ts) / [ISURL](../../../univer/packages/engine-formula/src/functions/information/isurl/index.ts) 严格一个参数：矩形先返回 #VALUE!，单格保留原错误，非文本为 false。邮件不 trim，最大254字符；网址 trim 后最大1000字符。本地沿同一引用矩阵读取和已有错误协议，使用固定 [topLevelDomainSet](../../../univer/packages/core/src/common/url.ts) 的全部269项；仅以 AST 拼接静态字符串，已核对域名值与顺序、邮件及网址两条正则均和固定来源一致，没有执行正则或求值器。

网址协议分支支持 http/https/ftp/sftp/ftps/nfs/ssh，不要求域名属于269项；无协议域名和 mailto 使用该清单，另保留 news/aim、端口、编码路径、空白和长度规则。不能把 core 的 isLegalUrl / isSafeUrl 当作这两个公式的契约。已读取 [邮件测试](../../../univer/packages/engine-formula/src/functions/information/isemail/__tests__/index.spec.ts) 和 [网址测试](../../../univer/packages/engine-formula/src/functions/information/isurl/__tests__/index.spec.ts)；示例使用 [信息描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/information.ts) 的 developer@univer.ai 与 univer.ai。

原公式草稿新增2声明、Facade 新增1声明，保留边界/类型/错误、源文本修改、manual 缓存、撤销重做及内部重开；原 Ribbon 和 Ant 参数表单补两个名称。七个信息名称合计目录200→207、信息13→20、菜单出现项207→214、含 picker 子项208→215，全部运行0。UI、模型、命令、内容历史与 Worker/WASM 协议仍沿原链路，完整 Office、真实浏览器与性能保持未验收。

<a id="information-interval-continuation"></a>

## 2026-11-27 ISBETWEEN 区间判断续接

[固定实现](../../../univer/packages/engine-formula/src/functions/information/isbetween/index.ts) 接受3–5参数，两个包含端点标志缺省 true；按 [checkVariantsErrorIsArray](../../../univer/packages/engine-formula/src/engine/utils/check-variant-error.ts) 的参数顺序先处理矩形/原错误，再拒绝文本标志，最后检查下限大于上限的 #NUM!。标志数值按真值判断，显式 Null 为不包含端点。本地在原引擎内增加分支，复用引用矩阵，不新建求值、服务或历史协议。

比较依据为 [primitive-object](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)、[math-kit](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts) 与 [compare](../../../univer/packages/engine-formula/src/engine/utils/compare.ts)：数值＜文本＜Boolean，数值文本不转换；Null 在左侧按右侧类型变为空文本/false/0，Null 在右侧按原 getValue 的0处理，并保留 Boolean 对 Null 的数值比较。普通数值使用固定12/15位和1e-10误差规则；文字保留 NFC、大小写和连字折叠及 localeCompare。右侧含 * / ? 时使用固定通配路径，包括 ~ 转义、dotAll、保持 UTF-16 长度的大小写折叠，以及 >/>= 与 </<= 的不同命中规则。原通用 compare 的数字文本排序不符合这些规则，本次没有改变它的其他调用方。

已读取[固定测试](../../../univer/packages/engine-formula/src/functions/information/isbetween/__tests__/index.spec.ts) 与[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/information.ts)，示例为7.9,1.2,12.45。原[公式草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增2声明，[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)新增1声明，覆盖边界/类型/空值方向/精度/通配/错误顺序、范围与标志修改、typed text、manual 缓存、历史及内部保存重开。原 [Ribbon](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 和 [Ant 参数表单](../../tests/browser/sheet-toolbar.spec.ts) 增加该名称；两个信息菜单名称断言均已与实际21项目录做 AST 对比，均未运行。

目录207→208、信息20→21、函数菜单出现项214→215、含 picker 子项215→216；常用七项、函数根9和各类新增根0保持。该批固定信息 map 23项仍缺 CELL / ISDATE，INFO / ISOMITTED 继续单列未注册枚举项。RuntimeValue 尚未保留数值格式身份，因此 ISBETWEEN 的双日期格式比较（固定源码舍入至1e-8）、ISDATE 及派生数组等当时仍待补协议，不按日期函数名称假推断类型；ISFORMULA 数组/si、ERROR.TYPE 数组、TYPE 对象和完整函数/Pro 继续未完成。该批公式82声明/82计划、calculation27/27、sheet-toolbar14/14，运行0；小型路径保持同步 TS，大样本 Worker/Rust/WASM、Office、两后端、系统 IME、主题及性能未验收。ISDATE 后续标量/引用进度见[日期判断续接](#information-isdate-continuation)。

<a id="information-predicate-arity-text-reference"></a>

## 2026-11-27 既有信息判断的参数与文本引用续修

固定 [ISBLANK](../../../univer/packages/engine-formula/src/functions/information/isblank/index.ts)、[ISERR](../../../univer/packages/engine-formula/src/functions/information/iserr/index.ts)、[ISERROR](../../../univer/packages/engine-formula/src/functions/information/iserror/index.ts)、[ISLOGICAL](../../../univer/packages/engine-formula/src/functions/information/islogical/index.ts)、[ISNA](../../../univer/packages/engine-formula/src/functions/information/isna/index.ts)、[ISNONTEXT](../../../univer/packages/engine-formula/src/functions/information/isnontext/index.ts)、[ISNUMBER](../../../univer/packages/engine-formula/src/functions/information/isnumber/index.ts) 和 [ISTEXT](../../../univer/packages/engine-formula/src/functions/information/istext/index.ts) 均声明一个参数。[FunctionNode](../../../univer/packages/engine-formula/src/engine/ast-node/function-node.ts) 在调用前将参数数目错误判为 #N/A；本地八个分支原为 #VALUE!，现统一改为 #N/A。错误参数与多余参数同时存在时保留该参数数目优先级，原多格数组输出缺口继续单独记录。

八个判断复用已有 evaluateArrayArgument 和 formulaReferenceCellValue，读取单格、冒号单格、INDEX 及跨表引用。固定 [getCellValueObject](../../../univer/packages/engine-formula/src/engine/reference-object/base-reference-object.ts) 将 number 类型的 @ / @@@ 格式引用视为文本；原本地判断直接用 resolve 丢失该身份，现按原共享读取执行，含数字缓存的公式引用同样处理。源数值、公式、格式、空单元格及 ISBLANK 的持久内容判断不变，没有新增求值引擎或 DTO。

固定 ISNA 比较 getValue() 是否等于 #N/A，没有要求 ErrorValueObject 身份；本地现在对字面文本和强制字符串 #N/A 同样返回 true。这是固定 v1.6.0 的明确契约，不能用 Office 常见的仅错误身份判断替代。ISERROR 仍对这种文本返回 false，ISERR 返回 false，ISTEXT 返回 true；其他错误类型和原命令/历史/保存链路沿用。已读取 [ISNA 固定测试](../../../univer/packages/engine-formula/src/functions/information/isna/__tests__/index.spec.ts)，未运行。

原 [公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs) 修正两处文本 ISNA 与八个函数的旧参数断言，新增一条真实引用回归，覆盖 @ / @@@、缓存数值公式、INDEX、跨表、强制错误文本与源值不变。原 [calculation 契约草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs) 新增一条真实 Facade 闭环，覆盖只改数字格式的依赖重算、undo/redo、强制文本错误、manual 缓存、内部保存重开及显式重算。全部草稿未运行；唯一目录仍为208项/8类，信息21，根与菜单子项增量0。

该批无 UI、模型、公共 API、命令、内容历史或 Worker/WASM 协议变更。数组广播/spill、当时未实现的 CELL / ISDATE、ISFORMULA 的 si、日期格式身份、完整函数/Pro、真实输入与主题、Office 文件及大样本性能继续未完成或未验收；不能由参数和标量引用修正认定八个函数完整验收。

<a id="logical-percentof-arity-continuation"></a>

## 2026-11-27 PERCENTOF 与逻辑参数数目续核

固定 [PERCENTOF](../../../univer/packages/engine-formula/src/functions/logical/percentof/index.ts) 接受两个参数，先执行 SUM(dataSubset)，成功后执行 SUM(dataAll)，再相除；子集原错误优先于全集错误，零分母返回 #DIV/0!。本地复用既有 SUM 与二元除法，支持直接标量、稀疏引用和原数值文本/Boolean 转换；不增加第二套求和。该名称在[逻辑 map](../../../univer/packages/engine-formula/src/functions/logical/function-map.ts)注册但没有描述项，本地示例为 A1:A3,A1:A10，中文短标签为“子集占比”。固定 GROUPBY 将它作为 aggregator token 的 AST 路径尚未实现。

参数数目依据 [FunctionNode](../../../univer/packages/engine-formula/src/engine/ast-node/function-node.ts) 及 [AND](../../../univer/packages/engine-formula/src/functions/logical/and/index.ts)、[OR](../../../univer/packages/engine-formula/src/functions/logical/or/index.ts)、[XOR](../../../univer/packages/engine-formula/src/functions/logical/xor/index.ts)、[NOT](../../../univer/packages/engine-formula/src/functions/logical/not/index.ts)、[IF](../../../univer/packages/engine-formula/src/functions/logical/if/index.ts)、[IFERROR](../../../univer/packages/engine-formula/src/functions/logical/iferror/index.ts)、[IFNA](../../../univer/packages/engine-formula/src/functions/logical/ifna/index.ts) 和 [IFS](../../../univer/packages/engine-formula/src/functions/logical/ifs/index.ts)。AND / OR / XOR 为1–255参数，NOT为1，IF为2–3，IFERROR / IFNA为2，IFS为2–254且必须成对；无效参数数目在读取参数错误前返回 #N/A，IF 的未选分支继续保持惰性。TRUE / FALSE 及既有 NOW / TODAY 的零参数行为沿用前批修正。

原[公式契约草稿](../../packages/office-studio/tests/formula-engine.contract.test.mjs)新增 PERCENTOF 与逻辑参数各1声明；原[Facade 草稿](../../packages/office-studio/tests/workbook-calculation-mode.contract.test.mjs)各新增1声明，覆盖编辑重算、参数错误、255上界、manual缓存、一次撤销重做和内部保存重开。[Ribbon](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)及[Ant 参数表单](../../tests/browser/sheet-toolbar.spec.ts)增加 PERCENTOF 名称派发与真实参数场景，全部未运行。唯一目录208→209，逻辑11→12；根入口不增加，数组/高阶逻辑及完整公式仍未验收。

<a id="information-isdate-continuation"></a>

## 2026-11-27 ISDATE 日期判断续接

[固定 ISDATE](../../../univer/packages/engine-formula/src/functions/information/isdate/index.ts) 接受一个参数；多格矩形返回 #VALUE!，原错误保持。带 pattern 的数值按[日期格式判断](../../../univer/packages/engine-formula/src/basics/date.ts)识别，文本按固定日期系统解析日期或时间，其余类型返回 false。本地复用引用矩阵、[格式描述](../../packages/office-studio/src/features/sheet-workbench/number-format-render.ts)、已有 parseDate / parseTime，读取单格、冒号单格、跨表与 INDEX 的真实格式，并固定文本解析 locale 为 en，避免 UI 语言改变公式含义。

来源另含[固定测试](../../../univer/packages/engine-formula/src/functions/information/isdate/__tests__/index.spec.ts)、[信息 map](../../../univer/packages/engine-formula/src/functions/information/function-map.ts)及[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/information.ts)。原公式与 Facade 草稿各新增1声明，覆盖日期/时间文本、1900/1904边界、数值格式修改后的依赖重算、typed date text、manual缓存、历史和内部保存重开；原 Ribbon 信息目录断言和 Ant 参数循环补该名称，全部未运行。唯一目录209→210，信息21→22；固定23注册信息项尚缺 CELL。

本地仅对真实引用读取数字格式；嵌套表达式的数值 RuntimeValue 尚未携带 pattern，因此 ISDATE(IF(...)) 等格式身份传播仍缺。ISBETWEEN 双日期格式精度、ISFORMULA si、完整数组/spill、Office文件与大样本仍未完成或未验收，不能把该标量分支记为完整日期类型协议。

<a id="statistical-deviation-means-mode-continuation"></a>

## 2026-11-27 五个统计聚合续接

固定来源为 [AVEDEV](../../../univer/packages/engine-formula/src/functions/statistical/avedev/index.ts)、[DEVSQ](../../../univer/packages/engine-formula/src/functions/statistical/devsq/index.ts)、[GEOMEAN](../../../univer/packages/engine-formula/src/functions/statistical/geomean/index.ts)、[HARMEAN](../../../univer/packages/engine-formula/src/functions/statistical/harmean/index.ts)、[MODE.SNGL](../../../univer/packages/engine-formula/src/functions/statistical/mode-sngl/index.ts)、[统计 map](../../../univer/packages/engine-formula/src/functions/statistical/function-map.ts)及[描述目录](../../../univer/packages/engine-formula/src/services/formula/function-list/statistical.ts)。五者均为1–255参数，数量错误返回 #N/A；本地复用 evaluateArrayArgument、共享引用读取和行列顺序，不新建求值或历史服务。

| 函数 | 本地真实行为与固定差异点 |
| --- | --- |
| AVEDEV | 直接数字文本和Boolean参与，引用文字/Boolean/空值忽略；按每个引用参数分组保留求和、均值和偏差的错误阶段。无数值返回 #NUM!。 |
| DEVSQ | 直接及引用数字文本按原始 Number 参与，Boolean、空值和非数值文字忽略；返回偏差平方和，无数值返回 #NUM!。 |
| GEOMEAN / HARMEAN | 直接文字先做共享数值转换检查，再按原字符串 Number 参与；引用数字文本参与。非正数在原参数错误读取完成后返回 #NUM!；无数值分别返回 #NUM! / #N/A。 |
| MODE.SNGL | 直接文字先检查转换后忽略，引用文字/Boolean/空值忽略；并列频率取首次出现，引用按行列顺序。无重复数值返回 #N/A。 |

数值与错误阶段依据 [primitive-object](../../../univer/packages/engine-formula/src/engine/value-object/primitive-object.ts)、[base-value-object](../../../univer/packages/engine-formula/src/engine/value-object/base-value-object.ts)、[ArrayValueObject.sum](../../../univer/packages/engine-formula/src/engine/value-object/array-value-object.ts)、[math-kit](../../../univer/packages/engine-formula/src/engine/utils/math-kit.ts)、[isRealNum](../../../univer/packages/core/src/shared/generate.ts)及[getModeSnglResult](../../../univer/packages/engine-formula/src/basics/statistical.ts)。DEVSQ / GEOMEAN / HARMEAN 的最终结果沿固定 NumberValueObject.create 接收原始数值，不增加非有限结果守卫；引用非有限值沿共享入口返回 #NUM!。AVEDEV 的标量与引用分组错误不能统一改为 #NUM!：两个1e308标量产生 #VALUE!，后加 #REF! 则保留 #REF!；两个1e308同组引用后加 #REF! 返回 #VALUE!，分成两个单格引用则返回 #NUM!。

已读取五份固定测试：[AVEDEV](../../../univer/packages/engine-formula/src/functions/statistical/avedev/__tests__/index.spec.ts)、[DEVSQ](../../../univer/packages/engine-formula/src/functions/statistical/devsq/__tests__/index.spec.ts)、[GEOMEAN](../../../univer/packages/engine-formula/src/functions/statistical/geomean/__tests__/index.spec.ts)、[HARMEAN](../../../univer/packages/engine-formula/src/functions/statistical/harmean/__tests__/index.spec.ts)、[MODE.SNGL](../../../univer/packages/engine-formula/src/functions/statistical/mode-sngl/__tests__/index.spec.ts)，未执行。原公式草稿新增3声明、Facade草稿新增2声明，覆盖转换/空值/错误/255边界、INDEX、引用行序、并列众数、非有限结果、格式与源格编辑、manual缓存、undo/redo和内部保存重开。原 Ribbon 名称派发及 Ant 参数表单循环补五项，全部未运行。

目录210→215，统计21→26；常用七项、函数根9保持，菜单函数出现项222、picker另1、全部子项223。只读 AST 解析主线7个代码/草稿文件无语法诊断；公式89、calculation33、Sheets Ribbon32、sheet-toolbar14均为声明数，内部数据循环不另计声明，运行0。模型/API/历史与 Worker/WASM 协议无公式批增量；统计其余83个注册名称、完整数组/spill、派生值格式、Office及性能继续未完成或未验收。

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

NA 批补 [既有公式契约](../../packages/office-studio/tests/formula-engine.contract.test.mjs) 中的构造、#N/A、IFNA、IFERROR 和错误传播草稿；全目录契约明确 NA 应返回 #N/A，其他示例仍要求正常求值。该批信息根最初只有 NA；当时 Ribbon 契约明确九个信息子项、Info 图标、分类顺序和函数派发，并继续覆盖权限禁用。浏览器草稿将 NA 加入既有零参数编辑/取消循环。当前22项见本轮续接记录，全部草稿仍未运行。

NA 批静态检查通过：3 个既有契约 `node --check`；catalog/schema 两个 TypeScript AST 解析 0 条语法诊断；本证据 38 个链接均存在；26 个固定 map/description 文件中的枚举引用无缺失；按已有 CRLF 规则执行限定空白检查通过。RIGHT 小修在同一公式契约追加零长度、默认长度、正长度、超出文本长度及负数参数的真实实现回归草稿，未运行。追加后该契约 `node --check`、引擎 TypeScript AST（0 条语法诊断）、38 个证据链接和限定空白检查均通过。

前批主线静态检查覆盖当时引擎与两份契约草稿的语法、引用及本证据链接/锚点；目录 AST 清点确认 126 项、9 类，无重复名称。当时 15 个修改文件的限定结果见[字体、story 与公式收尾](../UNIVER_TOOLBAR_PARITY.md#font-story-formula-continuation)，不覆盖本次单格/文本/UTF-16 追加。语法解析未执行求值器或 Facade，草稿不作为行为通过证据。

入口两批、RIGHT 零长度及本次文本取整/WEEKDAY 小修均未运行测试、类型检查、构建、Rust/WASM、dev、浏览器或性能基准。实际菜单 hover/选中/禁用、明暗主题、焦点恢复、系统 IME、宽窄屏和 Canvas 两后端仍未验收。完整目录、全部函数/定义名称交互、自动范围与批量插入、迭代计算、SHEET-10 Pro、真实 Office 文件与大样本性能继续保持未完成或未验收，不能由这些小修推断能力闭环。
