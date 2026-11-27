# SHEET-06 地区货币、会计与格式预览

2026-11-26，参考 Univer v1.6.0 固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。按 [顶部入口清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)与 [Sheets 源码清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)复核「开始 → 数字」。

| Classic 顺序 | Grid 位置 | Univer 命令 | 本项目入口 |
| --- | --- | --- | --- |
| 数字格式（order 9） | r1c1 跨 4 列，160px | `sheet.operation.open.numfmt.panel`，选择预设执行 `sheet.command.numfmt.set.numfmt` | `sheet.command.set-number-format`，更多格式打开 `sheet.operation.open-custom-number-format` |
| 百分比（order 9.1） | r2c1 | `sheet.command.numfmt.set.percent` | `sheet.command.set-percent-format` |
| 货币（order 9.2） | r2c2 | `sheet.command.numfmt.set.currency` | `sheet.command.set-currency-format`，接受上游命令别名 |
| 增加小数位（order 9.3） | r2c3 | `sheet.command.numfmt.add.decimal.command` | `sheet.command.increase-decimals` |
| 减少小数位（order 9.4） | r2c4 | `sheet.command.numfmt.subtract.decimal.command` | `sheet.command.decrease-decimals` |

固定来源为 `packages/sheets-numfmt-ui/src/menu/schema.ts`、`menu.ts` 与 `number-format-options.ts`。数字组 5 个常驻根，格式菜单 13 个预设、4 个分隔、1 个面板入口。Classic 和 Grid 复用入口，More 收纳同组。扩展、独立溢出、上下文和常驻根的本批增量均为 0。固定版仅在 Sheets 显示，禁用条件读取 workbook 可编辑、worksheet 可编辑及设置样式权限和范围编辑权限。本项目保留 Sheets 工作台、工作表/范围保护和编辑会话状态判断。

## 区域、草稿与持久化

- 固定版 `packages/core/src/services/locale/currency-symbol.ts` 列 19 个区域及固定符号。`packages/sheets-numfmt/src/commands/commands/set-currency.command.ts` 在执行时读 RegionService，按所有当前选区提交两位小数、千分位和红色括号负数格式。`packages/sheets-numfmt-ui/src/menu/number-format-options.ts` 同样用区域符号生成货币、会计和整数货币三个菜单格式。本项目沿用这些模板。精确格式码保存在 `CellStyle.numberFormat`，切换区域后更新快捷按钮的图标和预览的分隔符，现有单元格保留保存的符号与数值。工作台使用实际区域符号作为图标，例如 NT$ 或 Rp。固定版使用各地区分组的货币图标。
- `packages/sheets-numfmt-ui/src/views/components/{Currency,Accounting}.tsx` 从当前模式和使用习惯初始化，再用 `packages/sheets-numfmt/src/base/const/formatdetail.ts` 的六种负数模式或会计填充模式。固定版高级面板缺省取最近使用或第一个通用符号。本项目设计为当前格式的符号优先，其次最近使用，二者缺省时用当前区域。面板仍保留原有 27 个符号，并加入 6 个地区符号（NT$、HK$、R$、Rp、zł、﷼），共 33 个去重选项。选择来源里已有的 OOXML 方括号符号也会临时列入。选项及已确认偏好共用同一目录。
- `sheet-currency.ts` 从正值段的符号文本或 OOXML `[$符号-地区]` 识别货币，忽略填充、留位和其他分段里的同形字符。多字符符号优先识别完整值。本地 `sheet-number-format.ts` 的会计分类与高级面板共用，覆盖固定版以 `_(` 开头的会计格式，以及本地已有的菜单会计格式和带 `*` 填充的会计格式。只改小数位时保留原始分段、负值颜色、符号位置和地区标记。显式选择货币或负数样式仍切换到该选择对应的模板。
- `CustomNumberFormatPanel.vue` 负数选项调用与网格同一格式渲染器，按当前区域、草稿小数位显示文本和模式指定颜色。面板沿用 [本地组件文档](../llm.md)中的 antdv-next 1.8.2 Form/FormItem、Select、InputNumber、Menu 和反馈控件，以及现有 ConfigProvider 主题与焦点恢复。预览仅更新 Canvas 的格式覆写，取消丢弃覆盖，确认后才更新内容及偏好。
- 工作台既有 `runNumberFormatCommand` 和 `FRange.setNumberFormat` / `setNumberFormats` 共用 Studio 数字格式命令及计划。单次格式提交只写选区的显示格式并进一次工作簿历史，数值、公式文本及已知缓存类型继续由原模型保留。重复应用现有格式不新增历史。内部快照与 `crates/office-writer/src/xlsx.rs` 的 numFmt 注册读取同一格式码。面板草稿创建参数现分别接收界面语言、区域和最近使用符号，公开 Facade 签名与 WASM 协议不变。

## 验证状态

已静态核对固定源码、地区目录及现有命令、快照和写出路径。给 `sheet-number-format.contract.test.mjs` 和 `sheet-number-format.browser.test.mjs` 各补两条草稿，涵盖 19 个地区模板、符号识别、会计分段保持、快捷按钮激活、负数地区预览、Facade、撤销重做及内部快照重开。

测试、类型检查、构建、浏览器及 Rust/WASM 留待统一阶段。真实 XLSX 导出与独立 Office 重开、全部地区实际预览、主题和焦点的真实复核，以及大范围格式应用性能仍为未验收。

<a id="number-group-review-2026-11-27"></a>

## 2026-11-27 数字组续核

本次按开始页从左到右继续追踪同一组 5 个常驻根、13 个预设和 1 个更多格式入口，入口数量没有增加。领域分类、科学数值格式限制，以及主线的工作台主格状态、预设应用范围和浏览器草稿均已落盘，运行验证后置。

### 固定来源与可观察契约

- [数字组 schema](../../../univer/packages/sheets-numfmt-ui/src/menu/schema.ts)与 [菜单状态](../../../univer/packages/sheets-numfmt-ui/src/menu/menu.ts)确认格式选择器跨第 1 行四列，第二行依次为百分比、区域货币、增加小数位、减少小数位。显隐由 Sheets 类型判断，禁用由 workbook/worksheet/范围编辑和 worksheet 设置样式权限判断；本项目继续通过既有工作台能力和保护模型判断。
- [numfmt-menu-util.ts](../../../univer/packages/sheets-numfmt-ui/src/menu/numfmt-menu-util.ts)读取末段 selection 的 `primary ?? range` 起点格式；selector 标签、百分比和货币 activated 都从该模式派生。货币 activated 要求 `getPatternType(pattern) === 'currency'` 且正值段符号等于当前区域符号。固定版的会计、整数货币预设仍属于引擎的 currency 类，本项目保留更细的菜单分类。
- [MoreNumfmtType.tsx](../../../univer/packages/sheets-numfmt-ui/src/views/components/MoreNumfmtType.tsx)的预设仅遍历末段 selection.range。[百分比命令](../../../univer/packages/sheets-numfmt/src/commands/commands/set-percent.command.ts)、[货币命令](../../../univer/packages/sheets-numfmt/src/commands/commands/set-currency.command.ts)、[增加小数位](../../../univer/packages/sheets-numfmt/src/commands/commands/add-decimal.command.ts)和 [减少小数位](../../../univer/packages/sheets-numfmt/src/commands/commands/subtract-decimal.command.ts)遍历全部当前 selections，均未过滤筛选隐藏行；百分比直接写 `0%`，菜单预设写 `0.00%`。显式 Facade/ranges/values 继续使用调用方给出的目标。
- [pattern.ts](../../../univer/packages/sheets-numfmt/src/utils/pattern.ts)的 `getPatternType` 复用 [format-info.ts](../../../univer/packages/core/src/shared/numfmt/format-info.ts)，General 带条件 token 时归为 number。标准数值元数据以正值段取数字类型和精度，`isDate` / `isPercent` 读取全部分段；日期与文本排除货币判断，货币优先于科学计数和百分比，正值段的科学计数优先于百分比。标签不能仅凭首段字符正则推断。
- [SetNumfmtCommand](../../../univer/packages/sheets-numfmt/src/commands/commands/set-numfmt.command.ts)与 [scientific-notation.ts](../../../univer/packages/sheets-numfmt/src/utils/scientific-notation.ts)限制科学数值可用格式；[core numfmt utils](../../../univer/packages/core/src/shared/numfmt/utils.ts)中的 `isDefaultFormat` / `isTextFormat` 精确接受缺省/`General`、`@`/`@@@`，不能用同类标签代替这些格式码。格式 mutation 与 undo/redo 一起提交，格式删除也进入历史。

### 本地修正、保持行为与差异

- [sheet-number-format.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-number-format.ts)复用 [number-format-render.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-render.ts)现有的格式解析缓存与 `describeSheetNumberFormat`。分类使用真实引擎元数据，再保留本地完整地区符号、会计、财务及持续时间分类，不建立第二套格式引擎。带单位的普通数字、多段百分比/日期、纯文本及条件 General 使用同一分类结果。
- 原先 `"$"0.00E+00` 被首段正则判成 scientific，科学数值因此被错误放行；现在判为 currency。[sheet-number-format-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-number-format-command.ts)分别检查精确 `General`、`@`/`@@@` 与 scientific 类型。`"item "@`、`[Red]General` 可以分别显示文本、常规类标签，但不能因此绕过科学数值的格式限制；`0.00E+00%` 按固定引擎优先级仍是 scientific。限制只跳过目标单元格的格式 patch，保留其他可应用单元格、原始数值、文本类型和公式缓存。
- 小数位仍由现有 [number-format-pattern.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-pattern.ts)和同一个 model/view 计划执行。[上游 decimal.ts](../../../univer/packages/sheets-numfmt/src/utils/decimal.ts)确认 0–30 位、补位复制末个可选占位符、忽略引号/转义/方括号/留位/填充操作数、保留日期时间段。本地继续保留负值颜色、段顺序、字面量和地区符号；General 科学值先按尾数精度生成至少两位的科学格式，再统一增减，普通数字保持普通格式。达到上限或仅选日期时间时不产生 patch。
- 已记录的算法差异：固定源码对无格式普通数值仅在 `!maxDecimals` 时读取原值的小数位，结果会依赖遍历顺序；显式存储的 `General` 又会由引擎报告 9 位。本地继续从所有有效 General 数值取最大原值/尾数精度，显式 General 与缺省采用同一计划。无小数点的格式也只在整数部分补位，避免修改分母或指数。本批没有复制这些上游内部差异，尚无 Microsoft Office 多选区样本实测证明最终效果一致。
- [OfficeStudio 数字命令](../../packages/office-studio/src/api/office-studio.ts)仍通过 `commitWorkbook(..., { recalculate: false })`提交；[Facade](../../packages/office-studio/src/api/facade.ts)的 `setNumberFormat` / `setNumberFormats` 复用同一注册和计划。全部目标被跳过或格式已相同不会新增内容历史。格式码继续存入 `CellStyle.numberFormat`，内部快照保存原始模型；[XLSX writer](../../crates/office-writer/src/xlsx.rs)以完整格式码注册 numFmt 并写出转义后的 formatCode，未修改协议或写出路径。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的选择器标签、百分比和货币激活态读取活动主格的实际格式；13 个预设只应用于活动选区段，百分比、货币和小数位快捷命令继续应用全部选区。更多格式面板保留首个选区起点作为草稿基准，确认仍应用全部目标范围；显式 API/ranges/values 不改目标。
- [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的地区货币图标使用 24×24 SVG viewBox；文字中心为 12/12，单字符字号 20，多字符及 `﷼` 字号 16、文字宽度限制 20。外框按 classic/grid 的实际图标尺寸缩放，避免 NT$/HK$/﷼ 的字形超过悬停按钮。浏览器草稿测量 SVG 文字与按钮矩形，此处不声称实际字体边界已通过。

### 静态核对与未验收项

[sheet-number-format.contract.test.mjs](../../packages/office-studio/tests/sheet-number-format.contract.test.mjs)增加 3 个合同草稿：16 个真实格式的分类与菜单 key；domain/view 对科学数值、带类型文本和普通数值采用相同限制；上游小数位别名在混合普通/科学/可选占位符/字面量/日期/百分比/文本选区中复用计划，重叠范围只执行一次，日期和 30 位上限返回无操作。测试直接调用既有模型、视图适配器和命令，作者格式通过现有 setter 构造。

[sheet-number-format.browser.test.mjs](../../packages/office-studio/tests/sheet-number-format.browser.test.mjs) 补活动主格、活动段预设与全选区快捷命令、首范围面板基准的真实 Studio 草稿；另补 classic/grid 中 en-US、zh-TW、zh-HK、fa-IR 的字形中心、按钮内边界、悬停命中、明暗主题和无内容/历史变更的草稿。全部未运行。

只执行本批文件静态 diff、链接和固定源码路径检查。合同、浏览器、类型、构建及 Rust/WASM 均未运行；主格与预设范围的浏览器闭环、历史/保存重开、真实 XLSX 与 Office 样本、主题/焦点/IME、格式大范围性能与内存继续保留未验收。

### 数字组命令与边界复查

再次读取上述四条快捷命令、SetNumfmtCommand、菜单状态及本地 Studio/model/view 调用，没有新增入口或生产修改。快捷货币按执行时区域生成模板，百分比固定写 `0%`，小数位继续读取全部目标；本地显式 `ranges / values` 由调用方确定，重叠范围在计划入口去重。工作表及范围保护在应用计划前拒绝，Studio 提交继续使用 `recalculate:false` 和既有文档历史；相同格式或所有目标被跳过时不发布内容变化。上述结论是静态调用链核对，未新增运行验收证据。

[Currency.tsx](../../../univer/packages/sheets-numfmt-ui/src/views/components/Currency.tsx) 和 [Accounting.tsx](../../../univer/packages/sheets-numfmt-ui/src/views/components/Accounting.tsx) 的 InputNumber 明确为 0–20 位；工具栏 [decimal.ts](../../../univer/packages/sheets-numfmt/src/utils/decimal.ts) 的算法上限则为 30 位。[CustomNumberFormatPanel.vue](../../packages/office-studio/src/components/spreadsheet/CustomNumberFormatPanel.vue) 与 [number-format-pattern.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-pattern.ts) 采用对应两级限制，不能把面板 20 位上限认定为与固定版本不一致。表单及控件仍使用 [本地文档](../llm.md) 所述 Form、Select、InputNumber、Menu 的受控值、禁用和变更事件。

新增记录的比较差异：[core isPatternEqualWithoutDecimal](../../../univer/packages/core/src/shared/numfmt/utils.ts) 按引擎 token 忽略小数、COLOR 和 SKIP，并对颜色后的 MINUS 特殊处理；本地 `numberFormatWithoutDecimals` 保留颜色、留位、字面量和段序后转小写。因此带附加颜色或留位的作者格式可能得到不同的预设标签/负数选项，不能仅凭函数注释宣称比较语义完全相同，也不能以内部 token 规则代替 Office 可观察行为验证。现有完整格式码保存和小数位改写行为继续保留，该比较差异仍待浏览器与文件样本验收。

<a id="number-format-input-session-review"></a>

## 2026-11-27 数字组输入会话续修

本批仍为数字格式、百分比、货币、增加小数位、减少小数位五个根及原格式菜单，保持上表 classic order 9–9.4 和 grid 位置。常驻、插件扩展、上下文、独立溢出根和菜单子项增量均为 0，未增加按钮或格式种类。

固定 [FactoryOtherMenuItem:151](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-numfmt-ui/src/menu/menu.ts#L151)、[PercentMenuItem:116](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-numfmt-ui/src/menu/menu.ts#L116)、[CurrencySymbolIconMenuItem:64](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-numfmt-ui/src/menu/menu.ts#L64)、[AddDecimalMenuItem:80](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-numfmt-ui/src/menu/menu.ts#L80) 和 [SubtractDecimalMenuItem:96](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-numfmt-ui/src/menu/menu.ts#L96) 全部复用 `getCurrentRangeDisable$`。[menu-util:163/191–192](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L163) 缺省 supportCellEdit=false，禁用可见单元格编辑器及 FX 焦点，[225–226](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/menu-util.ts#L225) 继续处理绘图。先前证据只列权限条件，没有完整记录普通输入禁用，现补齐真实来源。

原本地 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 的 numberFormatDisabled 复用 fontEditable，普通单元格/公式输入时仍可用；[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 executeRibbon 先 finishEdit，会先提交草稿再改格式或打开更多格式。现在复用已有 canQuickFormatSelection，将活动工作台、完整选区权限、editingKey、组字及绘图状态传为可选内部 numberFormatEditable；五根共同禁用，省略可选状态的原调用方保留现有行为。executeRibbon 在 finishEdit 前检查既有 SHEET_NUMBER_FORMAT_COMMANDS 的 6 个定义 ID（5 个 Ribbon 命令及 setCells）和两个已有开面板 operation，能力不可用时直接返回，只判断 ID，不重复调用 numberFormatActionForCommand 做参数解析。

[RibbonItem.vue](../../packages/office-studio/src/platform/ribbon/RibbonItem.vue) 已将 disabled 传给专有组件，且 forward / execute 使用相同 schema 状态；[SheetNumberFormatControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetNumberFormatControl.vue) 已将 disabled 传入 `antdv-next` Button/Menu，选择/键盘入口拒绝禁用状态并在能力变化时关闭 Popover。本批复用这些现有行为，没有改组件、原生桥或弹层协议。普通字体草稿格式仍依原能力；更多格式面板继续按原 watcher 在开始单元格输入、权限/文档变化时关闭。

输入取消后的预设仍应用活动段，快捷命令和更多格式确认仍应用全部选区；同一 model/view 计划保留精确格式码、地区符号、数值/公式缓存，Studio 以 recalculate:false 进入既有文档历史和内部快照。公开 API、格式解析/分类、小数位算法、事件字段、Office writer 和 Worker/WASM 协议均无新增，只有可选内部 UI 状态。程序化 Studio/Facade 及其上游别名保持现有域命令路径；本批没有给原 UI dispatcher 新增这些别名处理器，也不把 setCells 的提交前拦截登记为新的可执行 UI 功能。

本批五文件为上述两份生产文件、[原 Ribbon 合同草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs)、[原数字格式浏览器草稿](../../packages/office-studio/tests/sheet-number-format.browser.test.mjs) 和本页。生产增量 Excel `+3/-0`、Ribbon `+2/-1`，合计 `+5/-1`。合同新增 1 个声明，检查五根禁用、可选/省略状态、缺表/字体权限、原顺序/位置、预设与更多格式/四快捷命令 invocation，并确认普通字体能力保持原状态。浏览器新增 1 个声明及 fixture 的版本观察，含 cell/formula × classic/grid 四个真实 Studio 循环样本，覆盖原生 disabled 点击、8 个定义 ID/operation 的迟到真实 Vue execute emitter、普通输入、合成 composition 和明转暗的草稿/模型/版本/历史/命令记录不变；composition 结束仍输入时保持禁用，Esc 后真实菜单应用、一次历史、undo/redo/内部 JSON 重开以及 Canvas 宿主保持。没有新增产品测试接口，合成事件不等于系统 IME 验收。

当前静态声明为 Ribbon 合同 27（计划 27）、数字格式浏览器 31（5 个 test.each 数组与 1 个外层 for 各注册两个变体，按源码数组预计 37）；本批循环样本不另计声明，全部运行 0。本批仅做源码/diff 阅读和 rg 计数，主线统一静态语法/引用收尾；未运行测试、类型、构建、dev、浏览器、Rust/WASM、Office 或基准。可替换控件仍为原 `antdv-next` Button/Menu/Popover/Input/TextArea 和 Form 面板，沿统一 ConfigProvider/主题；实际明暗主题、焦点、真实 IME、standalone 浏览器、WebGPU/Canvas2D、Office 独立重开和大范围性能/内存保持 **未验收**，SHEET-06 整体状态不变。
