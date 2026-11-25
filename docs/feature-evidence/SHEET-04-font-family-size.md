# SHEET-04 字体、字号与字号步进（待统一验证）

记录日期：2026-11-23；2026-11-25 续作见[富文本字体与 Canvas 编辑](SHEET-04-05-rich-text-font-editing.md)和[默认样式继承](SHEET-04-default-style-inheritance.md)。固定参考根 `E:\rust_test\univer`，Univer v1.6.0、提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。这份记录覆盖字体组第一行的四个入口，与 [第二行字体开关和颜色](SHEET-04-color-reset.md)、[边框](SHEET-04-borders.md)互补。仅完成源码复核、实现与用例编写，本批未执行测试、类型检查、构建、Rust/WASM 或浏览器；SHEET-04 保持“未验收”。

## 来源与入口

| 标签 | 字体组 classic 次序 | grid 位置 | 本项目命令 | 固定版入口及来源 |
| --- | --- | --- | --- | --- |
| 字体 | 1 | 1/1，跨 4 列，宽 180、最小 120 | `sheet.command.set-font-family` | `SetRangeFontFamilyCommand` / `sheet.command.set-range-font-family`；`sheets-ui/src/menu/schema.ts`、`menu/menu.ts` 的 `FontFamilySelectorMenuItemFactory`。 |
| 字号 | 2 | 1/5，宽 72 | `sheet.command.set-font-size` | `SetRangeFontSizeCommand` / `sheet.command.set-range-fontsize`；`schema.ts`、`menu/font.menu.ts` 的 `FontSizeSelectorMenuItemFactory`。 |
| 增大字号 | 3 | 1/6 | `sheet.command.increase-font-size` | `SetRangeFontIncreaseCommand` / `sheet.command.set-range-font-increase`；`schema.ts`、`font.menu.ts`。 |
| 减小字号 | 4 | 1/7 | `sheet.command.decrease-font-size` | `SetRangeFontDecreaseCommand` / `sheet.command.set-range-font-decrease`；`schema.ts`、`font.menu.ts`。 |

表中的短路径均相对固定参考根的 `packages`。四项属于 START / FORMAT 下的常驻根声明；overflow 是同一入口的替代呈现，本批不增加插件或上下文根入口。字体组共有 11 个现有根入口：本行 4 项、第二行 7 项。下拉字体与字号值不计作额外根入口，固定别名也不计为新功能。

固定版 factory 在 Sheets 外隐藏，禁用检查涉及工作簿编辑、工作表编辑/设置样式、范围 EditPoint，并允许行内编辑上下文。`packages/sheets-ui/src/commands/commands/inline-format.command.ts` 的四个 Range 命令根据编辑器焦点分派到整格或行内文本命令。`packages/sheets/src/commands/commands/set-style.command.ts` 的 `SetFontFamilyCommand` / `SetFontSizeCommand` 将 `ff` / `fs` 交给 `SetStyleCommand`，后者走范围 mutation、筛除行迭代和按 unitID 隔离的撤销重做。

`packages/sheets/src/facade/f-range.ts` 的 `setFontFamily(string | null)` / `setFontSize(number | null)` 直接调用 SetStyleCommand；公开注释明确 null 重置。固定 UI 字号输入 min=6、max=400，而这两个 Facade setter 没有 UI 数值收窄。`packages/core/src/types/const/const.ts` 的固定默认值为 Arial / 11。

## 本项目实现

`sheet-font-format-command.ts` 在已有 mark/color 动作上追加 family、size、size-step，并注册本地四个命令及上述四个固定别名。对象输入支持 `unitId`、`sheetId/subUnitId`、两种行列坐标的 `range/ranges`、可选 `primary: { row, col }`；原始字符串、数字、布尔值和 null 经同一输入归一化，已有开关的 raw boolean 合同保留。显式 null 范围、无效坐标、空字体和非法字号在修改之前拒绝。

字体接受非空字符串或 null，保留原始字体名称；字号接受有限正数或 null，保留既有 API 的非 UI 范围，例如 5、500 和小数。UI 仍通过 `RibbonEditableSelect.vue` 的 antdv-next AutoComplete 与 `confirmRibbonDraft` 限定字号输入 6–400，字体草稿选择现有配置中的字体；API 写入的自定义字体加入当前显示值。输入、Enter/Escape、失焦、外部属性同步和组合输入事件沿用统一可编辑选择器。

字号步进先读取活动单元格，再将一个确定字号应用于选区并集，重叠格只处理一次。UI 传入 `selectedKey` 对应主格；明确 API 范围未传 primary 时，使用最后范围的起点。显式 primary 必须是有效整数坐标并位于选区内；null、非对象、缺失坐标或选区外坐标拒绝整个操作。与固定命令相同，增大仅检查上界、减小仅检查下界：`500 → 499` 是减小字号，`500 → 400` 是增大字号；`399.5 → 400.5` 的增大行为也保留单侧检查，不能擅自改成双侧 clamp。

family/size 为 null 时仅扫描已有单元格，跳过筛除行，不新增稀疏空白。2026-11-25 已改为清直接/run 作者属性、屏蔽相应引用字段并恢复有效继承；没有上层来源才回退 Aptos / 11。非重置操作保留无关字体属性、样式、值、公式、缓存、metadata，并同步整格对应 rich-run 覆盖。model 和 standalone view 路径使用相同动作语义；Studio 合并工作表与本工作表表格的筛除行，保留手动隐藏行，提交时关闭公式重算。

`office-studio.ts` 通过同一注册项验证权限、提交模型和内容历史；`FRange.setFontFamily/setFontSize` 改为派发该命令。`ExcelWorkbench.vue` 的 dispatcher 使用统一字体注册表，原字体/字号局部 handler 已移除；稀疏 null 重置绕过直接枚举选区门槛。`src/index.ts` 导出 font action、command input/params、primary、range 和 mark 类型。成功但未改内容的命令返回 false，仍有通用 CommandExecuted 事件；不会添加内容历史。取消、非法和权限失败没有成功事件或部分模型修改。

工作台字体控件由当前选区反推；混合字号为空，字号步进仍依据活动格。内容历史按工作簿隔离，API 明确指定的非活动工作簿/工作表不因当前 UI 选择而改目标。字体/字号模型通过现有 `excelViewAdapter.ts` 与 XLSX writer 字体 registry 保存；保存快照后的重开及 Office 文件读写须分别验证。

## 用例与验证状态

新增 [sheet-font-properties.contract.test.mjs](../../packages/office-studio/tests/sheet-font-properties.contract.test.mjs) 25 组：固定别名、raw/object/null/false、两种坐标和显式非法范围、字体/字号参数、单侧步进边界、model/view 的值/缓存/metadata/richText 保留、百万行稀疏重置、筛除/手动隐藏行、default/no-op、活动格/多选/重叠、无效 primary 原子拒绝、保护、Facade 与注册事件、取消/非法输入、历史/快照重开、工作表/表格筛除合并和工作簿隔离。

新增 [sheet-font-properties.browser.test.mjs](../../packages/office-studio/tests/sheet-font-properties.browser.test.mjs) 13 组：AutoComplete 字体/小数字号提交、Canvas 保留、撤销重做/快照、Escape/失焦/非法草稿、混合选区步进、UI/API 数值范围、外部字体同步、全表 sparse null、取消和控件禁用、工作簿切换、筛除/隐藏行、活动输入草稿、grid overflow、合成 IME 键及 portal 明暗/RTL。仅已编写，未运行；合成 IME 事件不替代真实输入法验收。

本轮静态复核限定在本批文件及已协调的共享接线区段。统一验证时需要执行这些新用例、既有字体/颜色/边框回归、类型与构建、相关 Rust/WASM 和真实文件链路；编写用例及静态审阅不作为通过证据。

## 仍未验收的差异

- 2026-11-25 已补 workbook/worksheet/行列/theme/direct 组合、目录引用/reset、空格输入/复制/格式刷/填充和来源持久基础，见[默认样式证据](SHEET-04-default-style-inheritance.md)。null 已参与继承；本地底层默认仍 Aptos/11，与固定 Arial/11 不同。完整范围主题、字体/CJK/复杂脚本及原生文件映射仍未完成或未验收。
- 2026-11-25 已补同一字体命令的行内范围/typing 草稿、整格运行覆盖、Facade 部分样式与 Canvas 编辑接线，见[续作证据](SHEET-04-05-rich-text-font-editing.md)。复杂 shaping、真实 IME、富剪贴板、全部继承场景、浏览器和 Office 往返仍未验收，不能仅以运行保留计完成。
- 独立工作簿只读、设置样式权限与绘图内文字上下文仍需整体能力验收；当前工作表/范围编辑保护及 UI 禁用不能推导完整权限等价。
- 字体文件加载、字体回退/度量、真实中文输入法、每个语言与主题下的焦点、XLSX 导出/独立 Excel 重开、源主题字体及默认样式保真没有本批实测结果。
- 非重置大范围设置仍需 Worker/批量模型与取消路径。本批没有同数据集前后耗时、帧时间、主线程长任务、传输或内存基准，不声明性能优化或大样本达标。
