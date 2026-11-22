# SHEET-04 字体组第二行与单元格颜色（待统一验证）

2026-11-22，固定源码 `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`。`packages/sheets-ui/src/menu/schema.ts` 将两个重置入口分别放在文字色和背景色菜单下；`packages/sheets/src/commands/commands/set-style.command.ts` 的 `ResetTextColorCommand` / `ResetBackgroundColorCommand` 只将 `cl.rgb` 或 `bg.rgb` 设为 null。

项目使用相同 command ID。工作台与 `FRange.setFontColor(null)`、`FRange.setBackground(null)` 在带 Studio 的情况下共用按工作簿隔离的命令、权限及历史。只修改选区内已有单元格，不因重置而新增空单元格。`CellStyle.fill` 和 `FontStyle.color` 可为 null；视图的 undefined 表示未覆盖，null 表示清除。显示层仍提供回退颜色；Facade 的 getBackground 保持返回显示色。旧快照中已有的白色保留为显式白色，不猜测它原先是用户指定还是旧默认值。

`crates/office-writer/src/xlsx.rs` 的单元格样式将无填充映射到 none，显式白色保留为 solid；未指定字体色不写固定 RGB。未重建 WASM。待统一执行 `sheet-color-reset.contract.test.mjs`、writer 专项、浏览器操作及 XLSX 导出/重导入测试；真实 Excel 重开、主题色/tint 与局部富文本颜色仍未验收。

2026-11-23：继续清点 Sheets Home 字体组第二行。以下来源均是上面的固定提交，参考根为 `E:\rust_test\univer`。本次只读复核、实现与测试编写分开记录，集中验证之前不标记产品能力已验收。

| 标签 | grid 位置 | classic 次序（字体组内） | 项目入口 / 固定源码入口 | 状态与命令去向 |
| --- | --- | --- | --- | --- |
| 加粗 | 第 2 行第 1 列 | 5 | `sheet.command.toggle-bold` / `sheet.command.set-range-bold` | 选区一致为 true 时按下；混合再次设置为 true。固定版由 `SetRangeBoldCommand` 在选区/行内编辑之间分派，底层 `SetBoldCommand` 依据主单元格翻转。 |
| 斜体 | 第 2 行第 2 列 | 6 | `sheet.command.toggle-italic` / `sheet.command.set-range-italic` | 同上；底层 `SetItalicCommand`。 |
| 下划线 | 第 2 行第 3 列 | 7 | `sheet.command.toggle-underline` / `sheet.command.set-range-underline` | 同上；底层 `SetUnderlineCommand`。 |
| 删除线 | 第 2 行第 4 列 | 8 | `sheet.command.toggle-strike` / `sheet.command.set-range-stroke` | 同上；固定版符号是 `SetRangeStrickThroughCommand`，底层 `SetStrikeThroughCommand`。 |
| 边框 | 第 2 行第 5 列 | 11 | `sheet.menu.borders` / `sheet.command.set-border-basic` | 分体按钮主区重用最近预设；15 个预设、颜色子面板、12 个线型。后续已补同一注册命令和 Facade 接线，详见 [边框记录](SHEET-04-borders.md)。 |
| 填充颜色 | 第 2 行第 6 列 | 10 | `sheet.menu.fill` / `sheet.command.set-background-color` | 分体主区重用最近执行颜色；面板来自当前选区；清除经独立重置命令。 |
| 文字颜色 | 第 2 行第 7 列 | 9 | `sheet.menu.text-color` / `sheet.command.set-range-text-color` | 同上；固定版主区默认来自主题 gray.900，填充来自 primary.600；面板值与主区最近值分别订阅。 |

位置与顺序证据：`packages/sheets-ui/src/menu/schema.ts` 的 `RibbonStartGroup.FORMAT`；入口和显隐/禁用证据：`packages/sheets-ui/src/menu/menu.ts` 的四个字体 factory、`TextColorSelectorMenuItemFactory` / `BackgroundColorSelectorMenuItemFactory` / 两个 Reset factory，以及 `packages/sheets-ui/src/menu/border.menu.ts`。固定版禁用条件包含 WorkbookEditablePermission、WorksheetEditPermission、WorksheetSetCellStylePermission、范围 EditPoint；Sheets 外隐藏，填充/边框在普通绘图上下文额外隐藏。本项目工作台本身限定 Sheets，当前沿用工作表/范围保护判断；独立的工作簿只读/设置样式权限和绘图内文字上下文仍须按整体权限/绘图能力核实，不能据此声称完整权限等价。

分类计数：这一行是 7 个常驻根入口（可响应式移动至 overflow，不新增功能）；其中 4 个开关、3 个分体入口。文字/填充各有 45 个预设值、1 个“更多颜色”阶段入口、1 个独立重置操作。边框有 15 个预设、12 个线型、45 个颜色预设和 1 个“更多颜色”阶段入口。自定义阶段提供谱/色相/颜色输入及确认/取消，它们不与预设阶段同时显示，也不把交互控件数当成功能完成数；此局部统计没有覆盖其他组、插件扩展或上下文工具栏。

颜色契约证据：`packages/design/src/components/color-picker/presets.ts` 是 5×9 的 45 色，`ColorPresets.tsx` 使用明确选择状态，`ColorPicker.tsx` 预设立即提交；“更多颜色”在同一弹层切换为草稿阶段，确认/输入 Enter 才提交，取消/Escape 丢弃草稿并返回“更多颜色”焦点。`ColorInput.tsx` 提供 HEX 与 RGB；`ColorPicker` 的默认格式是 hex，只有 rgba 使用 AlphaSlider。对应 `__tests__/ColorPicker.spec.tsx` 验证未设置色与显式黑色、粘贴 HEX、取消、Enter 和焦点。顶部 Sheets 经 `packages/docs-ui/src/views/EditorColorPicker.tsx` 使用默认 hex，没有 alpha 或 transparent 按钮；`ColorSelect.tsx` 的“unset 与 transparent 不同”是另外一类可配置选择器契约，本批不将其误算成顶部入口。

行为差异：固定 schema 的文字重置节点键是 `ResetTextColorCommand.id`，但其 factory 实际返回 `ResetRangeTextColorCommand.id`；`inline-format.command.ts` 在普通选区写主题 gray.900，在行内编辑传 null。项目已有独立清除覆写的 null 契约与 Facade `setFontColor(null)`；本批保留此保存行为并补固定版别名，不把清除改成显式黑色。项目使用整体选区判断开关混合，而固定版开关主要跟随主单元格；混合反馈采用工具提示，颜色面板采用独立混合提示。项目自定义色使用 `antdv-next` ColorPicker，可选 HEX/RGB/HSB，扩展 HSB 输入并不证明任何新 Office 能力。

历史与保存复核：本批已将四个字体开关、设置/重置文字/填充色接入 `sheet-font-format-command.ts` 的共享动作和 Studio 注册命令，工作台经 `runFontFormatCommand` 调用相同注册入口；Facade 的粗细、斜体、文字色和填充 setter 也改走该入口。standalone 工作台使用同语义的 view 适配。命令提交关闭公式重算，跳过工作表/表格筛除行，保留手动隐藏行；重叠范围去重，混合开关设置为 true，未改变时不新增内容历史。重置只扫描已有单元格，不扩展空白稀疏区域。参考底层 `packages/sheets/src/commands/commands/set-style.command.ts` 使用 SetRangeValuesMutation 和 unitID 隔离的 undo/redo；其 `createRangeIteratorWithSkipFilteredRows` 明确跳过筛除行。项目模型的字体/填充色经 `excelViewAdapter.ts`、本地 XLSX writer 的字体和填充 registry 保存，后续统一浏览器与文件样本验证才可判定闭环。

颜色阶段的边框工作包括完整颜色面板、禁用守卫与选区显示状态；2026-11-23 后续边框批次已将 15 预设/12 线型、API/Facade、邻格及合并边段接入共享 plan 和注册命令，具体差异与未执行用例见 [边框记录](SHEET-04-borders.md)。内容内的局部富文本/绘图文字格式、真实 IME、Office 重开、主题/tint 文件保真、首屏与大样本格式操作性能均保留“未验收”。

字体组第一行的字体、字号及增大/减小字号已接入同一字体命令注册和 Facade，另写 25 组契约、13 组浏览器用例，尚未执行，详见 [字体与字号记录](SHEET-04-font-family-size.md)。UI 字号保持 6–400，API 保留有限正数；null 目前恢复 Aptos/11，没有据此声称原始默认样式继承或局部富文本字体命令已实现。

本批落盘记录（2026-11-23）：`packages/office-studio/src/features/sheet-workbench/sheet-colors.ts` 提供 45 色、严格 HEX/RGB/HSB 草稿校验、主题不透明色及边框色归一化；`sheet-font-format-command.ts` 提供动作解析、固定版别名、范围/权限验证和 model/view 两条适配。`components/spreadsheet/SheetColorControl.vue` 使用 `antdv-next@1.8.2` 的 Button、Tooltip、ColorPicker、Form/Form.Item 和 Divider，复用库内 Input/InputNumber/Select，不新增原生交互控件。由于该版本 ColorPicker 不暴露通道草稿或输入引用，组件只查询其库内输入，用于补可访问标签、焦点恢复和原始草稿校验，避免不完整输入提交旧标量。谱/色相的键盘调整同样刷新草稿。顶部弹层使用共享主题 token，边框颜色子面板挂在父弹层内并使用 child popup scope，父层关闭时子层一并关闭；这些交互仍须浏览器验证。

`components/ExcelWorkbench.vue` 已接入选区的文字/填充/边框颜色状态，以及与选区分开的最近执行颜色；带 ID 的工作簿按 ID 隔离记忆，关闭工作簿清理其记忆，取消/权限失败不收到成功事件，重置不清除最近色。当前记忆只接收有效 HEX，旧 API 的其他非空颜色字符串仍原样保留在模型内，不据此声称 named color/theme/tint 的最近色语义已验收。默认文字色取 ConfigProvider 文本 token 的 RGB 部分，默认填充色取 primary token；这与固定版 gray.900/primary.600 的具体 token 值存在实现差异。`i18n/sheets/locales/toolbar-colors.ts` 和 `office-i18n.ts` 接入 19 种语言的颜色阶段、混合、重置和验证文案，尚未逐地区浏览器验收。

快捷键来源：`packages/sheets-ui/src/controllers/shortcuts/style.shortcut.ts` 将粗体、斜体、下划线绑定为 Ctrl/Cmd+B、I、U，删除线是 Ctrl/Cmd+Shift+X；`utils.ts` 的 `whenSheetFocusedInlineFormat` 也允许上游行内编辑。本项目已把四个非编辑 Canvas 选区快捷键接入相同命令，同时忽略 IME、Alt 和已处理事件。当前编辑中的富文本快捷键尚未实现，顶部整格格式操作仍先完成当前单元格编辑，不能记为上游行内命令等价。

新增验证用例：`packages/office-studio/tests/sheet-font-color.contract.test.mjs` 共 24 组，覆盖 45 色、主题色、严格通道输入、颜色记忆、边框混合态、19 语言键、命令别名/参数、快捷键、范围转换、混合/重叠选区、筛除行、稀疏重置、model/view 保真、no-op、权限原子拒绝、Ribbon 布局、Facade/注册命令、撤销重做和快照重开。`sheet-font-color.browser.test.mjs` 共 19 组，覆盖 Canvas 保留、两阶段颜色、焦点、Enter/Escape、合成输入事件、滑块键盘/RGB/HSB、最近色/选区/重置、API 同步、历史和工作簿隔离、取消、保护、主题/RTL、grid overflow、嵌套边框及外部颜色命令期间的编辑草稿。全部仅已编写，尚未执行；合成 composition 事件不能代替真实输入法验收，快照重开不能代替 XLSX 导出和独立 Office 重开。

检查状态：本批只做源码静态审阅及所改 tracked 文件的 `git diff --check`（退出码 0），未运行测试、类型检查、前端构建、Rust/WASM 检查或浏览器。统一验证之前维持“待统一验证”；没有真实文件/浏览器结果和同数据集前后性能记录，不标记完整支持或性能达标。非重置大范围操作仍需落实 Worker/批量模型路径并完成基准，现有 UI 的直接选区门槛不能充当大范围功能完成证明。无 ID 的 standalone 工作簿沿用现有按名称判定的会话边界；重名匿名工作簿的身份隔离仍须由统一模型身份解决。
