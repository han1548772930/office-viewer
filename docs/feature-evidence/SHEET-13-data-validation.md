# SHEET-13 数据验证进度记录

状态：部分实现，功能未验收。基线为 `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`，公开行为按 2026-11-20 可访问的 Univer Sheets 数据验证文档 v1.4.0 核对。

2026-11-27 快捷入口核对与 CSV 小修见 [快捷下拉与快捷日期](SHEET-13-quick-dropdown-date.md)，日期/时间 Canvas 控件的待确认改动见 [实施方案](SHEET-13-date-picker-entry-plan.md)。本轮运行检查全部后置，以下历史通过记录不覆盖本轮差异。

复选框自定义布尔值与默认候选验证的小修、固定默认值差异和 Canvas 勾选缺口见 [复选框入口证据](SHEET-13-checkbox-entry-review.md)。本轮仅新增真实命令草稿，未验收。

多选列表的单元格 CSV 校验小修、固定版 JSON 表示差异和未完成的多选弹层见 [多选列表证据](SHEET-13-list-multiple-csv.md)。本轮草稿未运行，未验收。

## 参考依据

**数据页与画布取区（待统一测试）：** 固定版 `sheets-data-validation-ui/src/menu/schema.ts` 注册管理与添加两个子项；`sheets-data-validation/src/utils/create.ts` 的默认规则为 decimal/equal/100。工作台的 Data/RULES 菜单与规则列表的添加操作共用创建和按 ID 编辑路径。对照 `DataValidationDetail.tsx` 的 RangeSelector，本地以同工作表 Canvas 选区加显式确认回填草稿：应用范围保留多个矩形，列表源和比较操作数接受单一连续引用。取消不改草稿，切换工作表关闭面板。本地编辑规则采用显式应用，与上游回调提交时机不同；焦点、明暗主题、实际拖选、跨表来源和 Office 互操作尚未验收。

**2026-11-22 续作，待统一验证：** 对照固定提交 `packages/sheets-data-validation-ui/src/menu/dv.menu.ts`、`views/components/DropdownPresets.tsx` 及 `commands/commands/data-validation-ui.command.ts`。Insert/CELL 的复选框与日期分别使用 split 主动作和分类清除；下拉菜单含三列预设、编辑与清除。`QuickDropdownPicker.vue` 用 antdv-next Button/Divider 与主题 token，窄屏 More 复用相同命令参数。工作台一次将全部选区 ranges 交给 `sheet.command.set-data-validation` 或 `sheet.command.clear-quick-data-validation`；清除下拉同时匹配 list/listMultiple。编辑下拉、编辑日期先创建规则再以该规则 ID 打开详情。中文和英文预设文案已接入；其它语言的本地化、焦点、真实多选区、XLSX 往返与受保护主体留待统一测试，未验收。

- 固定版菜单注册：`packages/sheets-data-validation-ui/src/menu/schema.ts`，规则面板入口位于 RibbonDataGroup.RULES；快速复选框、下拉和日期入口属于 RibbonInsertGroup.CELL。
- 固定版命令入口：`packages/sheets-data-validation-ui/src/commands/commands/data-validation-ui.command.ts`，快捷规则绑定当前选择并转发至统一数据验证命令。
- 固定版规则服务：`packages/sheets-data-validation/src/commands/commands/data-validation.command.ts`、`models/rule-matrix.ts`、`services/dv-validator.service.ts` 和 `validators/`。覆盖规则/范围更新、选区冲突、单元格写入、复选框默认值和公式验证。
- 固定版工作台装配：`examples/src/sheets/mount.ts` 注册 `UniverSheetsDataValidationPreset()`。
- 官方公开行为：`https://docs.univer.ai/zh-CN/guides/sheets/features/data-validation` 列出数字、整数、文本长度、日期、复选框、单/多选下拉与自定义公式；规则选项包含空白策略、下拉显示、输入提示、错误提示和 STOP/WARNING/INFO。

## 本地实现

- 领域模型在 `packages/office-studio/src/domain/model.ts` 的 `SheetDataValidationRule`，按工作表保存稳定 ID、类型、0 基范围、条件公式、错误策略、提示和列表呈现选项。
- 引擎在 `packages/office-studio/src/features/sheet-validation/validation-engine.ts`。实现整数、小数、日期、时间、文本长度、复选框、内联/区域来源单选与多选列表、自定义公式；实现规则范围覆盖替换/清除、日期系统、错误消息及选区结果查询。
- `sheet.command.set-data-validation`、`sheet.command.clear-data-validation` 和 `sheet.command.set-range-values` 在 `packages/office-studio/src/api/office-studio.ts` 注册并进入按工作簿隔离的历史。注册值命令对整批新值预检，STOP 失败时不会提交部分值或创建历史。
- `FWorksheet` 与 `FRange` 在 `packages/office-studio/src/api/facade.ts` 提供设置、查询、清除、范围错误查询和列表选项查询。Facade 批量写入也预检，不绕过工作表规则。
- Data Ribbon 在 `packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts` 打开实际规则面板。`SheetDataValidationPanel.vue` 配置类型、操作符、公式/边界、内联列表、空值、下拉、错误及输入提示并通过统一内容命令保存。列表及复选框候选项按当前编辑主格查询，进入 Canvas 的 antdv-next AutoComplete / Input.TextArea 和既有输入会话。
- `ExcelWorkbench.vue` 在直接键入、文本/TSV 粘贴和清除操作前检查整批变更。STOP 拒绝写入并显示错误；WARNING/INFO 允许变更并显示相应消息。API 写入的 WARNING/INFO 仍可通过 Facade 错误查询获得明细。
- `excelViewAdapter.ts`、`workbookNormalizer.ts` 和 workbook 输入归一化保留规则字段。Rust `office-writer/src/xlsx_validation.rs` 写出 SpreadsheetML `dataValidations`、多范围、比较类型、列表、复选框、公式、下拉和提示/错误设置。AutoFilter 仍由原 `xlsx_filter.rs` 路径写出。
- `office-parser` 将 `rxls::Sheet::data_validations()` 映射为稳定的 `dataValidations` DTO。XLSX sidecar 按 workbook relationship 定位 worksheet，补回 `showDropDown`（包含 OOXML 反向语义）和 `errorStyle`，按 range 与 rxls 读回规则配对；不支持、坏范围和未能配对的规则进入 parser warnings。

<a id="list-color-preservation-review"></a>

## 2026-11-27 管理入口保留列表颜色

状态：已写入修复及浏览器草稿，运行验证后置，SHEET-13 保持未验收。

固定 v1.6.0 来源：

- [DataValidationDetail.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/components/DataValidationDetail.tsx) 的 `baseRule` 保留 `formula1/formula2`；`handleChangeType` 在 LIST 与 LIST_MULTIPLE 之间保留规则设置，切到其它类型时重置条件公式。
- [ListFormulaInput.tsx](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/components/formula-input/ListFormulaInput.tsx) 用 `formula2.split(',')` 读取内联与引用列表的颜色，引用选项颜色更新写回 `newColors.join(',')`，修改引用公式时同样携带 `formula2`。颜色按选项索引对应，空颜色位置属于原数据。
- Data/RULES 根入口仍为 `sheet.menu.data-validation`，grid r1c1，两个子项依次是管理验证规则、添加规则；根、子项和 classic 顺序增量均为 0。固定 schema 及默认添加链沿用上文依据。

本地 [SheetDataValidationPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataValidationPanel.vue) 已读取既有 `formula2`，但保存草稿只为比较规则和自定义复选框写出第二个操作数，导致编辑列表范围后删除原颜色。现有对象构造补一项：当前与原规则都属于 list/listMultiple 时保留原 `formula2`，包括单选与多选互换；原本没有颜色时保持缺省，整数的 `10` 和复选框的 FALSE 不成为颜色。生产增量 +1 行，无新增 helper、模型、命令、Facade、Worker 或 WASM 协议。

接线保持 `apply → createSheetDataValidationRule → sheet.command.replace-data-validation → commitWorkbook/replaceSheetDataValidationRule`。命令同时校验原范围与目标范围的编辑权限，并使用工作簿隔离的内容历史；原规则 ID、metadata 及列表源沿原路径保留。本地仍采用显式应用，上游按回调提交。面板已有 Ant Form/FormItem、Input、Select、Checkbox、Collapse 与 Button，主题来自统一 ConfigProvider/token；现有消息段落的 Ant 反馈迁移尚未完成。

原 [sheet-format-painter.browser.test.mjs](../../packages/office-studio/tests/sheet-format-painter.browser.test.mjs) 追加 2 个 `test.each` 声明、16 个计划场景，复用真实 ExcelWorkbench/Studio、命令和原 `painterFixture`，未新增 fixture 或生产测试入口：

- classic/grid × list/listMultiple × 内联/引用，共 8 例：范围编辑、主题切换保留 Ant 输入焦点与草稿、列表类型互换、颜色字符串/空位置/ID/metadata、两次独立编辑历史、撤销重做与内部快照重开。
- classic/grid × 整数/自定义复选框/无色单选/无色多选，共 8 例：转换为列表后不写入数字、布尔第二操作数或默认 `10` 颜色，并记录一次转换历史。

本轮仅读取固定源码、本地实现和 diff，未执行草稿、类型、构建、浏览器、Rust/WASM、Office 或基准。逐项颜色编辑、选项重排时的颜色映射、Canvas 彩色标签、多选控件及真实 Office 颜色往返仍待实现或验收；当前 CSV 文本框沿用颜色数组顺序与原字符串。原生 IME、面板关闭回焦、真实明暗主题和两种渲染后端均未验收，内部快照重开草稿不构成 XLSX 导出重开证据。

## 验证

以下为历史执行记录；2026-11-27 新增命令及浏览器草稿均未运行。

- `node --experimental-strip-types --test packages/office-studio/tests/sheet-validation-engine.contract.test.mjs packages/office-studio/tests/office-api.contract.test.mjs packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs`：20/20 通过，包含更新规则后保留未覆盖区域的回归。
- `pnpm.cmd --filter @office-studio/office-studio check:types`：通过。
- `pnpm.cmd exec playwright test --config tests/browser/sheet-toolbar.playwright.config.ts`：7/7 通过。验证 Data Ribbon 面板创建整数/比较规则、有效键入、STOP 拒绝且保留原值、内联下拉候选输入，以及模型保存/重开后下拉候选保留。
- `cargo test --manifest-path crates/office-writer/Cargo.toml`：39/39 通过，包含新写出模块测试、原 AutoFilter 测试和 DOCX run formatting 测试。
- `cargo fmt --manifest-path crates/office-parser/Cargo.toml -- --check`：通过。
- `cargo test --manifest-path crates/office-parser/Cargo.toml`：8/8 通过。真实 XLSX writer fixture 覆盖基础验证 DTO、公式/范围/提示/错误字段、多范围列表、下拉反转及 INFO 样式映射；sidecar 单元测试覆盖边界和坏范围诊断。

## 未完成与未核实

- parser 的 round-trip fixture 由本地 `rxls` writer 生成，并修改 worksheet XML 来覆盖其 writer 未暴露的 OOXML 属性；尚未用 Excel、LibreOffice 或独立 XLSX 读取器验证第三方文件及导出文件往返。
- SHEET-13 仍未验收。公式引擎/UI 覆盖和结构编辑、协作验收缺口见上文及功能对齐清单。
- 本地公式求值能力少于 Univer 完整公式引擎；跨表/命名范围、复杂相对/绝对引用及公开函数覆盖仍需扩展样本和边界测试。
- 日期规则已有快捷创建，Canvas 日期/时间控件和复选框单元格勾选仍未实现。列表搜索/彩色标签和多选弹出控件的完整验收仍待补齐；多选文本当前采用 CSV，带逗号或引号的标签必须按 CSV 字段转义，Canvas 候选点击仍整段替换。
- 没有为规则修改/数据写入组合开展大范围结构编辑、行列移动、筛选隐藏、权限与协作资源变换验收。SHEET-13 保持未验收。
