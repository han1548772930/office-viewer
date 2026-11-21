# SHEET-13 数据验证进度记录

状态：部分实现，功能未验收。基线为 `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`，公开行为按 2026-11-20 可访问的 Univer Sheets 数据验证文档 v1.4.0 核对。

## 参考依据

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
- Data Ribbon 在 `packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts` 打开实际规则面板。`SheetDataValidationPanel.vue` 配置类型、操作符、公式/边界、内联列表、空值、下拉、错误及输入提示并通过 Facade 命令保存。列表及复选框候选项进入 Canvas 最小输入桥的 datalist。
- `ExcelWorkbench.vue` 在直接键入、文本/TSV 粘贴和清除操作前检查整批变更。STOP 拒绝写入并显示错误；WARNING/INFO 允许变更并显示相应消息。API 写入的 WARNING/INFO 仍可通过 Facade 错误查询获得明细。
- `excelViewAdapter.ts`、`workbookNormalizer.ts` 和 workbook 输入归一化保留规则字段。Rust `office-writer/src/xlsx_validation.rs` 写出 SpreadsheetML `dataValidations`、多范围、比较类型、列表、复选框、公式、下拉和提示/错误设置。AutoFilter 仍由原 `xlsx_filter.rs` 路径写出。
- `office-parser` 将 `rxls::Sheet::data_validations()` 映射为稳定的 `dataValidations` DTO。XLSX sidecar 按 workbook relationship 定位 worksheet，补回 `showDropDown`（包含 OOXML 反向语义）和 `errorStyle`，按 range 与 rxls 读回规则配对；不支持、坏范围和未能配对的规则进入 parser warnings。

## 验证

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
- 本地 UI 暂无日期快捷选择、复选框单元格的绘制交互、列表搜索/彩色标签、多选弹出控件；多选文本当前以逗号分隔。
- 没有为规则修改/数据写入组合开展大范围结构编辑、行列移动、筛选隐藏、权限与协作资源变换验收。SHEET-13 保持未验收。
