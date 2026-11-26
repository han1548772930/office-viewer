# SHEET-13：复选框快捷入口与勾选链路

日期：2026-11-27。固定版本：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本轮只修验证与真实候选的错接，测试和工程检查后置；Canvas 勾选尚未实现，SHEET-13 保持未验收。

<a id="checkbox-entry-review-2026-11-27"></a>

## 固定入口和行为

来源：[schema](../../../univer/packages/sheets-data-validation-ui/src/menu/schema.ts)、[quickCheckboxMenuFactory](../../../univer/packages/sheets-data-validation-ui/src/menu/dv.menu.ts)、[快捷内容命令](../../../univer/packages/sheets-data-validation-ui/src/commands/commands/data-validation-ui.command.ts)。Insert / CELL 的复选框位于 order 0，是 BUTTON_SELECTOR；这个注册节点没有 gridLayout。分类计数为 1 个常驻根、1 个默认主动作、2 个菜单动作（设置、清除），没有扩展或上下文入口；窄屏溢出仍是同一入口。

主动作和“设置复选框”提交 `{ rule: { type: CHECKBOX, operator: undefined, formula1: undefined, formula2: undefined } }`，绑定当前全部 selections 的 ranges；“清除复选框”提交 `types: [CHECKBOX]`，只裁剪选中范围内该类型的规则。菜单仅在 Sheets 显示，禁用条件包含 WorkbookEditable、WorksheetSetCellStyle、WorksheetEdit、RangeProtectionEdit。

[CheckboxValidatorView](../../../univer/packages/sheets-data-validation-ui/src/views/validator-views/checkbox-validator-view.ts) 注册 [CheckboxRender](../../../univer/packages/sheets-data-validation-ui/src/views/widgets/checkbox-widget.ts)。固定控件按字体大小的 1.6 倍确定方框尺寸，按水平/垂直对齐定位，在合并单元格使用主格边界；只在方框内命中，右键不切换，指针进入使用 pointer cursor。空格和与勾选/未勾选值相符的格绘制方框，其他值继续显示普通文本。点击只写命中的实际主格，不把当前全部选区一起切换；写入经过 `sheet.command.set-range-values` 及真实历史。

[CheckboxValidator](../../../univer/packages/sheets-data-validation/src/validators/checkbox-validator.ts) 的默认勾选/未勾选值为 **1 / 0**；`transformCheckboxValue` 只把 TRUE/FALSE 布尔词映射为 1/0，其他自定义文字保持原值。自定义值的原始表示也参与验证。

[getDataValidationDiffMutations](../../../univer/packages/sheets-data-validation/src/commands/commands/data-validation.command.ts) 在添加 checkbox 时给符合条件的空格填 originFormula2，范围扩展也填未勾选值，规则和内容变更进入同一 undo/redo。更新自定义值时把旧勾选/未勾选值映射为新值，保留不匹配的内容；不会用空格填值覆盖富文本。清除入口裁剪规则范围，不能仅凭删除菜单证明这些内容行为已对齐。

## 本地链路与本轮小修

[quick-validation.ts](../../packages/office-studio/src/features/sheet-validation/quick-validation.ts) 的 checkbox plan 创建一条覆盖全部选区的规则；[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 一次交给统一内容命令。[office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的设置/清除入口在 `commitWorkbook` 内检查目标范围权限，写入按工作簿隔离的历史。[FWorksheet.getCellValidationOptions](../../packages/office-studio/src/api/facade.ts) 查询当前编辑格的真实规则。

本地 [validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts) 默认候选仍为 **TRUE / FALSE**，默认规则此前已兼容布尔及 1/0。原自定义分支直接比较 `String(actual)` 与公式字面值，既拒绝输入解析得到的 true/false，也拒绝仅配置一侧时另一侧的默认候选。本轮改为复用 `sheetDataValidationOptions`，比较时仅把整个 TRUE/FALSE 布尔词映射为 1/0；其他自定义文字仍精确匹配。候选、保存数据和接口形状不变。

[parseBool](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/parse-value.ts) 实际把输入 TRUE/FALSE 解析为布尔值；[sheet-cell-input-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-input-command.ts) 对该真实结果进行合并主格归一化、会话冲突、工作表/范围权限和数据验证。[sheet-cell-input-commands.ts](../../packages/office-studio/src/api/sheet-cell-input-commands.ts) 成功后一次提交内容并发布 SheetValueChanged，拒绝时返回 accepted false，保留原值与历史。

## 草稿和静态检查

[sheet-quick-validation.contract.test.mjs](../../packages/office-studio/tests/sheet-quick-validation.contract.test.mjs) 新增 4 个参数化真实命令草稿：显式 TRUE/FALSE、显式 1/0、只配置 DONE、只配置 NO。测试调用 `sheet.command.commit-cell-input`，覆盖候选、解析后的布尔值、单次内容历史、拒绝不改快照、自定义文字大小写精确匹配、undo/redo 和原生 WorkbookModel 重开。另补只清除 A1 后保留 A2 复选框及 B1 列表规则的真实命令草稿。全部未运行。

静态检查限定本轮差异、文档文件链接及 `git diff --check`。原生快照重开草稿不代表 XLSX 文件互操作验收；本轮没有测试、类型、构建、Rust/WASM、浏览器或性能运行结果。

## 已确认缺口与待确认 Canvas 方案

[scene.ts](../../packages/office-studio/src/rendering/spreadsheet/scene.ts) 目前仍按普通单元格文本绘制，没有 checkbox 控件状态；[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 的格命中路径只选择/拖选或打开超链接，没有 checkbox 绘制、命中和切换事件。当前候选进入 Ant AutoComplete / Input.TextArea，不能把候选文字编辑称为可点击的画布勾选。默认空格填值和自定义值迁移也未进入本地规则应用链路。

画布勾选需要新增 UI，以下仅为待确认方案。路径相对 `packages/office-studio/src/`，7 个生产文件约 248–383 行；测试与证据另列，当前没有实施这些改动。

| 文件 | 行数 | 具体改动 |
| --- | --- | --- |
| `features/sheet-validation/validation-engine.ts` | 40–70 | 复用真实候选，提供空格/合法值的 checkbox 状态和下一值；供验证与显示共用；保留非布尔自定义文字和非法现有内容 |
| 新增 `api/sheet-checkbox-input-commands.ts` | 55–85 | 单格 typed 切换命令，绑定目标和当前源值，复用权限、验证及一次内容历史；直接写 CellScalar，保留自定义文字，避免输入 parser 改写标签 |
| 新增 `components/spreadsheet/SheetCheckboxValidationOverlay.vue` | 70–100 | 使用 Ant Checkbox、统一主题 token 与受控状态；只创建当前可见格控件；按单元格边界裁剪，处理点击/空格键和焦点返回 |
| `rendering/spreadsheet/scene.ts` | 8–15 | 在受控 checkbox 覆盖层对应的空格/合法格抑制普通值文字，保留网格、背景、边框和其他内容的 Canvas 绘制 |
| `components/spreadsheet/SpreadsheetCanvas.vue` | 45–70 | 使用既有 geometry/panes 生成可见主格位置，处理合并、冻结、滚动和缩放；绑定 Ant 覆盖层并发送单格切换 |
| `components/ExcelWorkbench.vue` | 25–35 | 校验当前工作台与目标，路由统一内容命令，协调编辑会话/IME 和焦点；多选区点击只切换所属主格 |
| `api/office-studio.ts` | 5–8 | 注册/释放切换命令 |

上述表只描述画布勾选的新 UI 改动。固定版默认空格填值、自定义值变更迁移及其大选区计算/内存路径仍是独立模型缺口，尚无可交付实现；不能借本轮布尔候选小修将其标为完成。画布行为、主题、焦点、输入法、保护、文件样本和大选区性能均保持未验收。
