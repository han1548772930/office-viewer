# SHEET-13：复选框快捷入口与勾选链路

日期：2026-11-27。固定版本：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。续接已写入 Ant Checkbox 单主格切换、共享公式读值、Canvas 文字抑制及跨冻结 pane 裁剪；测试和工程运行检查后置，SHEET-13 保持未验收。规则默认空格填值与自定义值迁移仍待实施。

<a id="checkbox-entry-review-2026-11-27"></a>

## 固定入口和行为

来源：[schema](../../../univer/packages/sheets-data-validation-ui/src/menu/schema.ts)、[quickCheckboxMenuFactory](../../../univer/packages/sheets-data-validation-ui/src/menu/dv.menu.ts)、[快捷内容命令](../../../univer/packages/sheets-data-validation-ui/src/commands/commands/data-validation-ui.command.ts)。Insert / CELL 的复选框位于 order 0，是 BUTTON_SELECTOR；这个注册节点没有 gridLayout。分类计数为 1 个常驻根、1 个默认主动作、2 个菜单动作（设置、清除），该 menu factory 没有扩展或上下文根；窄屏溢出仍是同一入口。单元格 Checkbox 是另 1 类 widget 上下文，不能把可见格数量计作新增根。

主动作和“设置复选框”提交 `{ rule: { type: CHECKBOX, operator: undefined, formula1: undefined, formula2: undefined } }`，绑定当前全部 selections 的 ranges；“清除复选框”提交 `types: [CHECKBOX]`，只裁剪选中范围内该类型的规则。菜单仅在 Sheets 显示，禁用条件包含 WorkbookEditable、WorksheetSetCellStyle、WorksheetEdit、RangeProtectionEdit。

[CheckboxValidatorView](../../../univer/packages/sheets-data-validation-ui/src/views/validator-views/checkbox-validator-view.ts) 注册 [CheckboxRender](../../../univer/packages/sheets-data-validation-ui/src/views/widgets/checkbox-widget.ts)。固定控件按字体大小的 1.6 倍确定方框尺寸，按水平/垂直对齐定位，在合并单元格使用主格边界；只在方框内命中，右键不切换，指针进入使用 pointer cursor。空格和与勾选/未勾选值相符的格绘制方框，其他值继续显示普通文本。点击只写命中的实际主格，不把当前全部选区一起切换；写入经过 `sheet.command.set-range-values` 及真实历史。

[CheckboxValidator](../../../univer/packages/sheets-data-validation/src/validators/checkbox-validator.ts) 的默认勾选/未勾选值为 **1 / 0**；`transformCheckboxValue` 只把 TRUE/FALSE 布尔词映射为 1/0，其他自定义文字保持原值。自定义值的原始表示也参与验证。

[getDataValidationDiffMutations](../../../univer/packages/sheets-data-validation/src/commands/commands/data-validation.command.ts) 在添加 checkbox 时给符合条件的空格填 originFormula2，范围扩展也填未勾选值，规则和内容变更进入同一 undo/redo。更新自定义值时把旧勾选/未勾选值映射为新值，保留不匹配的内容；不会用空格填值覆盖富文本。清除入口裁剪规则范围，不能仅凭删除菜单证明这些内容行为已对齐。

## 本地链路与本轮实现

[quick-validation.ts](../../packages/office-studio/src/features/sheet-validation/quick-validation.ts) 的 checkbox plan 创建一条覆盖全部选区的规则；[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 一次交给统一内容命令。[office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的设置/清除入口在 `commitWorkbook` 内检查目标范围权限，写入按工作簿隔离的历史。[FWorksheet.getCellValidationOptions](../../packages/office-studio/src/api/facade.ts) 查询当前编辑格的真实规则。

本地 [validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts) 与 [validation-checkbox.ts](../../packages/office-studio/src/features/sheet-validation/validation-checkbox.ts) 共用规则值和完整 TRUE/FALSE 别名比较；缺省候选已改为 **1 / 0**，自定义大小写、空白和 scalar 类型保留，部分配置使用另一侧缺省值。规则公式在保存 anchor 或首范围经共享公式 reader 求值；隐藏依赖、错误和无效规则不显示 checkbox。

[parseBool](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/parse-value.ts) 实际把输入 TRUE/FALSE 解析为布尔值；[sheet-cell-input-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-input-command.ts) 对该真实结果进行合并主格归一化、会话冲突、工作表/范围权限和数据验证。[sheet-cell-input-commands.ts](../../packages/office-studio/src/api/sheet-cell-input-commands.ts) 成功后一次提交内容并发布 SheetValueChanged，拒绝时返回 accepted false，保留原值与历史。

## 草稿和静态检查

[sheet-quick-validation.contract.test.mjs](../../packages/office-studio/tests/sheet-quick-validation.contract.test.mjs) 新增 4 个参数化真实命令草稿：显式 TRUE/FALSE、显式 1/0、只配置 DONE、只配置 NO。测试调用 `sheet.command.commit-cell-input`，覆盖候选、解析后的布尔值、单次内容历史、拒绝不改快照、自定义文字大小写精确匹配、undo/redo 和原生 WorkbookModel 重开。另补只清除 A1 后保留 A2 复选框及 B1 列表规则的真实命令草稿。全部未运行。

静态检查限定本轮差异、文档文件链接及 `git diff --check`。原生快照重开草稿不代表 XLSX 文件互操作验收；本轮没有测试、类型、构建、Rust/WASM、浏览器或性能运行结果。

## 已落盘的 Canvas 勾选与保留缺口

[SheetCheckboxValidationOverlay.vue](../../packages/office-studio/src/components/spreadsheet/SheetCheckboxValidationOverlay.vue) 使用受控 Ant Checkbox、ConfigProvider theme token 和真实控件事件；[validation-checkbox-layout.ts](../../packages/office-studio/src/rendering/spreadsheet/validation-checkbox-layout.ts) 按可见稀疏轴、合并主格、align/vertical 和 zoom 定位，汇集全部冻结 pane fragments，单个控件以 clip-path 裁剪。跨两条冻结线时仍保持一个方框；没有自绘 Checkbox 或 Table 编辑器替代。

[use-sheet-validation-checkbox.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-checkbox.ts) 捕获主格、规则、用户、版本与输入会话。点击、Checkbox Space/Enter 和 Canvas Space 路由 [validation-checkbox-input.ts](../../packages/office-studio/src/features/sheet-validation/validation-checkbox-input.ts) 的 `sheet.command.commit-checkbox-input`，由 [sheet-validation-input-commands.ts](../../packages/office-studio/src/api/sheet-validation-input-commands.ts) 注册并释放。一次 typed 内容写入只切换命中的主格，直接保存原 checked/unchecked scalar，复用完整 merge View/Edit、验证、一次历史和 SheetValueChanged；旧 pointer/Space、输入或 composition、命令取消与重入不覆盖新内容。无效现有文本/富文本/图片仍是 Canvas 原内容。

[scene.ts](../../packages/office-studio/src/rendering/spreadsheet/scene.ts) 仅对真实可见合法 Checkbox 控件抑制普通 scalar 文字；背景、边框、选择和 fill handle 保留。View-denied 格不生成控件或内容反馈；只允许 View 的格显示禁用 Ant Checkbox。输入提示/无效 marker、公开 show/hide 和动态源的实际链路、静态结果及未运行草稿见 [续接专项记录](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)。

固定版默认空格填值、自定义值变更迁移已进入本地同一次规则/cells 事务；同步共享 classifier 和有界 Worker/Rust 计划、异步 owner/取消/内存 admission 见 [联合事务专项](SHEET-13-checkbox-rule-cells-worker-2026-11-28.md)。这些代码已落盘，运行验收仍为 0。画布行为、clip-path、主题、焦点、真实 Windows IME、保护、独立 Office 样本和大范围性能全部保持未验收；内部保存重开与合成 composition 草稿不能替代 Office/IME 验收。
