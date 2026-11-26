# SHEET-13：快捷下拉与快捷日期

日期：2026-11-27。固定对照：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

本轮修正快捷下拉和规则面板的列表转义，并补充真实命令及浏览器测试草稿。所有运行检查按本轮约定后置；快捷日期的 Canvas 日期/时间控件尚未实现，SHEET-13 保持未验收。

<a id="quick-dropdown-date-entry-review-2026-11-27"></a>

## 固定入口与计数

来源：[固定 schema](../../../univer/packages/sheets-data-validation-ui/src/menu/schema.ts)、[菜单 factory](../../../univer/packages/sheets-data-validation-ui/src/menu/dv.menu.ts)、[预设组件](../../../univer/packages/sheets-data-validation-ui/src/views/components/DropdownPresets.tsx)。

| 入口 | 固定分组、顺序和形态 | 子动作 | 固定命令 |
| --- | --- | --- | --- |
| 复选框，相邻入口 | Insert / CELL，order 0 | 本轮只记录邻项，未重新验收 | `data-validation.command.insert-quick-rule` |
| `sheet.menu.quick-dropdown` | Insert / CELL，order 1，SUBITEMS | 是/否、未开始/进行中/已完成、ABC；编辑下拉列表；清除下拉列表 | 插入 `data-validation.command.insert-quick-rule`；编辑 `data-validation.command.addRuleAndOpen`；清除 `data-validation.command.clear-quick-rule` |
| `sheet.menu.quick-date` | Insert / CELL，order 2，BUTTON_SELECTOR | 主按钮默认日期；菜单含日期、日期和时间、编辑、移除 | 同上三条命令 |

这三个 Insert / CELL 注册节点均未声明 `gridLayout`，不把本地网格坐标写成固定来源事实。下拉预设容器宽 360px、三列、gap 1.5、按钮最小高度 80px；本地 [QuickDropdownPicker.vue](../../packages/office-studio/src/components/spreadsheet/QuickDropdownPicker.vue) 对应三列和 80px，另按视口限制容器宽度。

可复查分类计数：本轮有 2 个常驻根入口；下拉有 3 个预设动作和 2 个普通菜单动作，共 5 个；日期有 1 个默认主动作和 4 个菜单动作。固定下拉声明的一行自定义 selection 与其内部 3 个按钮只计一次。此次 2 个根入口不属于扩展或上下文入口；本地窄屏移入共享 Ribbon 溢出区时仍是同一入口，不新增计数。计数只说明入口覆盖，不证明行为完成。

固定菜单仅在 Sheets 显示；禁用条件同时包含 WorkbookEditable、WorksheetSetCellStyle、WorksheetEdit、RangeProtectionEdit。实际定义见 `dv.menu.ts` 的 `getQuickDataValidationMenuState`。本地 [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 另包含当前工作台、编辑会话和输入法状态约束，不能把本地全部状态条件称为上游同源。

## 规则与选区契约

来源：[UI 内容命令](../../../univer/packages/sheets-data-validation-ui/src/commands/commands/data-validation-ui.command.ts)、[默认规则](../../../univer/packages/sheets-data-validation/src/utils/create.ts)、[规则新增与范围清除](../../../univer/packages/sheets-data-validation/src/commands/commands/data-validation.command.ts)。

快捷下拉固定 payload：

```ts
{ rule: {
  type: LIST,
  operator: undefined,
  formula1: serializeListOptions(values),
  formula2: '',
  renderMode: CUSTOM,
} }
```

“编辑下拉列表”先用本地化 Option1 / Option2 创建新规则，再打开这条新规则；不是查找并复用原规则。“清除下拉列表”传入 `types: [LIST, LIST_MULTIPLE]`。

快捷日期固定 payload：

```ts
{ rule: { type: DATE, operator: undefined, formula1: undefined, formula2: undefined } }
```

“日期和时间”在同一规则上增加 `bizInfo: { showTime: true }`；移除传入 `types: [DATE]`。无 operator 的 date 规则表示任意合法日期，不附加范围条件。普通“添加规则”的默认值另外来自 `createDefaultNewRule`，为 decimal / equal / formula1 `'100'`，不能套用快捷日期默认值。

三个 UI 内容命令都读取当前**全部 selection ranges**。`addRuleAndOpen` 总是创建新 uid，新增成功后打开 `{ ruleId: rule.uid, isAdd: true }`。核心新增经过 rule matrix clone / add / diff、validator.normalizeFormula 和真实 undo/redo；范围清除按 types 筛选规则并裁剪所选范围，保留范围外残余。

本地链路：

1. [quick-validation.ts](../../packages/office-studio/src/features/sheet-validation/quick-validation.ts) 使用 `planQuickValidation` 生成全部目标范围的一条规则或带类型的清除计划。普通添加使用 decimal / equal / `'100'`；快捷日期以 `metadata.showTime` 表达显示时间。
2. [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 `applyQuickDataValidation` 一次把全部 ranges 交给内容命令；仅明确的 edit 动作在成功后按新规则 ID 打开面板。
3. [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 的 `sheet.command.set-data-validation`、`replace-data-validation`、`clear-quick-data-validation` 复用 `commitWorkbook` 和 [validation-engine.ts](../../packages/office-studio/src/features/sheet-validation/validation-engine.ts)，进入文档隔离的真实内容历史。
4. 后续单元格编辑以 editingKey / selectedKey 对应的主格调用 [FWorksheet.getCellValidationOptions](../../packages/office-studio/src/api/facade.ts)，由 `sheetDataValidationOptions` 获取这格的真实候选；[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 使用现有 Ant AutoComplete / Input.TextArea 和输入会话。规则写入范围与候选查询目标分别是全部选区和当前编辑主格。

上述持久化草稿验证本项目原生 WorkbookModel 快照重开，不代表 XLSX 文件写出/重开的验收。

## 本轮小修与表示差异

固定 [serializeListOptions / deserializeListOptions](../../../univer/packages/sheets/src/basics/utils.ts) 当前用 JSON 数组字符串，读取时先解析 JSON，旧数据才回退到逗号分隔。本地 `formula1` 使用已有 CSV 列表表达式：逐项处理逗号、双引号和 CR/LF，双引号双写，再加 formula1 外层引号。这是领域模型表示差异，本轮未新增 Univer JSON 兼容层或商业 SDK 依赖。

本轮仅修两处确定错接：

- `createQuickSheetDataValidationRule` 的旧手写拼接会把 `Green, dark` 拆成两个候选；现复用 `serializeValidationListItems`，同时保留引号和换行所在的单项。
- `SheetDataValidationPanel.vue` 原来提前反转双引号，再按逗号 split，重新保存会破坏单项边界；现加载时只移除 formula1 外层引号，保存复用 `parseValidationListItems` 和 `serializeValidationListItems`。原有单行 Ant Input、Form、错误反馈和提交命令不变。

规则面板原来的 `apply` 已复制 `rule.metadata`，普通保存不会丢失 `showTime`。本轮没有把这个已存在行为当作缺陷修改。日期类型目前也没有编辑 `showTime` 的勾选项。

## 草稿与静态检查

| 文件 | 本轮真实断言 | 执行状态 |
| --- | --- | --- |
| [sheet-quick-validation.contract.test.mjs](../../packages/office-studio/tests/sheet-quick-validation.contract.test.mjs) | 快捷默认值与显示时间；含逗号、引号和换行的 list 写入 B2 / B4；B3 无规则；候选完整；错误子项拒绝且快照不变；正确单项可编辑；单次规则历史、内容 undo/redo、原生重开 | 未运行 |
| [sheet-quick-validation.spec.ts](../../tests/browser/sheet-quick-validation.spec.ts) | 3 个真实 Ribbon 下拉预设；全部不连续 A1 / A4；一个规则、一次历史；规则面板真实输入编辑保留逗号/引号；日期及日期和时间规则；undo/redo 和原生重开 | 未运行 |

[sheet-toolbar.playwright.config.ts](../../tests/browser/sheet-toolbar.playwright.config.ts) 已将新增浏览器草稿纳入 testMatch，配置变更由主任务持有。

浏览器夹具使用既有 runtime 资源路径选择方式。`OfficeStudio.createWorkbook` 的默认激活和 `history.clear` 都建立 UnitHistory，故夹具读取当前书的 undo 队列，不增加假设性的历史 fallback。菜单定位依据当前共享 Ant Menu 的 menuitem 角色；面板定位依据实际 aside、Ant Button 和 Input。

静态交付只检查本轮差异、文件链接和限定文件 `git diff --check`；这不替代类型、前端、Rust/WASM、浏览器、文件样本和性能检查。测试、构建、dev 服务均未运行。

## 快捷预设尺寸复核

固定 [DropdownPresets.tsx](../../../univer/packages/sheets-data-validation-ui/src/views/components/DropdownPresets.tsx) 的预设按钮使用 `border-0 / p-3 / gap-1.5 / text-sm / leading-5`，每个候选值占满内容宽度，分别对应无边框、12px 内边距、6px 行间距、14px 字号和 20px 行高。

[QuickDropdownPicker.vue](../../packages/office-studio/src/components/spreadsheet/QuickDropdownPicker.vue) 本批只改三行：预设 Button 通过 `styles.root` 固定零边框；按钮使用 border-box、12px 内边距和 14/20px 文本；候选列表占满内容宽度并使用 6px 行间距。原 360px 面板上限、三列、最小 80px 高度、完整可访问名称、禁用属性和命令接口继续沿用。

依据本地 [Button API](../llm.md) 的 `styles` 语义样式及安装版 [Button.js](../../node_modules/antdv-next/dist/button/Button.js)，`styles.root` 合并到真实按钮的 inline style；[buttonHelper.js](../../node_modules/antdv-next/dist/button/buttonHelper.js) 保留候选列表 VNode，没有额外内容包装。背景仍由现有主题 token 和库 hover/disabled 状态处理。该改动属于 Vue/antdv-next 展示层，不涉及计算或跨 Worker/WASM 协议。

边框和数字格式仅复查真实 Ribbon 调用、居中布局与打开/禁用监听，未增加生产修改。[既有 popup 布局草稿](../../packages/office-studio/tests/sheet-ribbon-popup-layout.browser.test.mjs) 的预设内边距、行间距和 hover 前后真实边界断言由主线持有。SFC、脚本、样式语法、专项链接和限定空白检查已通过；草稿未运行，箭头中心、命中范围、焦点、主题切换和实际尺寸继续未验收。

## 保留缺口

- 日期验证规则、菜单参数和 `metadata.showTime` 有实际链路；Canvas 日期/时间 picker 和面板显示时间选项尚未实现。具体待确认改动见 [日期控件实施方案](SHEET-13-date-picker-entry-plan.md)。
- 列表换行仅补了真实命令草稿。单行 Input 中的多行编辑交互未验收，不能由 parser 草稿推断完整 UI 支持。
- 规则编辑、日期提交、主题、焦点恢复、输入法、权限以及 XLSX 文件闭环的本轮运行证据均未取得；既有 SHEET-13 历史通过结果不覆盖本轮差异。
- 未执行新大样本测量，不声称本轮性能优化或达标。
