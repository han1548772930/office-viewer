# SHEET-14 快捷条件格式与管理菜单

更新日期：2026-11-26。状态：开发中、未验收。固定源码为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地只读目录 `E:\rust_test\univer`。

本批复核已提交的快捷数据条、图标集和类型限定清除，并补齐管理菜单的十个入口及清除整表命令。测试、类型检查、构建、浏览器、Rust/WASM 和基准按用户要求后置；这里记录静态实现，不宣称产品能力已经验收。

## 固定源码与入口清点

| 源码 | 核对内容 |
| --- | --- |
| [menu/schema.ts][schema] | INSERT/CELL 数据条 order 3、图标集 order 4；DATA/RULES 管理根 order 1、grid r2c1 |
| [menu/quick-insert.menu.ts][quick] | 两个 BUTTON_SELECTOR、固定参数、3/5 个子项及工作簿/工作表/范围权限 |
| [menu/manage-rule.ts][manage] | 十项 SELECTOR；区域清除判断交叠规则的完整旧范围，整表清除判断全部旧范围 |
| [commands/operations/open-conditional-formatting-panel.ts][panel] | 六类新建、普通新建、列表管理及两类清除的路由 |
| [commands/commands/add-data-bar-cf.command.ts][bar]、[add-icon-set-cf.command.ts][icon] | 创建规则、目标范围及默认 stopIfTrue |
| [clear-range-cf.command.ts][clear-range]、[clear-worksheet-cf.command.ts][clear-sheet] | 类型限定差集、整表删除及单次内容历史 |

静态声明口径为 3 个常驻根、18 个子项：快捷数据条 3，快捷图标集 5，管理菜单 10。它们全部来自 OSS；本批没有新增 Pro 根、上下文根或独立溢出根。Classic/grid 和 More 的同一入口不重复计数。菜单数量只说明清点覆盖度。

| 根 | 本地入口 | Classic / grid | 子项顺序 |
| --- | --- | --- | --- |
| 快捷数据条 | `sheet.command.add-data-bar-conditional-rule` | INSERT/CELL order 3；随既有单元格工具组自动排入第二行 | 应用固定数据条、新建数据条规则、清除区域的数据条规则 |
| 快捷图标集 | `sheet.command.add-icon-set-conditional-rule` | INSERT/CELL order 4；随既有单元格工具组自动排入第二行 | 等级、形状、方向、新建图标规则、清除区域的图标规则 |
| 管理菜单 | `sheet.operation.open-conditional-formatting` | DATA/RULES order 1，grid r2c1 | 突出显示、排名和平均值、公式、色阶、数据条、图标集、管理、新建、清除区域、清除整表 |

根在没有工作表或当前格式权限不足时禁用。区域清除还需存在交叠规则且拥有其全部旧范围权限；整表清除还需存在规则且拥有所有规则的旧范围权限。Quick 根的类型限定清除单独判断相应类型，不因其他类型的受保护规则而禁用。

## 规则、命令和持久化

快捷数据条使用 min/max、正值 `#ffbe38`、负值 `#abd91a`、实色和显示数值。快捷图标选择 `3Stars`、`3TrafficLights1`、`3Arrows`，前两个门槛为大于或等于 67%、33%，其余值使用第三个图标。上游第三项的 MAX_SAFE_INTEGER/percent 是兜底哨兵，本地转为 num 哨兵，保留相同兜底行为；不把它当作可编辑的超界百分比。

`quickConditionalFormatRules` 生成完整领域规则或差集，不修改单元格底色和值。清除部分区域保留规则 ID、剩余范围及配置，复用 `reanchorConditionalFormatRule` 调整公式基准。整表清除使用固定命令 `sheet.command.clear-worksheet-conditional-rule`，允许不提供参数、没有当前选区，支持显式 `unitId/subUnitId`；空规则表返回 no-op。

API 与工作台继续使用 `planSheetConditionalFormatsCommand`、`commitWorkbook` 和既有文档历史。提交判断改变规则的旧范围和新范围，失败不发布部分列表；无变化不增加内容版本或历史。撤销、重做、内部快照保存/重开保留规则、顺序、来源及单元格原值/缓存。独立工作台复用同一 planner 和权限检查，通过原有模型更新事件发布。

面板 `initialType` 扩展为 `highlightCell/top10/expression/colorScale/dataBar/iconSet/create`。前六项直接进入相应编辑器，`create` 进入默认新建，未传值进入管理列表。公式新建保留 `=` 输入起点，须填写有效条件后提交；打开或取消草稿不产生内容历史。所有新建与列表编辑继续提交同一 `sheet.command.set-conditional-formats`。

本地管理 operation 名称保留现有公开入口，与上游 `sheet.operation.open.conditional.formatting.panel` 不同；菜单参数使用可读类型字符串，对应上游 `CF_MENU_OPERATION` 的六类和新建/管理动作。清除仍使用上游固定命令 ID。新增的编辑意图和菜单状态是 UI 状态，不进入文档快照。

## 修改文件与测试草稿

| 文件 | 本批实际变化 |
| --- | --- |
| `features/conditional-formatting/quick-conditional-formatting.ts` | 整表清除、六类面板意图、区域/整表清除状态 |
| `api/sheet-conditional-formatting-commands.ts` | 注册固定整表清除命令，支持无选区和显式目标，复用完整规则事务 |
| `components/spreadsheet/sheet-ribbon.ts` | 十项管理菜单、真实命令参数及禁用状态 |
| `components/spreadsheet/ConditionalFormattingPanel.vue` | 六类/普通新建入口，复用现有 Ant 编辑表单 |
| `components/ExcelWorkbench.vue` | 主任务接线面板意图和整表清除；文件由主任务负责 |

新增未运行草稿为 7 个声明：

| 文件 | 本批新增 | 场景 |
| --- | ---: | --- |
| `sheet-conditional-formatting-commands.contract.test.mjs` | 4 | 无选区整表清除、非可视规则/旧范围权限状态、真实 API 历史与重开、受保护/非活动目标隔离 |
| `sheet-ui-ribbon.contract.test.mjs` | 1 | 十项顺序、参数、grid 位置和清除状态 |
| `sheet-format-painter.browser.test.mjs` | 2 | 七种新建意图取消不修改内容、部分清除/整表清除和两次撤销 |

既有快捷规则草稿已更新新增状态字段；既有管理面板浏览器入口改为先打开菜单，再选择管理。没有执行这些草稿，不以断言数量作为功能验收。

## UI、计算域和剩余缺口

本批沿用 `antdv-next` 的 Ribbon 菜单、Drawer、Form/Form.Item、Input、Select、InputNumber、ColorPicker 和 Checkbox；主题继续由统一 ConfigProvider/token 管理。没有新建原生输入、独立内容历史或第二套 Canvas 规则引擎。这里只新增轻量规则规划与命令路由，无 Worker/WASM 协议变更；清除仍受现有范围数量/计算预算约束。

真实浏览器下的菜单/More、两种布局、明暗主题、禁用态、焦点恢复和 Windows IME 尚未运行验证。复杂条件计算、Office 原生往返、真实大样本性能仍按 [条件格式 Office 专项](SHEET-14-conditional-format-office-roundtrip.md) 和 [规则/格式刷证据](SHEET-04-format-painter.md) 保留未验收状态，本批没有关闭 SHEET-14 的完整目标。

[schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/schema.ts
[quick]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/quick-insert.menu.ts
[manage]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/manage-rule.ts
[panel]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/commands/operations/open-conditional-formatting-panel.ts
[bar]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/commands/commands/add-data-bar-cf.command.ts
[icon]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/commands/commands/add-icon-set-cf.command.ts
[clear-range]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/commands/commands/clear-range-cf.command.ts
[clear-sheet]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/commands/commands/clear-worksheet-cf.command.ts
