# SHEET-04 单元格颜色重置（待统一验证）

2026-11-22，固定源码 `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`。`packages/sheets-ui/src/menu/schema.ts` 将两个重置入口分别放在文字色和背景色菜单下；`packages/sheets/src/commands/commands/set-style.command.ts` 的 `ResetTextColorCommand` / `ResetBackgroundColorCommand` 只将 `cl.rgb` 或 `bg.rgb` 设为 null。

项目使用相同 command ID。工作台与 `FRange.setFontColor(null)`、`FRange.setBackground(null)` 在带 Studio 的情况下共用按工作簿隔离的命令、权限及历史。只修改选区内已有单元格，不因重置而新增空单元格。`CellStyle.fill` 和 `FontStyle.color` 可为 null；视图的 undefined 表示未覆盖，null 表示清除。显示层仍提供回退颜色；Facade 的 getBackground 保持返回显示色。旧快照中已有的白色保留为显式白色，不猜测它原先是用户指定还是旧默认值。

`crates/office-writer/src/xlsx.rs` 的单元格样式将无填充映射到 none，显式白色保留为 solid；未指定字体色不写固定 RGB。未重建 WASM。待统一执行 `sheet-color-reset.contract.test.mjs`、writer 专项、浏览器操作及 XLSX 导出/重导入测试；真实 Excel 重开、主题色/tint 与局部富文本颜色仍未验收。
