# SHEET-04 边框面板与应用位置（待统一验证）

参考 `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546` 的 `packages/sheets-ui/src/menu/schema.ts`、`views/border-panel/BorderPanel.tsx`、`views/border-panel/interface.ts` 以及 `packages/sheets/src/commands/commands/set-border.command.ts` 和对应单测。

`sheet-ribbon.ts` 将格式组 `r2c5` 边框接到 `SheetBorderSelector.vue`。五列位置面板含 15 项，颜色/线型由 antdv-next ColorPicker 与 Select 承载；主动作重复应用当前预设。内侧横线仅更新内部上下相邻边，内侧竖线同理；15 种预设支持逐选区应用。复合斜线使用角到边中点的线段，并复用 Canvas 场景的线型与颜色。More 菜单复用位置和线型，颜色目前为固定色板，未对齐任意颜色选择。

XLSX writer 输出标准 diagonalUp/diagonalDown 标志及共享 diagonal 样式，把连字符线型映射为 OOXML 名称。复合斜线或不同样式的双向对角线在 XLSX 单元格边框无法完全表示时产生 `LIMITED_DIAGONAL_BORDER` 导出诊断；内部快照保留独立边。解析器的对角线重导入、合并单元格、选区外围邻居的相反边、所有语言、焦点/键盘、暗色与真实文件仍待对齐或统一测试。未运行测试、未重建 WASM，本记录不改变验收状态。
