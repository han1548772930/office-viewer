# SHEET-04 文字方向与缩小适应切片

状态：部分实现，SHEET-04 整体未验收。参考版本为 `E:\rust_test\univer` 的 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

## 参考依据

- `packages/sheets-ui/src/menu/menu.ts` 的 `TEXT_ROTATE_CHILDREN` 定义水平、左斜 45°、右斜 45°、竖排、左转 90°、右转 90°，顺序分别对应 `0`、`-45`、`45`、`v`、`-90`、`90`。
- 同文件的 shrink-to-fit 菜单项在自动换行时禁用，激活态要求 `stf=true` 且 `tb` 不为 `WRAP`。
- `packages/sheets/src/commands/commands/set-style.command.ts` 的 `SetTextRotationCommand` 将 `v` 转成 `{ a: 0, v: TRUE }`，数字角度写入 `a`；`SetShrinkToFitCommand` 写入 `stf`。

## 本地实现

- `sheet-ribbon.ts` 按上游顺序提供六个文字方向状态和缩小字体填充控件。`ExcelWorkbench.vue` 把数值角度与 `v` 命令接入选区变更和单次历史记录，并由混合选区状态计算激活项。
- `CellStyle` / `ExcelCellView` 使用 `rotationVertical` 保存竖排语义。兼容旧数值 `rotation=255`，导入归一为 `rotation=0, rotationVertical=true`。`office-parser` DTO 明确输出此布尔状态，`office-writer` 将其写成 SpreadsheetML `textRotation="255"`。
- `scene.ts` 对角度模式发出带旋转的文本绘制命令。竖排模式按 Unicode 字素逐字纵向布局。两种模式都按单元格 clip 裁切，选区和命中几何继续使用网格单元格坐标。
- shrink-to-fit 先测量完整文本宽度，只在超过单元格可用宽度时设置 `maxWidth`。Canvas2D 使用原生文本宽度约束，WebGPU 文本纹理光栅化使用同一约束并重算对齐锚点。启用自动换行会清除 shrink-to-fit，启用 shrink-to-fit 会关闭自动换行；导入文件两者同时为真时 wrap 优先。

## 验证

- `node --experimental-strip-types --test packages/office-studio/tests/sheet-rotation-shrink.contract.test.mjs`：6/6 通过，覆盖菜单命令、混合旋转状态、wrap/shrink 互斥、scene 文本命令、选区高亮、竖排字素、parser DTO 归一与 model/view 往返、Canvas2D 旋转和宽度约束调用。
- `cargo test --manifest-path crates/office-parser/Cargo.toml maps_the_ooxml_vertical_text_sentinel`：1/1 通过。
- `cargo test --manifest-path crates/office-writer/Cargo.toml vertical_text_uses_the_ooxml_255_rotation_sentinel`：1/1 通过，检查输出 `styles.xml` 中的 `textRotation="255"`。
- `pnpm.cmd exec vue-tsc -b`：通过。
- `pnpm.cmd exec playwright test --config tests/browser/toolbar-parity.playwright.config.ts --grep "text rotation, vertical text and shrink-to-fit" --reporter=list`：测试用例 1/1 通过，截图比较确认旋转、竖排、缩小适应产生重绘，撤销/重做恢复对应样式。Playwright runner 在测试完成后停留于 Vite webServer 清理阶段。

## 未验收

- 浏览器闭环使用单元格选择；混合范围的状态只覆盖契约测试，范围内所有单元格的浏览器级变更与历史尚未覆盖。
- 编辑输入仍是未旋转的 DOM 输入框，竖排/旋转编辑中的光标方向和输入布局未核实。单元格选择命中不依赖旋转字形坐标。
- XLSX 导入测试覆盖 OOXML 255 语义映射，写出测试检查生成的 XML 属性；未使用独立第三方 XLSX 文件验证导入、编辑、导出、重开往返。
- Canvas2D 有命令级旋转/宽度约束测试。Playwright 截图未记录实际后端；WebGPU 设备级像素输出未单独核验。
- SHEET-04 的主题、默认样式、完整边框、格式继承、范围主题及全部 Office 往返仍未验收。
