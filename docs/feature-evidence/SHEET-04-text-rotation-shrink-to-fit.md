# SHEET-04 文字方向与缩小适应切片

状态：开发中，SHEET-04 整体未验收。2026-11-22 静态复核；参考版本为 `E:\rust_test\univer` 的 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本批按用户要求保留到全部对齐后统一测试，下面的历史结果不覆盖当前改动。

## 参考依据

- `packages/sheets-ui/src/menu/menu.ts` 的 `TEXT_ROTATE_CHILDREN` 定义水平、左斜 45°、右斜 45°、竖排、左转 90°、右转 90°，顺序分别对应 `0`、`-45`、`45`、`v`、`-90`、`90`。
- 同文件的 shrink-to-fit 菜单项在自动换行时禁用，激活态要求 `stf=true` 且 `tb` 不为 `WRAP`。
- `packages/sheets/src/commands/commands/set-style.command.ts` 的 `SetTextRotationCommand` 将 `v` 转成 `{ a: 0, v: TRUE }`，数字角度写入 `a`；`SetShrinkToFitCommand` 写入 `stf`。
- `packages/sheets-ui/src/controllers/shortcuts/editor.shortcut.ts` 的 `EditorBreakLineShortcut` 将 Alt+Enter 接到 `BreakLineCommand`；Ctrl+Enter 提交原格。`packages/docs-ui/src/commands/commands/break-line.command.ts` 的 `doc.command.break-line` 由 Sheets 的 Docs 编辑桥处理。本地使用同一单元格草稿及独立输入规则实现这些行为。

## 本地实现

- `sheet-ribbon.ts` 按上游顺序提供六个文字方向状态、overflow/clip/wrap 和缩小字体填充控件。`sheet-layout-format-command.ts` 统一命令参数、重叠选区去重、过滤行处理、权限检查和混合状态；Studio / Facade 与无 Studio 工作台复用该逻辑，成功变更进入文档历史。
- `CellStyle` / `ExcelCellView` 使用 `rotationVertical` 保存竖排语义。兼容旧数值 `rotation=255`，导入归一为 `rotation=0, rotationVertical=true`。`office-parser` DTO 明确输出此布尔状态，`office-writer` 将其写成 SpreadsheetML `textRotation="255"`。
- `sheet-cell-text-layout.ts` 保留显式换行、空行、UTF-16 文本位置和字素，按文字度量换行；`cell-text.ts` 将整块多行文字旋转，竖排按字素及显式列断点布局。`scene.ts` 依据内容、类型、邻格阻挡和 pane clip 计算水平溢出；合并区域不向外溢出，选区和命中继续使用网格坐标。Canvas2D / WebGPU 共用同一布局和绘制命令。
- shrink-to-fit 按旋转后的整块边界计算比例，统一减小字体、行高及装饰线几何，不再用 Canvas `maxWidth` 单轴压缩文字。启用自动换行清除 shrink-to-fit；启用 shrink-to-fit 关闭 wrap，导入两者同时为真时 wrap 优先。`sheet-auto-fit.ts` 共用文字度量、换行、旋转和竖排布局；即时自适应尚不等于持久自动行高。
- `SpreadsheetCanvas.vue` 保留 `antdv-next AutoComplete` 的数据验证选项；wrap / 多行草稿通过其自定义输入 slot 使用 `Input.TextArea`，普通单行使用库默认 Input。`sheet-cell-edit-input.ts` 共用 Alt+Enter、Ctrl/Cmd+Enter、方向提交、Escape 和 IME 规则。多行粘贴只替换草稿选区，控件切换恢复光标，输入框按草稿内容增高并裁到当前 pane；公式栏仍共用同一草稿。当前未核验焦点、真实输入法和所有弹层行为。

## 当前待统一验证

本批未运行测试、类型检查、前端或 Rust/WASM 构建，未执行浏览器、真实 Office 文件或性能测量。新增 / 更新用例包括 `sheet-cell-text-layout.contract.test.mjs`、`sheet-layout-format.contract.test.mjs`、`sheet-cell-edit-input.contract.test.mjs`、`sheet-cell-layout.browser.test.mjs`、`rendering/spreadsheet-editor-layout.test.mjs` 和 `sheet-rotation-shrink.contract.test.mjs`。用例存在不表示测试通过或行为已验收。

## 早期实现的历史验证

以下是本次统一文字布局及输入桥改写前留下的记录，仅说明当时的版本。旧 `maxWidth` 和截图用例结果不能作为当前字号缩小、多行旋转、溢出和 TextArea 实现的验证。

- `node --experimental-strip-types --test packages/office-studio/tests/sheet-rotation-shrink.contract.test.mjs`：6/6 通过，覆盖菜单命令、混合旋转状态、wrap/shrink 互斥、scene 文本命令、选区高亮、竖排字素、parser DTO 归一与 model/view 往返、Canvas2D 旋转和宽度约束调用。
- `cargo test --manifest-path crates/office-parser/Cargo.toml maps_the_ooxml_vertical_text_sentinel`：1/1 通过。
- `cargo test --manifest-path crates/office-writer/Cargo.toml vertical_text_uses_the_ooxml_255_rotation_sentinel`：1/1 通过，检查输出 `styles.xml` 中的 `textRotation="255"`。
- `pnpm.cmd exec vue-tsc -b`：通过。
- `pnpm.cmd exec playwright test --config tests/browser/toolbar-parity.playwright.config.ts --grep "text rotation, vertical text and shrink-to-fit" --reporter=list`：测试用例 1/1 通过，截图比较确认旋转、竖排、缩小适应产生重绘，撤销/重做恢复对应样式。Playwright runner 在测试完成后停留于 Vite webServer 清理阶段。

## 未验收

- 当前浏览器用例尚未运行；多选、过滤 / 隐藏轴、冻结 pane、混合状态、权限失败、主题 / classic / grid / More、编辑历史和重开均须统一验证。
- 编辑输入仍是未旋转的 DOM 输入框，竖排/旋转编辑中的光标方向和输入布局未核实。单元格选择命中不依赖旋转字形坐标。
- XLSX 导入测试覆盖 OOXML 255 语义映射，写出测试检查生成的 XML 属性；未使用独立第三方 XLSX 文件验证导入、编辑、导出、重开往返。
- 新 Canvas2D 行为用例记录实际文字绘制和变换，尚未执行；WebGPU 设备级渲染、设备丢失和恢复未单独核验。没有同数据集的前后性能 / 内存测量。
- 完整 Unicode 行断、shaping、双向文字、字体加载后的受影响布局、富文本范围格式和全部水平 / 垂直分布对齐仍需补齐；`Intl.Segmenter` 不可用时的组合字素回退正在补充。旋转 / 竖排 general 对齐尚须沿上游实际绘制和 Docs skeleton 追踪，不能仅由 overflow clip 分支推断。
- SHEET-04 的主题、默认样式、完整边框、格式继承、范围主题及全部 Office 往返仍未验收。
