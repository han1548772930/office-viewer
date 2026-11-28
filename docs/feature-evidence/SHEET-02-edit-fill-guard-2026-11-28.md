# SHEET-02 编辑填充后的完整网格与鼠标命中

日期：2026-11-28。状态：已修改，未验收。用户截图与进一步说明确认：填充值正常，但网格线与右侧、下方空白行列一起消失，显示范围恰好截至填充区域。用户同时要求保留输入中的右下角方块，以及同一次鼠标按下提交输入后继续拖动的交互。

## 维度根因与固定来源

固定版本为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。[sheet-snapshot-utils.ts](../../../univer/packages/core/src/sheets/sheet-snapshot-utils.ts) 中 `DEFAULT_WORKSHEET_ROW_COUNT` / `DEFAULT_WORKSHEET_COLUMN_COUNT` 为 1000 / 20，`mergeWorksheetSnapshotWithDefault` 物化缺省尺寸；[selection-shape-extension.ts](../../../univer/packages/sheets-ui/src/services/selection/selection-shape-extension.ts) 的填充拖动基于现有工作表和选区。填充范围不是工作表的完整尺寸。

本地旧 `createBlankSheet()` 未设置 `rowCount` / `columnCount`，显示适配层另以 100×26 缺省展示。[sheet-auto-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fill.ts) 填充时却按 `Math.max(sheet.rowCount ?? 1, target.endRow + 1)` 和列的相同公式物化尺寸。因此省略维度的工作表从 C5 填至 C17 后，会持久化成 17×3；下一次视图与布局按这个显式尺寸裁剪，正好截断右侧和下方空白网格。这是静态可追踪的维度错误，不能描述成目标外内容被自动清空。

主线在 [domain/model.ts](../../packages/office-studio/src/domain/model.ts) 导出 `DEFAULT_SHEET_ROW_COUNT=1000` / `DEFAULT_SHEET_COLUMN_COUNT=20`，新建、公开快照、解析归一化和视图投影共用缺省值。Sheets 侧修改 [sheet-auto-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-auto-fill.ts) 与 [sheet-fill.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-fill.ts)，仅在原维度缺省时使用这些常量，再与目标坐标取最大值。显式指定的小尺寸继续保持。

## 输入柄的真实鼠标命中

[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 原可视 `canvas-cell-fill-handle` 设置 `pointer-events:none`，小方块自身不能成为鼠标目标，完全依赖穿过 Ant 输入层后的几何命中。本批让现有编辑器内核方块接收 pointer、使用 crosshair 和 `touch-action:none`，直接复用 `onEditorFillPointerDown`。实际目标为方块时，来源取当前编辑单元格或合并格；输入层转发的事件继续使用原几何命中。

同一次按下先 `preventDefault/stopPropagation`，再通过 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 `prepare-fill` 和原单元格命令提交输入，收到正向确认后继续同一 pointer 的拖动。组字、验证拒绝、事件改选、指针丢失和内容版本变化仍保留既有保护。没有隐藏或删除编辑柄；正常四方向填充、向内清除、四种选项、权限和内容历史仍沿用原实现。

用户随后确认“可以命中，但不能向下拖动”。主线查到第二个确定的误取消条件：[sheet-fill-gesture.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-fill-gesture.ts) 严格比较 `rangeType`，而 UI 的普通选区可以省略该属性，提交后 API 选区投影为显式 0。相同范围在提交后的 watcher 和 `canResumeSheetFillGesture` 中被当作切换，提前清除手势。固定 [rectangle.ts](../../../univer/packages/core/src/shared/rectangle.ts) 的 `Rectangle.equals` 明确将缺省与 `RANGE_TYPE.NORMAL` 视为等价，[typedef.ts](../../../univer/packages/core/src/sheets/typedef.ts) 的 NORMAL 为 0。现按 `??0` 比较，保留行/列/全表类型之间的区别，以及真正选择、版本和输入会话变化的取消。

Canvas 另在处理 `lostpointercapture` 前核对真实事件目标，子输入或控件冒泡的捕获丢失不再取消根宿主持有的手势。确认输入成功且上下文合法后，在稳定根宿主上检查并必要时重新建立 pointer capture，再转为 `fillDrag`；真实 pointerup/cancel 和根捕获失效仍结束手势。

方块属于 Canvas 编辑器命中内核，使用原有宿主 span；输入继续使用 `antdv-next` 的 AutoComplete / Input.TextArea 与统一主题。本批没有新增 Ribbon 常驻、扩展、上下文、溢出根或子菜单，也没有持久字段、公开 Facade 或 Worker/WASM 协议变化。

## 草稿与验证边界

- [sheet-grid-extents.contract.test.mjs](../../packages/office-studio/tests/sheet-grid-extents.contract.test.mjs)：主线覆盖缺省、公开/领域/解析输入、旧无维度模型、填充后的 view/scene 空白网格、显式小尺寸和内部快照重开。
- [sheet-selection-input.browser.test.mjs](../../packages/office-studio/tests/sheet-selection-input.browser.test.mjs)：新增省略维度、C5 输入并同指针填到 C17，检查 1000×20、Canvas CSS/backing/viewport 尺寸、两条独立撤销和重做、F21 空白格可选择及内部快照重开；另对普通、多行和富文本输入执行真实 `elementFromPoint` 与鼠标按下/移动/抬起，检查方块可命中、根宿主持有 capture，以及子元素冒泡捕获丢失后仍能同指针续拖。既有组字与校验拒绝草稿继续保留。
- [sheet-fill-gesture.contract.test.mjs](../../packages/office-studio/tests/sheet-fill-gesture.contract.test.mjs)：主线补普通选区类型省略/显式 0 等价与行/列/全表区别的草稿。
- [sheet-auto-fill-model.contract.test.mjs](../../packages/office-studio/tests/sheet-auto-fill-model.contract.test.mjs)：新增沿两轴填充时目标外单元格、元数据、原模型与显式行列规模保持不变。

草稿均未运行。只读 SFC / TypeScript 语法解析、草稿 `node --check` 和限定 `git diff --check` 不构成验收。未执行测试、类型检查、前端或 Rust/WASM 构建、dev、浏览器、Office 二进制往返和性能基准；真实输入法、两渲染后端、用户场景与最终闭环仍未验收。内部 JSON 快照保存/重开不等于 XLSX 导出或独立 Office 重开。
