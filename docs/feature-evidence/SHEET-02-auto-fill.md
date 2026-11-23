# SHEET-02 Auto-Fill Evidence

Review date: 2026-11-21  
Reference revision: `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`

## Fixed-source contract

The reference exposes `COPY`, `SERIES`, `ONLY_FORMAT`, and `NO_FORMAT` in `packages/sheets/src/services/auto-fill/type.ts`. `AutoFillService` defaults to `SERIES` and exposes these four choices in its apply menu. Fill rules in `packages/sheets/src/services/auto-fill/rules.ts` infer number/date and other recognized sequences, while formula references translate with the destination. `packages/sheets-ui/src/controllers/auto-fill-ui.controller.ts` runs the same fill command for drag and double-click. Double-click uses adjacent data to detect a downward range, prefers the left neighboring column, falls back to the right, and stops before an already occupied source cell.

## Local implementation

- `packages/office-studio/src/features/sheet-workbench/sheet-fill.ts` now defines the four apply modes. `SERIES` extends a regular one-dimensional numeric pattern and numeric date serials, translates formulas, and falls back to repeating the source pattern when numeric continuation cannot be inferred. `COPY` repeats values/formulas and translates formulas. `ONLY_FORMAT` overlays source formatting while retaining destination values/formulas. `NO_FORMAT` copies source values/formulas and retains destination formatting.
- `detectSheetFillDownEndRow()` implements downward adjacent-data detection with left-first/right-fallback selection and stops before occupied cells in the fill columns.
- `SpreadsheetCanvas.vue` recognizes double-click on the selection fill handle and emits a `SERIES` fill for the detected range. Pointer-drag also emits a fill event.
- `ExcelWorkbench.vue` snapshots the target cell grid for the plan, validates non-format fills against data validation, checks protected ranges, applies the plan to the workbook view, recalculates formulas, and commits through the existing workbook history.
- Contract tests cover numeric and date serial continuation, irregular pattern repetition, formula translation, format-only/value-only behavior, double-click range detection, drag preview and event wiring.

## Gaps and unverified behavior

- There is no visible post-fill options menu yet, so users cannot choose COPY, ONLY_FORMAT, or NO_FORMAT from the canvas. The planner accepts these modes for the shared command path, but only SERIES is currently emitted by the canvas.
- Sequence detection currently covers regular numeric values and numeric date serials only. Weekday/month names, textual numeric suffixes, Chinese numeral sequences, date increments by calendar unit, and formula series rules remain unimplemented.
- Fill range detection and fill preview have not been browser-verified. Viewport-edge auto-scroll was added in the 2026-11-23 continuation below, and remains unverified.
- Undo/redo is committed through existing workbook history, but no browser or Office file round-trip validation was performed for this operation.

## Verification

- `pnpm.cmd exec vitest run packages/office-studio/tests/sheet-fill.contract.test.mjs`: 11 tests passed.
- `pnpm.cmd --filter @office-studio/office-studio check:types`: passed.
- `git diff --check` on changed implementation and test files: passed.

SHEET-02 remains **not accepted** until the fill-options UI, additional sequence families, browser interaction, and persistence/Office round-trip cases are covered.

## 2026-11-23 编辑态填充柄续作（未执行验证）

固定 v1.6.0 参考仍为上述提交。再次核对 `packages/sheets-ui/src/services/selection/selection-shape-extension.ts` 的 `_initialAutoFill()`、`_autoFillForPointerdown()`、`_autoFillForPointermove()` 和 `_autoFillMoving()`：填充柄独立命中，进入时显示 crosshair，拖动保留起始范围，源/目标坐标共用缩放和滚动变换，释放后提交，滚动与权限检查由手势和命令链共同处理。本次不推断上游编辑输入框与填充柄衔接的内部行为；本项目编辑态本来就画有填充方块，下面修复的是该可见入口无法操作的问题。

### 缺陷与本次代码

- 原实现在编辑态继续显示 6px 方块，但指针命中使用 `!editingKey` 排除编辑态；方块的 `pointer-events: none` 还使靠内的一半落在 Ant Input 上，被普通指针处理器的 input/textarea 早退忽略。
- `rendering/spreadsheet/editor-layout.ts` 新增可见填充柄 clip 和 `hitCellEditorFillHandle()`。只命中可见方块，返回真实单格或合并源范围；不会把已滚出视口或被冻结分区裁掉的柄迁到另一单元格。
- `SpreadsheetCanvas.vue` 在根 `pointerdown.capture` 中先处理编辑柄，调用私有 `prepareFill` 回调。`ExcelWorkbench.prepareCanvasFillGesture()` 复用 `finishEdit()`，只有提交成功、未组合输入、未导航/切换宿主时才返回当前 sheet/revision 确认。原 `finish-edit` 事件继续用于输入失焦。
- `sheet-fill-gesture.ts` 独立保存选区快照，确认本次提交后的 revision、选区和已结束的输入会话；Canvas 使用原 pointer ID 穿过 Vue props 更新。失败提交、重入选区改变、新会话、sheet/revision 变化、Escape、pointercancel/lostcapture 均不能继续写目标范围。
- 正常和编辑态填充共用现有 `fill-selection → ExcelWorkbench.fillSelection → createSheetFillPlan` 数据/公式/验证/保护/内容历史路径。源编辑确认和目标填充是两个内容历史动作；取消拖拽保留已确认的源编辑。实际移动达到 3 CSS px 后才产生填充，单击方块不写邻格。
- 拖拽期间在视口边缘调用现有滚动调度，并重新计算填充目标和预览；pointerup 使用最终坐标。该接线尚无浏览器或性能测量。
- 输入仍为 `antdv-next` AutoComplete/Input/TextArea；方块是编辑器命中内核的视觉标记，不新增原生表单控件。悬浮状态可透过 Input 显示 crosshair。ExcelWorkbench 的内部 ConfigProvider 同时复用 `WORKBENCH_CONTROL_THEME`，保留外部主题 token 并统一菜单分隔线设置。

### 新草稿与边界

- `tests/sheet-fill-gesture.contract.test.mjs`：选区快照隔离、提交 revision 确认、未结束/替换输入会话、相同工作表重入导航、活动拖拽失效。
- `tests/rendering/spreadsheet-editor-layout.test.mjs`：方块内外两半、不同缩放、合并源范围、冻结裁剪、离屏和非有限坐标。
- `tests/sheet-selection-input.browser.test.mjs`：同 pointer 提交/拖填、单击不填、公式引用、撤销/重做/内部快照重开、验证/命令取消/组合输入阻止填充、观察者重入、Escape 和后续选区改变。

按用户要求后置统一验证，本轮仅落盘及静态回读，**未启动 dev/浏览器，未执行测试、类型检查、构建、Rust/WASM 或基准**。前一节列出的 2026-11-21 历史验证不覆盖本轮修改。真实 Windows 输入法、边缘自动滚动、冻结/合并后的填充模型保真、明暗主题与焦点恢复、Office 写出/独立重开仍为未验收；上述其他序列家族和填充后选项菜单缺口继续保留。
