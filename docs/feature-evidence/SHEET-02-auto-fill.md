# SHEET-02 Auto-Fill Evidence

Review date: 2026-11-25（代码续作；未执行本批验证）  
Reference revision: `dream-num/univer@eaf4cfa04fabdbcfc922c692f3450459f8a29546`

## Fixed-source contract

The reference exposes `COPY`, `SERIES`, `ONLY_FORMAT`, and `NO_FORMAT` in `packages/sheets/src/services/auto-fill/type.ts`. `AutoFillService` starts with `SERIES`, then uses active hooks to choose the actual default; a single ordinary numeric seed prefers `COPY`. `packages/sheets/src/commands/commands/auto-fill.command.ts` registers auto-fill, copy-down/right and shrink-to-clear; `refill.command.ts` accepts `{ type }`. `AutoFillService.fillData()` returns success after applying a prepared fill and updates the menu even when content is equivalent. `rules.ts` and `tools.ts` define numeric/date, text-number, Chinese-number/weekday and cycling names. The Chinese multi-seed branch recognizes 一/二 and 五/六 as weekdays, while 十/十一 continue as Chinese numbers.

`packages/sheets-ui/src/controllers/auto-fill-ui.controller.ts` runs the same fill command for drag and double-click. Double-click uses adjacent data to detect a downward range, prefers the left neighboring column, falls back to the right, and stops before an already occupied source cell. `selection-shape-extension.ts` supplies the handle gesture and preview contract. CF and validation resource hooks follow the same fixed-version clipboard resource contracts; the local implementation reuses its existing clipboard resource projection.

## Local implementation

- `features/sheet-workbench/sheet-auto-fill.ts` prepares one detached model candidate for upwards/downwards/left/right extension or trailing-content clear. `sheet-fill-target.ts` supplies the same geometry to command, Facade and Canvas preview; ranges cannot expand both axes or both directions at once.
- `sheet-fill-series.ts` infers each row/column lane independently: numeric least-squares continuation rounded to five decimal places, geometric continuation with at least three matching seeds, numeric date serials, the last text-number segment and padding, Chinese numbers/weekdays, English weekdays/months, Chinese months, 时辰/天干/季度 loops. Unrecognized groups repeat their seed. A single ordinary number defaults to COPY; explicit SERIES remains available.
- COPY repeats complete content/style and translates relative formulas. SERIES generates recognized values and repeats unrecognized seeds. ONLY_FORMAT replaces the complete source style and applies CF/merge resources while retaining destination content and validation. NO_FORMAT generates/copies content and validation while retaining destination styles, CF and merge structure. The planner reuses `sheet-clipboard-cell.ts` and `sheet-clipboard-resources.ts`; formula cache, rich-text, sparkline, partial/backward resource tiles and per-cell formula anchors follow those shared contracts.
- Source/target filter-hidden rows are excluded; manually hidden and outline-hidden rows remain in the fill. Permissions are checked before preparation and publication. Partial merges and merges crossing filter holes are rejected. Non-format candidates are validated atomically; stop issues reject and warning/information issues remain in the detailed result.
- `api/sheet-fill-commands.ts` registers `sheet.command.auto-fill`, `refill`, `copy-down`, `copy-right` and `auto-clear-content`. `FRange.autoFill(target, applyType?)` uses the same command; `FOffice.Enum.AUTO_FILL_APPLY_TYPE` and public command/types are exported. Both `type` and `applyType` are normalized for refill; conflicting values are rejected.
- `sheet-fill-store.ts` keeps the transient original sheet, selection, user, unit/sheet IDs, revision and history receipt in Pinia. Refill replaces the exact applied top history entry and preserves the initial undo. Equivalent mode switches update the menu without manufacturing a content edit. An initially unchanged fill succeeds and updates the selection; its later changing refill creates a new entry, so it cannot amend an unrelated earlier edit. Later content revisions release the obsolete session. The transient session is absent from saved snapshots.
- `use-sheet-auto-fill.ts` captures host/revision before executing and only publishes a current prepared view. `ExcelWorkbench.vue` uses this bridge for Canvas, Ribbon and Ctrl/Cmd+D/R. `SpreadsheetCanvas.vue` displays `SheetFillOptions.vue`, an antdv-next Button/Dropdown with four modes, SERIES capability, selected state and focus restoration. Frozen-pane/viewport placement reuses `sheetClipboardOverlayPosition`; internal/external ConfigProvider and theme tokens continue to apply.
- `detectSheetFillDownEndRow()` retains left-first/right-fallback and treats raw 0/false as occupied. Double-click uses the shared inferred default. Shrink gestures clear only trailing content, and the existing viewport-edge scroll scheduler recalculates the same preview/target.
- The old `createSheetFillPlan()` is now a compatibility adapter to this same planner. The standalone controlled workbench uses the same model planner and its existing change/history callback; it does not create a private undo stack.

## Gaps and unverified behavior

- The synchronous source-plus-destination budget is 100,000 cells. Larger fills explicitly require the still-pending chunked Worker/Rust/WASM path; no large-range support or performance target is claimed.
- The standalone controlled workbench has no refill options because its external history currently does not supply an amendment receipt. This remains an interface/behavior gap, not a final scope reduction.
- All sequence families, mixed groups, reverse/wrapped boundaries, date calendar-unit behavior and formula edge cases still require unified comparison against samples. Numeric date serial continuation does not establish full calendar filling semantics.
- True Windows IME, theme/popup focus, frozen/merged placement, edge auto-scroll, permission changes, observer reentry and memory lifetime require browser verification. Model/Facade snapshot drafts are not independent Office export/reopen evidence.

## Historical verification — 2026-11-21 only

- `pnpm.cmd exec vitest run packages/office-studio/tests/sheet-fill.contract.test.mjs`: 11 tests passed.
- `pnpm.cmd --filter @office-studio/office-studio check:types`: passed.
- `git diff --check` on changed implementation and test files: passed.

These historical results apply to the earlier implementation only. They do not validate the 2026-11-23 or 2026-11-25 changes.

## 2026-11-25 verification drafts — none executed

- `tests/sheet-auto-fill-model.contract.test.mjs`: independent lanes, numeric/geometric/reverse series, Chinese/text/name families, formula/rich-text translation, format contracts, filter/manual/outline hiding, shrink clear, merge rejection, validation, permissions and budget; partial/backward CF/DV tiles in both axes preserve per-cell formulas.
- `tests/sheet-auto-fill-api.contract.test.mjs`: Facade/Enum/canonical command results, four-mode amendment and undo/redo, snapshots, cancellation/reentry, stale revisions/receipts, initially unchanged fill and equivalent switches.
- `tests/sheet-auto-fill-ui.contract.test.mjs`: prepared publication, current selection/options, shortcuts, equivalent switches without a second content callback, standalone host and preparation/reentry failures.
- `tests/sheet-auto-fill.browser.test.mjs` plus the updated selection-input fixture/browser draft: gesture, options, shortcuts, shrink, frozen positioning, theme and focus behavior. `sheet-fill.contract.test.mjs` now checks the shared planner behavior.

Per the user's deferred verification instruction, no tests, type checks, builds, Cargo/WASM, dev server, browser or benchmark were run. SHEET-02 remains **未验收**, with the gaps above preserved.

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

按用户要求后置统一验证，本轮仅落盘及静态回读，**未启动 dev/浏览器，未执行测试、类型检查、构建、Rust/WASM 或基准**。前一节列出的 2026-11-21 历史验证不覆盖本轮修改。真实 Windows 输入法、边缘自动滚动、冻结/合并后的填充模型保真、明暗主题与焦点恢复、Office 写出/独立重开仍为未验收；当时保留的其他序列家族和填充后选项菜单缺口已在 2026-11-25 补代码，当前边界以上面的更新为准，仍未验收。
