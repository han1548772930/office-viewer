# SHEET-02 Auto-Fill Evidence

Review date: 2026-11-21  
Reference revision: `dream-num/univer@6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`

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
- Fill range detection and fill preview have not been browser-verified. Auto-scroll while dragging beyond the viewport is still missing.
- Undo/redo is committed through existing workbook history, but no browser or Office file round-trip validation was performed for this operation.

## Verification

- `pnpm.cmd exec vitest run packages/office-studio/tests/sheet-fill.contract.test.mjs`: 11 tests passed.
- `pnpm.cmd --filter @office-studio/office-studio check:types`: passed.
- `git diff --check` on changed implementation and test files: passed.

SHEET-02 remains **not accepted** until the fill-options UI, additional sequence families, browser interaction, and persistence/Office round-trip cases are covered.
