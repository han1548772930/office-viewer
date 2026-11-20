# Univer Facade and API Inventory

Audit date: 2026-11-20

Reference repository: `D:\rust_test\univer-reference`

Reference commit: `6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`

This inventory records the public Univer Facade surface that is relevant to the
Sheets and Docs workbench, then maps it to the current project. It is an
evidence document, not a declaration that the target feature set is complete.
The fixed source commit and the source paths in this file are the authority for
the signatures. Public documentation is used as a secondary check:

- Context7: `https://reference.univer.ai/en-US`
- API stability policy: `D:\rust_test\univer-reference\docs\API_STABILITY.md`
- Facade contribution policy: `D:\rust_test\univer-reference\docs\CONTRIBUTING-FACADE.md`
- Project reference rules: `docs/UNIVER_REFERENCE_RULES.md`

## Status vocabulary

- **implemented**: the project exposes the same contract and the behavior is
  connected to the model and runtime.
- **partial**: a related model, UI action, or service exists, but the public
  contract, semantics, or lifecycle is different or incomplete.
- **planned**: the capability is a stated target, but no connected project
  implementation was found.
- **not available**: no project counterpart was found for the referenced
  public surface.

An internal function, a Pinia store property, a component emit, or a source
file export is not counted as a stable Facade API unless a documented public
entry point exposes it and the operation is connected to the model.

## Public entry points in the reference

The fixed commit exports these core Facade symbols from
`packages/core/src/facade/index.ts`:

- `FUniver`, `FEventName`, `IEventParamConfig`, `FBlob`, `FBase`, and
  `FBaseInitialable`.

Sheets exports `FWorkbook`, `FWorksheet`, `FRange`, `FSelection`, the Sheets
event names, and the Sheets mixin from `packages/sheets/src/facade/index.ts`.
Docs exports `FDocument`, `FDocumentParagraph`, `FDocumentTextRange`, and
`FDocumentSection` from `packages/docs/src/facade/index.ts`.

The project exposes the public barrel at
`packages/office-studio/src/index.ts`. `FUniver.newAPI(studio)` creates the
root facade used by `examples/workbench/src/runtime.ts`; the Vue and Pinia
stores remain implementation details behind that boundary.

The public contract is covered by
`packages/office-studio/tests/office-api.contract.test.mjs`, which is run by
`pnpm test:api`.

## Core facade mapping

| Univer surface | Fixed source evidence | Project counterpart | Status |
| --- | --- | --- | --- |
| `FUniver.newAPI(wrapped)` | `packages/core/src/facade/f-univer.ts`, `FUniver.newAPI` | `FUniver.newAPI(studio)` from `packages/office-studio/src/index.ts` | implemented |
| `disposeUnit(unitId)` | `f-univer.ts`, `FUniver.disposeUnit` | `FUniver.disposeUnit` delegates to the unit registry and emits disposal events | implemented |
| `getCurrentLifecycleStage()` | `f-univer.ts`, `FUniver.getCurrentLifecycleStage` | `FUniver.getCurrentLifecycleStage()` returns `ready` or `disposed` | implemented |
| `undo()` / `redo()` | `f-univer.ts`, async `Promise<boolean>` methods | `FWorkbook.undo/redo` and `FDocument.undo/redo` use the per-unit history store; root facade exposes the command/history events | partial |
| `executeCommand(id, params, options)` and `syncExecuteCommand` | `f-univer.ts`, typed command id/params/options | `FUniver.executeCommand` and `executeCommandSync` use the registered command map; options beyond cancellation are not implemented | partial |
| `Event` and `addEvent(event, callback)` | `f-univer.ts`; `FEventName` and `IEventParamConfig` in `packages/core/src/facade/f-event.ts`; returns `IDisposable` | `FUniver.Event`, `FUniver.addEvent`, `EventHub`, and `Disposable` are exported from the package barrel | implemented |
| `DocCreated`, `DocDisposed`, `LifeCycleChanged` | `f-event.ts`, `IEventParamConfig` | `OfficeStudio` emits the typed lifecycle and document events through the public event hub | implemented |
| `BeforeCommandExecute`, `CommandExecuted`, `BeforeUndo`, `Undo`, `BeforeRedo`, `Redo` | `f-event.ts`, `ICommandEvent` and cancelable before events | `OfficeStudio` emits these events and honors `cancel` before command, undo, and redo mutations | implemented |
| theme, locale, region, direction getters/setters | `f-univer.ts`, `setTheme`, `getCurrentTheme`, `setLocale`, `setRegion`, `setDirection`, and related getters | No equivalent application facade contract found | not available |
| `FBlob` byte/string/content-type methods | `packages/core/src/facade/f-blob.ts` | `FBlob` is exported from the public barrel and implements byte/string/content-type operations | implemented |

## Sheets facade mapping

| Univer surface | Fixed source evidence | Project counterpart | Status |
| --- | --- | --- | --- |
| `createWorkbook(data, options)` | `packages/sheets/src/facade/f-univer.ts`, `IFUniverSheetsMixin` | `FUniver.createWorkbook(data, { makeCurrent })` creates and registers a model | implemented |
| `getActiveWorkbook()`, `getWorkbook(id)` | `f-univer.ts` | `FUniver.getActiveWorkbook` and `getWorkbook` return `FWorkbook` wrappers | implemented |
| `getSheetCommandTarget`, `getActiveSheet`, `setFreezeSync` | `f-univer.ts` | `FUniver.getSheetCommandTarget/getActiveSheet` are implemented; the synchronous freeze helper is not | partial |
| workbook identity, name, `save()` | `packages/sheets/src/facade/f-workbook.ts`, `getId`, `getName`, `setName`, `save` | `FWorkbook.getId/getName/setName/save` return detached model snapshots and mark the unit saved | implemented |
| sheet collection and lifecycle: `getSheets`, `insertSheet`, `deleteSheet`, `setActiveSheet` | `f-workbook.ts` | `FWorkbook` implements all four operations with model mutations and history entries | implemented |
| workbook `undo`, `redo` | `f-workbook.ts` | `FWorkbook.undo/redo` route through the per-unit history store | implemented |
| workbook command/selection hooks | `f-workbook.ts`, `onBeforeCommandExecute`, `onCommandExecuted`, `onSelectionChange`, each returning `IDisposable` | Global `FUniver.addEvent` provides command and selection subscriptions; workbook-specific convenience methods are not exposed | partial |
| sheet duplication/movement, defined names, calculation settings, permissions | `f-workbook.ts` | No connected project service or facade found | planned |
| `FWorksheet.getRange(...)` overloads | `packages/sheets/src/facade/f-worksheet.ts`, `getRange` by row/column, A1 text, or `IRange` | `FWorksheet.getRange` supports row/column, A1, and coordinate object overloads | implemented |
| row/column insert, delete, move, hide, show, size | `f-worksheet.ts`, `insertRows`, `deleteRows`, `moveRows`, `hideRows`, `showRows`, `setRowHeight`, and column counterparts | `ExcelWorkbench.vue` has direct insert/delete row and column actions; model stores widths/heights and optional hidden arrays; there is no worksheet command/facade layer | partial |
| merge, freeze, gridlines, tab color, hidden state, naming | `f-worksheet.ts`, `setFreeze`, `setFrozenRows/Columns`, `setHiddenGridlines`, `setTabColor`, `hideSheet`, `showSheet`, `setName` | `FWorksheet` implements merge-related range operations, freeze rows/columns, gridline visibility, and naming; tab color and hide/show helpers are not exposed | partial |
| clear/data range/append/count/metadata | `f-worksheet.ts`, `clear`, `getDataRange`, `appendRow`, `setRowCount`, `setColumnCount`, metadata methods | `FRange.clear` and worksheet row/column counts are exposed; data-region, append, and metadata helpers remain incomplete | partial |

## Range and selection mapping

| Univer surface | Fixed source evidence | Project counterpart | Status |
| --- | --- | --- | --- |
| range coordinates and A1 notation | `packages/sheets/src/facade/f-range.ts`, `getRow`, `getLastRow`, `getColumn`, `getLastColumn`, `getA1Notation` | `FRange` implements zero-based coordinates and A1 conversion | implemented |
| values, raw values, display values, cell data | `f-range.ts`, `getValue(s)`, `getRawValue(s)`, `getDisplayValue(s)`, `getCellData(s)` | `FRange` exposes all listed reads over sparse `SpreadsheetCell` data | implemented |
| formulas | `f-range.ts`, `getFormula(s)`, `setFormula(s)` | `FRange` preserves and reads/writes formula text; calculation evaluation is outside this contract | partial |
| styles, alignment, wrap, font, background, borders | `f-range.ts`, style getters/setters and fluent methods | `FRange` exposes font, fill, alignment, and wrap setters; border and mixed-state semantics remain incomplete | partial |
| `setValue(s)`, `setFormula(s)`, style setters | `f-range.ts`, mutating methods return `FRange`/`this` | `FRange` implements fluent value, formula, and style mutations through commands and history | implemented |
| merge/unmerge, clear, insert/delete cells | `f-range.ts`, `merge`, `clear`, `insertCells`, `deleteCells` | `FRange.merge` and `clear` are implemented; unmerge and shift-cell operations are not | partial |
| `activate`, `activateAsCurrentCell`, `getDataRegion`, `autoFill` | `f-range.ts` | `FRange.activate` is implemented; current-cell, data-region, and auto-fill methods are not | partial |
| `FSelection` active range/list/current cell | `packages/sheets/src/facade/f-selection.ts` | `FWorksheet.getSelection` exposes active range, current cell, and active range list | implemented |

## Sheets event mapping

The fixed commit extends core `FEventName` in
`packages/sheets/src/facade/f-event.ts`. The event parameter map includes:

- workbook lifecycle: `WorkbookCreated`, `WorkbookDisposed`;
- sheet lifecycle: `BeforeSheetCreate`, `SheetCreated`,
  `BeforeActiveSheetChange`, `ActiveSheetChanged`, `BeforeSheetDelete`,
  `SheetDeleted`, `BeforeSheetMove`, `SheetMoved`,
  `BeforeSheetNameChange`, `SheetNameChanged`,
  `BeforeSheetTabColorChange`, `SheetTabColorChanged`,
  `BeforeSheetHideChange`, `SheetHideChanged`;
- gridline and data changes: `BeforeGridlineEnableChange`,
  `BeforeGridlineColorChange`, `GridlineChanged`, `SheetValueChanged`.

The project implements the common lifecycle, command, undo/redo, selection,
sheet-name, and value-change events in `packages/office-studio/src/api/events.ts`.
Before command, undo, redo, and sheet-name events carry `cancel`, and
`FUniver.addEvent` returns a disposable subscription. The full Sheets event
catalog (sheet move/hide/tab-color and all plugin-specific payloads) is still
not available; status for the complete catalog is **partial**.

## Docs facade mapping

| Univer surface | Fixed source evidence | Project counterpart | Status |
| --- | --- | --- | --- |
| `FDocument` identity, name, `save`, `undo`, `redo` | `packages/docs/src/facade/f-document.ts`, methods at `getId`, `getName`, `setName`, `save`, `undo`, `redo` | `FDocument` implements identity, naming, detached save snapshots, and per-unit undo/redo | implemented |
| document text and block operations | `f-document.ts`, `insertText`, `getTextRange`, `insertParagraph`, `appendParagraph`, `deleteRange` | `FDocument` implements text reads, insertion, replacement, append paragraph, and text-range access; advanced block operations remain incomplete | partial |
| `FDocumentParagraph` | `packages/docs/src/facade/f-document-paragraph.ts`, text/style/find/task methods | `FDocumentParagraph` supports stable identity/info and text reads, text-range replacement, append text, character-style reads/writes, and paragraph-style reads/writes; find/task and the remaining block-style contract are incomplete | partial |
| `FDocumentTextRange` | `packages/docs/src/facade/f-document-text-range.ts`, `getText`, style-run reads, `setText`, `setTextStyle` | `FDocumentTextRange` implements half-open offset reads, clipped style-run reads, common-style reads, replacement, and character-style patches through the document history mutation path; full reference range coverage remains incomplete | partial |
| `FDocumentSection` and page/header/footer methods | `packages/docs/src/facade/f-document-section.ts` and `f-document.ts` section methods | `DocumentModel.page` stores one page setup; no section collection, section breaks, or header/footer facade | planned |
| Docs facade events | Core `FEventName` plus Docs extensions in the reference facade packages | Core lifecycle/command events are available through `FUniver`; the full Docs-specific event catalog is not | partial |

## Import/export mapping

| Univer/public concept | Fixed source evidence | Project counterpart | Status |
| --- | --- | --- | --- |
| `FBlob` as a reusable byte/string stream facade | `packages/core/src/facade/f-blob.ts` | `FBlob` is exported and implements byte/string/content-type operations | implemented |
| workbook/document snapshot save | `FWorkbook.save()` and `FDocument.save()` in their respective facade files | `FWorkbook.save` and `FDocument.save` return detached snapshots and mark the corresponding unit saved | implemented |
| Office import and export API | Public Univer Facade class list and import/export guides; exact IO implementation depends on registered plugins | `useEditorStore.importFile` delegates to `parserStore`; `exportOfficeModel` uses `src/parser/export.worker.ts` and returns XLSX/DOCX bytes | partial |
| export format enumeration, facade events, cancellation/disposable contract | No matching public project entry point found; reference plugin APIs are separate from `FBlob` | Export has worker progress and timeout internally, but no format enum, add-event API, or exported cancellation handle | not available |

## Explicit gaps and implementation order

The audit now has a connected public Facade for the core workbook/document
editing path. The remaining contract gaps are:

1. Extend the Sheets facade with row/column operations, data regions, auto-fill,
   tab colors, hidden sheets, defined names, calculation settings, permissions,
   and the remaining plugin-specific events.
2. Complete formula calculation and range style semantics, including borders,
   mixed values, unmerge, and shift-cell operations.
3. Extend Docs with sections, headers/footers, full rich-text style coverage,
   tables, page layout operations, and the Docs-specific event catalog.
4. Connect the public save/export facade to the full format enumeration and
   cancellation handles used by the parser/writer workers.

These gaps align with the still-unaccepted COMMON, SHEET, and DOC rows in
`docs/UNIVER_FEATURE_PARITY.md`. The Docs text/style subset is covered by
`packages/office-studio/tests/document-facade.contract.test.mjs` (10 tests),
and the public unit wiring is covered by
`packages/office-studio/tests/office-api.contract.test.mjs` (3 tests). These
tests support the **partial** rows above; they do not replace the browser and
file-sample acceptance evidence required before any DOC row is marked
**implemented**. Pro and plugin rows remain unaccepted until their required
evidence exists.

