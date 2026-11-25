import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, ref, shallowRef } from 'vue'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { useSheetCellShift } from '../src/components/spreadsheet/use-sheet-cell-shift.ts'
import { SHEET_CELL_SHIFT_COMMANDS, SHEET_CELL_SHIFT_CONFIRM_COMMANDS } from '../src/features/sheet-workbench/sheet-cell-shift-command.ts'

// Drafts only; no test runner, browser, build/type checks or benchmarks were started.
const range = (row = 1, col = 1) => ({ startRow: row, endRow: row, startCol: col, endCol: col })
function sample(merged = true) {
  const model = createBlankWorkbook(); model.id = 'cell-shift-host'; model.calculationMode = 'manual'
  model.sheets[0].id = 'data'; model.sheets[0].name = 'Data'; model.activeSheetId = 'data'
  model.sheets[0].rowCount = 30; model.sheets[0].columnCount = 12; model.sheets[0].defaultRowAutoHeight = false
  model.sheets[0].cells['1:1'] = { ...createCell(1, 1, 'move'), authoredStyle: {} }
  if (merged) model.sheets[0].mergedRanges = [{ startRow: 1, endRow: 2, startCol: 0, endCol: 2 }]
  return model
}
function host(model = sample(), studio) {
  const view = ref(toExcelWorkbookView(studio ? studio.getWorkbookModel(model.id) : model)), active = ref(true), user = ref('local-user')
  const studioRef = shallowRef(studio), ranges = ref([range(), range(5, 3)]), activeIndex = ref(0)
  const calls = { applied: 0, selected: [], changed: [], focused: 0 }, scope = effectScope()
  const context = { studio: () => studioRef.value, workbook: () => view.value, sheetId: () => view.value.activeSheetId,
    ranges: () => ranges.value, activeRangeIndex: () => activeIndex.value, userId: () => user.value, region: () => 'en-US', active: () => active.value,
    finishEdit: () => true, prepareView: value => value,
    applyView: value => { view.value = value; calls.applied++ }, selection: value => calls.selected.push(cloneModel(value)),
    changed: options => calls.changed.push(options), focus: () => calls.focused++ }
  const controller = scope.run(() => useSheetCellShift(context))
  return { view, active, user, ranges, activeIndex, studioRef, context, calls, controller, dispose: () => scope.stop() }
}

test('standalone confirmation captures the active primary range, publishes once and restores focus after closing', async () => {
  const state = host()
  try {
    const before = cloneModel(state.view.value)
    assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), false)
    assert.deepEqual(state.controller.request.value.input.range, range()); assert.equal(state.calls.applied, 0)
    assert.deepEqual(state.view.value, before); state.controller.confirmCellShift()
    assert.equal(state.view.value.sheets[0].cells.B3.value, 'move'); assert.deepEqual(state.view.value.sheets[0].merges, [])
    assert.equal(state.controller.request.value, null); assert.equal(state.calls.applied, 1); assert.equal(state.calls.changed.length, 1)
    assert.deepEqual(state.calls.selected, [range()]); assert.equal(state.calls.changed[0].current(), true)
    assert.equal(state.calls.focused, 0); state.controller.afterCellShiftClose(); await nextTick(); assert.equal(state.calls.focused, 1)
  } finally { state.dispose() }
})

test('in-place source edits preserve a pending request and show its error without replacing newer content', () => {
  const state = host()
  try {
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); const request = state.controller.request.value
    state.view.value.sheets[0].cells.A10 = { value: 'newer', rawValue: 'newer' }
    state.controller.confirmCellShift()
    assert.equal(state.controller.request.value, request); assert.match(state.controller.cellShiftError.value, /已更改/)
    assert.equal(state.view.value.sheets[0].cells.A10.value, 'newer'); assert.equal(state.view.value.sheets[0].cells.B2.value, 'move')
    assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('replacing a standalone workbook with identical IDs and JSON does not authorize a captured confirmation', () => {
  const state = host()
  try {
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); state.view.value = cloneModel(state.view.value)
    state.controller.confirmCellShift(); assert.match(state.controller.cellShiftError.value, /宿主已更改/)
    assert.equal(state.calls.applied, 0); assert.equal(state.view.value.sheets[0].cells.B2.value, 'move')
  } finally { state.dispose() }
})

test('switching the document or sheet cancels the old dialog and its late close never focuses the replacement', async () => {
  const state = host()
  try {
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); state.view.value = { ...state.view.value, id: 'replacement' }
    assert.equal(state.controller.request.value, null); state.controller.afterCellShiftClose(); await nextTick(); assert.equal(state.calls.focused, 0)
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); state.view.value.activeSheetId = 'other'
    assert.equal(state.controller.request.value, null); state.controller.afterCellShiftClose(); await nextTick(); assert.equal(state.calls.focused, 0)
  } finally { state.dispose() }
})

test('a close from the previous activation cannot focus a reactivated host with the same IDs', async () => {
  const state = host()
  try {
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); state.controller.cancelCellShift()
    state.active.value = false; state.active.value = true
    state.controller.afterCellShiftClose(); await nextTick(); assert.equal(state.calls.focused, 0)
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); state.controller.cancelCellShift()
    state.controller.afterCellShiftClose(); await nextTick(); assert.equal(state.calls.focused, 1)
  } finally { state.dispose() }
})

test('user changes cancel captured merge permission and unfinished input blocks dispatch', () => {
  const state = host()
  try {
    state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); state.user.value = 'owner'
    assert.equal(state.controller.request.value, null); state.context.finishEdit = () => false
    assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), false)
    assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('preparation that leaves and reenters the same activation cannot publish its old detached candidate', () => {
  const state = host(sample(false))
  try {
    const before = cloneModel(state.view.value)
    state.context.prepareView = value => { state.active.value = false; state.active.value = true; return value }
    assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), false)
    assert.deepEqual(state.view.value, before); assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('publication reentry preserves the observer content and suppresses stale selection and parent updates', () => {
  const state = host(sample(false))
  try {
    state.context.applyView = value => { state.view.value = value; state.view.value.sheets[0].cells.A10 = { value: 'observer' }; state.calls.applied++ }
    assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), true)
    assert.equal(state.view.value.sheets[0].cells.B3.value, 'move'); assert.equal(state.view.value.sheets[0].cells.A10.value, 'observer')
    assert.deepEqual(state.calls.selected, []); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('selection reentry invalidates the captured publication guard before the parent change callback', () => {
  const state = host(sample(false))
  try {
    state.context.selection = value => { state.calls.selected.push(value); state.active.value = false; state.active.value = true }
    assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), true)
    assert.equal(state.calls.applied, 1); assert.equal(state.calls.selected.length, 1); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('detached identity without a studio ID stays detached after the shared candidate is projected', () => {
  const model = sample(false); delete model.id; const state = host(model)
  try {
    delete state.view.value.id
    assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down, { range: range(), ranges: [range(), range(5, 3)] }), true)
    assert.equal(state.view.value.id, undefined); assert.equal(state.view.value.sheets[0].cells.B3.value, 'move')
  } finally { state.dispose() }
})

test('standalone right-click confirm dispatch rejects a filtered tail before preparing a merge prompt or publishing content', () => {
  const model = sample(); model.sheets[0].autoFilter = { range: { startRow: 0, endRow: 6, startCol: 0, endCol: 3 }, columns: {}, filteredRows: [5] }
  const state = host(model)
  try {
    const before = cloneModel(state.view.value)
    assert.throws(() => state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_CONFIRM_COMMANDS.down), /筛选隐藏行/)
    assert.deepEqual(state.view.value, before); assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
    assert.equal(state.controller.request.value, null)
  } finally { state.dispose() }
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function studioHost(t, merged = false) {
  const model = sample(merged), studio = new OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8 }), state = host(model, studio)
  t.onTestFinished(() => { state.dispose(); studio.dispose() })
  return { ...state, studio, api: FOffice.newAPI(studio) }
}

test('studio UI dispatch uses the active non-last range and relies on registered content history and selection publication', t => {
  const state = studioHost(t), id = state.view.value.id
  state.studio.setSelections(id, 'data', [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
    { startRow: 5, endRow: 5, startColumn: 3, endColumn: 3 }], { activeRangeIndex: 0 })
  assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), true)
  assert.equal(state.api.getActiveWorkbook().getActiveSheet().getRange('B3').getValue(), 'move')
  assert.equal(state.studio.editor.history.units[id].undo.length, 1); assert.equal(state.calls.changed.length, 1)
  assert.equal(state.calls.changed[0].syncSelection, false); assert.deepEqual(state.calls.selected, [])
})

test('a studio structure observer that edits again does not receive an obsolete view or extra parent history update', t => {
  const state = studioHost(t); let reentered = false
  state.studio.events.on('SheetStructureChanged', () => {
    if (reentered) return; reentered = true; state.api.getActiveWorkbook().getActiveSheet().getRange('A10').setValue('observer')
  })
  assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), true)
  assert.equal(state.api.getActiveWorkbook().getActiveSheet().getRange('B3').getValue(), 'move')
  assert.equal(state.api.getActiveWorkbook().getActiveSheet().getRange('A10').getValue(), 'observer')
  assert.equal(state.studio.editor.history.units[state.view.value.id].undo.length, 2)
  assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
})

test('studio confirmation rejects a replaced content model and retains the visible request error', t => {
  const state = studioHost(t, true)
  state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down); const request = state.controller.request.value
  state.api.getActiveWorkbook().getActiveSheet().getRange('A10').setValue('newer')
  state.controller.confirmCellShift(); assert.equal(state.controller.request.value, request)
  assert.match(state.controller.cellShiftError.value, /工作簿.*更改/); assert.equal(state.calls.applied, 0)
  assert.equal(state.api.getActiveWorkbook().getActiveSheet().getRange('B2').getValue(), 'move')
})

test('studio view publication reentry cannot publish an uncommitted observer view back to its content model', t => {
  const state = studioHost(t)
  state.context.applyView = value => { state.view.value = value; state.view.value.sheets[0].cells.A10 = { value: 'uncommitted observer' }; state.calls.applied++ }
  assert.equal(state.controller.runCellShiftCommand(SHEET_CELL_SHIFT_COMMANDS.down), true)
  assert.equal(state.api.getActiveWorkbook().getActiveSheet().getRange('B3').getValue(), 'move')
  assert.equal(state.api.getActiveWorkbook().getActiveSheet().getRange('A10').getValue(), null)
  assert.equal(state.view.value.sheets[0].cells.A10.value, 'uncommitted observer')
  assert.equal(state.calls.applied, 1); assert.equal(state.calls.changed.length, 0)
})
