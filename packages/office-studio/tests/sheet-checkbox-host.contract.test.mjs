import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { COMMIT_SHEET_CHECKBOX_INPUT_COMMAND } from '../src/features/sheet-validation/validation-checkbox-input.ts'

// Real command/owner integration drafts. No runner, SSR server or Office sample has been executed.
let vite, OfficeStudio, FOffice, useSheetValidationCheckbox
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
  ;({ useSheetValidationCheckbox } = await vite.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/use-sheet-validation-checkbox.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
function host(standalone = false, configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'checkbox-host'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'; sheet.rowCount = 60; sheet.columnCount = 15
  sheet.cells['0:0'] = createCell(0, 0, 0); sheet.cells['1:0'] = createCell(1, 0, 1); sheet.cells['0:1'] = createCell(0, 1, 'existing')
  sheet.dataValidations = [{ id: 'checkbox', type: 'checkbox', ranges: [area(0, 0, 1, 0)] }]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = standalone ? undefined : new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), scope = effectScope()
  const state = shallowReactive({ studio, model: toExcelWorkbookView(studio?.getWorkbookModel(model.id) ?? model), sheetId: 'data', userId: 'local-user',
    active: true, blocked: false, inputSession: null, region: 'en-US', currentYear: 2026, inputOwner: 'checkbox-owner' })
  const refreshed = [], focused = [], snapshots = []
  const options = { state: () => state, refresh: (api, id, current) => {
    if (current()) { refreshed.push(id); state.model = toExcelWorkbookView(api.getWorkbookModel(id)) }
  }, applyStandalone: (plan, currentModel, current, published) => {
    if (!plan.result.accepted || !plan.result.changed || !plan.cell) return cloneModel(plan.result)
    assert.equal(current(), true)
    const next = cloneModel(currentModel); next.sheets.find(sheet => sheet.id === state.sheetId).cells[plan.cell.id] = cloneModel(plan.cell)
    state.model = toExcelWorkbookView(next); published(state.model); snapshots.push(cloneModel(state.model))
    return cloneModel(plan.result)
  }, focus: () => focused.push(state.sheetId), canRestoreFocus: () => !state.blocked }
  const control = scope.run(() => useSheetValidationCheckbox(options)), book = studio && FOffice.newAPI(studio).getActiveWorkbook()
  return { studio, scope, state, options, control, book, focused, refreshed, snapshots,
    current: () => studio?.getWorkbookModel(model.id), entry: (row = 0, column = 0) => control.entryAt(row, column),
    dispose: () => { scope.stop(); studio?.dispose() } }
}

test('the control reads authoritative values even when the projected Canvas view has an older scalar', async () => {
  const h = host()
  try {
    h.state.model.sheets[0].cells.A1.rawValue = 1
    const entry = h.entry(), before = cloneModel(h.current())
    assert.equal(entry.checked, false); assert.equal(entry.inputOwner, 'checkbox-owner')
    assert.equal(entry.activate().accepted, true); assert.equal(h.current().sheets[0].cells['0:0'].value, 1)
    assert.deepEqual(h.current().sheets[0].cells['1:0'], before.sheets[0].cells['1:0']); assert.deepEqual(h.refreshed, ['checkbox-host'])
    await nextTick(); assert.deepEqual(h.focused, ['data'])
  } finally { h.dispose() }
})

test('a captured pointer callback cannot retoggle the new value, while a fresh control can', () => {
  const h = host()
  try {
    const old = h.entry(); assert.equal(old.activate().accepted, true); const before = cloneModel(h.current())
    assert.equal(old.activate(false).accepted, false); assert.deepEqual(h.current(), before)
    const current = h.entry(); assert.equal(current.checked, true); assert.equal(current.activate().accepted, true)
    assert.equal(h.current().sheets[0].cells['0:0'].value, 0); assert.equal(h.studio.editor.history.units['checkbox-host'].undo.length, 2)
  } finally { h.dispose() }
})

test.each(['cell', 'rule', 'history', 'user', 'sheet', 'input', 'inactive', 'replacement'])('%s invalidates an old checkbox activation without a content write', async kind => {
  const h = host()
  try {
    const old = h.entry()
    if (kind === 'cell') h.current().sheets[0].cells['0:0'].value = 1
    else if (kind === 'rule') h.current().sheets[0].dataValidations[0].formula2 = 'TODO'
    else if (kind === 'history') h.studio.editor.history.clear('checkbox-host')
    else if (kind === 'user') h.state.userId = 'peer'
    else if (kind === 'sheet') h.state.sheetId = 'other'
    else if (kind === 'input') h.state.inputSession = { sessionId: 'competing' }
    else if (kind === 'inactive') h.state.active = false
    else h.studio.editor.workbook.replace(cloneModel(h.current()))
    const before = cloneModel(h.current()); assert.equal(old.activate().accepted, false); assert.deepEqual(h.current(), before)
    await nextTick(); assert.deepEqual(h.focused, []); assert.deepEqual(h.refreshed, [])
  } finally { h.dispose() }
})

test.each(['edit', 'view'])('complete merged %s protection disables or removes its control', kind => {
  const h = host(false, (_model, sheet) => {
    sheet.mergedRanges = [area(0, 0, 1, 1)]; sheet.rangeProtectionRules = [{ id: 'p', permissionId: 'p', ranges: [area(1, 1)],
      ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: kind === 'view' ? 'noOneElseCanView' : 'othersCanView' }]
  })
  try {
    if (kind === 'view') assert.equal(h.entry(1, 1), undefined)
    else { assert.equal(h.entry(1, 1).disabled, true); assert.equal(h.entry(1, 1).activate().accepted, false) }
    assert.equal(h.current().sheets[0].cells['0:0'].value, 0)
  } finally { h.dispose() }
})

test('a canceled checkbox command keeps its source and allows a later retry', async () => {
  const h = host(), cancel = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === COMMIT_SHEET_CHECKBOX_INPUT_COMMAND) event.cancel = true })
  try {
    const entry = h.entry(), before = cloneModel(h.current())
    assert.equal(entry.activate().accepted, false); assert.match(h.control.error.value, /取消/); assert.deepEqual(h.current(), before)
    await nextTick(); assert.deepEqual(h.focused, []); cancel.dispose()
    assert.equal(entry.activate().accepted, true); assert.equal(h.control.error.value, '')
  } finally { cancel.dispose(); h.dispose() }
})

test('observer edits prevent the original gesture from publishing a stale view or restoring focus', async () => {
  const h = host(), subscription = h.studio.events.addEvent('SheetValueChanged', () => {
    subscription.dispose(); h.book.getActiveSheet().getRange('B1').setValue('observer')
  })
  try {
    assert.equal(h.entry().activate().accepted, true); assert.equal(h.current().sheets[0].cells['0:0'].value, 1)
    assert.equal(h.current().sheets[0].cells['0:1'].value, 'observer'); assert.deepEqual(h.refreshed, [])
    await nextTick(); assert.deepEqual(h.focused, [])
  } finally { subscription.dispose(); h.dispose() }
})

test('standalone writes publish one actual snapshot and reject the old control after model replacement', async () => {
  const h = host(true)
  try {
    const old = h.entry(); assert.equal(old.activate().accepted, true)
    assert.equal(h.state.model.sheets[0].cells.A1.rawValue, 1); assert.equal(h.snapshots.length, 1)
    assert.equal(old.activate(false).accepted, false); await nextTick(); assert.deepEqual(h.focused, ['data'])
    const next = h.entry(); h.state.model = cloneModel(h.state.model)
    assert.equal(next.activate(false).accepted, false); assert.equal(h.snapshots.length, 1)
  } finally { h.dispose() }
})

test('a stopped scope leaves no usable keyboard or pointer callback', async () => {
  const h = host()
  try {
    const entry = h.entry(), before = cloneModel(h.current()); h.scope.stop()
    assert.equal(entry.activate().accepted, false); assert.deepEqual(h.current(), before); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})
