import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditSession } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'
import { SHEET_VALIDATION_DROPDOWN_OPERATIONS as operations } from '../src/features/sheet-validation/validation-dropdown.ts'

// One SSR graph owns the Studio, hooks and WeakMap coordinator. These implementation-phase drafts have not run.
let vite, OfficeStudio, FOffice, useSheetValidationCell, activeSheetValidationDropdown, attachSheetValidationDropdownController, registerSheetValidationDropdownOperations
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
  ;({ useSheetValidationCell } = await vite.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/use-sheet-validation-cell.ts'))
  ;({ activeSheetValidationDropdown, attachSheetValidationDropdownController, registerSheetValidationDropdownOperations } =
    await vite.ssrLoadModule('/packages/office-studio/src/api/sheet-validation-dropdown-operations.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
function host(configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'dropdown'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'; sheet.rowCount = 60; sheet.columnCount = 20
  sheet.cells['0:0'] = createCell(0, 0, 60.125); sheet.cells['0:1'] = createCell(0, 1, 'Red')
  sheet.dataValidations = [{ id: 'date', type: 'date', ranges: [area(0, 0)] },
    { id: 'list', type: 'listMultiple', ranges: [area(0, 1)], formula1: '"Red,Blue"' }]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), scope = effectScope(), book = FOffice.newAPI(studio).getActiveWorkbook()
  const state = shallowReactive({ studio, model: toExcelWorkbookView(studio.getWorkbookModel(model.id)), sheetId: 'data', key: 'A1',
    userId: 'local-user', active: true, blocked: false, inputSession: null, region: 'en-US', currentYear: 2026, inputOwner: 'dropdown-owner' })
  const revealed = [], focused = []
  const options = { state: () => state, refresh: (api, id, current) => { if (current()) state.model = toExcelWorkbookView(api.getWorkbookModel(id)) },
    applyStandalone: () => { throw new Error('Unexpected standalone path') }, endEditing: () => {}, openRule: () => {},
    canRestoreFocus: () => !state.blocked, focus: () => focused.push(state.key), reveal: key => { revealed.push(key); return true } }
  const controller = scope.run(() => useSheetValidationCell(options))
  return { studio, scope, book, state, options, controller, revealed, focused,
    current: () => studio.getWorkbookModel(model.id), count: () => studio.editor.history.units[model.id]?.undo.length ?? 0,
    show: (row = 0, column = 0, extra = {}) => studio.executeCommandSync(operations.show, { unitId: model.id, subUnitId: sheet.id, row, column, ...extra }),
    hide: input => studio.executeCommandSync(operations.hide, input), active: () => activeSheetValidationDropdown(studio),
    dispose: () => { scope.stop(); studio.dispose() } }
}

test('a public show at another cell reveals its popup without moving selection or writing content/history', () => {
  const h = host()
  try {
    const before = cloneModel(h.current())
    assert.equal(h.show(0, 1), true); assert.equal(h.state.key, 'A1'); assert.equal(h.controller.popup.value.key, 'B1')
    assert.deepEqual(h.active(), { unitId: 'dropdown', sheetId: 'data', row: 0, column: 1 }); assert.deepEqual(h.revealed, ['B1'])
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test.each([{ showDropDown: false }, { renderMode: 'text' }])('explicit show opens a hidden trigger using the same rule and command', settings => {
  const h = host((_model, sheet) => Object.assign(sheet.dataValidations[1], settings))
  try {
    h.state.key = 'B1'; assert.equal(h.controller.entry.value, undefined)
    assert.equal(h.show(0, 1), true); assert.equal(h.controller.entry.value.triggerVisible, false)
    assert.deepEqual(h.controller.popup.value.options.map(option => option.label), ['Red', 'Blue']); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('show resolves a merged covered coordinate to its main cell and validates complete View/Edit', () => {
  const h = host((_model, sheet) => { sheet.mergedRanges = [area(0, 0, 1, 1)] })
  try {
    assert.equal(h.show(1, 1), true); assert.deepEqual(h.active(), { unitId: 'dropdown', sheetId: 'data', row: 0, column: 0 })
    h.hide({}); h.controller.popup.value.afterOpenChange(false)
    h.current().sheets[0].rangeProtectionRules = [{ id: 'p', ranges: [area(1, 1)], ownerUserId: 'peer', allowedUserIds: [],
      viewState: 'othersCanView', editState: 'onlyMe' }]
    assert.throws(() => h.show(1, 1), /permission|权限|保护/i); assert.equal(h.active(), undefined); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('repeated show is idempotent and retains the same draft and composition owner', () => {
  const h = host()
  try {
    assert.equal(h.show(), true); const owner = h.controller.popup.value; owner.setComposing(true)
    assert.equal(h.show(), true); assert.equal(h.controller.popup.value, owner); assert.equal(owner.composing, true)
    assert.deepEqual(h.revealed, ['A1']); assert.equal(h.count(), 0)
    owner.setComposing(false)
  } finally { h.dispose() }
})

test('hide requires a parameter object, honors an optional target filter and is idempotent once closed', async () => {
  const h = host()
  try {
    h.show(0, 1); const owner = h.controller.popup.value
    assert.equal(h.hide(undefined), false); assert.equal(h.hide({ unitId: 'other' }), false); assert.equal(owner.visible, true)
    assert.equal(h.hide({ sheetId: 'other' }), false); assert.equal(owner.visible, true)
    assert.equal(h.hide({ unitId: 'dropdown', sheetId: 'data' }), true); assert.equal(owner.visible, false); assert.equal(h.active(), undefined)
    assert.equal(h.hide({}), true); owner.afterOpenChange(false); await nextTick()
    assert.equal(h.state.key, 'A1'); assert.deepEqual(h.focused, ['A1']); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test.each(['false', 'throw'])('a reveal returning %s preserves the old popup instance and draft', failure => {
  const h = host()
  try {
    h.show(); const owner = h.controller.popup.value, before = cloneModel(h.current())
    h.options.reveal = () => { if (failure === 'throw') throw new Error('Reveal rejected'); return false }
    if (failure === 'throw') assert.throws(() => h.show(0, 1), /Reveal rejected/)
    else assert.equal(h.show(0, 1), false)
    assert.equal(h.controller.popup.value, owner); assert.equal(owner.visible, true); assert.equal(owner.disabled, false)
    assert.equal(owner.date.originalSerial, 60.125); assert.equal(h.active().column, 0); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('a source revision changed during reveal rejects the new owner while retaining the previous valid draft', () => {
  const h = host()
  try {
    h.show(); const owner = h.controller.popup.value
    h.options.reveal = () => { h.book.getActiveSheet().getRange('F10').setValue('observer'); return true }
    assert.equal(h.show(0, 1), false); assert.equal(h.controller.popup.value, owner); assert.equal(owner.visible, true); assert.equal(owner.disabled, false)
    assert.equal(h.current().sheets[0].cells['9:5'].value, 'observer'); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test('a missing rule or a sheet outside the active workbench does not close an existing popup', () => {
  const h = host()
  try {
    h.show(); const owner = h.controller.popup.value
    assert.equal(h.show(10, 10), false); assert.equal(h.show(0, 0, { subUnitId: 'other' }), false); assert.equal(h.show(60, 0), false)
    assert.equal(h.controller.popup.value, owner); assert.equal(owner.visible, true); assert.equal(h.count(), 0)
    assert.throws(() => h.show(0, 0, { expectedRevision: h.studio.getWorkbookRevision('dropdown') + 1 }), /更改/)
    assert.equal(h.controller.popup.value, owner)
  } finally { h.dispose() }
})

test('UI entry cancellation returns false; public show/hide cancellations retain every previous owner', () => {
  const h = host(), cancel = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === operations.show) event.cancel = true })
  try {
    assert.equal(h.controller.entry.value.open(), false); assert.equal(h.controller.popup.value, undefined)
    cancel.dispose(); h.show(); const owner = h.controller.popup.value
    const cancelShow = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === operations.show) event.cancel = true })
    assert.throws(() => h.show(0, 1), /canceled/); cancelShow.dispose(); assert.equal(h.controller.popup.value, owner)
    const cancelHide = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === operations.hide) event.cancel = true })
    assert.throws(() => h.hide({}), /canceled/); cancelHide.dispose(); assert.equal(owner.visible, true); assert.equal(h.count(), 0)
  } finally { cancel.dispose(); h.dispose() }
})

test('composition rejects hide and a different target, including a callback that starts composition during reveal', () => {
  const h = host()
  try {
    h.show(); const owner = h.controller.popup.value; owner.setComposing(true)
    assert.equal(h.hide({}), false); assert.equal(h.show(0, 1), false); assert.equal(h.controller.popup.value, owner)
    owner.setComposing(false); h.options.reveal = () => { owner.setComposing(true); return true }
    assert.equal(h.show(0, 1), false); assert.equal(h.controller.popup.value, owner); assert.equal(owner.composing, true); assert.equal(h.count(), 0)
    owner.setComposing(false)
  } finally { h.dispose() }
})

test('a competing input session cannot be reused for a different popup cell', () => {
  const h = host()
  try {
    h.state.inputSession = createSheetCellEditSession({ documentId: 'dropdown', sheetId: 'data', key: 'A1', date1904: false,
      cell: h.state.model.sheets[0].cells.A1 }, { region: 'en-US', currentYear: 2026 })
    assert.equal(h.show(0, 1), false); assert.equal(h.controller.popup.value, undefined); assert.equal(h.state.inputSession.key, 'A1')
    assert.equal(h.show(), true); assert.equal(h.controller.popup.value.session, h.state.inputSession); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('reentrant show/hide requests during reveal are rejected without losing the outer popup owner', () => {
  const h = host()
  try {
    h.show(); const reentrant = []
    h.options.reveal = () => { reentrant.push(h.show(), h.hide({})); return true }
    assert.equal(h.show(0, 1), true); assert.deepEqual(reentrant, [false, false]); assert.equal(h.controller.popup.value.key, 'B1'); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('the currently open scope remains the sole coordinator owner while blocked', () => {
  const h = host(); let shows = 0
  const subscription = attachSheetValidationDropdownController(h.studio, {
    supports: () => true, show: () => { shows++; return true }, hide: () => true, active: () => undefined,
  })
  try {
    subscription.dispose(); h.show(); const owner = h.controller.popup.value
    const competing = attachSheetValidationDropdownController(h.studio, {
      supports: () => true, show: () => { shows++; return true }, hide: () => true, active: () => undefined,
    })
    h.state.blocked = true; assert.equal(h.show(0, 1), false); assert.equal(shows, 0); assert.equal(h.controller.popup.value, owner)
    competing.dispose(); h.state.blocked = false; assert.equal(h.hide({}), true)
  } finally { subscription.dispose(); h.dispose() }
})

test('a disposed workbench controller leaves Studio registration intact and cannot restore stale focus', async () => {
  const h = host()
  try {
    h.show(); const owner = h.controller.popup.value; h.scope.stop()
    assert.equal(h.active(), undefined); assert.equal(h.show(), false); assert.equal(h.hide({}), true)
    owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, [])
    assert.throws(() => registerSheetValidationDropdownOperations(h.studio), /注册一次/)
  } finally { h.dispose() }
})

test('same-ID replacement and a changed selection invalidate programmatic popup ownership', async () => {
  const h = host()
  try {
    h.show(0, 1); const old = h.controller.popup.value; h.studio.editor.workbook.replace(cloneModel(h.current()))
    assert.equal(h.controller.popup.value, undefined); assert.equal(h.active(), undefined); old.afterOpenChange(false)
    h.show(0, 1); const next = h.controller.popup.value; h.state.key = 'B1'
    assert.equal(h.controller.popup.value, undefined); assert.equal(h.active(), undefined); next.afterOpenChange(false)
    await nextTick(); assert.deepEqual(h.focused, []); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})
