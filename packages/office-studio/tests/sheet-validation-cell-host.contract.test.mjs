import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { createSheetCellEditSession } from '../src/features/sheet-workbench/sheet-cell-edit-session.ts'
import { COMMIT_SHEET_DATE_INPUT_COMMAND } from '../src/features/sheet-validation/validation-date-input.ts'

// Implementation-phase drafts. The real OfficeStudio/SSR/effectScope harness has not been run.
let vite, OfficeStudio, FOffice, useSheetValidationCell
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
  ;({ useSheetValidationCell } = await vite.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/use-sheet-validation-cell.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row, column, endRow = row, endCol = column) => ({ startRow: row, startCol: column, endRow, endCol })
const changedDate = { kind: 'date', year: 1900, month: 3, day: 1 }
function host(standalone = false, configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'validation-cell-host'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 60; sheet.columnCount = 20; sheet.cells['0:0'] = createCell(0, 0, 60.125); sheet.cells['0:1'] = createCell(0, 1, 'Red')
  sheet.cells['0:3'] = createCell(0, 3, 'Red'); sheet.cells['1:3'] = createCell(1, 3, 'Green, dark')
  sheet.dataValidations = [{ id: 'date-rule', type: 'date', ranges: [area(0, 0)], metadata: { showTime: true, retained: ['date'] } },
    { id: 'list-rule', type: 'listMultiple', ranges: [area(0, 1)], formula1: '"Red,"Green, dark",It""s blue"', formula2: '#ff0000,#00ff00,#0000ff' }]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = standalone ? undefined : new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), scope = effectScope()
  const state = shallowReactive({ studio, model: toExcelWorkbookView(studio?.getWorkbookModel(model.id) ?? model), sheetId: 'data', key: 'A1',
    userId: 'local-user', active: true, blocked: false, inputSession: null, region: 'en-US', currentYear: 2026, inputOwner: 'validation-owner' })
  const refreshed = [], focused = [], ended = [], openedRules = [], echoes = []
  const options = { state: () => state, refresh: (api, id, current) => {
    if (!current()) return
    refreshed.push(id); state.model = toExcelWorkbookView(api.getWorkbookModel(id))
  }, applyStandalone: (plan, model, current, expectEcho) => {
    if (!plan.result.accepted || !plan.result.changed || !plan.cell) return cloneModel(plan.result)
    assert.equal(current(), true)
    const next = cloneModel(model), sheet = next.sheets.find(item => item.id === state.sheetId)
    sheet.cells[plan.cell.id] = cloneModel(plan.cell); state.model = toExcelWorkbookView(next)
    const snapshot = cloneModel(state.model); expectEcho(snapshot); echoes.push(snapshot); refreshed.push(next.id)
    return cloneModel(plan.result)
  }, endEditing: id => { ended.push(id); if (state.inputSession?.sessionId === id) state.inputSession = null },
  openRule: id => openedRules.push(id), canRestoreFocus: () => !state.blocked, focus: session => focused.push(session?.sessionId ?? state.model.id) }
  const controller = scope.run(() => useSheetValidationCell(options)), current = () => studio?.getWorkbookModel(model.id) ?? state.model
  const book = studio ? FOffice.newAPI(studio).getActiveWorkbook() : undefined
  return { studio, scope, state, options, controller, refreshed, focused, ended, openedRules, echoes, current, book,
    count: () => studio?.editor.history.units[model.id]?.undo.length ?? echoes.length,
    open: () => { assert.equal(controller.entry.value?.open(), true); return controller.popup.value },
    dispose: () => { scope.stop(); studio?.dispose() } }
}

test('a date popup reads the authoritative serial/calendar/rule even when the projected Canvas view is older', async () => {
  const h = host()
  try {
    h.state.model.sheets[0].cells.A1.rawValue = 1; h.state.model.sheets[0].dataValidations[0].metadata.showTime = false
    const before = cloneModel(h.current()), owner = h.open()
    assert.equal(owner.date.originalSerial, 60.125); assert.equal(owner.date.parts.day, 29); assert.equal(owner.date.mode, 'datetime')
    assert.equal(owner.region, 'en-US'); assert.equal(owner.currentYear, 2026); assert.equal(owner.inputOwner, 'validation-owner')
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    owner.setOpen(false); owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, ['validation-cell-host'])
  } finally { h.dispose() }
})

test('one date confirmation refreshes once and restores focus only after its own popup has closed', async () => {
  const h = host()
  try {
    const owner = h.open(), result = owner.submitDate(changedDate)
    assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(owner.visible, false)
    assert.equal(h.count(), 1); assert.deepEqual(h.refreshed, ['validation-cell-host']); assert.deepEqual(h.focused, [])
    assert.equal(h.current().sheets[0].cells['0:0'].value, 61.125)
    owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, ['validation-cell-host'])
    owner.afterOpenChange(false); await nextTick(); assert.equal(h.focused.length, 1)
    h.book.undo(); assert.equal(h.current().sheets[0].cells['0:0'].value, 60.125); h.book.redo(); assert.equal(h.current().sheets[0].cells['0:0'].value, 61.125)
  } finally { h.dispose() }
})

test('an untouched date confirmation preserves the exact model/history and cancellation retains an existing cell input session', async () => {
  const h = host()
  try {
    const session = createSheetCellEditSession({ documentId: h.state.model.id, sheetId: 'data', key: 'A1', date1904: false,
      cell: h.state.model.sheets[0].cells.A1 }, { region: 'en-US', currentYear: 2026 })
    h.state.inputSession = session
    const before = cloneModel(h.current()), owner = h.open()
    owner.setOpen(false); owner.afterOpenChange(false); await nextTick()
    assert.equal(h.state.inputSession, session); assert.deepEqual(h.ended, []); assert.deepEqual(h.focused, [session.sessionId])
    const next = h.open(), result = next.submitDate({ kind: 'unchanged' })
    assert.equal(result.accepted, true); assert.equal(result.changed, false); assert.equal(h.count(), 0); assert.deepEqual(h.current(), before)
    assert.deepEqual(h.ended, [session.sessionId]); assert.equal(h.state.inputSession, null); assert.equal(next.visible, false)
    next.afterOpenChange(false); await nextTick(); assert.equal(h.focused.at(-1), 'validation-cell-host')
  } finally { h.dispose() }
})

test('multiple choices commit immediately in candidate order, rebase one owner and keep its popup open', async () => {
  const h = host()
  try {
    h.state.key = 'B1'; const entry = h.controller.entry.value, owner = h.open()
    assert.deepEqual(owner.selected, ['Red']); assert.equal(owner.multiple, true)
    assert.equal(owner.chooseList(['Green, dark', 'Red']).accepted, true); assert.equal(owner.visible, true)
    assert.equal(h.controller.popup.value, owner); assert.deepEqual(owner.selected, ['Red', 'Green, dark']); assert.equal(h.count(), 1)
    assert.equal(owner.chooseList(['Green, dark']).accepted, true); assert.equal(h.current().sheets[0].cells['0:1'].value, '"Green, dark"')
    assert.deepEqual(owner.selected, ['Green, dark']); assert.equal(h.count(), 2); assert.equal(owner.visible, true); assert.equal(entry.open(), false)
    owner.setOpen(false); owner.afterOpenChange(false); await nextTick(); assert.equal(h.focused.length, 1)
    const reopened = h.open(); assert.deepEqual(reopened.selected, ['Green, dark']); assert.equal(h.count(), 2)
  } finally { h.dispose() }
})

test('STOP rejection and normal command cancellation retain the date draft owner and permit a later retry', () => {
  const h = host(false, (_model, sheet) => { sheet.dataValidations[0].operator = 'equal'; sheet.dataValidations[0].formula1 = '60.125' })
  try {
    const owner = h.open(), before = cloneModel(h.current()), result = owner.submitDate(changedDate)
    assert.equal(result.accepted, false); assert.equal(result.reason, 'validation'); assert.equal(owner.visible, true); assert.match(owner.message, /日期/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    const canceled = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === COMMIT_SHEET_DATE_INPUT_COMMAND) event.cancel = true })
    assert.equal(owner.submitDate({ kind: 'unchanged' }).accepted, false); assert.match(owner.message, /取消/); canceled.dispose()
    assert.equal(owner.submitDate({ kind: 'unchanged' }).accepted, true); assert.equal(owner.visible, false); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('old submit, clear, close and after-close callbacks cannot affect a new popup for the same or another target', async () => {
  const h = host()
  try {
    const old = h.open(); old.setOpen(false); old.afterOpenChange(false)
    h.state.key = 'B1'; const next = h.open(), before = cloneModel(h.current())
    assert.equal(old.submitDate(changedDate).accepted, false); assert.equal(old.chooseList([]).accepted, false)
    old.setOpen(false); old.afterOpenChange(false); await nextTick()
    assert.equal(h.controller.popup.value, next); assert.equal(next.visible, true); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test.each(['cell', 'sheet', 'document', 'user', 'region', 'session', 'activation'])('%s changes permanently end the old popup owner, even if the host immediately returns', async kind => {
  const h = host()
  try {
    const owner = h.open(), before = cloneModel(h.current())
    if (kind === 'cell') { h.state.key = 'B1'; h.state.key = 'A1' }
    else if (kind === 'sheet') { h.state.sheetId = 'other'; h.state.sheetId = 'data' }
    else if (kind === 'document') { h.state.model = { ...h.state.model, id: 'replaced' }; h.state.model = toExcelWorkbookView(h.current()) }
    else if (kind === 'user') { h.state.userId = 'other'; h.state.userId = 'local-user' }
    else if (kind === 'region') { h.state.region = 'de-DE'; h.state.region = 'en-US' }
    else if (kind === 'session') { h.state.inputSession = { sessionId: 'competing' }; h.state.inputSession = null }
    else { h.state.active = false; h.state.active = true }
    assert.equal(h.controller.popup.value, undefined); assert.equal(owner.submitDate(changedDate).accepted, false)
    owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test.each(['clear-history', 'same-id-replacement'])('%s invalidates a popup without a content transaction', async change => {
  const h = host()
  try {
    const owner = h.open(), before = cloneModel(h.current())
    if (change === 'clear-history') h.studio.editor.history.clear('validation-cell-host')
    else h.studio.editor.workbook.replace(cloneModel(h.current()))
    assert.equal(h.controller.popup.value, undefined); assert.equal(owner.submitDate(changedDate).accepted, false)
    owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.current().sheets[0].cells, before.sheets[0].cells); assert.equal(h.count(), 0); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('an unrelated content edit can advance the owner, while a changed target keeps the original date draft blocked', () => {
  const h = host()
  try {
    const owner = h.open(); h.book.getActiveSheet().getRange('F10').setValue('unrelated')
    assert.equal(h.controller.popup.value, owner); assert.equal(owner.disabled, false); assert.equal(owner.submitDate(changedDate).accepted, true); assert.equal(h.count(), 2)
    owner.afterOpenChange(false); const next = h.open(); h.book.getActiveSheet().getRange('A1').setValue(62.125)
    const before = cloneModel(h.current()), count = h.count()
    assert.equal(next.visible, true); assert.equal(next.date.originalSerial, 61.125); assert.equal(next.disabled, true)
    assert.equal(next.submitDate(changedDate).accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), count); assert.match(next.message, /更改/)
  } finally { h.dispose() }
})

test('live reference candidates, casing and source permissions update a multiple popup without writing the cell', () => {
  const h = host(false, (_model, sheet) => { sheet.dataValidations[1].formula1 = '=$D$1:$D$2' })
  try {
    h.state.key = 'B1'; const owner = h.open(); assert.deepEqual(owner.options.map(option => option.label), ['Red', 'Green, dark'])
    h.book.getActiveSheet().getRange('D2').setValue('Approved'); const count = h.count()
    assert.deepEqual(owner.options.map(option => option.label), ['Red', 'Approved']); assert.deepEqual(owner.selected, ['Red']); assert.equal(h.current().sheets[0].cells['0:1'].value, 'Red')
    assert.equal(owner.chooseList(['approved', 'RED']).accepted, true); assert.equal(h.count(), count + 1); assert.deepEqual(owner.selected, ['Red', 'Approved'])
    h.current().sheets[0].rangeProtectionRules = [{ id: 'private-source', ranges: [area(0, 3, 1, 3)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' }]
    assert.deepEqual(owner.options, []); assert.equal(owner.chooseList(['Approved']).accepted, false)
  } finally { h.dispose() }
})

test('dynamic branch changes update one popup through the shared reference engine and the subsequent choice records one history item', () => {
  const h = host(false, (_model, sheet) => {
    sheet.cells['0:4'] = createCell(0, 4, true); sheet.cells['2:3'] = createCell(2, 3, 'Blue'); sheet.cells['3:3'] = createCell(3, 3, 'Approved')
    sheet.dataValidations[1].formula1 = '=IF($E$1,$D$1:$D$2,$D$3:$D$4)'
  })
  try {
    h.state.key = 'B1'; const owner = h.open(); assert.deepEqual(owner.options.map(option => option.label), ['Red', 'Green, dark'])
    h.book.getActiveSheet().getRange('E1').setValue(false)
    assert.equal(h.controller.popup.value, owner); assert.deepEqual(owner.options.map(option => option.label), ['Blue', 'Approved'])
    assert.equal(h.current().sheets[0].cells['0:1'].value, 'Red'); assert.equal(h.count(), 1)
    assert.equal(owner.chooseList(['Approved']).accepted, true); assert.equal(h.current().sheets[0].cells['0:1'].value, 'Approved')
    assert.equal(h.count(), 2); assert.equal(h.controller.popup.value, owner); assert.equal(owner.visible, true)
    owner.setOpen(false); owner.afterOpenChange(false)
    h.book.undo(); assert.equal(h.current().sheets[0].cells['0:1'].value, 'Red')
    h.book.redo(); assert.equal(h.current().sheets[0].cells['0:1'].value, 'Approved')
    assert.deepEqual(h.open().selected, ['Approved'])
  } finally { h.dispose() }
})

test('rule changes and protection changes block the original owner; editing a rule requires every covered range', () => {
  const h = host(false, (_model, sheet) => {
    sheet.dataValidations[0].ranges.push(area(0, 13)); sheet.rangeProtectionRules = [{ id: 'remote-lock', ranges: [area(0, 13)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'othersCanView', editState: 'onlyMe' }]
  })
  try {
    const owner = h.open(); assert.equal(owner.disabled, false); assert.equal(owner.canEditRule, false); owner.editRule(); assert.deepEqual(h.openedRules, [])
    h.current().sheets[0].dataValidations[0].metadata.showTime = false
    assert.equal(owner.disabled, true); assert.match(owner.message, /规则/); assert.equal(owner.submitDate(changedDate).accepted, false); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('a view-denied cell has no entry, and blocked input cannot open or submit a widget', () => {
  const h = host(false, (_model, sheet) => { sheet.rangeProtectionRules = [{ id: 'private-cell', ranges: [area(0, 0)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }] })
  try { assert.equal(h.controller.entry.value, undefined); assert.equal(h.count(), 0) } finally { h.dispose() }
  const next = host()
  try {
    next.state.blocked = true; assert.equal(next.controller.entry.value.disabled, true); assert.equal(next.controller.entry.value.open(), false)
    next.state.blocked = false; const owner = next.open(), before = cloneModel(next.current()); next.state.blocked = true
    assert.equal(owner.disabled, true); assert.equal(owner.submitDate(changedDate).accepted, false); assert.deepEqual(next.current(), before); assert.equal(next.count(), 0)
  } finally { next.dispose() }
})

test('an owned successful edit can end its cell input session, while a host leaving during refresh cannot reclaim focus', async () => {
  const h = host()
  try {
    h.state.inputSession = createSheetCellEditSession({ documentId: h.state.model.id, sheetId: 'data', key: 'A1', date1904: false,
      cell: h.state.model.sheets[0].cells.A1 }, { region: 'en-US', currentYear: 2026 })
    const id = h.state.inputSession.sessionId, owner = h.open()
    h.options.refresh = () => { h.state.active = false; h.state.active = true }
    assert.equal(owner.submitDate(changedDate).accepted, true); assert.deepEqual(h.ended, [id]); assert.equal(h.state.inputSession, null)
    assert.equal(h.controller.popup.value, undefined); assert.equal(h.count(), 1); owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('an observer edit after widget publication cannot count as the popup commit or overwrite its newer host view', async () => {
  const h = host()
  try {
    const subscription = h.studio.events.addEvent('CommandExecuted', event => { if (event.id === COMMIT_SHEET_DATE_INPUT_COMMAND) h.book.getActiveSheet().getRange('F10').setValue('listener edit') })
    const owner = h.open(); assert.equal(owner.submitDate(changedDate).accepted, true); subscription.dispose()
    assert.equal(h.count(), 2); assert.equal(h.controller.popup.value, undefined); assert.deepEqual(h.refreshed, [])
    assert.equal(h.current().sheets[0].cells['9:5'].value, 'listener edit'); owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('standalone multiple choices accept one exact emitted model echo, while equal external replacement callbacks remain stale', async () => {
  const h = host(true)
  try {
    h.state.key = 'B1'; const owner = h.open()
    assert.equal(owner.chooseList(['Red', 'Green, dark']).accepted, true); assert.equal(owner.visible, true); assert.equal(h.count(), 1)
    const snapshot = h.echoes[0]
    h.controller.receiveView(snapshot, () => { h.state.model = cloneModel(snapshot) })
    assert.equal(h.controller.popup.value, owner); assert.equal(owner.disabled, false)
    assert.equal(owner.chooseList(['Green, dark']).accepted, true); assert.equal(h.count(), 2)
    const old = owner, equalReplacement = cloneModel(h.state.model)
    h.controller.receiveView(equalReplacement, () => { h.state.model = equalReplacement })
    assert.equal(h.controller.popup.value, undefined); assert.equal(old.chooseList(['Red']).accepted, false)
    old.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, [])
    const next = h.open(); assert.deepEqual(next.selected, ['Green, dark'])
  } finally { h.dispose() }
})

test('inline candidates keep immutable identity across unrelated revisions while query context follows the current owner revision', () => {
  const h = host()
  try {
    h.state.key = 'B1'; const owner = h.open(), candidates = owner.options
    assert.equal(Object.isFrozen(candidates), true); assert.equal(candidates.every(Object.isFrozen), true)
    assert.deepEqual(owner.queryContext(), { unitId: 'validation-cell-host', baseRevision: 0 })
    h.book.getActiveSheet().getRange('F10').setValue('unrelated')
    assert.equal(owner.options, candidates); assert.equal(owner.current(), true); assert.equal(owner.queryContext().baseRevision, 1)
    assert.equal(owner.chooseList(['Red', 'Green, dark']).accepted, true)
    assert.equal(owner.options, candidates); assert.equal(owner.queryContext().baseRevision, 2)
    assert.deepEqual(owner.selected, ['Red', 'Green, dark']); assert.equal(h.count(), 2)
  } finally { h.dispose() }
})

test.each([false, true])('standalone=%s: a region change after close permanently cancels delayed focus and the original entry lease', async standalone => {
  const h = host(standalone)
  try {
    h.state.key = 'B1'; const entry = h.controller.entry.value, owner = h.open(), before = cloneModel(h.current())
    owner.setOpen(false); owner.afterOpenChange(false)
    h.state.region = 'de-DE'; h.state.region = 'en-US'; await nextTick()
    assert.deepEqual(h.focused, []); assert.equal(owner.current(), false); assert.equal(owner.queryContext(), undefined)
    assert.deepEqual(owner.options, []); assert.deepEqual(owner.selected, []); assert.equal(entry.open(), false)
    assert.equal(owner.chooseList(['Green, dark']).accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    const next = h.open(); assert.equal(next.region, 'en-US'); assert.equal(next.current(), true); assert.notEqual(next.ownerKey, owner.ownerKey)
  } finally { h.dispose() }
})

test.each([[false, 'assignment'], [false, 'receiveView'], [true, 'assignment'], [true, 'receiveView']])(
  'standalone=%s, projection=%s: replacing and restoring a closed host view cannot revive its delayed focus or entry lease', async (standalone, method) => {
    const h = host(standalone)
    try {
      h.state.key = 'B1'
      const originalView = h.state.model, entry = h.controller.entry.value, owner = h.open(), before = cloneModel(h.current())
      owner.setOpen(false); owner.afterOpenChange(false)
      const replacement = cloneModel(originalView)
      if (method === 'receiveView') {
        h.controller.receiveView(replacement, () => { h.state.model = replacement })
        h.controller.receiveView(originalView, () => { h.state.model = originalView })
      } else {
        h.state.model = replacement; h.state.model = originalView
      }
      await nextTick()
      assert.equal(h.state.model, originalView); assert.deepEqual(h.focused, [])
      assert.equal(h.controller.popup.value, undefined); assert.equal(owner.current(), false); assert.equal(owner.queryContext(), undefined)
      assert.deepEqual(owner.options, []); assert.deepEqual(owner.selected, []); assert.equal(entry.open(), false)
      assert.equal(owner.chooseList(['Green, dark']).accepted, false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
      owner.afterOpenChange(false); await nextTick(); assert.deepEqual(h.focused, [])
      const next = h.open(); assert.equal(next.current(), true); assert.notEqual(next.ownerKey, owner.ownerKey)
      assert.deepEqual(next.selected, ['Red']); assert.equal(h.count(), 0)
    } finally { h.dispose() }
  })

test('revoked target View ends the popup lease and never exposes inline labels or saved selection through an old owner', () => {
  const h = host()
  try {
    h.state.key = 'B1'; const owner = h.open(); assert.deepEqual(owner.selected, ['Red']); assert.equal(owner.options.length, 3)
    h.studio.commitWorkbook(h.state.model.id, 'Revoke target View', model => { model.sheets[0].rangeProtectionRules = [{ id: 'private-target', ranges: [area(0, 1)],
      ownerUserId: 'peer', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' }] })
    const before = cloneModel(h.current()), deniedCount = h.count()
    assert.equal(h.controller.popup.value, undefined); assert.equal(owner.current(), false); assert.equal(owner.queryContext(), undefined)
    assert.deepEqual(owner.options, []); assert.deepEqual(owner.selected, []); assert.equal(owner.chooseList(['Green, dark']).accepted, false)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), deniedCount)
    h.studio.commitWorkbook(h.state.model.id, 'Restore target View', model => { model.sheets[0].rangeProtectionRules = [] })
    const restoredCount = h.count()
    const next = h.open(); assert.equal(next.current(), true); assert.notEqual(next.ownerKey, owner.ownerKey)
    assert.deepEqual(owner.options, []); assert.equal(h.count(), restoredCount); assert.equal(h.current().sheets[0].cells['0:1'].value, 'Red')
  } finally { h.dispose() }
})
