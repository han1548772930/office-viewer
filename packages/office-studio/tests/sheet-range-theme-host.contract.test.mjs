import assert from 'node:assert/strict'
import { test } from 'vitest'
import { effectScope, nextTick, ref } from 'vue'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { useSheetRangeTheme } from '../src/components/spreadsheet/use-sheet-range-theme.ts'
import { SHEET_RANGE_THEME_COMMANDS } from '../src/features/sheet-workbench/sheet-range-theme-command.ts'

// Drafts only; these model/controller contracts have not been executed.
const range = (startRow = 0, endRow = startRow) => ({ startRow, endRow, startCol: 0, endCol: 1 })
function host(override = {}) {
  const model = createBlankWorkbook(); model.id = 'detached'; model.sheets[0].id = 'data'; model.activeSheetId = 'data'
  const view = ref(toExcelWorkbookView(model)), active = ref(true), changes = [], focus = [], scope = effectScope()
  const context = { studio: () => undefined, workbook: () => view.value, sheetId: () => view.value.activeSheetId,
    ranges: () => [range(0, 2)], userId: () => 'local-user', active: () => active.value, finishEdit: () => true,
    applyView: value => { view.value = value }, changed: options => changes.push(options), focus: () => focus.push(view.value.id), ...override }
  const controller = scope.run(() => useSheetRangeTheme(context))
  return { view, active, changes, focus, context, controller, dispose: () => scope.stop() }
}

test('detached UI application publishes shared author rules once and restores focus after the modal closes', async () => {
  const state = host()
  try {
    state.controller.openRangeThemeDialog(); assert.equal(state.controller.request.value.themeNames.length, 32)
    state.controller.confirmRangeTheme('apply', 'default')
    assert.equal(state.view.value.sheets[0].rangeThemeRules[0].themeName, 'default'); assert.equal(state.changes.length, 1)
    assert.equal(state.controller.request.value, null); assert.equal(state.focus.length, 0)
    state.controller.afterRangeThemeClose(); await nextTick(); assert.deepEqual(state.focus, ['detached'])
  } finally { state.dispose() }
})

test('in-place changes after dialog capture reject application and retain the visible error and original request', () => {
  const state = host()
  try {
    state.controller.openRangeThemeDialog(); const request = state.controller.request.value
    state.view.value.sheets[0].cells.A1 = { value: 'external', rawValue: 'external' }
    state.controller.confirmRangeTheme('apply', 'default')
    assert.equal(state.changes.length, 0); assert.equal(state.view.value.sheets[0].rangeThemeRules, undefined)
    assert.equal(state.controller.request.value, request); assert.match(state.controller.rangeThemeError.value, /已更改/)
  } finally { state.dispose() }
})

test('switching document or workbench cancels the pending modal and never focuses the replacement host', async () => {
  const state = host()
  try {
    state.controller.openRangeThemeDialog(); state.view.value = { ...state.view.value, id: 'replacement' }
    assert.equal(state.controller.request.value, null); state.controller.afterRangeThemeClose(); await nextTick()
    assert.deepEqual(state.focus, []); assert.equal(state.changes.length, 0)
    state.controller.openRangeThemeDialog(); state.active.value = false
    assert.equal(state.controller.request.value, null); state.controller.afterRangeThemeClose(); await nextTick(); assert.deepEqual(state.focus, [])
  } finally { state.dispose() }
})

test('a late modal close cannot focus a reactivated session with the same workbook and sheet', async () => {
  const state = host()
  try {
    state.controller.openRangeThemeDialog(); state.controller.cancelRangeTheme()
    state.active.value = false; state.active.value = true
    state.controller.afterRangeThemeClose(); await nextTick()
    assert.deepEqual(state.focus, []); assert.equal(state.changes.length, 0)
    state.controller.openRangeThemeDialog(); state.controller.cancelRangeTheme()
    state.controller.afterRangeThemeClose(); await nextTick(); assert.deepEqual(state.focus, ['detached'])
  } finally { state.dispose() }
})

test('standalone removal checks protected cells across the full rule and leaves history and content untouched', () => {
  const state = host()
  try {
    const sheet = state.view.value.sheets[0]
    sheet.rangeThemeRules = [{ id: 'rule', themeName: 'default', range: range(0, 8) }]
    sheet.rangeProtectionRules = [{ id: 'protected', ranges: [range(8)], allowedUserIds: ['owner'] }]
    assert.throws(() => state.controller.runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { range: range(1), themeName: null }), /保护/)
    assert.equal(sheet.rangeThemeRules.length, 1); assert.equal(state.changes.length, 0)
  } finally { state.dispose() }
})

test('an unresolved registered name remains removable from the UI without selecting a missing catalog entry', () => {
  const state = host()
  try {
    state.view.value.sheets[0].rangeThemeRules = [{ id: 'gone', themeName: 'unregistered', range: range(0, 2) }]
    state.controller.openRangeThemeDialog()
    assert.equal(state.controller.request.value.initialTheme, 'default'); assert.equal(state.controller.request.value.canRemove, true)
    state.controller.confirmRangeTheme('remove'); assert.deepEqual(state.view.value.sheets[0].rangeThemeRules, [])
  } finally { state.dispose() }
})

test('reentrant preparation or publication cannot emit a change for a different or newly edited detached host', () => {
  const state = host()
  try {
    state.context.prepareView = value => { state.view.value = { ...state.view.value, id: 'other' }; return value }
    assert.equal(state.controller.runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { themeName: 'default' }), false)
    assert.equal(state.changes.length, 0); assert.equal(state.view.value.sheets[0].rangeThemeRules, undefined)
    delete state.context.prepareView
    state.context.applyView = value => { state.view.value = value; state.view.value.sheets[0].cells.B1 = { value: 'observer' } }
    assert.equal(state.controller.runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { themeName: 'default' }), true)
    assert.equal(state.changes.length, 0); assert.equal(state.view.value.sheets[0].cells.B1.value, 'observer')
  } finally { state.dispose() }
})
