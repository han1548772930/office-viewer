import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, ref, shallowRef } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { normalizeSheetRangeThemeStyle } from '../src/features/sheet-workbench/sheet-range-theme-style.ts'
import { SHEET_RANGE_THEME_COMMANDS } from '../src/features/sheet-workbench/sheet-range-theme-command.ts'
import { useSheetRangeTheme } from '../src/components/spreadsheet/use-sheet-range-theme.ts'

// Drafts only. No test runner, dev/browser, build/type checks or benchmarks were executed.
const range = (row = 0, endRow = row) => ({ startRow: row, endRow, startCol: 0, endCol: 1 })
const definition = (name = 'edited') => normalizeSheetRangeThemeStyle({ name, wholeStyle: { ff: 'Georgia', fs: 22 }, headerRowStyle: { bl: 0 } })
function sample() {
  const model = createBlankWorkbook(); model.id = 'range-theme-editor-host'; model.calculationMode = 'manual'; model.activeSheetId = 'data'
  model.sheets[0].id = 'data'; model.sheets[0].defaultRowAutoHeight = false
  model.sheets[0].cells['0:0'] = { ...createCell(0, 0, 'source'), authoredStyle: {} }
  const other = createBlankSheet('Other'); other.id = 'other'; other.defaultRowAutoHeight = false; model.sheets.push(other)
  return model
}
function host(model = sample(), studio) {
  const view = ref(toExcelWorkbookView(studio ? studio.getWorkbookModel(model.id) : model)), active = ref(true), user = ref('local-user')
  const owner = shallowRef(studio), scope = effectScope(), calls = { applied: 0, changed: [], focused: 0 }
  const context = { studio: () => owner.value, workbook: () => view.value, sheetId: () => view.value.activeSheetId,
    ranges: () => [range(0, 3)], userId: () => user.value, active: () => active.value, finishEdit: () => true,
    prepareView: value => value, applyView: value => { view.value = value; calls.applied++ }, changed: options => calls.changed.push(options),
    focus: () => calls.focused++ }
  const controller = scope.run(() => useSheetRangeTheme(context))
  return { view, active, user, owner, context, controller, calls, dispose: () => scope.stop() }
}

test('a detached custom draft registers and applies in one publication without materializing blank cells', async () => {
  const state = host()
  try {
    state.controller.openRangeThemeDialog(); state.controller.confirmRangeTheme('apply', 'edited', definition())
    assert.equal(state.calls.applied, 1); assert.equal(state.calls.changed.length, 1); assert.equal(state.calls.changed[0].current(), true)
    assert.equal(state.view.value.rangeThemes.edited.wholeStyle.font.size, 22); assert.equal(state.view.value.sheets[0].rangeThemeRules[0].themeName, 'edited')
    assert.equal(Object.keys(state.view.value.sheets[0].cells).length, 1); assert.deepEqual(state.view.value.sheets[0].cells.A1.authoredStyle, {})
    assert.equal(state.controller.request.value, null); state.controller.afterRangeThemeClose(); await nextTick(); assert.equal(state.calls.focused, 1)
  } finally { state.dispose() }
})

test('request styles are detached and an identical-ID replacement cannot authorize the captured draft', () => {
  const model = sample(); model.rangeThemes = { custom: definition('custom') }; const state = host(model)
  try {
    state.controller.openRangeThemeDialog(); state.controller.request.value.themeStyles.custom.wholeStyle.font.size = 1
    assert.equal(state.view.value.rangeThemes.custom.wholeStyle.font.size, 22)
    state.view.value = cloneModel(state.view.value); state.controller.confirmRangeTheme('apply', 'edited', definition())
    assert.match(state.controller.rangeThemeError.value, /宿主已更改/); assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('detached edits reject a protected reference on another worksheet without leaving a partial catalog update', () => {
  const model = sample(); model.rangeThemes = { edited: definition() }
  model.sheets[1].rangeThemeRules = [{ id: 'other-rule', themeName: 'edited', range: range(0, 5) }]
  model.sheets[1].rangeProtectionRules = [{ id: 'locked', ranges: [range(5)], allowedUserIds: ['owner'] }]
  const state = host(model)
  try {
    state.controller.openRangeThemeDialog(); const before = cloneModel(state.view.value), next = definition(); next.wholeStyle.font.size = 30
    state.controller.confirmRangeTheme('apply', 'edited', next)
    assert.match(state.controller.rangeThemeError.value, /保护/); assert.deepEqual(state.view.value, before)
    assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0); assert.ok(state.controller.request.value)
  } finally { state.dispose() }
})

test('preparation that leaves and reenters the same host cannot publish the old theme candidate', () => {
  const state = host()
  try {
    const before = cloneModel(state.view.value)
    state.context.prepareView = value => { state.active.value = false; state.active.value = true; return value }
    assert.equal(state.controller.runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { themeName: 'edited', rangeThemeStyleJson: definition() }), false)
    assert.deepEqual(state.view.value, before); assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
  } finally { state.dispose() }
})

test('a publication observer cannot send its uncommitted detached edits back through the parent callback', () => {
  const state = host()
  try {
    state.context.applyView = value => { state.view.value = value; state.calls.applied++; state.view.value.sheets[0].cells.B10 = { value: 'observer' } }
    assert.equal(state.controller.runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { themeName: 'edited', rangeThemeStyleJson: definition() }), true)
    assert.equal(state.calls.applied, 1); assert.equal(state.calls.changed.length, 0)
    assert.equal(state.view.value.sheets[0].cells.B10.value, 'observer'); assert.equal(state.view.value.sheets[0].rangeThemeRules[0].themeName, 'edited')
  } finally { state.dispose() }
})

test('the parent publication guard expires after in-place edits and after a new activation', () => {
  const state = host()
  try {
    state.controller.runRangeThemeCommand(SHEET_RANGE_THEME_COMMANDS.set, { themeName: 'edited', rangeThemeStyleJson: definition() })
    const current = state.calls.changed[0].current; assert.equal(current(), true)
    state.view.value.sheets[0].cells.B10 = { value: 'newer' }; assert.equal(current(), false)
    delete state.view.value.sheets[0].cells.B10; assert.equal(current(), true)
    state.active.value = false; state.active.value = true; assert.equal(current(), false)
  } finally { state.dispose() }
})

test('a changed permission user cancels the old dialog and its late close does not steal focus', async () => {
  const state = host()
  try {
    state.controller.openRangeThemeDialog(); state.user.value = 'owner'
    assert.equal(state.controller.request.value, null); state.controller.afterRangeThemeClose(); await nextTick(); assert.equal(state.calls.focused, 0)
    state.context.finishEdit = () => false; state.controller.openRangeThemeDialog(); assert.equal(state.controller.request.value, null)
  } finally { state.dispose() }
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run) {
  const model = sample(), studio = new OfficeStudio({ workbook: model }), api = FOffice.newAPI(studio), book = api.getActiveWorkbook(), state = host(model, studio)
  try { return run(state, studio, book) } finally { state.dispose(); studio.dispose() }
}

test('Studio draft confirmation commits one history step and publishes only the matching candidate', () => withStudio((state, studio, book) => {
  state.controller.openRangeThemeDialog(); const revision = studio.getWorkbookRevision(book.getId())
  state.controller.confirmRangeTheme('apply', 'edited', definition())
  assert.equal(studio.getWorkbookRevision(book.getId()), revision + 1); assert.equal(studio.editor.history.units[book.getId()].undo.length, 1)
  assert.equal(state.calls.applied, 1); assert.equal(state.calls.changed.length, 1); assert.equal(state.calls.changed[0].current(), true)
  assert.equal(studio.getWorkbookModel(book.getId()).rangeThemes.edited.wholeStyle.font.size, 22)
}))

test('history reset after opening leaves Studio content intact and reports a stale draft', () => withStudio((state, studio, book) => {
  state.controller.openRangeThemeDialog(); const before = cloneModel(book.save()); studio.editor.history.clear(book.getId())
  state.controller.confirmRangeTheme('apply', 'edited', definition())
  assert.match(state.controller.rangeThemeError.value, /历史已更改/); assert.deepEqual(book.save(), before)
  assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0); assert.ok(state.controller.request.value)
}))

test('replacing a Studio model under the same ID invalidates a captured custom-theme confirmation', () => withStudio((state, studio, book) => {
  state.controller.openRangeThemeDialog(); studio.editor.workbook.replace(cloneModel(studio.getWorkbookModel(book.getId())))
  state.controller.confirmRangeTheme('apply', 'edited', definition())
  assert.match(state.controller.rangeThemeError.value, /实例已更改/); assert.equal(Object.hasOwn(book.save().rangeThemes ?? {}, 'edited'), false)
  assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0)
}))

test('a newer edit from a Studio command observer is preserved without publishing the previous revision view', () => withStudio((state, studio, book) => {
  const listener = studio.events.addEvent('CommandExecuted', event => {
    if (event.id === SHEET_RANGE_THEME_COMMANDS.set) book.getActiveSheet().getRange('F1').setValue('newer observer')
  })
  try {
    state.controller.openRangeThemeDialog(); const revision = studio.getWorkbookRevision(book.getId()); state.controller.confirmRangeTheme('apply', 'edited', definition())
    assert.equal(studio.getWorkbookRevision(book.getId()), revision + 2); assert.equal(book.getActiveSheet().getRange('F1').getValue(), 'newer observer')
    assert.equal(state.calls.applied, 0); assert.equal(state.calls.changed.length, 0); assert.equal(book.save().rangeThemes.edited.wholeStyle.font.size, 22)
  } finally { listener.dispose() }
}))

test('Studio applyView reentry cannot contaminate the committed model through a parent view update', () => withStudio((state, _studio, book) => {
  state.context.applyView = value => { state.view.value = value; state.calls.applied++; state.view.value.sheets[0].cells.F1 = { value: 'uncommitted observer' } }
  state.controller.openRangeThemeDialog(); state.controller.confirmRangeTheme('apply', 'edited', definition())
  assert.equal(state.calls.applied, 1); assert.equal(state.calls.changed.length, 0); assert.equal(book.getActiveSheet().getRange('F1').getValue(), null)
  assert.equal(state.view.value.sheets[0].cells.F1.value, 'uncommitted observer'); assert.equal(book.save().rangeThemes.edited.wholeStyle.font.size, 22)
}))
