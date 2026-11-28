import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { useSheetTableWorkbench } from '../src/components/spreadsheet/use-sheet-table-workbench.ts'
import { SHEET_TABLE_COMMANDS } from '../src/features/sheet-tables/sheet-table-create.ts'

// Implementation-phase drafts; no test runner or dev server has been started.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
})
afterAll(async () => { await vite?.close() })
const range = { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }
function host() {
  const model = createBlankWorkbook(), sheet = model.sheets[0], other = createBlankSheet('Other')
  model.id = 'table-host'; model.calculationMode = 'manual'; model.activeSheetId = sheet.id = 'data'; other.id = 'other'; model.sheets.push(other)
  sheet.cells['0:0'] = createCell(0, 0, '部门'); sheet.cells['0:1'] = createCell(0, 1, '数量')
  sheet.cells['1:0'] = createCell(1, 0, '研发'); sheet.cells['1:1'] = createCell(1, 1, 0)
  sheet.cells['2:0'] = createCell(2, 0, '财务'); sheet.cells['2:1'] = createCell(2, 1, false)
  const studio = new OfficeStudio({ workbook: model }), book = FOffice.newAPI(studio).getActiveWorkbook(), scope = effectScope()
  const state = shallowReactive({ studio, unitId: model.id, sheetId: sheet.id, active: true, blocked: false, userId: studio.currentUserId, ranges: [range] })
  const refreshed = [], focused = [], options = { state: () => state, refresh: (_studio, id) => refreshed.push(id),
    focus: () => focused.push(state.unitId), canRestoreFocus: () => !state.blocked }
  const controller = scope.run(() => useSheetTableWorkbench(options))
  return { studio, book, scope, state, options, controller, refreshed, focused,
    model: () => studio.getWorkbookModel(model.id), history: () => studio.editor.history.units[model.id]?.undo.length ?? 0,
    dispose: () => { scope.stop(); studio.dispose() } }
}

test('opening captures an independent request and blocked API requests cannot replace the active draft', () => {
  const h = host()
  try {
    const before = cloneModel(h.model())
    h.state.ranges = [{ startRow: 7, endRow: 8, startCol: 4, endCol: 5 }, range]
    assert.equal(h.controller.openSelector(), true); assert.deepEqual(h.controller.request.value.range, range)
    const request = h.controller.request.value
    h.state.ranges = [{ startRow: 9, endRow: 10, startCol: 6, endCol: 7 }]
    assert.equal(h.controller.openSelector(), false); assert.equal(h.controller.request.value, request)
    assert.deepEqual(h.model(), before); assert.equal(h.history(), 0)
    h.controller.cancel(); h.state.blocked = true
    assert.equal(h.controller.openSelector(), false)
    h.studio.executeCommandSync(SHEET_TABLE_COMMANDS.openSelector, { range })
    assert.equal(h.controller.open.value, false); assert.equal(h.controller.request.value, undefined)
  } finally { h.dispose() }
})

test('confirmation publishes one real table/history item and focuses only after the close callback', async () => {
  const h = host()
  try {
    h.controller.openSelector(); assert.equal(h.controller.confirm('A1:B3'), true)
    assert.equal(h.model().tables.length, 1); assert.equal(h.history(), 1); assert.deepEqual(h.refreshed, ['table-host'])
    assert.equal(h.controller.open.value, false); assert.deepEqual(h.focused, [])
    h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, ['table-host'])
    h.controller.afterClose(); await nextTick(); assert.equal(h.focused.length, 1)
    const added = cloneModel(h.model().tables); h.book.undo(); assert.equal(h.model().tables?.length ?? 0, 0)
    h.book.redo(); assert.deepEqual(h.model().tables, added)
  } finally { h.dispose() }
})

test('cancel writes no content/history and defers legal focus while a competing input blocks focus', async () => {
  const h = host()
  try {
    const before = cloneModel(h.model()); h.controller.openSelector(); h.controller.cancel()
    assert.deepEqual(h.focused, []); assert.equal(h.history(), 0); assert.deepEqual(h.model(), before)
    h.state.blocked = true; h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, [])
    h.state.blocked = false; h.controller.openSelector(); h.controller.cancel(); h.controller.afterClose(); await nextTick()
    assert.deepEqual(h.focused, ['table-host']); assert.deepEqual(h.refreshed, [])
  } finally { h.dispose() }
})

test('invalid ranges retain the draft and visible error; successful correction uses the same captured owner', () => {
  const h = host()
  try {
    h.controller.openSelector(); const request = h.controller.request.value
    assert.equal(h.controller.confirm('A1:B1'), false); assert.equal(h.controller.open.value, true)
    assert.equal(h.controller.request.value, request); assert.ok(h.controller.error.value); assert.equal(h.history(), 0)
    assert.equal(h.controller.confirm('A1:B3'), true); assert.equal(h.history(), 1); assert.equal(h.controller.error.value, '')
  } finally { h.dispose() }
})

test('sheet, workbook, user and active-session changes cancel pending requests without focusing a replacement', async () => {
  for (const change of [state => { state.sheetId = 'other' }, state => { state.unitId = 'replacement' },
    state => { state.userId = 'replacement-user' }, state => { state.active = false; state.active = true }]) {
    const h = host()
    try {
      h.controller.openSelector(); change(h.state)
      assert.equal(h.controller.open.value, false); assert.equal(h.controller.request.value, undefined)
      h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, []); assert.equal(h.history(), 0)
    } finally { h.dispose() }
  }
})

test('history replacement and content revision changes reject old drafts and their delayed focus', async () => {
  for (const change of [h => h.studio.editor.history.clear('table-host'),
    h => h.book.getActiveSheet().getRange('G10').setValue('external revision')]) {
    const h = host()
    try {
      h.controller.openSelector(); change(h); const before = cloneModel(h.model()), history = h.history()
      assert.equal(h.controller.confirm('A1:B3'), false); assert.match(h.controller.error.value, /更改/)
      assert.deepEqual(h.model(), before); assert.equal(h.history(), history)
      h.controller.cancel(); h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, [])
    } finally { h.dispose() }
  }
})

test('a replacement model with the same workbook ID cannot receive the previous dialog commit or focus', async () => {
  const h = host()
  try {
    h.controller.openSelector(); h.studio.editor.workbook.replace(cloneModel(h.model()))
    assert.equal(h.controller.confirm('A1:B3'), false); assert.equal(h.model().tables?.length ?? 0, 0)
    h.controller.cancel(); h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('command cancellation retains the request and disposal removes its event subscription and late focus', async () => {
  const h = host()
  try {
    const stop = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_TABLE_COMMANDS.add) event.cancel = true })
    h.controller.openSelector(); const request = h.controller.request.value
    assert.equal(h.controller.confirm('A1:B3'), false); assert.equal(h.controller.request.value, request); assert.equal(h.history(), 0)
    stop.dispose(); h.controller.cancel(); h.scope.stop(); h.controller.afterClose(); await nextTick()
    assert.deepEqual(h.focused, []); assert.equal(h.studio.events.hasListeners('SheetTableSelectorRequested'), false)
  } finally { h.dispose() }
})

test('reentrant CommandExecuted host switches leave the committed table on its owner and skip stale UI refresh/focus', async () => {
  const h = host()
  try {
    h.studio.events.addEvent('CommandExecuted', event => {
      if (event.id === SHEET_TABLE_COMMANDS.add) h.studio.createWorkbook({ id: 'replacement-host' })
    })
    h.controller.openSelector(); assert.equal(h.controller.confirm('A1:B3'), true)
    assert.equal(h.model().tables.length, 1); assert.equal(h.studio.getWorkbookModel('replacement-host').tables?.length ?? 0, 0)
    assert.equal(h.controller.open.value, false); assert.deepEqual(h.refreshed, [])
    h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})

test('same-ID model replacement during UI refresh closes the obsolete request without repeating the completed commit', async () => {
  const h = host()
  try {
    h.options.refresh = () => h.studio.editor.workbook.replace(cloneModel(h.model()))
    h.controller.openSelector(); assert.equal(h.controller.confirm('A1:B3'), true)
    assert.equal(h.model().tables.length, 1); assert.equal(h.controller.open.value, false)
    h.controller.afterClose(); await nextTick(); assert.deepEqual(h.focused, [])
  } finally { h.dispose() }
})
