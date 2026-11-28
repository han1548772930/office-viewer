import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { createSheetTableCommandCandidate, SHEET_TABLE_COMMANDS, sheetTableCreationRange, sheetTableSelectorRange } from '../src/features/sheet-tables/sheet-table-create.ts'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'

// Implementation-phase drafts. No test runner, dev server or browser has been started.
const range = (startRow = 0, endRow = 2, startCol = 0, endCol = 1) => ({ startRow, endRow, startCol, endCol })
function sample() {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'table-command'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.name = 'Data'; sheet.rowCount = 50; sheet.columnCount = 16
  sheet.cells['0:0'] = createCell(0, 0, '部门'); sheet.cells['0:1'] = createCell(0, 1, true)
  sheet.cells['1:0'] = createCell(1, 0, '研发'); sheet.cells['1:1'] = createCell(1, 1, 0)
  sheet.cells['2:0'] = createCell(2, 0, '财务'); sheet.cells['2:1'] = createCell(2, 1, false)
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other)
  return model
}
const labels = { tablePrefix: 'Table', columnPrefix: 'Column' }

test('single-cell selector expands connected zero/false data and multiple-cell selections retain their exact range', () => {
  const sheet = sample().sheets[0]
  assert.deepEqual(sheetTableSelectorRange(sheet, range(1, 1, 1, 1)), range())
  assert.deepEqual(sheetTableSelectorRange(sheet, range(1, 3, 0, 2)), range(1, 3, 0, 2))
  assert.deepEqual(sheetTableSelectorRange(sheet, range(10, 10, 10, 10)), range(10, 10, 10, 10))
})

test('creation accepts a single local address or public range DTO and rejects malformed/cross-sheet/oversize input', () => {
  const sheet = sample().sheets[0]
  assert.deepEqual(sheetTableCreationRange(sheet, '$A$1:$B$3'), range())
  assert.deepEqual(sheetTableCreationRange(sheet, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }), range())
  for (const input of ['Data!A1:B3', 'A1:B3,D1:E3', 'A0:B3', 'A1:ZZ99', null, {}, { ...range(), endCol: 16_384 }]) {
    assert.throws(() => createSheetTableCommandCandidate(sample(), 'data', { range: input }, labels))
  }
})

test('default names avoid case-insensitive table and defined-name collisions while invalid public names fall back', () => {
  const model = sample(); model.metadata.definedNames = [{ name: 'TABLE1' }, { name: 'Table7' }]
  for (const name of [undefined, '', 'A1', 'R1C1', 'space name', 'Table7']) {
    const next = createSheetTableCommandCandidate(model, 'data', { name, range: range() }, labels)
    assert.equal(next.tables[0].name, 'Table8')
  }
  const next = createSheetTableCommandCandidate(model, 'data', { name: '部门表', range: range() }, labels)
  assert.equal(next.tables[0].name, '部门表'); assert.deepEqual(next.tables[0].columns.map(column => column.name), ['部门', 'TRUE'])
  assert.equal(model.tables?.length ?? 0, 0)
})

test('the fixed table root occupies Data organization row1/column2 and keeps the enhanced outline menu separate', () => {
  const ribbon = createSheetRibbon({ hasSheet: true, tableAvailable: true, tableEditable: true,
    canUndo: false, canRedo: false, canMerge: true, canUnmerge: false, canDeleteRow: true, canDeleteColumn: true,
    chartAvailable: false, dataValidationAvailable: false, canSort: false, fontFamily: 'Aptos', fontSize: 11,
    numberFormat: 'general', gridlines: true, frozen: false, zoom: 100, fullscreen: false })
  const group = ribbon.find(tab => tab.id === 'data').groups.find(group => group.id === 'organization')
  const root = group.items.find(item => item.id === SHEET_TABLE_COMMANDS.openSelector)
  assert.equal(root.commandId, SHEET_TABLE_COMMANDS.openSelector); assert.equal(root.order, 2)
  assert.deepEqual(root.gridLayout, { row: 1, column: 2 }); assert.equal(root.disabled, false)
  assert.deepEqual(group.items.find(item => item.id === 'sheet.menu.outline-details').gridLayout, { row: 1, column: 4 })
  assert.equal(createSheetRibbon({ ...{ hasSheet: true }, tableAvailable: true, tableEditable: false }).find(tab => tab.id === 'data')
    .groups.find(group => group.id === 'organization').items.find(item => item.id === SHEET_TABLE_COMMANDS.openSelector).disabled, true)
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
})
afterAll(async () => { await vite?.close() })
function host() {
  const studio = new OfficeStudio({ workbook: sample() }), book = FOffice.newAPI(studio).getActiveWorkbook()
  return { studio, book, model: () => studio.getWorkbookModel('table-command'), history: () => {
    const unit = studio.editor.history.units['table-command']; return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0 }
  } }
}

test('opening emits a detached request for the last selection and cancelling by omitting add writes no content/history', () => {
  const state = host(), requests = []
  try {
    state.studio.setSelections('table-command', 'data', [
      { startRow: 8, endRow: 9, startColumn: 6, endColumn: 7 }, { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
    ])
    const before = cloneModel(state.model()), history = state.history()
    const listener = state.studio.events.addEvent('SheetTableSelectorRequested', value => requests.push(value))
    assert.equal(state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.openSelector), true)
    assert.deepEqual(requests[0].range, range()); assert.equal(requests[0].sheetId, 'data')
    requests[0].range.endRow = 40
    assert.deepEqual(state.model(), before); assert.deepEqual(state.history(), history)
    listener.dispose(); assert.equal(state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.openSelector), false)
  } finally { state.studio.dispose() }
})

test('the registered add command creates one content history item and undo/redo plus internal save/reopen preserve the table', () => {
  const state = host()
  try {
    const before = cloneModel(state.model())
    assert.equal(state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { unitId: 'table-command', subUnitId: 'data', id: 'table-1', range: 'A1:B3' }), true)
    assert.equal(state.history().undo, 1); assert.equal(state.model().tables[0].id, 'table-1')
    assert.equal(state.model().tables[0].showHeaderRow, true); assert.equal(state.model().tables[0].showTotalsRow, false)
    assert.deepEqual(state.model().sheets[0].cells, before.sheets[0].cells)
    const added = cloneModel(state.model().tables)
    state.book.undo(); assert.equal(state.model().tables?.length ?? 0, 0)
    state.book.redo(); assert.deepEqual(state.model().tables, added)
    const saved = state.book.save(), reopened = state.studio.createWorkbook({ ...saved, id: 'table-reopened' }, false)
    assert.deepEqual(reopened.tables, added)
  } finally { state.studio.dispose() }
})

test('single-row, merge, duplicate-ID and overlapping table failures preserve content and history', () => {
  const state = host()
  try {
    for (const input of [{ range: 'A1:B1' }, { range: 'A1:Q3' }, { id: '', range: 'A1:B3' }]) {
      const before = cloneModel(state.model()), history = state.history()
      assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { unitId: 'table-command', sheetId: 'data', ...input }))
      assert.deepEqual(state.model(), before); assert.deepEqual(state.history(), history)
    }
    state.model().sheets[0].mergedRanges = [range(1, 1, 0, 1)]
    assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { range: 'A1:B3' }))
    state.model().sheets[0].mergedRanges = []
    state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { id: 'existing', range: 'A1:B3' })
    const before = cloneModel(state.model()), history = state.history()
    for (const input of [{ id: 'existing', range: 'D1:E3' }, { range: 'B2:C4' }]) {
      assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, input))
      assert.deepEqual(state.model(), before); assert.deepEqual(state.history(), history)
    }
  } finally { state.studio.dispose() }
})

test('command-layer view/edit protection rejects the full proposed range even when the original selection is editable', () => {
  const state = host()
  try {
    state.model().sheets[0].rangeProtectionRules = [{ id: 'locked-tail', ranges: [range(2, 2, 1, 1)], allowedUserIds: ['owner'] }]
    const before = cloneModel(state.model()), history = state.history()
    assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { range: 'A1:B3' }), /保护|权限/)
    assert.deepEqual(state.model(), before); assert.deepEqual(state.history(), history)
  } finally { state.studio.dispose() }
})

test('stale revision/history/user and switched workbench owners reject without creating a table', () => {
  const state = host()
  try {
    const revision = state.studio.getWorkbookRevision('table-command'), generation = state.studio.editor.history.captureSavepoint('table-command').generation
    for (const guard of [{ expectedRevision: revision + 1 }, { expectedHistoryGeneration: generation + 1 }, { expectedUserId: 'other-user' }]) {
      assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { unitId: 'table-command', sheetId: 'data', range: 'A1:B3', ...guard }), /更改/)
    }
    state.book.getSheetBySheetId('other').activate()
    assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { unitId: 'table-command', sheetId: 'data', range: 'A1:B3', host: 'workbench' }), /宿主/)
    assert.equal(state.model().tables?.length ?? 0, 0); assert.equal(state.history().undo, 0)
    assert.equal(state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { unitId: 'table-command', sheetId: 'data', range: 'A1:B3' }), true)
  } finally { state.studio.dispose() }
})

test('malformed context and cancelled BeforeCommandExecute never reach the model mutation', () => {
  const state = host()
  try {
    for (const input of [null, [], { unitId: '' }, { host: 'other' }, { expectedHistoryGeneration: 0 }, { sheetId: 'data', subUnitId: 'other' }]) {
      assert.throws(() => state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.openSelector, input))
    }
    const before = cloneModel(state.model()), stop = state.studio.events.addEvent('BeforeCommandExecute', event => {
      if (event.id === SHEET_TABLE_COMMANDS.add) event.cancel = true
    })
    assert.equal(state.studio.executeCommandSync(SHEET_TABLE_COMMANDS.add, { range: 'A1:B3' }), false)
    assert.deepEqual(state.model(), before); assert.equal(state.history().undo, 0); stop.dispose()
  } finally { state.studio.dispose() }
})
