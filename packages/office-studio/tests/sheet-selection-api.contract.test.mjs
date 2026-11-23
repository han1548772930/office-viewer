import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { parseSheetA1Range, splitSheetA1Reference } from '../src/features/sheet-workbench/sheet-a1-reference.ts'
import { parseSheetSelectionShortcut } from '../src/api/sheet-selection-commands.ts'

let server, module
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  module = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const coords = (startRow, startColumn, endRow = startRow, endColumn = startColumn, rangeType) => ({ startRow, startColumn, endRow, endColumn,
  ...(rangeType === undefined ? {} : { rangeType }) })
function runtime(t, extra = {}) {
  const model = module.createBlankWorkbook(); model.id = 'selection-book'; model.calculationMode = 'manual'
  model.sheets = ['First', "销售 ! O'Brien"].map((name, index) => ({ ...module.createBlankSheet(name), id: `selection-sheet-${index}`,
    rowCount: 20, columnCount: 8, defaultRowHeight: 26, defaultColumnWidth: 90 }))
  model.activeSheetId = model.sheets[0].id
  const studio = new module.OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8, ...extra })
  t.onTestFinished(() => studio.dispose())
  const api = module.FOffice.newAPI(studio), book = api.getActiveWorkbook(), [first, second] = book.getSheets()
  return { studio, api, book, first, second }
}
function history(studio, id = studio.activeWorkbookId) {
  const unit = studio.editor.history.units[id]
  return { undo: unit?.undo.length ?? 0, redo: unit?.redo.length ?? 0, revision: studio.getWorkbookRevision(id) }
}

test('A1 boundaries cover quoted sheet names, escaped apostrophes, whole axes and Office limits', () => {
  assert.deepEqual(splitSheetA1Reference("'销售 ! O''Brien'!$C:$E"), { reference: '$C:$E', sheetName: "销售 ! O'Brien" })
  assert.deepEqual(parseSheetA1Range('$3:$7', { rowCount: 20, columnCount: 8 }), coords(2, 0, 6, 7, 1))
  assert.deepEqual(parseSheetA1Range("'销售 ! O''Brien'!$C:$E", { rowCount: 20, columnCount: 8 }), coords(0, 2, 19, 4, 2))
  assert.deepEqual(parseSheetA1Range('XFD1048576'), coords(1_048_575, 16_383))
  for (const reference of ['', 'A0', 'XFE1', 'A1048577', "'bad'quote'!A1", "'Sheet 1'!", 'Sheet 1!A1',
    '[Other.xlsx]Sheet1!A1', 'Sheet1:Sheet2!A1', 'A1,B2']) assert.throws(() => parseSheetA1Range(reference))
  for (const bounds of [{ rowCount: 0 }, { columnCount: 16_385 }, { rowCount: Infinity }]) assert.throws(() => parseSheetA1Range('A1', bounds), RangeError)
})

test('sheet-scoped API operations retain the worksheet identity resolved from A1 and copy external coordinate objects', async t => {
  const { api, book, first, second, studio } = runtime(t)
  const crossSheet = first.getRange("'销售 ! O''Brien'!C:C")
  assert.equal(crossSheet.getSheetId(), second.getSheetId())
  assert.equal(crossSheet.getA1Notation(true), "'销售 ! O''Brien'!C:C")
  const coordinates = coords(1, 1, 2, 2), range = first.getRange(coordinates)
  coordinates.endRow = 18
  assert.deepEqual(range.getRange(), coords(1, 1, 2, 2))
  const exported = range.getRange(); exported.startColumn = 7
  assert.deepEqual(range.getRange(), coords(1, 1, 2, 2))
  const other = api.createWorkbook({ id: 'foreign-book' }, { makeCurrent: false }).getActiveSheet().getRange('A1')
  const initial = copy(studio.getWorkbookModel(book.getId())), initialHistory = history(studio)
  for (const action of [() => first.setActiveRange(crossSheet), () => first.setActiveRange(other), () => first.setActiveRangeList([range, crossSheet]),
    () => first.setAutoFilter(crossSheet), () => first.clearSparklines(crossSheet), () => first.setDataValidation(crossSheet, null)]) {
    assert.throws(action, /当前工作表/)
  }
  await assert.rejects(() => first.sortRange(crossSheet, [{ column: 0, direction: 'ascending' }]), /当前工作表/)
  assert.deepEqual(copy(studio.getWorkbookModel(book.getId())), initial)
  assert.deepEqual(history(studio), initialHistory)
})

test('complete multiple selections and the non-last primary round-trip through Facade and events without replacing content', t => {
  const { studio, api, book, first } = runtime(t)
  const model = studio.getWorkbookModel(book.getId()), before = history(studio), events = [], legacy = []
  api.addEvent(api.Event.SelectionsChanged, event => events.push(copy(event)))
  api.addEvent(api.Event.SelectionChanged, event => legacy.push(copy(event)))
  first.setActiveRangeList([first.getRange('A1:C3'), first.getRange('B2:D4'), first.getRange('F10')],
    { activeRangeIndex: 0, primary: { row: 1, column: 1 } })
  const state = studio.getSelectionState()
  assert.equal(state.activeRangeIndex, 0)
  assert.deepEqual(state.primary, { row: 1, column: 1 })
  assert.deepEqual(book.getActiveRangeList().map(range => range.getA1Notation()), ['A1:C3', 'B2:D4', 'F10'])
  assert.equal(book.getActiveRange().getA1Notation(), 'A1:C3')
  assert.equal(book.getActiveCell().getA1Notation(), 'B2')
  assert.deepEqual(first.getSelection().getCurrentCell(), { ...coords(1, 1, 1, 1, 0), actualRow: 1, actualColumn: 1, isMerged: false, isMergedMainCell: false })
  assert.equal(studio.getWorkbookModel(book.getId()), model)
  assert.deepEqual(history(studio), before)
  assert.equal(events.length, 1); assert.equal(legacy.length, 1)
  studio.setSelections(book.getId(), first.getSheetId(), state.ranges, state)
  assert.equal(events.length, 1)
  state.ranges[0].endRow = 19; state.primary.column = 7
  assert.equal(studio.getSelection().endRow, 2)
  assert.deepEqual(studio.getSelectionState().primary, { row: 1, column: 1 })
})

test('selection snapshots are immutable; primary updates and navigation match the fixed Facade contract', t => {
  const { api, first, studio } = runtime(t)
  first.getRange('A1:B2').activate()
  const original = first.getSelection(), before = history(studio)
  assert.equal(original.getNextDataRange(api.Enum.Direction.RIGHT).getA1Notation(), 'B1')
  const moved = original.updatePrimaryCell(first.getRange('B1'))
  assert.deepEqual(original.getCurrentCell(), { ...coords(0, 0, 0, 0, 0), actualRow: 0, actualColumn: 0, isMerged: false, isMergedMainCell: false })
  assert.equal(moved.getCurrentCell().actualColumn, 1)
  assert.equal(moved.getNextDataRange(api.Enum.Direction.RIGHT).getA1Notation(), 'A2')
  const outside = moved.updatePrimaryCell(first.getRange('C3'))
  assert.equal(outside.getActiveRange().getA1Notation(), 'C3')
  assert.equal(original.getActiveRange().getA1Notation(), 'A1:B2')
  const mutableList = original.getActiveRangeList(); mutableList.length = 0
  const mutableCell = original.getCurrentCell(); mutableCell.actualColumn = 7
  assert.equal(original.getActiveRangeList().length, 1); assert.equal(original.getCurrentCell().actualColumn, 0)
  assert.deepEqual(history(studio), before)
})

test('an explicit active index distinguishes identical overlapping ranges and merged primaries expose merge bounds', t => {
  const { studio, api, book, first } = runtime(t)
  first.setActiveRangeList([first.getRange('A1'), first.getRange('A1'), first.getRange('D4')], { activeRangeIndex: 1 })
  assert.equal(first.getSelection().getNextDataRange(api.Enum.Direction.RIGHT).getA1Notation(), 'D4')
  const model = copy(studio.getWorkbookModel(book.getId()))
  model.sheets[0].mergedRanges = [{ startRow: 5, endRow: 6, startCol: 1, endCol: 2 }]
  studio.editor.workbook.replace(model)
  first.setActiveRangeList([first.getRange('B6:C7')], { primary: { row: 6, column: 2 } })
  const primary = first.getSelection().getCurrentCell()
  assert.deepEqual(primary, { ...coords(5, 1, 6, 2, 0), actualRow: 6, actualColumn: 2, isMerged: true, isMergedMainCell: false })
  assert.equal(book.getActiveCell().getA1Notation(), 'B6:C7')
})

test('shared formatting targets every selected rectangle and undo restores the entire selection without editing the gap', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('A1:D2').setValues([[1, 2, 3, 4], [5, 6, 7, 8]])
  studio.editor.history.clear(book.getId())
  first.setActiveRangeList([first.getRange('A1:A2'), first.getRange('D1:D2')], { activeRangeIndex: 0, primary: { row: 1, column: 0 } })
  const before = studio.getSelectionState()
  api.executeCommandSync('sheet.command.toggle-bold')
  assert.deepEqual(first.getRange('A1:D2').getCellDatas().map(row => row.map(cell => cell.style.font.bold)), [[true, false, false, true], [true, false, false, true]])
  assert.equal(history(studio).undo, 1)
  first.getRange('G7').activate()
  book.undo()
  assert.deepEqual(studio.getSelectionState(), before)
  assert.deepEqual(first.getRange('A1:D2').getCellDatas().map(row => row.map(cell => cell.style.font.bold)), [[false, false, false, false], [false, false, false, false]])
  book.redo()
  assert.deepEqual(studio.getSelectionState(), before)
})

test('each worksheet and workbook retains its complete selection; navigation never advances content revisions', t => {
  const { studio, api, book, first, second } = runtime(t)
  first.setActiveRangeList([first.getRange('A1:B2'), first.getRange('F10')], { activeRangeIndex: 0, primary: { row: 1, column: 1 } })
  const saved = studio.getSelectionState(), initial = history(studio)
  second.activate(); second.getRange('C4:D5').activate()
  const secondSaved = studio.getSelectionState()
  const neighbor = api.createWorkbook({ id: 'selection-neighbor' }); neighbor.getActiveSheet().getRange('E8').activate()
  const neighborRevision = studio.getWorkbookRevision(neighbor.getId())
  studio.activateWorkbook(book.getId())
  assert.deepEqual(studio.getSelectionState(), secondSaved)
  first.activate()
  assert.deepEqual(studio.getSelectionState(), saved)
  studio.activateWorkbook(neighbor.getId())
  assert.equal(studio.getSelectionState().primary.row, 7)
  assert.deepEqual(history(studio, book.getId()), initial)
  assert.equal(studio.getWorkbookRevision(neighbor.getId()), neighborRevision)
})

test('structural history saves a complete before selection and preserves a later active worksheet on undo and redo', t => {
  const { studio, api, book, first, second } = runtime(t)
  first.getRange('A1:A4').setValues([['header'], ['one'], ['two'], ['three']]); studio.editor.history.clear(book.getId())
  first.setActiveRangeList([first.getRange('2:3'), first.getRange('6:7')], { activeRangeIndex: 0, primary: { row: 1, column: 2 } })
  const before = studio.getSelectionState()
  api.executeCommandSync('sheet.command.insert-multi-rows-above', { unitId: book.getId(), subUnitId: first.getSheetId(), range: coords(1, 0, 2, 7, 1), value: 2 })
  const after = studio.getSelectionState()
  assert.equal(first.getMaxRows(), 22)
  assert.equal(first.getRange('A4').getValue(), 'one')
  second.getRange('C4:D5').activate(); const later = studio.getSelectionState()
  book.undo()
  assert.equal(book.getActiveSheet().getSheetId(), second.getSheetId())
  assert.deepEqual(studio.getSelectionState(), later)
  assert.deepEqual(studio.getSelectionState(book.getId(), first.getSheetId()), before)
  assert.equal(first.getMaxRows(), 20); assert.equal(first.getRange('A2').getValue(), 'one')
  book.redo()
  assert.equal(book.getActiveSheet().getSheetId(), second.getSheetId())
  assert.deepEqual(studio.getSelectionState(book.getId(), first.getSheetId()), after)
  assert.equal(history(studio).undo, 1)
})

test('invalid, denied and canceled selections leave model identity, view state and content history untouched', t => {
  const { studio, api, book, first, second } = runtime(t)
  first.getRange('B2:C3').activate()
  const source = studio.getWorkbookModel(book.getId()), before = studio.getSelectionState(), depth = history(studio)
  const unchanged = () => { assert.equal(studio.getWorkbookModel(book.getId()), source); assert.deepEqual(studio.getSelectionState(), before); assert.deepEqual(history(studio), depth) }
  for (const ranges of [[coords(20, 0)], [coords(-1, 0)], [{ ...coords(0, 0), sheetId: second.getSheetId() }], [coords(0.5, 0)]]) {
    assert.throws(() => studio.setSelections(book.getId(), first.getSheetId(), ranges)); unchanged()
  }
  const canceled = api.addEvent(api.Event.BeforeCommandExecute, event => { if (event.id === 'sheet.operation.set-selections') event.cancel = true })
  assert.throws(() => second.getRange('D4').activate(), /canceled/); unchanged(); canceled.dispose()
  const locked = copy(source)
  locked.sheets[1].worksheetProtection = { id: 'private', permissionId: 'private-p', ownerUserId: 'owner', unitType: 'worksheet',
    editState: 'onlyMe', viewState: 'noOneElseCanView', allowedUserIds: [] }
  studio.editor.workbook.replace(locked)
  assert.equal(studio.getSelectionState(book.getId(), second.getSheetId()), null)
  const lockedBefore = studio.getSelectionState()
  assert.throws(() => studio.setSelections(book.getId(), second.getSheetId(), [coords(0, 0)]), /查看权限/)
  assert.equal(book.getActiveSheet().getSheetId(), first.getSheetId())
  assert.deepEqual(studio.getSelectionState(), lockedBefore)
})

test('empty selections, subscriber mutation and reentrant events cannot leave a stale legacy primary', t => {
  const { studio, api, book, first } = runtime(t)
  studio.setSelections(book.getId(), first.getSheetId(), [])
  assert.equal(book.getActiveRange(), null); assert.equal(book.getActiveCell(), null); assert.deepEqual(book.getActiveRangeList(), [])
  const legacy = [], complete = []
  api.addEvent(api.Event.SelectionChanged, event => legacy.push(event.range.startRow))
  api.addEvent(api.Event.SelectionsChanged, event => {
    if (event.selection.primary?.row === 0) {
      event.selection.ranges[0].endRow = 19
      studio.setSelections(book.getId(), first.getSheetId(), [coords(3, 3)])
    }
  })
  api.addEvent(api.Event.SelectionsChanged, event => complete.push(event.selection.primary?.row))
  studio.setSelections(book.getId(), first.getSheetId(), [coords(0, 0)])
  assert.deepEqual(legacy, [3])
  assert.deepEqual(complete, [3])
  assert.equal(studio.getSelection().endRow, 3)
})

test('removed sheets and disposed workbooks discard selection state, including identifiers that are object prototype names', t => {
  const { studio, api, book, first, second } = runtime(t)
  second.getRange('F10').activate()
  studio.commitWorkbook(book.getId(), 'Remove selected worksheet', model => { model.sheets = model.sheets.filter(sheet => sheet.id !== second.getSheetId()); model.activeSheetId = first.getSheetId() })
  assert.equal(studio.sheetSelections.get(book.getId(), second.getSheetId()), undefined)
  const source = copy(studio.getWorkbookModel(book.getId()))
  source.sheets.push({ ...module.createBlankSheet('Recreated'), id: second.getSheetId(), rowCount: 20, columnCount: 8 })
  studio.editor.workbook.replace(source)
  assert.deepEqual(studio.getSelectionState(book.getId(), second.getSheetId()).primary, { row: 0, column: 0 })
  const special = module.createBlankWorkbook(); special.id = '__proto__'; special.sheets[0].id = 'toString'; special.activeSheetId = 'toString'
  const specialBook = api.createWorkbook(special); specialBook.getActiveSheet().getRange('E8').activate()
  assert.deepEqual(studio.getSelectionState().primary, { row: 7, column: 4 })
  api.disposeUnit(specialBook.getId())
  assert.equal(studio.sheetSelections.get('__proto__', 'toString'), undefined)
  assert.equal(studio.getSelectionState('__proto__'), null)
})

test('Enter and Tab use the shared command, retain multiple ranges and do not create content history', t => {
  const { studio, api, book, first } = runtime(t)
  first.setActiveRangeList([first.getRange('A1:B2'), first.getRange('F6')], { activeRangeIndex: 0, primary: { row: 0, column: 1 } })
  const before = history(studio), ranges = copy(studio.getSelectionState().ranges)
  const move = (direction, keycode) => api.executeCommandSync('sheet.command.move-selection-enter-tab', { direction, keycode })
  move(api.Enum.Direction.RIGHT, 9)
  assert.deepEqual(studio.getSelectionState().primary, { row: 1, column: 0 })
  assert.deepEqual(studio.getSelectionState().ranges, ranges)
  move(api.Enum.Direction.DOWN, 13)
  assert.deepEqual(studio.getSelectionState().primary, { row: 0, column: 1 })
  first.setActiveRangeList([first.getRange('A1'), first.getRange('F6')], { activeRangeIndex: 0 })
  move(api.Enum.Direction.RIGHT, 9)
  assert.equal(studio.getSelectionState().activeRangeIndex, 1)
  assert.deepEqual(studio.getSelectionState().primary, { row: 5, column: 5 })
  assert.deepEqual(history(studio), before)
  assert.equal(book.getActiveRangeList().length, 2)
})

test('a single-cell Tab origin survives an unrelated edit, resumes on Enter and resets on explicit selection or undo', t => {
  const { studio, api, book, first, second } = runtime(t)
  const move = (direction, keycode) => api.executeCommandSync('sheet.command.move-selection-enter-tab', { direction, keycode })
  first.getRange('B2').activate()
  move(api.Enum.Direction.RIGHT, 9); move(api.Enum.Direction.RIGHT, 9)
  assert.equal(book.getActiveCell().getA1Notation(), 'D2')
  const origin = studio.sheetSelections.getTabStart(book.getId(), first.getSheetId())
  origin.primary.column = 7
  assert.equal(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()).primary.column, 1)
  first.getRange('H20').setValue('unrelated')
  move(api.Enum.Direction.DOWN, 13)
  assert.equal(book.getActiveCell().getA1Notation(), 'B3')
  assert.equal(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), undefined)
  move(api.Enum.Direction.RIGHT, 9)
  second.activate(); second.getRange('F6').activate(); first.activate()
  move(api.Enum.Direction.DOWN, 13)
  assert.equal(book.getActiveCell().getA1Notation(), 'B4')
  move(api.Enum.Direction.RIGHT, 9); first.getRange('D4').activate(); move(api.Enum.Direction.DOWN, 13)
  assert.equal(book.getActiveCell().getA1Notation(), 'D5')
  move(api.Enum.Direction.RIGHT, 9)
  book.undo()
  assert.equal(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), undefined)
  assert.equal(first.getRange('H20').getValue(), null)
})

test('canceling a navigation operation keeps the current primary and Tab origin; Before failures remain explicit', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('B2').activate()
  api.executeCommandSync('sheet.command.move-selection-enter-tab', { direction: api.Enum.Direction.RIGHT, keycode: 9 })
  const before = studio.getSelectionState(), origin = studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), depth = history(studio)
  const veto = api.addEvent(api.Event.BeforeCommandExecute, event => { if (event.id === 'sheet.operation.set-selections') event.cancel = true })
  assert.throws(() => api.executeCommandSync('sheet.command.move-selection-enter-tab', { direction: api.Enum.Direction.DOWN, keycode: 13 }), /canceled/)
  assert.deepEqual(studio.getSelectionState(), before)
  assert.deepEqual(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), origin)
  assert.deepEqual(history(studio), depth)
  veto.dispose()
  const throwing = api.addEvent(api.Event.BeforeCommandExecute, () => { throw new Error('before observer refused') })
  assert.throws(() => first.getRange('D4').activate(), /before observer refused/)
  throwing.dispose()
  assert.deepEqual(studio.getSelectionState(), before)
  assert.deepEqual(history(studio), depth)
})

test('selection notifications follow completed history and isolate subscriber mutation and failure on execute, undo and redo', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('B2').activate()
  const delivered = [], reported = [], failure = new Error('selection observer failed')
  api.addEvent(api.Event.EventListenerError, event => reported.push(event))
  api.addEvent(api.Event.SelectionsChanged, event => { event.selection.primary.row = 19; event.selection.ranges[0].endRow = 19; throw failure })
  api.addEvent(api.Event.SelectionsChanged, event => delivered.push({ row: event.selection.primary.row, ...history(studio) }))
  studio.commitWorkbook(book.getId(), 'hide selected row', model => { model.sheets[0].hiddenRows = [1] })
  assert.equal(studio.getSelectionState().primary.row, 2)
  assert.deepEqual(delivered.map(({ row, undo, redo }) => ({ row, undo, redo })), [{ row: 2, undo: 1, redo: 0 }])
  book.undo(); book.redo()
  assert.deepEqual(delivered.map(({ row, undo, redo }) => ({ row, undo, redo })),
    [{ row: 2, undo: 1, redo: 0 }, { row: 1, undo: 0, redo: 1 }, { row: 2, undo: 1, redo: 0 }])
  assert.equal(reported.length, 3)
  assert.ok(reported.every(event => event.eventName === 'SelectionsChanged' && event.error === failure))
  assert.equal(studio.getEventListenerErrors().length, 3)
  const errors = studio.getEventListenerErrors(); errors.length = 0
  assert.equal(studio.getEventListenerErrors().length, 3)
})

test('an edit triggered by a selection observer is a later transaction and each undo preserves the other transaction', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('B2').activate()
  let observed = false
  api.addEvent(api.Event.SelectionsChanged, event => {
    if (observed || event.selection.primary?.row !== 2) return
    observed = true
    assert.equal(history(studio).undo, 1)
    first.getRange('D4').setValue('later edit')
  })
  studio.commitWorkbook(book.getId(), 'hide selected row', model => { model.sheets[0].hiddenRows = [1] })
  assert.equal(history(studio).undo, 2)
  assert.equal(first.getRange('D4').getValue(), 'later edit')
  book.undo()
  assert.equal(first.getRange('D4').getValue(), null)
  assert.deepEqual(studio.getWorkbookModel(book.getId()).sheets[0].hiddenRows, [1])
  book.undo()
  assert.deepEqual(studio.getWorkbookModel(book.getId()).sheets[0].hiddenRows, [])
  assert.equal(studio.getSelectionState().primary.row, 1)
})

test('navigation resolves an explicitly named workbook to its own active worksheet when the sheet is omitted', t => {
  const { studio, api, book, first } = runtime(t)
  const neighborModel = module.createBlankWorkbook(); neighborModel.id = 'navigation-neighbor'
  neighborModel.sheets[0].id = 'neighbor-sheet'; neighborModel.activeSheetId = 'neighbor-sheet'
  neighborModel.sheets[0].rowCount = 20; neighborModel.sheets[0].columnCount = 8
  const neighbor = api.createWorkbook(neighborModel, { makeCurrent: false })
  first.getRange('B2').activate()
  const before = history(studio, book.getId()), otherBefore = history(studio, neighbor.getId())
  api.executeCommandSync(module.MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND,
    { unitId: neighbor.getId(), direction: api.Enum.Direction.RIGHT, keycode: 9 })
  assert.equal(studio.activeWorkbookId, neighbor.getId())
  assert.equal(studio.getWorkbookModel(neighbor.getId()).activeSheetId, 'neighbor-sheet')
  assert.deepEqual(studio.getSelectionState().primary, { row: 0, column: 1 })
  assert.deepEqual(history(studio, book.getId()), before); assert.deepEqual(history(studio, neighbor.getId()), otherBefore)
  assert.throws(() => parseSheetSelectionShortcut({ unitId: 'foreign', direction: 1, key: 'Tab' },
    { unitId: 'current', sheetId: 'current-sheet' }), /工作表标识/)
  assert.equal(parseSheetSelectionShortcut({ unitId: 'foreign', direction: 1, key: 'Tab' },
    { unitId: 'current', sheetId: 'current-sheet' }, id => id === 'foreign' ? 'foreign-sheet' : undefined).sheetId, 'foreign-sheet')
})

test('history publication follows the final active workbook and suppresses pending sheet notices when Docs becomes active', t => {
  const { studio, api, book, first } = runtime(t)
  const neighbor = api.createWorkbook({ id: 'publication-neighbor' }, { makeCurrent: false })
  const received = []
  api.addEvent(api.Event.SelectionsChanged, event => received.push({ id: event.workbookId, row: event.selection.primary?.row,
    applying: studio.editor.history.isApplying, depth: history(studio, book.getId()).undo }))
  studio.editor.history.execute(book.getId(), 'publication scope', () => {
    first.getRange('B2').activate()
    neighbor.getActiveSheet().getRange('D4').activate()
  }, () => {})
  assert.deepEqual(received, [{ id: neighbor.getId(), row: 3, applying: false, depth: 1 }])
  received.length = 0
  studio.editor.history.execute(book.getId(), 'Docs publication scope', () => {
    first.getRange('C3').activate()
    studio.activateDocument(studio.activeDocumentId)
  }, () => {})
  assert.equal(studio.editor.state.activeKind, 'document'); assert.deepEqual(received, [])
})

test('a legacy selection subscriber that selects again stops delivery of the older primary to later subscribers', t => {
  const { studio, api, book, first } = runtime(t), received = []
  api.addEvent(api.Event.SelectionChanged, event => {
    if (event.range.startRow === 1) first.getRange('E5').activate()
  })
  api.addEvent(api.Event.SelectionChanged, event => received.push(event.range.startRow))
  first.getRange('B2').activate()
  assert.deepEqual(received, [4]); assert.deepEqual(studio.getSelectionState().primary, { row: 4, column: 4 })
  assert.equal(history(studio, book.getId()).undo, 0)
})

test('noncurrent workbooks select a visible viewable sheet at creation and preserve the source when none is available', t => {
  const { studio, api } = runtime(t)
  const input = module.createBlankWorkbook(); input.id = 'hidden-created'
  input.sheets = ['Hidden', 'Visible'].map((name, index) => ({ ...module.createBlankSheet(name), id: 'hidden-created-' + index, visible: index === 1 }))
  input.activeSheetId = input.sheets[0].id
  const book = api.createWorkbook(input, { makeCurrent: false })
  assert.equal(book.getActiveSheet().getSheetId(), input.sheets[1].id)
  assert.equal(studio.getSelectionState(book.getId(), input.sheets[0].id), null)
  assert.throws(() => book.getSheetBySheetId(input.sheets[0].id).activate(), /隐藏/)
  assert.equal(history(studio, book.getId()).undo, 0)
  const allHidden = copy(input); allHidden.id = 'all-hidden-created'; allHidden.sheets[1].visible = false
  const unavailable = api.createWorkbook(allHidden, { makeCurrent: false })
  assert.equal(studio.getWorkbookModel(unavailable.getId()).activeSheetId, allHidden.activeSheetId)
  assert.equal(studio.getSelectionState(unavailable.getId()), null)
  assert.equal(allHidden.sheets[0].visible, false)
})
