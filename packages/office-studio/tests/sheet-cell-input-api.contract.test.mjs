import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server, module
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  module = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
function runtime(t, options = {}) {
  const model = module.createBlankWorkbook(); model.id = 'input-api-book'; model.calculationMode = 'manual'
  model.sheets = ['First', 'Second'].map((name, index) => ({ ...module.createBlankSheet(name), id: 'input-api-sheet-' + index, rowCount: 20, columnCount: 8 }))
  model.activeSheetId = model.sheets[0].id
  const studio = new module.OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8, ...options })
  t.onTestFinished(() => studio.dispose())
  const api = module.FOffice.newAPI(studio), book = api.getActiveWorkbook(), [first, second] = book.getSheets()
  return { studio, api, book, first, second }
}
const history = studio => {
  const unit = studio.editor.history.units[studio.activeWorkbookId]
  return { undo: unit.undo.length, redo: unit.redo.length, revision: studio.getWorkbookRevision(studio.activeWorkbookId) }
}

test('Facade editor-text input and the registered command share typed parsing, one content history and real change events', t => {
  const { studio, api, book, first } = runtime(t), changed = [], commands = []
  api.addEvent(api.Event.SheetValueChanged, event => changed.push(copy(event)))
  api.addEvent(api.Event.CommandExecuted, event => commands.push(event.id))
  first.getRange('B2').activate()
  const result = first.getRange('B2').setInputText('12.5%')
  assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(first.getRange('B2').getValue(), .125)
  assert.equal(history(studio).undo, 1); assert.equal(changed.length, 1)
  assert.equal(changed[0].ranges[0].startRow, 1); assert.equal(changed[0].ranges[0].startColumn, 1)
  assert.ok(commands.includes(module.COMMIT_SHEET_CELL_INPUT_COMMAND))
  const before = history(studio)
  assert.equal(first.getRange('B2').setInputText('12.5%').changed, false); assert.deepEqual(history(studio), before)
  book.undo(); assert.equal(first.getRange('B2').getValue(), null)
  book.redo(); assert.equal(first.getRange('B2').getValue(), .125)
  const saved = book.save(), reopened = api.createWorkbook({ ...saved, id: 'input-api-reopened' })
  assert.equal(reopened.getActiveSheet().getRange('B2').getValue(), .125)
  assert.match(reopened.getActiveSheet().getRange('B2').getNumberFormat(), /%/)
})

test('single input retains complete multiple selections through Tab navigation and undo/redo', t => {
  const { studio, api, book, first } = runtime(t)
  first.setActiveRangeList([first.getRange('A1:C2'), first.getRange('F6:G7')], { activeRangeIndex: 0, primary: { row: 0, column: 1 } })
  const original = studio.getSelectionState()
  first.getRange('B1').setInputText('10')
  api.executeCommandSync(module.MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, { direction: api.Enum.Direction.RIGHT, keycode: 9 })
  assert.equal(studio.getSelectionState().primary.column, 2)
  assert.deepEqual(studio.getSelectionState().ranges, original.ranges); assert.equal(history(studio).undo, 1)
  book.undo(); assert.deepEqual(studio.getSelectionState(), original); assert.equal(first.getRange('B1').getValue(), null)
  book.redo(); assert.deepEqual(studio.getSelectionState(), original); assert.equal(first.getRange('B1').getValue(), 10)
})

test('canceling cell input preserves content, complete selection, Tab origin, revision and history', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('B2').activate()
  api.executeCommandSync(module.MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, { direction: api.Enum.Direction.RIGHT, keycode: 9 })
  const state = copy(studio.getWorkbookModel(book.getId())), selection = studio.getSelectionState(), before = history(studio)
  const origin = studio.sheetSelections.getTabStart(book.getId(), first.getSheetId())
  const veto = api.addEvent(api.Event.BeforeCommandExecute, event => { if (event.id === module.COMMIT_SHEET_CELL_INPUT_COMMAND) event.cancel = true })
  assert.throws(() => first.getRange('C2').setInputText('99'), /canceled/)
  veto.dispose()
  assert.deepEqual(copy(studio.getWorkbookModel(book.getId())), state); assert.deepEqual(studio.getSelectionState(), selection)
  assert.deepEqual(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), origin); assert.deepEqual(history(studio), before)
})

test('stop/warning validation and stale session rejection use the shared transaction before any mutation', t => {
  const { studio, book, first } = runtime(t)
  first.getRange('A1').setValue(5)
  first.setDataValidation('A1', { type: 'whole', operator: 'between', formula1: '1', formula2: '10', errorStyle: 'stop', error: '1 到 10' })
  const start = module.createSheetCellEditSession({ documentId: book.getId(), sheetId: first.getSheetId(), key: 'A1', date1904: false,
    cell: module.toExcelCellView(first.getRange('A1').getCellData()) }, { region: 'en-US', currentYear: 2026 })
  const before = history(studio), content = copy(studio.getWorkbookModel(book.getId()))
  const rejected = first.getRange('A1').setInputText('20', { session: start })
  assert.equal(rejected.accepted, false); assert.equal(rejected.reason, 'validation'); assert.deepEqual(history(studio), before)
  assert.deepEqual(copy(studio.getWorkbookModel(book.getId())), content)
  first.setDataValidation('A1', { type: 'whole', operator: 'between', formula1: '1', formula2: '10', errorStyle: 'warning', error: '1 到 10' })
  const warning = first.getRange('A1').setInputText('20', { session: start })
  assert.equal(warning.accepted, true); assert.equal(warning.issues[0].severity, 'warning'); assert.equal(first.getRange('A1').getValue(), 20)
  const after = history(studio)
  assert.equal(first.getRange('A1').setInputText('7', { session: start }).conflict, 'content'); assert.deepEqual(history(studio), after)
})

test('a SheetValueChanged subscriber failure cannot turn a committed cell into a canceled edit or block other observers', t => {
  const { studio, api, first } = runtime(t), received = [], failure = new Error('cell observer failed')
  api.addEvent(api.Event.SheetValueChanged, event => { event.ranges[0].startRow = 19; throw failure })
  api.addEvent(api.Event.SheetValueChanged, event => received.push(event.ranges[0].startRow))
  const result = first.getRange('A1').setInputText('7')
  assert.equal(result.accepted, true); assert.equal(first.getRange('A1').getValue(), 7); assert.equal(history(studio).undo, 1)
  assert.deepEqual(received, [0]); assert.ok(studio.getEventListenerErrors().some(item => item.eventName === 'SheetValueChanged' && item.error === failure))
})

test('automatic formula commits recalculate their caches and manual commits keep source formulas distinct', t => {
  const { studio, api, first, book } = runtime(t)
  api.executeCommandSync('sheet.command.set-calculation-mode', { unitId: book.getId(), mode: 'automatic' })
  assert.equal(first.getRange('A1').setInputText('=1+2').accepted, true)
  assert.equal(first.getRange('A1').getFormula(), '=1+2'); assert.equal(first.getRange('A1').getValue(), 3)
  book.setCalculationMode('manual')
  assert.equal(first.getRange('A1').setInputText('=2+3').accepted, true)
  assert.equal(first.getRange('A1').getFormula(), '=2+3'); assert.equal(first.getRange('A1').getValue(), null)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['0:0'].cachedDataType, null)
})

test('merged range input writes only its owner and multiple ordinary cells require the range-value API', t => {
  const { studio, book, first } = runtime(t)
  studio.commitWorkbook(book.getId(), 'merge fixture', model => { model.sheets[0].mergedRanges = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }] })
  assert.equal(first.getRange('A1:B2').setInputText('owner').accepted, true)
  assert.equal(first.getRange('A1').getValue(), 'owner'); assert.equal(first.getRange('B2').getValue(), null)
  assert.throws(() => first.getRange('C1:D2').setInputText('many'), /单个单元格/)
})

test('detached view preparation failure preserves the authoritative model, selection, Tab origin and history', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('B2').activate()
  api.executeCommandSync(module.MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, { direction: api.Enum.Direction.RIGHT, keycode: 9 })
  const model = studio.getWorkbookModel(book.getId()), before = copy(model), depth = history(studio), selection = studio.getSelectionState()
  const origin = studio.sheetSelections.getTabStart(book.getId(), first.getSheetId())
  assert.throws(() => studio.commitWorkbook(book.getId(), 'failing view', candidate => {
    candidate.sheets[0].cells['0:0'] = module.createCell(0, 0, 10)
  }, { prepareView: candidate => { candidate.sheets[0].cells['0:0'].value = 999; throw new Error('view refused') } }), /view refused/)
  assert.equal(studio.getWorkbookModel(book.getId()), model); assert.deepEqual(copy(model), before)
  assert.deepEqual(history(studio), depth); assert.deepEqual(studio.getSelectionState(), selection)
  assert.deepEqual(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), origin)
  assert.equal(studio.editor.history.isApplying, false)
  studio.commitWorkbook(book.getId(), 'prepared edit', candidate => {
    candidate.sheets[0].cells['0:0'] = module.createCell(0, 0, 10)
  }, { prepareView: candidate => { candidate.sheets[0].cells['0:0'].value = 999 } })
  assert.equal(first.getRange('A1').getValue(), 10); assert.equal(history(studio).undo, depth.undo + 1)
})

test('preparation detects a nested content edit while retaining navigation that occurs during preparation', t => {
  const { studio, book, first, second } = runtime(t)
  assert.throws(() => studio.commitWorkbook(book.getId(), 'stale prepared edit', candidate => {
    candidate.sheets[0].cells['0:0'] = module.createCell(0, 0, 10)
  }, { prepareView: () => { first.getRange('C3').setInputText('later') } }), /准备编辑期间工作簿内容已更改/)
  assert.equal(first.getRange('A1').getValue(), null); assert.equal(first.getRange('C3').getValue(), 'later')
  assert.equal(history(studio).undo, 1)
  studio.commitWorkbook(book.getId(), 'prepared navigation', candidate => {
    candidate.sheets[0].cells['0:0'] = module.createCell(0, 0, 10)
  }, { prepareView: () => { second.getRange('D4').activate() } })
  const selection = studio.getSelectionState()
  assert.equal(book.getActiveSheet().getSheetId(), second.getSheetId()); assert.equal(first.getRange('A1').getValue(), 10)
  book.undo(); book.redo()
  assert.equal(book.getActiveSheet().getSheetId(), second.getSheetId()); assert.deepEqual(studio.getSelectionState(), selection)
})

test('mutation command observations isolate nested parameter objects for both Sheets and Docs', t => {
  const { studio, api, book } = runtime(t), seen = []
  api.addEvent(api.Event.CommandExecuted, event => {
    if (event.id.endsWith('observer-copy')) event.params.nested.value = 'subscriber'
  })
  api.addEvent(api.Event.CommandExecuted, event => {
    if (event.id.endsWith('observer-copy')) seen.push(event.params.nested.value)
  })
  const params = { nested: { value: 'original' } }
  studio.mutateWorkbook(book.getId(), 'sheet.command.observer-copy', 'rename workbook', candidate => { candidate.name = 'changed' }, params)
  studio.mutateDocument(studio.activeDocumentId, 'doc.command.observer-copy', 'rename document', candidate => { candidate.name = 'changed' }, params)
  assert.deepEqual(seen, ['original', 'original']); assert.equal(params.nested.value, 'original')
})

test('BeforeUndo and BeforeRedo host switches leave both workbook histories and content intact', t => {
  for (const direction of ['undo', 'redo']) for (const nextKind of ['workbook', 'document']) {
    const { studio, api, book, first } = runtime(t)
    first.getRange('A1').setInputText('first edit')
    if (direction === 'redo') studio.undo('workbook', book.getId())
    const other = api.createWorkbook({ id: 'history-other-book', calculationMode: 'manual' }, { makeCurrent: false })
    other.getActiveSheet().getRange('A1').setInputText('other edit')
    const document = api.getActiveDocument()
    document.insertText(0, 'document edit')
    studio.activateWorkbook(book.getId())
    const content = copy(studio.getWorkbookModel(book.getId())), otherContent = copy(studio.getWorkbookModel(other.getId()))
    const documentContent = copy(document.getModel())
    const state = id => ({ ...studio.editor.history.captureSavepoint(id), undo: studio.editor.history.units[id].undo.length,
      redo: studio.editor.history.units[id].redo.length, dirty: studio.editor.history.isDirty(id) })
    const before = [state(book.getId()), state(other.getId()), state(document.getId())], completed = []
    api.addEvent(direction === 'undo' ? api.Event.Undo : api.Event.Redo, event => completed.push(event))
    api.addEvent(direction === 'undo' ? api.Event.BeforeUndo : api.Event.BeforeRedo, () => {
      if (nextKind === 'workbook') studio.activateWorkbook(other.getId())
      else studio.activateDocument(document.getId())
    })
    assert.equal(studio[direction]('workbook', book.getId()), false)
    assert.deepEqual(copy(studio.getWorkbookModel(book.getId())), content)
    assert.deepEqual(copy(studio.getWorkbookModel(other.getId())), otherContent); assert.deepEqual(copy(document.getModel()), documentContent)
    assert.deepEqual([state(book.getId()), state(other.getId()), state(document.getId())], before); assert.deepEqual(completed, [])
    assert.equal(studio.editor.state.activeKind, nextKind)
    assert.equal(studio.editor.history.activeUnit, nextKind === 'workbook' ? other.getId() : document.getId())
  }
})

test('a workbook edit in BeforeUndo or BeforeRedo supersedes the captured entry without replaying new content', t => {
  for (const direction of ['undo', 'redo']) {
    const { studio, api, book, first } = runtime(t)
    first.getRange('A1').setInputText('first edit')
    if (direction === 'redo') studio.undo('workbook', book.getId())
    let nested = false
    const listener = api.addEvent(direction === 'undo' ? api.Event.BeforeUndo : api.Event.BeforeRedo, () => {
      if (!nested) { nested = true; first.getRange('B2').setInputText('newer edit') }
    })
    assert.equal(studio[direction]('workbook', book.getId()), false)
    assert.equal(first.getRange('A1').getValue(), direction === 'undo' ? 'first edit' : null)
    assert.equal(first.getRange('B2').getValue(), 'newer edit')
    assert.equal(history(studio).undo, direction === 'undo' ? 2 : 1); assert.equal(history(studio).redo, 0)
    listener.dispose()
    assert.equal(studio.undo('workbook', book.getId()), true)
    assert.equal(first.getRange('B2').getValue(), null)
    assert.equal(first.getRange('A1').getValue(), direction === 'undo' ? 'first edit' : null)
  }
})

test('canceled workbook undo and redo retain complete selections, Tab origin, content revision and savepoint', t => {
  for (const direction of ['undo', 'redo']) {
    const { studio, api, book, first } = runtime(t)
    first.getRange('A1').setInputText('first edit')
    if (direction === 'redo') studio.undo('workbook', book.getId())
    first.setActiveRangeList([first.getRange('B2:C3'), first.getRange('F6:G7')], { activeRangeIndex: 0, primary: { row: 1, column: 1 } })
    api.executeCommandSync(module.MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND, { direction: api.Enum.Direction.RIGHT, keycode: 9 })
    const selection = studio.getSelectionState(), origin = studio.sheetSelections.getTabStart(book.getId(), first.getSheetId())
    const content = copy(studio.getWorkbookModel(book.getId())), before = history(studio), savepoint = studio.editor.history.captureSavepoint(book.getId())
    api.addEvent(direction === 'undo' ? api.Event.BeforeUndo : api.Event.BeforeRedo, event => { event.cancel = true })
    assert.throws(() => studio[direction]('workbook', book.getId()), /canceled/)
    assert.deepEqual(copy(studio.getWorkbookModel(book.getId())), content); assert.deepEqual(history(studio), before)
    assert.deepEqual(studio.getSelectionState(), selection)
    assert.deepEqual(studio.sheetSelections.getTabStart(book.getId(), first.getSheetId()), origin)
    assert.deepEqual(studio.editor.history.captureSavepoint(book.getId()), savepoint); assert.equal(studio.editor.history.isApplying, false)
  }
})
