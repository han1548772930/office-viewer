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
function runtime(t) {
  const initial = module.createBlankWorkbook(); initial.id = 'lifecycle-api-book'; initial.calculationMode = 'manual'
  initial.sheets = ['Data', 'Host'].map(name => ({ ...module.createBlankSheet(name), id: name.toLowerCase(), rowCount: 20, columnCount: 10 }))
  initial.activeSheetId = 'data'
  initial.sheets[0].cells['0:0'] = { ...module.createCell(0, 0), value: 7, display: '7' }
  initial.sheets[1].cells['0:0'] = { ...module.createCell(0, 0), formula: '=Data!A1', value: 7, display: '7' }
  initial.metadata = { definedNames: [{ name: 'Local', localSheetId: 1, refersTo: '=Host!A1' }] }
  const studio = new module.OfficeStudio({ workbook: initial })
  t.onTestFinished(() => studio.dispose())
  const api = module.FOffice.newAPI(studio), book = api.getActiveWorkbook(), [data, host] = book.getSheets()
  return { studio, api, book, data, host }
}
function history(studio, book) {
  const state = studio.editor.history.units[book.getId()]
  return { undo: state.undo.length, redo: state.redo.length, revision: studio.getWorkbookRevision(book.getId()) }
}

test('Facade and canonical rename command share one event pair, one content history and save/reopen references', t => {
  const { studio, api, book, data, host } = runtime(t), events = []
  api.addEvent(api.Event.BeforeSheetNameChange, event => events.push(['before', event.oldName, event.newName]))
  api.addEvent(api.Event.SheetNameChanged, event => events.push(['after', event.oldName, event.newName]))
  data.setName('Budget 2026')
  assert.equal(data.getSheetName(), 'Budget 2026'); assert.equal(host.getRange('A1').getFormula(), "='Budget 2026'!A1")
  assert.deepEqual(events, [['before', 'Data', 'Budget 2026'], ['after', 'Data', 'Budget 2026']])
  assert.equal(history(studio, book).undo, 1)
  book.undo(); assert.equal(data.getSheetName(), 'Data'); assert.equal(host.getRange('A1').getFormula(), '=Data!A1')
  book.redo(); assert.equal(host.getRange('A1').getFormula(), "='Budget 2026'!A1")
  const saved = book.save(), reopened = api.createWorkbook({ ...saved, id: 'lifecycle-api-reopened' })
  assert.equal(reopened.getSheetBySheetId('host').getRange('A1').getFormula(), "='Budget 2026'!A1")
  assert.equal(reopened.getSheetBySheetId('data').getSheetName(), 'Budget 2026')
})

test('canceled and reentrant rename preserve the correct committed content and reject an obsolete candidate', t => {
  const { studio, api, book, data } = runtime(t), before = copy(studio.getWorkbookModel(book.getId())), state = history(studio, book)
  const cancel = api.addEvent(api.Event.BeforeSheetNameChange, event => { event.cancel = true })
  assert.throws(() => data.setName('Canceled'), /canceled/)
  cancel.dispose()
  assert.deepEqual(copy(studio.getWorkbookModel(book.getId())), before); assert.deepEqual(history(studio, book), state)
  const reenter = api.addEvent(api.Event.BeforeSheetNameChange, () => data.getRange('A1').setValue(9))
  assert.throws(() => data.setName('Obsolete'), /上下文已更改/)
  reenter.dispose()
  assert.equal(data.getSheetName(), 'Data'); assert.equal(data.getRange('A1').getValue(), 9)
  assert.equal(history(studio, book).undo, state.undo + 1)
})

test('hide/show/veryHidden, custom color and reordering use shared commands and retain their saved state', t => {
  const { studio, api, book, data, host } = runtime(t)
  assert.deepEqual(api.Enum.WorksheetHiddenState, { VISIBLE: 0, HIDDEN: 1, VERY_HIDDEN: 2 })
  data.setHiddenState(api.Enum.WorksheetHiddenState.VERY_HIDDEN)
  assert.equal(data.getHiddenState(), 2); assert.equal(book.getActiveSheet().getSheetId(), 'host')
  const before = history(studio, book)
  assert.throws(() => host.hideSheet(), /唯一可见/); assert.deepEqual(history(studio, book), before)
  data.showSheet(); data.setTabColor('#123456'); data.setIndex(1)
  assert.equal(data.getHiddenState(), 0); assert.equal(data.getTabColor(), '#123456'); assert.equal(data.getIndex(), 1)
  assert.equal(studio.getWorkbookModel(book.getId()).metadata.definedNames[0].localSheetId, 0)
  book.undo(); assert.equal(data.getIndex(), 0)
  book.redo(); assert.equal(data.getIndex(), 1)
  const reopened = api.createWorkbook({ ...book.save(), id: 'lifecycle-api-state-reopened' })
  assert.equal(reopened.getSheetBySheetId('data').getTabColor(), '#123456')
  assert.equal(reopened.getSheetBySheetId('data').getIndex(), 1)
  assert.equal(reopened.getSheetBySheetId('data').isSheetHidden(), false)
})

test('insert and remove share the registered lifecycle command and undo complete reference changes', t => {
  const { studio, api, book, data, host } = runtime(t)
  const inserted = book.insertSheet('Data', { index: 1 })
  assert.equal(inserted.getSheetName(), 'Data1'); assert.equal(inserted.getIndex(), 1)
  assert.equal(history(studio, book).undo, 1)
  assert.equal(book.deleteSheet(data), true)
  assert.equal(host.getRange('A1').getFormula(), '=#REF!')
  assert.equal(book.getNumSheets(), 2)
  book.undo(); assert.equal(book.getNumSheets(), 3); assert.equal(host.getRange('A1').getFormula(), '=Data!A1')
  book.redo(); assert.equal(host.getRange('A1').getFormula(), '=#REF!')
  assert.equal(book.deleteSheet('absent'), false)
  const reopened = api.createWorkbook({ ...book.save(), id: 'lifecycle-api-delete-reopened' })
  assert.equal(reopened.getSheetBySheetId('host').getRange('A1').getFormula(), '=#REF!')
})

test('stale revision and a changed modal host reject UI publication without extra history', t => {
  const { studio, api, book, data, host } = runtime(t)
  const revision = studio.getWorkbookRevision(book.getId()), generation = studio.editor.history.captureSavepoint(book.getId()).generation
  data.setTabColor('#123456')
  const before = history(studio, book)
  assert.throws(() => api.executeCommandSync('sheet.command.set-tab-color', { unitId: book.getId(), subUnitId: 'data', value: '#abcdef', expectedRevision: revision, expectedHistoryGeneration: generation }), /已更改/)
  book.setActiveSheet(host)
  assert.equal(api.executeCommandSync('sheet.command.set-worksheet-order', { unitId: book.getId(), subUnitId: 'data', order: 1, host: { unitId: book.getId(), sheetId: 'data' } }), false)
  assert.equal(data.getTabColor(), '#123456'); assert.equal(data.getIndex(), 0)
  assert.deepEqual(history(studio, book), before)
})
