import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Draft only: run together with the later unified frontend/Rust/browser checks.
let server, sdk
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const range = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
function runtime(t, customize) {
  const model = sdk.createBlankWorkbook(); model.id = 'fill-api-book'; model.calculationMode = 'manual'
  const sheet = model.sheets[0]; sheet.id = 'data'; model.activeSheetId = sheet.id; sheet.rowCount = 30; sheet.columnCount = 12
  for (let row = 0; row < 6; row++) sheet.cells[`${row}:0`] = sdk.createCell(row, 0, row < 2 ? row + 1 : 100 + row)
  sheet.cells['0:0'].style.font.bold = true; sheet.cells['1:0'].style.font.bold = true
  sheet.cells['2:0'].style.fill = '#00ff00'; sheet.cells['2:0'].style.font.italic = true
  customize?.(sheet, model)
  const studio = new sdk.OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8 })
  t.onTestFinished(() => studio.dispose())
  const api = sdk.FOffice.newAPI(studio), book = api.getActiveWorkbook(), data = book.getActiveSheet()
  data.getRange('A1:A2').activate()
  const history = () => studio.editor.history.units[book.getId()]
  const cells = () => copy(studio.getWorkbookModel(book.getId()).sheets[0].cells)
  return { studio, api, book, data, history, cells, model }
}

test('Facade fill and all four modes amend one history entry, keeping its original undo and final redo', async t => {
  const h = runtime(t), before = h.cells(), beforeSelection = copy(h.studio.getSelectionState(h.book.getId(), 'data'))
  assert.deepEqual(h.api.Enum.AUTO_FILL_APPLY_TYPE, { COPY: 'COPY', SERIES: 'SERIES', ONLY_FORMAT: 'ONLY_FORMAT', NO_FORMAT: 'NO_FORMAT' })
  assert.equal(await h.data.getRange('A1:A2').autoFill(h.data.getRange('A1:A6')), true)
  assert.deepEqual(h.data.getRange('A1:A6').getValues(), [[1], [2], [3], [4], [5], [6]])
  assert.equal(h.history().undo.length, 1)
  const id = h.studio.sheetFill.getView(h.book.getId()).id
  for (const applyType of ['COPY', 'ONLY_FORMAT', 'NO_FORMAT', 'SERIES']) {
    assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: id, applyType }), true)
    assert.equal(h.history().undo.length, 1); assert.equal(h.history().redo.length, 0)
    assert.equal(h.studio.sheetFill.getView(h.book.getId()).applyType, applyType)
    if (applyType === 'ONLY_FORMAT') { assert.equal(h.data.getRange('A3').getValue(), 102); assert.equal(h.cells()['2:0'].style.font.italic, false) }
    if (applyType === 'NO_FORMAT') { assert.equal(h.data.getRange('A3').getValue(), 3); assert.equal(h.cells()['2:0'].style.fill, '#00ff00') }
  }
  const final = h.cells(); h.book.undo()
  assert.deepEqual(h.cells(), before); assert.deepEqual(copy(h.studio.getSelectionState(h.book.getId(), 'data')), beforeSelection)
  h.book.redo(); assert.deepEqual(h.cells(), final)
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: id, applyType: 'COPY' }), false)
})

test('saved snapshots reopen with authored values, styles and formulas but without an old transient refill receipt', async t => {
  const h = runtime(t); h.data.getRange('B1').setValue({ ...sdk.createCell(0, 1), formula: '=A1+$C$1', dataType: 'formula' })
  assert.equal(await h.data.getRange('B1').autoFill(h.data.getRange('B1:B4'), 'COPY'), true)
  assert.deepEqual(h.data.getRange('B1:B4').getFormulas(), [['=A1+$C$1'], ['=A2+$C$1'], ['=A3+$C$1'], ['=A4+$C$1']])
  const saved = h.book.save(), reopened = h.api.createWorkbook({ ...saved, id: 'fill-api-reopened' })
  assert.deepEqual(reopened.getActiveSheet().getRange('B1:B4').getFormulas(), h.data.getRange('B1:B4').getFormulas())
  assert.equal(h.studio.sheetFill.getView(reopened.getId()), undefined)
  assert.deepEqual(reopened.getActiveSheet().getRange('A3').getCellData().style, h.data.getRange('A3').getCellData().style)
})

test('canonical copy-down/right and auto-clear produce boolean results and share content/selection history', t => {
  const h = runtime(t)
  assert.equal(h.api.executeCommandSync('sheet.command.copy-down', { unitId: h.book.getId(), subUnitId: 'data', range: range(0, 0, 3) }), true)
  assert.deepEqual(h.data.getRange('A1:A4').getValues(), [[1], [1], [1], [1]])
  assert.equal(h.studio.sheetFill.getView(h.book.getId()), undefined)
  assert.equal(h.api.executeCommandSync('sheet.command.copy-right', { unitId: h.book.getId(), subUnitId: 'data', range: range(0, 0, 0, 2) }), true)
  assert.deepEqual(h.data.getRange('A1:C1').getValues(), [[1, 1, 1]])
  assert.equal(h.api.executeCommandSync('sheet.command.auto-clear-content', { unitId: h.book.getId(), subUnitId: 'data',
    clearRange: range(2, 0, 3), selectionRange: range(0, 0, 1) }), true)
  assert.deepEqual(h.data.getRange('A1:A4').getValues(), [[1], [1], [null], [null]])
  assert.equal(h.history().undo.length, 3)
  h.book.undo(); assert.deepEqual(h.data.getRange('A1:A4').getValues(), [[1], [1], [1], [1]])
  assert.equal(h.api.executeCommandSync('sheet.command.copy-down', { unitId: h.book.getId(), subUnitId: 'data', range: range(0, 0) }), false)
})

test('Facade rejects cross-sheet/controller, shrinking, both-axis and double-direction targets before history changes', async t => {
  const h = runtime(t), source = h.data.getRange('A1:A2'), other = h.book.insertSheet('Other'), depth = h.history().undo.length
  assert.throws(() => source.autoFill(other.getRange('A1:A5')), /同一工作簿/)
  assert.throws(() => source.autoFill(h.data.getRange('A1')), /包含源/)
  assert.throws(() => source.autoFill(h.data.getRange('A1:B6')), /一个方向/)
  assert.throws(() => h.data.getRange('A2:A3').autoFill(h.data.getRange('A1:A5')), /一个方向/)
  assert.throws(() => source.autoFill(h.data.getRange('A1:A5'), 'unknown'), /模式/)
  assert.equal(await source.autoFill(h.data.getRange('A1:A2')), false)
  assert.equal(h.history().undo.length, depth)
})

test('obsolete UI revision, wrong session, later editing and reopened history cannot rewrite the old fill', async t => {
  const h = runtime(t), revision = h.studio.getWorkbookRevision(h.book.getId())
  await h.data.getRange('A1:A2').autoFill(h.data.getRange('A1:A6'))
  const id = h.studio.sheetFill.getView(h.book.getId()).id, depth = h.history().undo.length
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: 'wrong', applyType: 'COPY' }), false)
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: id, applyType: 'COPY', expectedRevision: revision }), false)
  h.data.getRange('D4').setValue('later')
  const before = h.cells()
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: id, applyType: 'COPY' }), false)
  assert.deepEqual(h.cells(), before); assert.equal(h.history().undo.length, depth + 1)
  h.studio.editor.history.clear(h.book.getId())
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: id, applyType: 'COPY' }), false)
})

test('command cancellation and a reentrant post-fill edit preserve model/history and reject later refill', async t => {
  const h = runtime(t), before = h.cells()
  const cancel = h.api.addEvent(h.api.Event.BeforeCommandExecute, event => { if (event.id === 'sheet.command.auto-fill') event.cancel = true })
  await assert.rejects(() => h.data.getRange('A1:A2').autoFill(h.data.getRange('A1:A6')), /canceled/)
  cancel.dispose(); assert.deepEqual(h.cells(), before); assert.equal(h.history().undo.length, 0)
  let reacted = false
  const edit = h.api.addEvent(h.api.Event.SheetValueChanged, () => {
    if (reacted) return
    reacted = true; h.data.getRange('D4').setValue('listener')
  })
  assert.equal(await h.data.getRange('A1:A2').autoFill(h.data.getRange('A1:A6')), true)
  edit.dispose(); assert.equal(h.history().undo.length, 2); assert.equal(h.data.getRange('D4').getValue(), 'listener')
  assert.equal(h.studio.sheetFill.getView(h.book.getId()), undefined)
  const result = h.studio.sheetFill.getResult(h.book.getId())
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: result.sessionId, applyType: 'COPY' }), false)
})

test('an unchanged fill succeeds and selects its target; later refill creates its own entry after an unrelated edit', async t => {
  const h = runtime(t, sheet => {
    for (let row = 0; row < 6; row++) {
      sheet.cells[`${row}:0`] = sdk.createCell(row, 0, 1)
    }
  })
  h.data.getRange('D1').setValue('prior'); h.data.getRange('A1').activate()
  const revision = h.studio.getWorkbookRevision(h.book.getId()), savepoint = h.studio.editor.history.captureSavepoint(h.book.getId())
  const events = []; h.api.addEvent(h.api.Event.SheetValueChanged, event => events.push(event))
  assert.equal(await h.data.getRange('A1').autoFill(h.data.getRange('A1:A6'), 'COPY'), true)
  assert.equal(h.history().undo.length, 1); assert.equal(h.studio.getWorkbookRevision(h.book.getId()), revision)
  assert.deepEqual(h.studio.editor.history.captureSavepoint(h.book.getId()), savepoint); assert.deepEqual(events, [])
  assert.deepEqual(h.studio.getSelection().endRow, 5)
  const session = h.studio.sheetFill.getView(h.book.getId())
  assert.equal(session.hasHistoryEntry, false); assert.equal(h.studio.sheetFill.getResult(h.book.getId()).changed, false)
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: session.id, type: 'SERIES' }), true)
  assert.equal(h.history().undo.length, 2); assert.equal(h.studio.sheetFill.getView(h.book.getId()).hasHistoryEntry, true)
  assert.deepEqual(h.data.getRange('A1:A6').getValues(), [[1], [2], [3], [4], [5], [6]])
  h.book.undo(); assert.deepEqual(h.data.getRange('A1:A6').getValues(), [[1], [1], [1], [1], [1], [1]])
  assert.equal(h.data.getRange('D1').getValue(), 'prior'); assert.equal(h.studio.getSelection().endRow, 0)
  h.book.undo(); assert.equal(h.data.getRange('D1').getValue(), null)
})

test('equivalent mode switches update the recent menu without changing its content revision or history receipt', async t => {
  const h = runtime(t, sheet => {
    for (let row = 0; row < 6; row++) sheet.cells[`${row}:0`] = sdk.createCell(row, 0, row < 2 ? 'word' : 'old')
  })
  assert.equal(await h.data.getRange('A1:A2').autoFill(h.data.getRange('A1:A6')), true)
  const session = h.studio.sheetFill.getView(h.book.getId()), revision = h.studio.getWorkbookRevision(h.book.getId())
  assert.equal(session.applyType, 'COPY'); assert.equal(session.hasHistoryEntry, true)
  assert.equal(h.api.executeCommandSync('sheet.command.refill', { unitId: h.book.getId(), subUnitId: 'data', sessionId: session.id, type: 'NO_FORMAT' }), true)
  const after = h.studio.sheetFill.getView(h.book.getId())
  assert.equal(after.applyType, 'NO_FORMAT'); assert.deepEqual(after.savepoint, session.savepoint)
  assert.equal(h.studio.getWorkbookRevision(h.book.getId()), revision); assert.equal(h.history().undo.length, 1)
  assert.equal(h.studio.sheetFill.getResult(h.book.getId()).changed, false)
  h.book.undo(); assert.equal(h.data.getRange('A3').getValue(), 'old'); h.book.redo(); assert.equal(h.data.getRange('A3').getValue(), 'word')
})
