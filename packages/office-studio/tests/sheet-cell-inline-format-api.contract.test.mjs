import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Draft: real command registration, content history and snapshot integration; not executed.
let server, module
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  module = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
function runtime(t) {
  const model = module.createBlankWorkbook(); model.id = 'inline-api'; model.calculationMode = 'manual'
  model.sheets = ['First', 'Second'].map((name, index) => ({ ...module.createBlankSheet(name), id: 'inline-api-' + index, rowCount: 30, columnCount: 12 }))
  model.activeSheetId = model.sheets[0].id
  const cell = module.createCell(0, 0, 'abcdef')
  cell.richText = [{ id: 'source', text: 'abcdef', font: { italic: true }, hyperlink: 'https://example.test', vendor: { keep: true } }]
  model.sheets[0].cells['0:0'] = cell
  model.sheets[0].cells['0:1'] = module.createCell(0, 1, '123')
  model.sheets[0].cells['0:2'] = { ...module.createCell(0, 2, "'abc"), dataType: 'force_string', metadata: { keep: true } }
  const studio = new module.OfficeStudio({ workbook: model, sheetTextMeasure: text => [...text].length * 8 })
  t.onTestFinished(() => studio.dispose())
  const api = module.FOffice.newAPI(studio), book = api.getActiveWorkbook(), [first, second] = book.getSheets()
  return { studio, api, book, first, second }
}
const history = studio => {
  const id = studio.activeWorkbookId, unit = studio.editor.history.units[id]
  return { undo: unit.undo.length, redo: unit.redo.length, revision: studio.getWorkbookRevision(id) }
}
function draft(module, book, sheet, key = 'A1') {
  const cell = module.toExcelCellView(sheet.getRange(key).getCellData())
  const session = module.createSheetCellEditSession({ documentId: book.getId(), sheetId: sheet.getSheetId(), key, date1904: false, cell },
    { region: 'en-US', currentYear: 2026 })
  const value = module.sheetCellInputText(cell)
  return { session, editTrace: { version: 1, sessionId: session.sessionId, source: value, operations: [] }, value }
}

test('partial range API writes one content transaction and retains complete run data through undo, redo and reopen', t => {
  const { studio, api, book, first } = runtime(t), before = structuredClone(first.getRange('A1').getCellData()), commands = []
  api.addEvent(api.Event.CommandExecuted, event => commands.push(event.id))
  const result = first.getRange('A1').setRichTextStyle(1, 4, { bold: true, size: 18, color: '#ff0000' })
  assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(history(studio).undo, 1)
  const cell = structuredClone(first.getRange('A1').getCellData())
  assert.equal(cell.value, 'abcdef'); assert.deepEqual(cell.richText.map(run => run.text), ['a', 'bcd', 'ef'])
  assert.deepEqual(cell.richText.map(run => run.id), ['source', undefined, undefined])
  assert.ok(cell.richText.every(run => run.hyperlink === 'https://example.test' && run.vendor.keep))
  assert.ok(commands.includes(module.FORMAT_SHEET_CELL_RICH_TEXT_COMMAND))
  book.undo(); assert.deepEqual(first.getRange('A1').getCellData(), before)
  book.redo(); assert.deepEqual(first.getRange('A1').getCellData(), cell)
  const reopened = api.createWorkbook({ ...book.save(), id: 'inline-api-reopened' })
  assert.deepEqual(reopened.getActiveSheet().getRange('A1').getCellData().richText, cell.richText)
})

test('numeric-looking text and literal apostrophes retain their raw values, display and storage type', t => {
  const { studio, first } = runtime(t)
  for (const key of ['B1', 'C1']) {
    const before = structuredClone(first.getRange(key).getCellData())
    assert.equal(first.getRange(key).setRichTextStyle(0, 1, { underline: true }).changed, true)
    const cell = first.getRange(key).getCellData()
    assert.deepEqual({ ...cell, richText: before.richText }, before)
    assert.equal(cell.richText.map(run => run.text).join(''), before.value)
  }
  assert.equal(history(studio).undo, 2)
})

test('editing font operation publishes no content until the same input session is committed once', t => {
  const { studio, api, book, first } = runtime(t), start = draft(module, book, first, 'B1'), depth = history(studio)
  const operation = api.executeCommandSync(module.FORMAT_SHEET_CELL_INPUT_OPERATION, {
    unitId: book.getId(), sheetId: first.getSheetId(), row: 0, column: 1, input: start.value,
    session: start.session, editTrace: start.editTrace, commandId: 'sheet.command.set-font-family',
    params: { value: 'Arial' }, selection: { start: 1, end: 2 },
  })
  assert.equal(operation.accepted, true); assert.equal(first.getRange('B1').getCellData().richText, undefined)
  assert.deepEqual(history(studio), depth)
  const result = first.getRange('B1').setInputText(start.value, { session: start.session, editTrace: start.editTrace, inlineFormats: operation.formats })
  assert.equal(result.accepted, true); assert.equal(result.changed, true); assert.equal(first.getRange('B1').getValue(), '123')
  assert.equal(history(studio).undo, depth.undo + 1)
  book.undo(); assert.equal(first.getRange('B1').getCellData().richText, undefined)
  book.redo(); assert.equal(first.getRange('B1').getCellData().richText[1].font.family, 'Arial')
})

test('empty and repeated style edits avoid history while command cancellation preserves source and selection', t => {
  const { studio, api, book, first } = runtime(t)
  first.getRange('A1').activate()
  const depth = history(studio), before = structuredClone(studio.getWorkbookModel(book.getId())), selection = studio.getSelectionState()
  assert.equal(first.getRange('A1').setRichTextStyle(0, 3, {}).changed, false); assert.deepEqual(history(studio), depth)
  const veto = api.addEvent(api.Event.BeforeCommandExecute, event => { if (event.id === module.FORMAT_SHEET_CELL_RICH_TEXT_COMMAND) event.cancel = true })
  assert.throws(() => first.getRange('A1').setRichTextStyle(1, 3, { bold: true }), /canceled/u); veto.dispose()
  assert.deepEqual(studio.getWorkbookModel(book.getId()), before); assert.deepEqual(studio.getSelectionState(), selection); assert.deepEqual(history(studio), depth)
  first.getRange('A1').setRichTextStyle(1, 3, { bold: true })
  const after = history(studio)
  assert.equal(first.getRange('A1').setRichTextStyle(1, 3, { bold: true }).changed, false); assert.deepEqual(history(studio), after)
})

test('explicit inactive sheet targets and complete merged ranges use their content owner', t => {
  const { studio, book, first, second } = runtime(t)
  second.getRange('C3').setValue('second')
  const selection = studio.getSelectionState()
  second.getRange('C3').setRichTextStyle(0, 2, { italic: true })
  assert.equal(book.getActiveSheet().getSheetId(), first.getSheetId()); assert.deepEqual(studio.getSelectionState(), selection)
  assert.equal(first.getRange('A1').getValue(), 'abcdef'); assert.equal(second.getRange('C3').getCellData().richText[0].font.italic, true)
  studio.commitWorkbook(book.getId(), 'merge fixture', model => { model.sheets[0].mergedRanges = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }] })
  assert.equal(first.getRange('A1:B2').setRichTextStyle(0, 1, { strike: true }).accepted, true)
  assert.equal(first.getRange('B2').getValue(), null)
  assert.throws(() => first.getRange('D1:E2').setRichTextStyle(0, 1, { bold: true }), /单个文本单元格/u)
})

test('true numeric, formula, surrogate, invalid-style and protected targets reject atomically', t => {
  const { studio, book, first } = runtime(t)
  first.getRange('D1').setValue(123); first.getRange('E1').setInputText('=1+2'); first.getRange('F1').setValue('A🙂B')
  const before = structuredClone(studio.getWorkbookModel(book.getId())), depth = history(studio)
  for (const run of [() => first.getRange('D1').setRichTextStyle(0, 1, { bold: true }),
    () => first.getRange('E1').setRichTextStyle(0, 1, { bold: true }), () => first.getRange('F1').setRichTextStyle(2, 3, { bold: true }),
    () => first.getRange('A1').setRichTextStyle(0, 1, { size: NaN }), () => first.getRange('A1').setRichTextStyle(0, 1, { fill: '#ff0000' })]) assert.throws(run)
  assert.deepEqual(studio.getWorkbookModel(book.getId()), before); assert.deepEqual(history(studio), depth)
  studio.commitWorkbook(book.getId(), 'protection fixture', model => {
    model.sheets[0].rangeProtectionRules = [{ id: 'locked', permissionId: 'locked', ownerUserId: 'remote', allowedUserIds: [],
      editState: 'onlyMe', viewState: 'othersCanView', description: '', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }] }]
  })
  const locked = history(studio)
  assert.throws(() => first.getRange('A1').setRichTextStyle(0, 1, { bold: true }), /权限/u); assert.deepEqual(history(studio), locked)
})

test('stale edit sessions keep their style drafts and cannot overwrite a later content change', t => {
  const { studio, api, book, first } = runtime(t), start = draft(module, book, first)
  first.getRange('A1').setRichTextStyle(0, 1, { bold: true })
  const before = structuredClone(first.getRange('A1').getCellData()), depth = history(studio)
  const operation = api.executeCommandSync(module.FORMAT_SHEET_CELL_INPUT_OPERATION, {
    unitId: book.getId(), sheetId: first.getSheetId(), row: 0, column: 0, input: start.value,
    session: start.session, editTrace: start.editTrace, inlineFormats: [], commandId: 'sheet.command.toggle-italic', selection: { start: 0, end: 1 },
  })
  assert.equal(operation.accepted, false); assert.deepEqual(operation.formats, [])
  assert.deepEqual(first.getRange('A1').getCellData(), before); assert.deepEqual(history(studio), depth)
})
