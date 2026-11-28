import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Draft: run in the unified verification stage, together with the real editor.
let server, sdk
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  sdk = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
function setup(t, customize) {
  const model = sdk.createBlankWorkbook(); model.id = 'function-book'; const source = model.sheets[0]
  source.id = 'data'; model.activeSheetId = source.id; source.rowCount = 20; source.columnCount = 8
  source.cells['0:0'] = sdk.createCell(0, 0, 2); source.cells['1:0'] = sdk.createCell(1, 0, 3)
  source.cells['0:1'] = sdk.createCell(0, 1, 4); source.cells['1:1'] = sdk.createCell(1, 1, 5)
  customize?.(source, model)
  const studio = new sdk.OfficeStudio({ workbook: model, sheetTextMeasure: text => text.length * 8 })
  t.onTestFinished(() => studio.dispose())
  const api = sdk.FOffice.newAPI(studio), book = api.getActiveWorkbook(), sheet = book.getActiveSheet()
  const history = () => studio.editor.history.units[book.getId()]
  return { studio, api, book, sheet, history }
}
const item = (formula, row, column, endRow = row, endColumn = column) => ({ range: { startRow: row, startColumn: column, endRow, endColumn },
  primary: { row, column }, formula })

test('automatic numeric matrix insertion uses one transaction and survives undo, redo and native snapshot reopening', t => {
  const h = setup(t); h.sheet.getRange('A1:B3').activate()
  const depth = h.history().undo.length, events = []
  const listener = h.studio.events.addEvent('SheetValueChanged', event => events.push(event))
  t.onTestFinished(() => listener.dispose())
  assert.equal(h.studio.executeCommandSync(sdk.SHEET_INSERT_FUNCTION_OPERATION, { value: 'SUM' }), true)
  assert.deepEqual(h.sheet.getRange('A3:B3').getFormulas(), [['=SUM(A1:A2)', '=SUM(B1:B2)']])
  assert.deepEqual(h.sheet.getRange('A3:B3').getValues(), [[5, 9]]); assert.equal(h.history().undo.length, depth + 1)
  assert.equal(events.length, 1); assert.equal(events[0].ranges.length, 2)
  h.book.undo(); assert.deepEqual(h.sheet.getRange('A3:B3').getFormulas(), [[null, null]])
  h.book.redo(); assert.deepEqual(h.sheet.getRange('A3:B3').getValues(), [[5, 9]])
  const reopened = h.api.createWorkbook({ ...h.book.save(), id: 'function-reopen' })
  assert.deepEqual(reopened.getActiveSheet().getRange('A3:B3').getFormulas(), h.sheet.getRange('A3:B3').getFormulas())
})

test('the fixed public list DTO expands shared relative formulas and preserves target direct styles', t => {
  const h = setup(t, sheet => { const cell = sdk.createCell(2, 1, 'old'); cell.richText = [{ text: 'old', bold: true }]
    cell.authoredStyle = { fill: '#123456' }; sheet.cells[cell.id] = cell })
  const depth = h.history().undo.length
  assert.equal(h.studio.executeCommandSync(sdk.INSERT_SHEET_FUNCTION_COMMAND, { unitId: h.book.getId(), subUnitId: h.sheet.getSheetId(),
    list: [item('=SUM(A1:$A$2)', 2, 0, 2, 1)] }), true)
  assert.deepEqual(h.sheet.getRange('A3:B3').getFormulas(), [['=SUM(A1:$A$2)', '=SUM(B1:$A$2)']])
  assert.equal(h.sheet.getRange('B3').getCellData().style.fill, '#123456')
  assert.equal(h.sheet.getRange('B3').getCellData().richText, undefined)
  assert.equal(h.history().undo.length, depth + 1)
})

test('one protected destination rejects every planned write and preserves the redo branch', t => {
  const h = setup(t, sheet => { sheet.rangeProtectionRules = [{ id: 'blocked', permissionId: 'blocked', ranges: [{ startRow: 2, endRow: 2, startCol: 1, endCol: 1 }],
    description: '', ownerUserId: 'other', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }] })
  h.sheet.getRange('F1').setValue('temporary'); h.book.undo()
  const before = JSON.stringify(h.book.save()), history = JSON.stringify(h.history())
  assert.throws(() => h.studio.executeCommandSync(sdk.INSERT_SHEET_FUNCTION_COMMAND, { unitId: h.book.getId(), subUnitId: h.sheet.getSheetId(),
    list: [item('=SUM(A1:A2)', 2, 0, 2, 1)] }), /保护|权限/)
  assert.equal(JSON.stringify(h.book.save()), before); assert.equal(JSON.stringify(h.history()), history)
})

test('a stopping data-validation rule rejects the complete automatic batch before any cell is published', t => {
  const h = setup(t)
  h.sheet.getRange('B3').setDataValidation({ type: 'whole', operator: 'between', formula1: '1', formula2: '6',
    errorStyle: 'stop', error: '结果必须在一到六之间' })
  h.sheet.getRange('A1:B3').activate()
  const before = JSON.stringify(h.book.save()), depth = h.history().undo.length
  assert.throws(() => h.studio.executeCommandSync(sdk.SHEET_INSERT_FUNCTION_OPERATION, { value: 'SUM' }), /一到六|验证/)
  assert.equal(JSON.stringify(h.book.save()), before); assert.equal(h.history().undo.length, depth)
  assert.deepEqual(h.sheet.getRange('A3:B3').getFormulas(), [[null, null]])
})

test('validation observes every final formula in a candidate even when the referenced result is written later in the list', t => {
  const h = setup(t)
  h.sheet.getRange('B3').setDataValidation({ type: 'whole', operator: 'equal', formula1: '6', errorStyle: 'stop' })
  const depth = h.history().undo.length
  assert.equal(h.studio.executeCommandSync(sdk.INSERT_SHEET_FUNCTION_COMMAND, { list: [item('=A3+1', 2, 1), item('=SUM(A1:A2)', 2, 0)] }), true)
  assert.deepEqual(h.sheet.getRange('A3:B3').getValues(), [[5, 6]])
  assert.equal(h.history().undo.length, depth + 1)
})

test('a command observer changing selection invalidates an editor-owned insertion before any content changes', t => {
  const h = setup(t); h.sheet.getRange('A3').activate(); let opened = 0, released = 0
  const owner = sdk.registerSheetFunctionEditor(h.studio, h.book.getId(), { canInsert: () => true, open: () => { opened++; return true }, release: () => released++ })
  const before = JSON.stringify(h.book.save())
  const listener = h.studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === sdk.SHEET_INSERT_FUNCTION_OPERATION) h.sheet.getRange('F5').activate()
  })
  assert.throws(() => owner.execute('SUM'), /选区/); listener.dispose()
  assert.equal(opened, 0); assert.equal(JSON.stringify(h.book.save()), before)
  owner.dispose(); assert.equal(released, 1)
})

test('single-cell operations require a live editor, open real unclosed text and create no content history', t => {
  const h = setup(t); h.sheet.getRange('A3').activate()
  assert.equal(h.studio.executeCommandSync(sdk.SHEET_INSERT_FUNCTION_OPERATION, { value: 'SUM' }), false)
  let target, released = 0
  const owner = sdk.registerSheetFunctionEditor(h.studio, h.book.getId(), { canInsert: () => true,
    open: value => { target = value; return true }, release: () => released++ })
  const depth = h.history().undo.length, snapshot = JSON.stringify(h.book.save())
  assert.equal(owner.execute('SUM'), true); assert.equal(target.formula, '=SUM(A1:A2)')
  assert.equal(h.history().undo.length, depth); assert.equal(JSON.stringify(h.book.save()), snapshot)
  owner.dispose(); assert.equal(owner.execute('SUM'), false); assert.equal(released, 1)
})

test('obsolete expected versions reject a public batch while preserving intervening edits', t => {
  const h = setup(t), revision = h.studio.getWorkbookRevision(h.book.getId())
  h.sheet.getRange('G1').setValue('later')
  const before = JSON.stringify(h.book.save()), depth = h.history().undo.length
  assert.throws(() => h.studio.executeCommandSync(sdk.INSERT_SHEET_FUNCTION_COMMAND, { unitId: h.book.getId(), expectedRevision: revision,
    list: [item('=1', 2, 0)] }), /版本/)
  assert.equal(JSON.stringify(h.book.save()), before); assert.equal(h.history().undo.length, depth)
})

test('a reentrant edit before the inner batch is preserved and the outer plan does not overwrite it', t => {
  const h = setup(t); h.sheet.getRange('A1:B3').activate()
  const listener = h.studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === sdk.INSERT_SHEET_FUNCTION_COMMAND) h.sheet.getRange('G1').setValue('observer')
  })
  assert.throws(() => h.studio.executeCommandSync(sdk.SHEET_INSERT_FUNCTION_OPERATION, { value: 'SUM' }), /变化/)
  listener.dispose(); assert.equal(h.sheet.getRange('G1').getValue(), 'observer')
  assert.deepEqual(h.sheet.getRange('A3:B3').getFormulas(), [[null, null]])
  h.book.undo(); assert.equal(h.sheet.getRange('G1').getValue(), null)
})

test('replacement and studio disposal release owners, and invalid formulas/coordinates do not publish partial edits', t => {
  const h = setup(t); let first = 0, second = 0
  const one = sdk.registerSheetFunctionEditor(h.studio, h.book.getId(), { canInsert: () => true, open: () => true, release: () => first++ })
  sdk.registerSheetFunctionEditor(h.studio, h.book.getId(), { canInsert: () => true, open: () => true, release: () => second++ })
  assert.equal(first, 1); assert.equal(one.execute('SUM'), false)
  const before = JSON.stringify(h.book.save())
  assert.throws(() => h.studio.executeCommandSync(sdk.INSERT_SHEET_FUNCTION_COMMAND, { list: [item('=1', 2, 0), item('=2', 20, 0)] }), /边界/)
  assert.equal(JSON.stringify(h.book.save()), before)
  assert.throws(() => h.studio.executeCommandSync(sdk.INSERT_SHEET_FUNCTION_COMMAND, { list: [item('plain text', 2, 0)] }), /公式/)
  h.studio.dispose(); assert.equal(second, 1)
})
