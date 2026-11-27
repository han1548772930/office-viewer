import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell, defaultCellStyle } from '../src/domain/model.ts'
import { applyExcelWorkbookView, cellKey, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import {
  applySheetFontFormatCommand, applySheetViewFontFormatCommand, fontFormatActionForCommand,
  nextSheetFontFormatSize, normalizeSheetFontSize, SHEET_FONT_FORMAT_COMMAND_IDS,
  sheetFontFormatCommandForShortcut, sheetFontFormatCommandInput, sheetFontFormatInputRanges,
} from '../src/features/sheet-workbench/sheet-font-format-command.ts'

const range = (startRow, endRow = startRow, startCol = 0, endCol = startCol) => ({ startRow, endRow, startCol, endCol })
const content = sheet => Object.fromEntries(Object.entries(sheet.cells).map(([key, { value, formula, display, dataType, cachedDataType, richText, metadata }]) => [key, { value, formula, display, dataType, cachedDataType, richText, metadata }]))
const sample = () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'font-properties-contract'; workbook.calculationMode = 'manual'
  sheet.rowCount = 100; sheet.columnCount = 26
  sheet.cells['0:0'] = { ...createCell(0, 0, '00123'), dataType: 'text', metadata: { source: 'text' }, richText: { dataStream: '00123\r\n', textRuns: [{ st: 0, ed: 5, ts: { it: 1 } }] } }
  sheet.cells['1:0'] = { ...createCell(1, 0, .25), formula: '=1/4', dataType: 'formula', cachedDataType: 'number', metadata: { source: 'formula' } }
  sheet.cells['2:0'] = createCell(2, 0, 7)
  for (const [row, size] of [[0, 9], [1, 19], [2, 30]]) {
    const cell = sheet.cells[row + ':0']
    cell.style.font.family = 'Courier New'; cell.style.font.size = size
    cell.style.font.bold = true; cell.style.font.italic = true
    cell.style.font.color = '#123456'; cell.style.fill = '#abcdef'
    cell.style.border = { bottom: { style: 'double', color: '#333333' } }
    cell.style.metadata = { imported: true }
  }
  sheet.cells['4:4'] = createCell(4, 4, 'outside')
  return workbook
}

test('font-family range alias and toolbar command share raw, object and null contracts', () => {
  for (const id of ['sheet.command.set-font-family', 'sheet.command.set-range-font-family']) {
    assert.ok(SHEET_FONT_FORMAT_COMMAND_IDS.includes(id))
    assert.deepEqual(fontFormatActionForCommand(id, 'Arial'), { kind: 'family', value: 'Arial' })
    assert.deepEqual(fontFormatActionForCommand(id, { value: 'Times New Roman' }), { kind: 'family', value: 'Times New Roman' })
    assert.deepEqual(fontFormatActionForCommand(id, null), { kind: 'family', value: null })
    assert.deepEqual(fontFormatActionForCommand(id, { value: null }), { kind: 'family', value: null })
  }
})

test('font-size aliases and both stepper aliases are registered with the exact fixed IDs', () => {
  for (const id of ['sheet.command.set-font-size', 'sheet.command.set-range-fontsize']) {
    assert.ok(SHEET_FONT_FORMAT_COMMAND_IDS.includes(id))
    assert.deepEqual(fontFormatActionForCommand(id, 14.5), { kind: 'size', value: 14.5 })
    assert.deepEqual(fontFormatActionForCommand(id, { value: null }), { kind: 'size', value: null })
  }
  for (const [direction, ids] of [[1, ['sheet.command.increase-font-size', 'sheet.command.set-range-font-increase']], [-1, ['sheet.command.decrease-font-size', 'sheet.command.set-range-font-decrease']]]) {
    for (const id of ids) {
      assert.ok(SHEET_FONT_FORMAT_COMMAND_IDS.includes(id))
      assert.deepEqual(fontFormatActionForCommand(id), { kind: 'size-step', direction })
    }
  }
  assert.equal(new Set(SHEET_FONT_FORMAT_COMMAND_IDS).size, SHEET_FONT_FORMAT_COMMAND_IDS.length)
})

test('font-size shortcuts accept fixed punctuation key codes and shifted keys with Ctrl or Cmd', () => {
  const bindings = [
    [{ key: '.' }, 'sheet.command.increase-font-size'],
    [{ key: '>' }, 'sheet.command.increase-font-size'],
    [{ key: 'Unidentified', keyCode: 190 }, 'sheet.command.increase-font-size'],
    [{ key: '<', keyCode: 190 }, 'sheet.command.increase-font-size'],
    [{ key: ',' }, 'sheet.command.decrease-font-size'],
    [{ key: '<' }, 'sheet.command.decrease-font-size'],
    [{ key: 'Unidentified', keyCode: 188 }, 'sheet.command.decrease-font-size'],
    [{ key: '>', keyCode: 188 }, 'sheet.command.decrease-font-size'],
  ]
  for (const modifiers of [{ ctrlKey: true, metaKey: false }, { ctrlKey: false, metaKey: true }]) {
    for (const [key, commandId] of bindings) {
      const event = { ...modifiers, shiftKey: true, ...key }
      assert.equal(sheetFontFormatCommandForShortcut(event), commandId)
      assert.deepEqual(fontFormatActionForCommand(commandId), { kind: 'size-step', direction: commandId === 'sheet.command.increase-font-size' ? 1 : -1 })
      for (const extra of [{ shiftKey: false }, { ctrlKey: false, metaKey: false }, { altKey: true }, { isComposing: true }, { keyCode: 229 }, { defaultPrevented: true }]) {
        assert.equal(sheetFontFormatCommandForShortcut({ ...event, ...extra }), undefined)
      }
    }
  }
  assert.equal(sheetFontFormatCommandForShortcut({ key: 'X', ctrlKey: true, metaKey: false, shiftKey: true }), 'sheet.command.toggle-strike')
  for (const key of ['b', 'i', 'u', '[', ']']) {
    assert.equal(sheetFontFormatCommandForShortcut({ key, ctrlKey: true, metaKey: false, shiftKey: true }), undefined)
  }
})

test('command input preserves explicit objects and wraps all supported raw scalars without losing null or false', () => {
  const input = { unitId: 'other', subUnitId: 'target', ranges: [range(1)], value: 'Arial' }
  assert.equal(sheetFontFormatCommandInput(input), input)
  assert.deepEqual(sheetFontFormatCommandInput(), {})
  for (const value of ['Arial', 14, true, false, null]) assert.deepEqual(sheetFontFormatCommandInput(value), { value })
  assert.deepEqual(fontFormatActionForCommand('sheet.command.toggle-bold', sheetFontFormatCommandInput(false)), { kind: 'mark', field: 'bold', value: false })
})

test('mixed coordinate ranges normalize without mutation while explicit null and incomplete ranges reject', () => {
  const ranges = [range(0), { startRow: 2, endRow: 3, startColumn: 1, endColumn: 2 }], before = structuredClone(ranges)
  assert.deepEqual(sheetFontFormatInputRanges({ ranges }), [range(0), range(2, 3, 1, 2)])
  assert.deepEqual(ranges, before)
  assert.deepEqual(sheetFontFormatInputRanges({ ranges: [], range: range(0) }), [])
  for (const input of [{ ranges: null }, { range: null }, { ranges: {} }, { ranges: [null] }, { range: { startRow: 0, startColumn: 0 } }, { ranges: [range(1, 0)] }]) assert.throws(() => sheetFontFormatInputRanges(input))
})

test('font families preserve exact nonempty names and reject invalid input rather than reset it', () => {
  assert.deepEqual(fontFormatActionForCommand('sheet.command.set-font-family', 'Custom Office Font'), { kind: 'family', value: 'Custom Office Font' })
  assert.deepEqual(fontFormatActionForCommand('sheet.command.set-font-family', ' Arial '), { kind: 'family', value: ' Arial ' })
  for (const value of ['', '   ', 12, true, [], {}, undefined]) assert.throws(() => fontFormatActionForCommand('sheet.command.set-font-family', { value }), TypeError)
})

test('public font sizes retain finite positive values beyond the toolbar input limits', () => {
  for (const value of [.5, 5, 6, 11.5, 400, 500]) {
    assert.equal(normalizeSheetFontSize(value), value)
    assert.deepEqual(fontFormatActionForCommand('sheet.command.set-font-size', { value }), { kind: 'size', value })
  }
  assert.equal(normalizeSheetFontSize(null), null)
  for (const value of [0, -1, NaN, Infinity, -Infinity, '12', false, undefined]) assert.throws(() => normalizeSheetFontSize(value), RangeError)
})

test('font steppers use fixed one-sided limits including fractional and imported oversized values', () => {
  for (const [size, direction, expected] of [[6, -1, 6], [5, -1, 6], [400, 1, 400], [500, 1, 400], [500, -1, 499], [11.5, 1, 12.5], [399.5, 1, 400.5], [11.5, -1, 10.5]]) assert.equal(nextSheetFontFormatSize(size, direction), expected)
  for (const size of [null, 0, -1, NaN, Infinity]) assert.throws(() => nextSheetFontFormatSize(size, 1), RangeError)
  for (const direction of [0, 2, undefined]) assert.throws(() => nextSheetFontFormatSize(11, direction), RangeError)
})

test('model and view family and size edits preserve values, caches, rich text, metadata and other styles', () => {
  const workbook = sample(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0]
  const beforeModel = structuredClone(sheet), beforeView = structuredClone(view), targets = [range(0, 2)]
  for (const action of [{ kind: 'family', value: 'Arial' }, { kind: 'size', value: 14.5 }]) {
    assert.equal(applySheetFontFormatCommand(sheet, targets, action), true)
    assert.equal(applySheetViewFontFormatCommand(view, targets, action), true)
    assert.equal(applySheetFontFormatCommand(sheet, targets, action), false)
    assert.equal(applySheetViewFontFormatCommand(view, targets, action), false)
  }
  for (const row of [0, 1, 2]) {
    const key = row + ':0', viewKey = cellKey(row, 0), cell = sheet.cells[key], original = beforeModel.cells[key]
    assert.equal(cell.style.font.family, 'Arial'); assert.equal(cell.style.font.size, 14.5)
    assert.deepEqual({ ...cell, style: original.style }, original)
    assert.deepEqual({ ...cell.style, font: original.style.font }, original.style)
    assert.deepEqual({ ...cell.style.font, family: original.style.font.family, size: original.style.font.size }, original.style.font)
    assert.equal(view.cells[viewKey].fontFamily, 'Arial'); assert.equal(view.cells[viewKey].fontSize, 14.5)
    assert.deepEqual({ ...view.cells[viewKey], fontFamily: beforeView.cells[viewKey].fontFamily, fontSize: beforeView.cells[viewKey].fontSize }, beforeView.cells[viewKey])
  }
  assert.deepEqual(content(sheet), content(beforeModel))
  assert.deepEqual(sheet.cells['4:4'], beforeModel.cells['4:4'])
  const reopened = applyExcelWorkbookView({ ...workbook, sheets: [beforeModel] }, { ...toExcelWorkbookView(workbook), sheets: [view] })
  assert.deepEqual(content(reopened.sheets[0]), content(beforeModel))
  assert.equal(reopened.sheets[0].cells['1:0'].style.font.family, 'Arial')
  assert.equal(reopened.sheets[0].cells['1:0'].style.font.size, 14.5)
})

test('null family and size reset a full sparse worksheet without creating blank cells or clearing filtered content', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.rowCount = 1_048_576; sheet.columnCount = 16_384
  for (const [row, col] of [[0, 0], [1, 0], [1_048_575, 16_383]]) {
    const cell = createCell(row, col, 'kept')
    cell.style.font.family = 'Arial'; cell.style.font.size = 24; cell.style.font.bold = true
    sheet.cells[row + ':' + col] = cell
  }
  sheet.autoFilter = { range: range(0, 2), columns: {}, filteredRows: [1] }
  sheet.hiddenRows = [1_048_575]
  const view = toExcelWorkbookView(workbook).sheets[0], before = content(sheet), targets = [range(0, 1_048_575, 0, 16_383)]
  for (const kind of ['family', 'size']) {
    assert.equal(applySheetFontFormatCommand(sheet, targets, { kind, value: null }), true)
    assert.equal(applySheetViewFontFormatCommand(view, targets, { kind, value: null }), true)
    assert.equal(applySheetFontFormatCommand(sheet, targets, { kind, value: null }), false)
    assert.equal(applySheetViewFontFormatCommand(view, targets, { kind, value: null }), false)
  }
  assert.equal(Object.keys(sheet.cells).length, 3); assert.equal(Object.keys(view.cells).length, 3)
  for (const key of ['0:0', '1048575:16383']) {
    assert.equal(sheet.cells[key].style.font.family, 'Aptos'); assert.equal(sheet.cells[key].style.font.size, 11)
  }
  assert.equal(view.cells.XFD1048576.fontFamily, 'Aptos'); assert.equal(view.cells.XFD1048576.fontSize, 11)
  assert.equal(sheet.cells['1:0'].style.font.family, 'Arial'); assert.equal(sheet.cells['1:0'].style.font.size, 24)
  assert.deepEqual(content(sheet), before)
  assert.ok(Object.values(sheet.cells).every(cell => cell.style.font.bold))
})

test('family, size and step commands skip filtered rows while manually hidden rows remain in scope', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.autoFilter = { range: range(0, 2), columns: {}, filteredRows: [1] }; sheet.hiddenRows = [2]
  const view = toExcelWorkbookView(workbook).sheets[0], targets = [range(0, 2)]
  for (const action of [{ kind: 'family', value: 'Arial' }, { kind: 'size', value: 14 }, { kind: 'size-step', direction: 1 }]) {
    assert.equal(applySheetFontFormatCommand(sheet, targets, action), true)
    assert.equal(applySheetViewFontFormatCommand(view, targets, action), true)
  }
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.font.family), ['Arial', 'Courier New', 'Arial'])
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.font.size), [15, 19, 15])
  assert.deepEqual(['A1', 'A2', 'A3'].map(key => view.cells[key].fontSize), [15, 19, 15])
  assert.deepEqual(sheet.hiddenRows, [2])
})

test('default and null property no-ops keep empty model and view worksheets sparse', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0]
  for (const [targets, action] of [
    [[range(0, 1)], { kind: 'family', value: defaultCellStyle().font.family }],
    [[range(0, 1)], { kind: 'size', value: defaultCellStyle().font.size }],
    [[range(0, 1_048_575, 0, 16_383)], { kind: 'family', value: null }],
    [[range(0, 1_048_575, 0, 16_383)], { kind: 'size', value: null }],
  ]) {
    assert.equal(applySheetFontFormatCommand(sheet, targets, action), false)
    assert.equal(applySheetViewFontFormatCommand(view, targets, action), false)
  }
  assert.deepEqual(sheet.cells, {}); assert.deepEqual(view.cells, {})
})

test('direct property handlers reject malformed actions before changing either model or view', () => {
  const sheet = sample().sheets[0], view = toExcelWorkbookView({ ...sample(), sheets: [sheet] }).sheets[0]
  const beforeModel = structuredClone(sheet), beforeView = structuredClone(view)
  for (const action of [{ kind: 'family', value: '' }, { kind: 'family', value: 12 }, { kind: 'size', value: 0 }, { kind: 'size', value: NaN }, { kind: 'size', value: Infinity }, { kind: 'size-step', direction: 2 }]) {
    assert.throws(() => applySheetFontFormatCommand(sheet, [range(0, 2)], action))
    assert.throws(() => applySheetViewFontFormatCommand(view, [range(0, 2)], action))
    assert.deepEqual(sheet, beforeModel); assert.deepEqual(view, beforeView)
  }
})

test('size steps use the supplied active cell across a mixed selection and de-duplicate overlaps', () => {
  const workbook = sample(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0], before = content(sheet)
  const targets = [range(0, 2), range(1, 2)], action = { kind: 'size-step', direction: 1 }, primary = { row: 1, col: 0 }
  assert.equal(applySheetFontFormatCommand(sheet, targets, action, 'local-user', [], primary), true)
  assert.equal(applySheetViewFontFormatCommand(view, targets, action, 'local-user', primary), true)
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.font.size), [20, 20, 20])
  assert.deepEqual(['A1', 'A2', 'A3'].map(key => view.cells[key].fontSize), [20, 20, 20])
  assert.deepEqual(content(sheet), before)
})

test('implicit size-step primary comes from the last range while unselected cells retain their size', () => {
  const workbook = sample(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0], targets = [range(0), range(2)]
  assert.equal(applySheetFontFormatCommand(sheet, targets, { kind: 'size-step', direction: -1 }), true)
  assert.equal(applySheetViewFontFormatCommand(view, targets, { kind: 'size-step', direction: -1 }), true)
  assert.deepEqual([0, 1, 2].map(row => sheet.cells[row + ':0'].style.font.size), [29, 19, 29])
  assert.deepEqual(['A1', 'A2', 'A3'].map(key => view.cells[key].fontSize), [29, 19, 29])
})

test('a sparse primary resolves inherited size before applying one step to model and view targets', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.defaultStyle = { font: { family: 'Inter', size: 27 } }
  sheet.defaultStyle = { font: { size: 31.5 } }
  sheet.cells['0:0'] = createCell(0, 0, 'kept')
  sheet.cells['0:0'].style.font.size = 9
  const view = toExcelWorkbookView(workbook).sheets[0], targets = [range(0, 1)], action = { kind: 'size-step', direction: 1 }
  const primary = { row: 1, col: 0 }
  assert.equal(applySheetFontFormatCommand(sheet, targets, action, 'local-user', [], primary, workbook), true)
  assert.equal(applySheetViewFontFormatCommand(view, targets, action, 'local-user', primary), true)
  assert.deepEqual([0, 1].map(row => sheet.cells[row + ':0'].style.font.size), [32.5, 32.5])
  assert.deepEqual(['A1', 'A2'].map(key => view.cells[key].fontSize), [32.5, 32.5])
  assert.equal(sheet.cells['0:0'].value, 'kept')
  assert.equal(view.cells.A1.value, 'kept')
})

test('out-of-selection, fractional, missing or null primary positions reject an entire step transaction', () => {
  const workbook = sample(), sheet = workbook.sheets[0], view = toExcelWorkbookView(workbook).sheets[0]
  const beforeModel = structuredClone(sheet), beforeView = structuredClone(view), action = { kind: 'size-step', direction: 1 }
  for (const primary of [{ row: 4, col: 4 }, { row: -.5, col: 0 }, { row: 0, col: 16_384 }, { row: NaN, col: 0 }, {}, null, 1]) {
    assert.throws(() => applySheetFontFormatCommand(sheet, [range(0, 2)], action, 'local-user', [], primary))
    assert.throws(() => applySheetViewFontFormatCommand(view, [range(0, 2)], action, 'local-user', primary))
    assert.deepEqual(sheet, beforeModel); assert.deepEqual(view, beforeView)
  }
})

test('one protected target rejects family, size, null reset and step edits atomically', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'permission', ranges: [range(1)], ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' }]
  const view = toExcelWorkbookView(workbook).sheets[0], beforeModel = structuredClone(sheet), beforeView = structuredClone(view)
  for (const action of [{ kind: 'family', value: 'Arial' }, { kind: 'size', value: 24 }, { kind: 'family', value: null }, { kind: 'size', value: null }, { kind: 'size-step', direction: 1 }]) {
    assert.throws(() => applySheetFontFormatCommand(sheet, [range(0, 2)], action, 'guest'), /保护/)
    assert.throws(() => applySheetViewFontFormatCommand(view, [range(0, 2)], action, 'guest'), /保护/)
    assert.deepEqual(sheet, beforeModel); assert.deepEqual(view, beforeView)
  }
})

let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
function withStudio(run, workbook = sample()) {
  const studio = new OfficeStudio({ workbook, currentUserId: 'local-user' })
  try { return run(studio, FOffice.newAPI(studio)) } finally { studio.dispose() }
}

test('Facade family and size setters share commands, history, null reset and snapshot reopen', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('A1:A3'), before = content(studio.getWorkbookModel(book.getId()).sheets[0]), executed = []
  studio.events.addEvent('CommandExecuted', event => executed.push(event.id))
  assert.equal(target.setFontFamily('Arial'), target); assert.equal(target.setFontSize(14.5), target)
  assert.deepEqual(executed, ['sheet.command.set-font-family', 'sheet.command.set-font-size'])
  assert.equal(studio.editor.history.units[book.getId()].undo.length, 2)
  target.setFontFamily(null).setFontSize(null)
  let model = studio.getWorkbookModel(book.getId())
  assert.deepEqual([0, 1, 2].map(row => [model.sheets[0].cells[row + ':0'].style.font.family, model.sheets[0].cells[row + ':0'].style.font.size]), [['Aptos', 11], ['Aptos', 11], ['Aptos', 11]])
  assert.deepEqual(content(model.sheets[0]), before)
  book.undo(); book.undo()
  model = studio.getWorkbookModel(book.getId())
  assert.equal(model.sheets[0].cells['0:0'].style.font.family, 'Arial'); assert.equal(model.sheets[0].cells['0:0'].style.font.size, 14.5)
  book.redo(); book.redo()
  const reopened = api.createWorkbook({ ...book.save(), id: 'font-properties-reopened' })
  model = studio.getWorkbookModel(reopened.getId())
  assert.deepEqual(content(model.sheets[0]), before)
  assert.equal(model.sheets[0].cells['1:0'].style.font.family, 'Aptos'); assert.equal(model.sheets[0].cells['1:0'].style.font.size, 11)
}))

test('API aliases use the current selection for raw values and preserve explicit coordinate-style targets', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet()
  sheet.getRange('A1:A2').activate()
  assert.equal(studio.executeCommandSync('sheet.command.set-range-font-family', 'Arial'), true)
  assert.equal(studio.executeCommandSync('sheet.command.set-range-fontsize', 24), true)
  assert.equal(studio.executeCommandSync('sheet.command.set-range-font-family', null), true)
  assert.equal(studio.executeCommandSync('sheet.command.set-range-fontsize', { unitId: book.getId(), subUnitId: sheet.getSheetId(), ranges: [{ startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }], value: 32 }), true)
  const model = studio.getWorkbookModel(book.getId())
  assert.deepEqual([0, 1, 2].map(row => model.sheets[0].cells[row + ':0'].style.font.family), ['Aptos', 'Aptos', 'Courier New'])
  assert.deepEqual([0, 1, 2].map(row => model.sheets[0].cells[row + ':0'].style.font.size), [24, 24, 32])
}))

test('API step aliases honor an explicit primary and one-sided bounds without creating extra history', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), target = sheet.getRange('A1:A3')
  target.setFontSize(500)
  assert.equal(studio.executeCommandSync('sheet.command.set-range-font-decrease', { unitId: book.getId(), sheetId: sheet.getSheetId(), ranges: [range(0, 2)], primary: { row: 1, col: 0 } }), true)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['0:0'].style.font.size, 499)
  assert.equal(studio.executeCommandSync('sheet.command.set-range-font-increase', { unitId: book.getId(), sheetId: sheet.getSheetId(), ranges: [range(0, 2)] }), true)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['0:0'].style.font.size, 400)
  assert.equal(studio.executeCommandSync('sheet.command.increase-font-size', { unitId: book.getId(), sheetId: sheet.getSheetId(), ranges: [range(0, 2)] }), false)
  assert.equal(studio.editor.history.units[book.getId()].undo.length, 3)
  target.setFontSize(6)
  assert.equal(studio.executeCommandSync('sheet.command.decrease-font-size', { unitId: book.getId(), sheetId: sheet.getSheetId(), ranges: [range(0, 2)] }), false)
  assert.equal(studio.editor.history.units[book.getId()].undo.length, 4)
}))

test('canceled family and size commands publish no successful event and leave model and history untouched', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('A1'), before = structuredClone(studio.getWorkbookModel(book.getId())), events = []
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  studio.events.addEvent('BeforeCommandExecute', event => { if (['sheet.command.set-font-family', 'sheet.command.set-font-size'].includes(event.id)) event.cancel = true })
  assert.throws(() => target.setFontFamily('Arial'), /canceled/)
  assert.throws(() => target.setFontSize(24), /canceled/)
  assert.deepEqual(studio.getWorkbookModel(book.getId()), before)
  assert.deepEqual(events, []); assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))

test('invalid family, size, range or primary API input cannot partially commit an edit', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), sheet = book.getActiveSheet(), target = sheet.getRange('A1:A3'), before = structuredClone(studio.getWorkbookModel(book.getId())), events = []
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  for (const value of ['', '  ', 12]) assert.throws(() => target.setFontFamily(value), TypeError)
  for (const value of [0, -1, NaN, Infinity, '24']) assert.throws(() => target.setFontSize(value), RangeError)
  for (const input of [{ ranges: null }, { range: null }, { ranges: [range(-1)] }, { ranges: [range(0, 2)], primary: { row: 4, col: 4 } }, { ranges: [range(0, 2)], primary: null }]) assert.throws(() => studio.executeCommandSync('sheet.command.increase-font-size', { unitId: book.getId(), sheetId: sheet.getSheetId(), ...input }))
  assert.deepEqual(studio.getWorkbookModel(book.getId()), before)
  assert.deepEqual(events, []); assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))

test('API permission checks cover complete property ranges and null resets', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('A1:A3'), locked = structuredClone(studio.getWorkbookModel(book.getId())), events = []
  locked.sheets[0].rangeProtectionRules = [{ id: 'locked', permissionId: 'permission', ranges: [range(1)], ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView', description: '' }]
  studio.editor.workbook.replace(locked)
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  for (const run of [() => target.setFontFamily('Arial'), () => target.setFontSize(24), () => target.setFontFamily(null), () => target.setFontSize(null)]) assert.throws(run, /保护|permission/i)
  assert.deepEqual(studio.getWorkbookModel(book.getId()), locked)
  assert.deepEqual(events, []); assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
}))

test('successful property no-ops still emit CommandExecuted while preserving history and sparsity', () => withStudio((studio, api) => {
  const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('B1:B2'), events = []
  studio.events.addEvent('CommandExecuted', event => events.push(event.id))
  target.setFontFamily('Aptos').setFontSize(11).setFontFamily(null).setFontSize(null)
  assert.deepEqual(events, ['sheet.command.set-font-family', 'sheet.command.set-font-size', 'sheet.command.set-font-family', 'sheet.command.set-font-size'])
  assert.equal(studio.editor.history.units[book.getId()]?.undo.length ?? 0, 0)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['0:1'], undefined)
  assert.equal(studio.getWorkbookModel(book.getId()).sheets[0].cells['1:1'], undefined)
}))

test('registered font commands combine sheet and table filtered rows without losing manually hidden rows', () => {
  const workbook = sample(), sheet = workbook.sheets[0]
  sheet.cells['3:0'] = createCell(3, 0, 8); sheet.cells['3:0'].style.font.family = 'Courier New'; sheet.cells['3:0'].style.font.size = 9
  sheet.autoFilter = { range: range(0, 3), columns: {}, filteredRows: [1] }; sheet.hiddenRows = [3]
  workbook.tables = [{ id: 'font-table', name: 'FontTable', sheetId: sheet.id, range: range(0, 3), columns: [{ id: 'font-column', name: 'Value' }], showHeaderRow: true, showTotalsRow: false, styleName: 'TableStyleMedium2', autoFilter: { range: range(0, 3), columns: {}, filteredRows: [2] } }]
  withStudio((studio, api) => {
    const book = api.getActiveWorkbook(), target = book.getActiveSheet().getRange('A1:A4')
    target.setFontFamily('Arial').setFontSize(14)
    const current = studio.getWorkbookModel(book.getId()).sheets[0]
    assert.deepEqual([0, 1, 2, 3].map(row => current.cells[row + ':0'].style.font.family), ['Arial', 'Courier New', 'Courier New', 'Arial'])
    assert.deepEqual([0, 1, 2, 3].map(row => current.cells[row + ':0'].style.font.size), [14, 19, 30, 14])
    assert.deepEqual(current.hiddenRows, [3])
  }, workbook)
})

test('font history and explicit Facade targets stay isolated while a second workbook is active', () => withStudio((studio, api) => {
  const first = api.getActiveWorkbook(), firstRange = first.getActiveSheet().getRange('A1')
  firstRange.setFontFamily('Arial')
  const second = api.createWorkbook({ ...sample(), id: 'font-properties-second' }), secondRange = second.getActiveSheet().getRange('A1')
  secondRange.setFontSize(24); firstRange.setFontSize(32)
  assert.equal(studio.activeWorkbookId, second.getId())
  first.undo()
  assert.equal(studio.getWorkbookModel(first.getId()).sheets[0].cells['0:0'].style.font.size, 9)
  assert.equal(studio.getWorkbookModel(first.getId()).sheets[0].cells['0:0'].style.font.family, 'Arial')
  assert.equal(studio.getWorkbookModel(second.getId()).sheets[0].cells['0:0'].style.font.size, 24)
  first.redo()
  assert.equal(studio.getWorkbookModel(first.getId()).sheets[0].cells['0:0'].style.font.size, 32)
  assert.equal(studio.activeWorkbookId, second.getId())
  assert.equal(studio.editor.history.units[first.getId()].undo.length, 2)
  assert.equal(studio.editor.history.units[second.getId()].undo.length, 1)
}))
