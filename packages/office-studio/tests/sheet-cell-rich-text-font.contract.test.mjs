import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { toExcelCellView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { sheetCellFontValues, sheetCellRichTextRunView, setSheetCellRichTextFont } from '../src/features/sheet-workbench/sheet-cell-rich-text-style.ts'
import { applySheetFontFormatCommand, applySheetViewFontFormatCommand } from '../src/features/sheet-workbench/sheet-font-format-command.ts'

// Draft: these contracts have not been executed.
const range = { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }
function sample() {
  const book = createBlankWorkbook(), cell = createCell(0, 0, 'abcd')
  book.id = 'rich-font-contract'; book.calculationMode = 'manual'
  cell.style.font.bold = true
  cell.richText = [{ id: 'first', text: 'ab', fontFamily: 'Arial', bold: false, font: { size: 9, italic: true, vendor: 7 } },
    { id: 'second', text: 'cd', font: { family: 'Georgia', size: 20 }, hyperlink: 'https://example.test', vendor: { keep: true } }]
  book.sheets[0].cells['0:0'] = cell
  return book
}

test('run aliases override inherited fonts while links and unknown fields keep their source ownership', () => {
  const cell = toExcelCellView(sample().sheets[0].cells['0:0']), run = cell.richText[0], before = structuredClone(run)
  const view = sheetCellRichTextRunView(cell, run)
  assert.equal(view.fontFamily, 'Arial'); assert.equal(view.fontSize, 9); assert.equal(view.bold, false); assert.equal(view.italic, true)
  const result = setSheetCellRichTextFont(cell.richText, 'family', 'Aptos')
  assert.equal(result.changed, true); assert.equal(result.runs[0].fontFamily, 'Aptos'); assert.equal(result.runs[0].font.family, 'Aptos')
  assert.equal(result.runs[0].font.vendor, 7); assert.equal(result.runs[1].hyperlink, 'https://example.test')
  result.runs[1].vendor.keep = false
  assert.equal(cell.richText[1].vendor.keep, true); assert.deepEqual(run, before)
})

test('reset removes both authored representations and keeps unrelated font and vendor fields', () => {
  const runs = [{ id: 'source', text: 'x', color: '#ff0000', font: { color: '#ff0000', bold: false, vendor: 5 }, vendor: { keep: 1 } }]
  const result = setSheetCellRichTextFont(runs, 'color', null)
  assert.deepEqual(result.runs, [{ id: 'source', text: 'x', font: { bold: false, vendor: 5 }, vendor: { keep: 1 } }])
  assert.equal(setSheetCellRichTextFont(result.runs, 'color', null).changed, false)
  assert.equal(runs[0].color, '#ff0000')
})

test('mixed state reads selected run intervals and uses the preceding run at an insertion boundary', () => {
  const cell = toExcelCellView(sample().sheets[0].cells['0:0'])
  assert.deepEqual(sheetCellFontValues(cell, 'bold'), [false, true])
  assert.deepEqual(sheetCellFontValues(cell, 'size', { start: 2, end: 4 }), [20])
  assert.deepEqual(sheetCellFontValues(cell, 'size', { start: 2, end: 2 }), [9])
  assert.deepEqual(sheetCellFontValues(cell, 'size', { start: 3, end: 3 }), [20])
  assert.deepEqual(sheetCellFontValues({ ...cell, formula: '=1' }, 'size'), [11])
})

test('whole-cell model and standalone font commands replace visible run styles and leave source content intact', () => {
  const book = sample(), sheet = book.sheets[0], view = toExcelWorkbookView(book).sheets[0], before = structuredClone(sheet.cells['0:0'])
  for (const action of [{ kind: 'family', value: 'Verdana' }, { kind: 'size', value: 14.5 },
    { kind: 'mark', field: 'italic', value: false }, { kind: 'color', field: 'color', value: '#102030' }]) {
    assert.equal(applySheetFontFormatCommand(sheet, [range], action), true)
    assert.equal(applySheetViewFontFormatCommand(view, [range], action), true)
    assert.deepEqual(sheet.cells['0:0'].richText, view.cells.A1.richText)
    assert.equal(applySheetFontFormatCommand(sheet, [range], action), false)
    assert.equal(applySheetViewFontFormatCommand(view, [range], action), false)
  }
  const cell = sheet.cells['0:0']
  assert.equal(cell.value, before.value); assert.equal(cell.dataType, before.dataType); assert.equal(cell.display, before.display)
  assert.deepEqual(cell.richText.map(run => run.id), ['first', 'second']); assert.equal(cell.richText[1].hyperlink, before.richText[1].hyperlink)
  assert.equal(cell.richText[0].font.vendor, 7)
})

test('a mixed bold toggle changes the explicit false run even when the cell default is already bold', () => {
  const sheet = sample().sheets[0]
  assert.equal(applySheetFontFormatCommand(sheet, [range], { kind: 'mark', field: 'bold' }), true)
  assert.deepEqual(sheetCellFontValues(toExcelCellView(sheet.cells['0:0']), 'bold'), [true, true])
  assert.equal(applySheetFontFormatCommand(sheet, [range], { kind: 'mark', field: 'bold', value: true }), false)
  assert.equal(applySheetFontFormatCommand(sheet, [range], { kind: 'mark', field: 'bold' }), true)
  assert.deepEqual(sheetCellFontValues(toExcelCellView(sheet.cells['0:0']), 'bold'), [false, false])
})

test('background edits do not overwrite rich fonts while text-color reset clears run overrides', () => {
  const sheet = sample().sheets[0], cell = sheet.cells['0:0']
  cell.richText[0].color = '#aabbcc'; const before = structuredClone(cell.richText)
  applySheetFontFormatCommand(sheet, [range], { kind: 'color', field: 'background', value: '#ddeeff' })
  assert.deepEqual(cell.richText, before)
  applySheetFontFormatCommand(sheet, [range], { kind: 'color', field: 'color', value: null })
  assert.equal(cell.richText[0].color, undefined); assert.equal(cell.richText[0].font.italic, true)
  assert.equal(cell.richText[1].hyperlink, 'https://example.test')
})

test('filtered rows remain unchanged and protected range rejection is atomic across rich-text cells', () => {
  const sheet = sample().sheets[0]
  sheet.cells['1:0'] = { ...structuredClone(sheet.cells['0:0']), id: '1:0', row: 1 }
  sheet.autoFilter = { range: { ...range, endRow: 2 }, columns: {}, filteredRows: [1] }
  const filtered = structuredClone(sheet.cells['1:0'])
  applySheetFontFormatCommand(sheet, [{ ...range, endRow: 1 }], { kind: 'size', value: 30 })
  assert.deepEqual(sheet.cells['1:0'], filtered)
  sheet.rangeProtectionRules = [{ id: 'locked', permissionId: 'locked', ownerUserId: 'remote', allowedUserIds: [], editState: 'onlyMe',
    viewState: 'othersCanView', description: '', ranges: [{ ...range, startRow: 1, endRow: 1 }] }]
  const before = structuredClone(sheet)
  assert.throws(() => applySheetFontFormatCommand(sheet, [{ ...range, endRow: 1 }], { kind: 'family', value: 'Arial' }))
  assert.deepEqual(sheet, before)
})
