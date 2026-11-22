import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createServer } from 'vite'
import { createCell, createBlankWorkbook } from '../src/domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { RESET_TEXT_COLOR, RESET_BACKGROUND_COLOR, resetSheetColors } from '../src/features/sheet-workbench/sheet-color-reset.ts'

test('reset a sparse selection without affecting other color, other ranges, or cell contents', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  for (const row of [0, 1, 3]) {
    const cell = createCell(row, 0, 'kept')
    cell.style.font.color = '#123456'
    cell.style.font.bold = true
    cell.style.fill = '#ffffff'
    sheet.cells[`${row}:0`] = cell
  }
  const ranges = [0, 3].map(row => ({ startRow: row, endRow: row, startCol: 0, endCol: 16383 }))
  resetSheetColors(sheet, ranges, RESET_TEXT_COLOR)
  assert.equal(sheet.cells['0:0'].style.font.color, null)
  assert.equal(sheet.cells['0:0'].style.fill, '#ffffff')
  assert.equal(sheet.cells['1:0'].style.font.color, '#123456')
  resetSheetColors(sheet, ranges, RESET_BACKGROUND_COLOR)
  assert.equal(sheet.cells['3:0'].style.fill, null)
  assert.equal(sheet.cells['1:0'].style.fill, '#ffffff')
  assert.equal(sheet.cells['0:0'].style.font.bold, true)
  assert.equal(sheet.cells['0:0'].value, 'kept')
  assert.equal(Object.keys(sheet.cells).length, 3)
  const view = toExcelWorkbookView(workbook)
  view.sheets[0].cells.A2.color = null
  const roundtrip = applyExcelWorkbookView(workbook, view)
  assert.equal(roundtrip.sheets[0].cells['1:0'].style.font.color, null)
  assert.equal(roundtrip.sheets[0].cells['1:0'].style.fill, '#ffffff')
})

test('range API and Ribbon reset share workbook history across disjoint regions', async () => {
  const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  try {
    const { FOffice, OfficeStudio } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts')
    const studio = new OfficeStudio()
    try {
      const workbook = FOffice.newAPI(studio).getActiveWorkbook()
      const sheet = workbook.getActiveSheet()
      sheet.getRange('A1:A3').setBackground('#ffffff').setFontColor('#aa0033')
      const cell = row => studio.getWorkbookModel(workbook.getId()).sheets[0].cells[`${row}:0`]
      studio.executeCommandSync(RESET_BACKGROUND_COLOR, { unitId: workbook.getId(), sheetId: sheet.getSheetId(), ranges: [0, 2].map(row => ({ startRow: row, endRow: row, startCol: 0, endCol: 0 })) })
      assert.deepEqual([0, 1, 2].map(row => cell(row).style.fill), [null, '#ffffff', null])
      workbook.undo()
      assert.deepEqual([0, 1, 2].map(row => cell(row).style.fill), Array(3).fill('#ffffff'))
      workbook.redo()
      sheet.getRange('A1').setFontColor(null)
      assert.equal(cell(0).style.font.color, null)
      workbook.undo()
      assert.equal(cell(0).style.font.color, '#aa0033')
      assert.equal(cell(0).style.fill, null)
    } finally { studio.dispose() }
  } finally { await vite.close() }
})
