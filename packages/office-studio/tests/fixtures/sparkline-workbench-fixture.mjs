import { createBlankSheet, createBlankWorkbook, createCell } from '../../src/domain/model.ts'
import { createSparklinePlacementPatches } from '../../src/features/sparkline/sparkline-model.ts'
import { parseSparklineReference } from '../../src/features/sparkline/sparkline-reference.ts'

// Native model fixture for deferred contract/browser work. It is not an accepted Office file sample.
export const sparklineJson = value => JSON.parse(JSON.stringify(value))
export function sparklineCell(sheet, address) {
  const { range } = parseSparklineReference(address), key = `${range.startRow}:${range.startCol}`
  return sheet.cells[key] ??= createCell(range.startRow, range.startCol, null)
}
export function installSparklinePatches(sheet, patches) {
  for (const [address, spec] of Object.entries(patches)) {
    const cell = sparklineCell(sheet, address)
    if (spec === null) delete cell.sparkline
    else cell.sparkline = sparklineJson(spec)
  }
  return sheet
}
export function sparklineWorkbook({ withGroups = true, id = 'sparkline-book' } = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = id; workbook.name = 'Sparkline native fixture'; workbook.calculationMode = 'manual'
  sheet.id = 'data'; sheet.name = '迷你图'; workbook.activeSheetId = sheet.id
  sheet.rowCount = 100; sheet.columnCount = 20; sheet.defaultRowHeight = 30; sheet.defaultColumnWidth = 100
  sheet.defaultRowAutoHeight = false
  const rows = [[2, 8, -3, 4], [4, 5, 1, -2], [10, 30, 20, 40], [-6, 3, 8, 1]]
  for (const [row, values] of rows.entries()) for (const [index, value] of values.entries()) {
    sheet.cells[`${row}:${index + 1}`] = createCell(row, index + 1, value)
  }
  const other = createBlankSheet("Data, O'Neil"); other.id = 'other'; other.rowCount = 100; other.columnCount = 20
  other.defaultRowHeight = 30; other.defaultColumnWidth = 100
  for (const [row, values] of [[100, 300, 200, 400], [4, -1, 3, 2]].entries()) for (const [index, value] of values.entries()) {
    other.cells[`${row}:${index + 1}`] = createCell(row, index + 1, value)
  }
  workbook.sheets.push(other)
  if (withGroups) {
    installSparklinePatches(sheet, createSparklinePlacementPatches({ sourceRanges: ['B1:E2'], targetRanges: ['G1:G2'], groupId: 'north',
      settings: { type: 'line', color: '#246fce', themeType: 'custom', points: { highPoint: { visible: true, color: '#cc2244' } },
        extremumMin: { type: 'group', value: -20 }, extremumMax: { type: 'individual', value: 50 } } }, sheet))
    installSparklinePatches(sheet, createSparklinePlacementPatches({ sourceRanges: ['B3:E4'], targetRanges: ['G4:G5'], groupId: 'south',
      settings: { type: 'column', color: '#228855', themeType: 'custom', axisOptions: { visible: true, reverse: true, color: '#654321' },
        points: { lowPoint: { visible: true, color: '#dd9922' } }, extremumMin: { type: 'individual', value: -40 }, extremumMax: { type: 'group', value: 100 } } }, sheet))
  }
  return workbook
}
