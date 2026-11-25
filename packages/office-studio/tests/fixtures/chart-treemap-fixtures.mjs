import { createBlankWorkbook, createCell } from '../../src/domain/model.ts'
import { createChart, readChartData } from '../../src/features/charts/chart-engine.ts'

// Shared authoring sample for P26 drafts. Importing it starts no validation.
export const TREEMAP_RECORDS = [
  ['South', 'Leaf A', 8], ['North', 'Leaf B', 12], [null, 'Leaf C', 6],
  ['South', null, 3], ['South', 'Leaf A', -2], ['South', 'Dropped', -4], ['North', null, -5],
]
export const TREEMAP_TABLE = [['Region', 'Category', 'Weight'], ...TREEMAP_RECORDS]
export function treemapFixture(options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'treemap-draft'
  sheet.cells = Object.fromEntries(TREEMAP_TABLE.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'treemap', categoryField: 0, valueFields: [2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: TREEMAP_TABLE.length - 1, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 560, height: 400 }, palette: ['#112233', '#ddeeff'],
    options: { ...options, mapping: { hierarchyIndexes: [0, 1], ...options.mapping } } }, 'treemap-draft')
  return { workbook, sheet, chart, data: readChartData(workbook, chart) }
}
