import { createBlankWorkbook, createCell } from '../../src/domain/model.ts'
import { createChart, readChartData } from '../../src/features/charts/chart-engine.ts'
import { TREEMAP_TABLE } from './chart-treemap-fixtures.mjs'

// One signed authoring sample shared with P26. No validation runs on import.
export const SUNBURST_TABLE = TREEMAP_TABLE
export function sunburstFixture(options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'sunburst-draft'
  sheet.cells = Object.fromEntries(SUNBURST_TABLE.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'sunburst', categoryField: 0, valueFields: [2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: SUNBURST_TABLE.length - 1, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 560, height: 400 }, palette: ['#112233', '#ddeeff'],
    options: { ...options, mapping: { hierarchyIndexes: [0, 1], ...options.mapping } } }, 'sunburst-draft')
  return { workbook, sheet, chart, data: readChartData(workbook, chart) }
}
