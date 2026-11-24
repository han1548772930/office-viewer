import { createBlankWorkbook, createCell } from '../../src/domain/model.ts'
import { createChart, readChartData } from '../../src/features/charts/chart-engine.ts'

export const CANDLE_FIELDS = ['Date', 'Open', 'High', 'Low', 'Close']
export const CANDLE_ROWS = [['Rising', 10, 15, 8, 12], ['Falling', 12, 14, 7, 9],
  ['Doji', 11, 13, 10, 11], ['Gap', null, 13, 9, 10]]
export const CANDLE_HLC_MAPPING = { highIndex: 2, lowIndex: 3, closeIndex: 4 }
export function candlestickFixture(options = {}, rows = CANDLE_ROWS) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'candlestick-draft'
  const table = [CANDLE_FIELDS, ...rows]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [r + ':' + c, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type: 'candlestick', categoryField: 0, valueFields: [1, 2, 3, 4],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: table.length - 1, endCol: 4 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 320 }, yAxis: { min: 0, max: 20 }, options }, 'candlestick-draft')
  workbook.charts = [chart]
  return { workbook, chart, data: readChartData(workbook, chart) }
}
