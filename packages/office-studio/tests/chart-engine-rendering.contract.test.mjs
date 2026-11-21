import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let createBlankWorkbook
let createChart
let chartLayout
let chartRenderCommands

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  const model = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  const engine = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts')
  createBlankWorkbook = model.createBlankWorkbook
  createChart = engine.createChart
  chartLayout = engine.chartLayout
  chartRenderCommands = engine.chartRenderCommands
})

afterAll(async () => {
  await server?.close()
})

test('chart scene renders category labels, value ticks, gridlines and axis titles', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells = {
    '0:0': { value: 'Region' },
    '0:1': { value: 'Sales' },
    '1:0': { value: 'North' },
    '1:1': { value: 10 },
    '2:0': { value: 'South' },
    '2:1': { value: 20 },
  }
  const chart = createChart(workbook, sheet.id, {
    type: 'line',
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, orientation: 'columns' },
    position: { x: 20, y: 30 },
    size: { width: 420, height: 280 },
    legend: { visible: false },
    xAxis: { title: 'Region' },
    yAxis: { title: 'Sales' },
  })

  const layout = chartLayout(workbook, chart)
  const commands = chartRenderCommands(workbook, chart)
  const texts = commands.filter((command) => command.kind === 'text')
  const horizontalGridlines = commands.filter((command) => command.kind === 'line' && command.x1 === layout.plot.x && command.x2 === layout.plot.x + layout.plot.width)

  assert.deepEqual(texts.filter((command) => ['North', 'South'].includes(command.text)).map((command) => command.text), ['North', 'South'])
  assert.deepEqual(texts.filter((command) => ['Region', 'Sales'].includes(command.text)).map((command) => command.text), ['Region', 'Sales'])
  assert.equal(texts.find((command) => command.text === 'Sales')?.rotation, -90)
  assert.equal(texts.filter((command) => ['0', '5', '10', '15', '20'].includes(command.text)).length, 5)
  assert.equal(horizontalGridlines.length, 6)
  assert.equal(layout.series[0].points[0].x, texts.find((command) => command.text === 'North').x)
})

test('logarithmic chart axes use geometrically spaced labels', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells = {
    '0:0': { value: 'Month' },
    '0:1': { value: 'Users' },
    '1:0': { value: 'Jan' },
    '1:1': { value: 1 },
    '2:0': { value: 'Feb' },
    '2:1': { value: 10000 },
  }
  const chart = createChart(workbook, sheet.id, {
    type: 'line',
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, orientation: 'columns' },
    position: { x: 0, y: 0 },
    size: { width: 360, height: 240 },
    legend: { visible: false },
    yAxis: { scaleType: 'log', logBase: 10 },
  })

  const labels = chartRenderCommands(workbook, chart)
    .filter((command) => command.kind === 'text')
    .map((command) => command.text)
    .filter((text) => /^\d/.test(text))

  assert.deepEqual(labels, ['1', '10', '100', '1000', '10000'])
})

test('logarithmic chart axes stay finite for constant or non-positive data', () => {
  const workbook = createBlankWorkbook()
  const sheet = workbook.sheets[0]
  sheet.cells = {
    '0:0': { value: 'Month' },
    '0:1': { value: 'Users' },
    '1:0': { value: 'Jan' },
    '1:1': { value: 10 },
    '2:0': { value: 'Feb' },
    '2:1': { value: 10 },
  }
  const config = {
    type: 'line',
    source: { range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 }, orientation: 'columns' },
    position: { x: 0, y: 0 },
    size: { width: 360, height: 240 },
    legend: { visible: false },
    yAxis: { scaleType: 'log', logBase: 10 },
  }
  const constantChart = createChart(workbook, sheet.id, config)
  assert.deepEqual([chartLayout(workbook, constantChart).min, chartLayout(workbook, constantChart).max], [10, 100])

  sheet.cells['1:1'].value = -5
  sheet.cells['2:1'].value = 0
  const emptyLogChart = createChart(workbook, sheet.id, config)
  const layout = chartLayout(workbook, emptyLogChart)
  assert.deepEqual([layout.min, layout.max], [1, 10])
  assert.ok(chartRenderCommands(workbook, emptyLogChart).every((command) => {
    if (command.kind === 'text') return Number.isFinite(command.x) && Number.isFinite(command.y)
    if (command.kind === 'line') return [command.x1, command.y1, command.x2, command.y2].every(Number.isFinite)
    if (command.kind === 'rect') return [command.x, command.y, command.width, command.height].every(Number.isFinite)
    return true
  }))
})
