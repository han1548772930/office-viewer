import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { applyExcelWorkbookView, toExcelWorkbookView } from '../src/services/excelViewAdapter.ts'

let server
let apiModule
let chartEngine
let pivotPanelState

before(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  chartEngine = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-engine.ts')
  pivotPanelState = await server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-panel-state.ts')
})

after(async () => {
  await server?.close()
})

function setup() {
  const studio = new apiModule.OfficeStudio()
  const officeApi = apiModule.FOffice.newAPI(studio)
  const workbook = officeApi.getActiveWorkbook()
  const source = workbook.getActiveSheet()
  const target = workbook.insertSheet('Pivot')
  source.getRange('A1:C5').setValues([
    ['Region', 'Quarter', 'Sales'],
    ['East', 'Q1', 100],
    ['East', 'Q2', 150],
    ['West', 'Q1', 80],
    ['West', 'Q2', 120],
  ])
  return { studio, officeApi, workbook, source, target }
}

async function addBasicPivot(workbook, source, target) {
  const pivot = await workbook.addPivotTable(
    {
      unitId: workbook.getId(),
      subUnitId: source.getSheetId(),
      sheetName: source.getSheetName(),
      range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
    },
    apiModule.PositionTypeEnum.Existing,
    { unitId: workbook.getId(), subUnitId: target.getSheetId(), row: 0, col: 0 },
  )
  await pivot.addField(0, 'row')
  await pivot.addField(1, 'column')
  await pivot.addField(2, 'value')
  return pivot
}

test('creates and renders a pivot with row, column, value fields and grand totals', async () => {
  const { studio, workbook, source, target } = setup()
  const rendered = []
  const subscription = apiModule.FOffice.newAPI(studio).addEvent('PivotTableRendered', (event) => rendered.push(event))
  const pivot = await addBasicPivot(workbook, source, target)

  assert.deepEqual(pivot.getSourceFieldsInfo().map((field) => field.fieldName), ['Region', 'Quarter', 'Sales'])
  assert.deepEqual(pivot.getFieldIdsByArea('row'), ['0'])
  assert.deepEqual(pivot.getFieldIdsByArea('column'), ['1'])
  assert.deepEqual(pivot.getFieldIdsByArea('value'), ['2'])
  assert.deepEqual(pivot.getData(), [
    [null, 'Q1', 'Q2', 'Grand total'],
    ['Region', 'Sales', 'Sales', 'Sales'],
    ['East', 100, 150, 250],
    ['West', 80, 120, 200],
    ['Grand total', 180, 270, 450],
  ])
  assert.equal(workbook.getPivotTableByCell(target.getSheetId(), 2, 1)?.getPivotTableId(), pivot.getPivotTableId())
  assert.equal(target.getPivotTableByCell(2, 1)?.getPivotTableId(), pivot.getPivotTableId())
  assert.ok(rendered.length >= 3)
  subscription.dispose()
  studio.dispose()
})

test('supports label filtering, sorting, value settings, refresh and drill down', async () => {
  const { studio, officeApi, workbook, source, target } = setup()
  const pivot = await addBasicPivot(workbook, source, target)
  await pivot.setLabelManualFilter('0', ['West'])
  await pivot.setLabelSort('0', { type: 'descending' })
  await pivot.setFieldSetting('2', { displayName: 'Total Sales', subtotalType: 'sum', format: '$#,##0' })
  assert.equal(pivot.getData()[2][0], 'West')
  assert.equal(pivot.getData()[2][1], 80)
  assert.equal(pivot.getData()[2][2], 120)
  const detail = pivot.drillDown(2, 1)
  assert.deepEqual(detail.headers, ['Region', 'Quarter', 'Sales'])
  assert.deepEqual(detail.rows, [['West', 'Q1', 80]])
  const beforeRevision = studio.getWorkbookModel(workbook.getId()).pivotTables[0].revision
  await pivot.refresh()
  assert.ok(studio.getWorkbookModel(workbook.getId()).pivotTables[0].revision > beforeRevision)
  assert.equal(officeApi.getPivotTableByCell(workbook.getId(), target.getSheetId(), 2, 1)?.getPivotTableId(), pivot.getPivotTableId())
  studio.dispose()
})

test('linked pivot charts follow pivot fields and filtering, and become available again after same-id restoration', async () => {
  const { studio, workbook, source, target } = setup()
  const pivot = await addBasicPivot(workbook, source, target)
  const chart = await pivot.insertChart(pivot.newChart('line').setPosition('F2').setSize(480, 300).setTitle('Sales by region').build())
  assert.equal(workbook.save().charts[0].pivotTableId, pivot.getPivotTableId())
  assert.deepEqual(chart.getData().categories, ['East', 'West'])
  assert.deepEqual(chart.getData().series.map((series) => series.values), [[100, 80], [150, 120]])
  assert.deepEqual(chart.getData().series.map((series) => series.name), ['Q1 · Sales', 'Q2 · Sales'])
  await chart.update({ type: 'bar', title: 'Filtered sales', position: { x: 720, y: 48 }, legend: { visible: false, position: 'top' } })
  assert.equal(chart.getInfo().type, 'bar')
  assert.equal(chart.getInfo().title, 'Filtered sales')
  assert.deepEqual(chart.getInfo().position, { x: 720, y: 48 })
  assert.deepEqual(chart.getInfo().legend, { visible: false, position: 'top' })
  await chart.update({ type: 'line', title: 'Sales by region', position: { x: 480, y: 24 }, legend: { visible: true, position: 'bottom' } })
  await pivot.setLabelManualFilter('0', ['West'])
  assert.deepEqual(chart.getData().categories, ['West'])
  assert.deepEqual(chart.getData().series.map((series) => series.values), [[80], [120]])
  source.getRange('C5').setValue(130)
  await pivot.refresh()
  assert.deepEqual(chart.getData().series.map((series) => series.values), [[80], [130]])
  await pivot.updateValuePosition('row')
  assert.equal(pivot.getData()[1][0], 'West')
  assert.deepEqual(chart.getData().categories, ['West'])
  assert.deepEqual(chart.getData().series.map((series) => series.values), [[80], [130]])
  await pivot.updateValuePosition('column')

  const movedTarget = workbook.insertSheet('Moved pivot')
  await source.setName('Renamed source')
  const model = studio.getWorkbookModel(workbook.getId())
  const draft = pivotPanelState.createPivotPanelDraft(model, source.getSheetId(), 'A1:C5', pivot.getPivotTableId())
  draft.targetSheetId = movedTarget.getSheetId()
  draft.targetAddress = 'H6'
  studio.executeCommandSync(pivotPanelState.CONFIGURE_PIVOT_COMMAND, pivotPanelState.preparePivotPanelInput(model, draft, pivot.getPivotTableId()))
  assert.deepEqual(chart.getData().categories, ['West'])
  assert.deepEqual(chart.getData().series.map((series) => series.values), [[80], [130]])

  const current = studio.getWorkbookModel(workbook.getId())
  const view = toExcelWorkbookView(current)
  const exported = applyExcelWorkbookView(current, view)
  assert.equal(exported.pivotTables[0].id, pivot.getPivotTableId())
  assert.equal(exported.pivotTables[0].anchor.subUnitId, movedTarget.getSheetId())
  assert.deepEqual(chartEngine.readChartData(exported, exported.charts.find((item) => item.id === chart.getChartId())).categories, ['West'])
  const chartId = chart.getChartId()
  await pivot.remove()
  assert.throws(() => chart.getData(), /Pivot table not found/)
  workbook.undo()
  assert.deepEqual(target.getChartById(chartId).getData().categories, ['West'])
  studio.dispose()
})

test('rejects pivot output conflicts and records pivot operations in workbook history', async () => {
  const { studio, workbook, source, target } = setup()
  target.getRange('A1').setValue('occupied')
  await assert.rejects(() => addBasicPivot(workbook, source, target), /overlaps existing content/)
  target.getRange('A1').clear()
  const pivot = await addBasicPivot(workbook, source, target)
  assert.ok(pivot)
  await pivot.removeField('0')
  assert.equal(pivot.getFieldIdsByArea('row').length, 0)
  workbook.undo()
  assert.equal(workbook.getPivotTables()[0].getFieldIdsByArea('row').length, 1)
  workbook.redo()
  assert.equal(workbook.getPivotTables()[0].getFieldIdsByArea('row').length, 0)
  studio.dispose()
})

