import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Drafts only; server/test execution belongs to the later unified phase.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

function setup() {
  const studio = new api.OfficeStudio(), facade = api.FOffice.newAPI(studio), workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C3').setValues([['Region', 'Revenue', 'Costs'], ['North', 7, 3], ['South', 4, 2]])
  return { studio, facade, workbook, sheet, builder: sheet.newChart('line').setSource('A1:C3').setCategoryField(0).setValueFields([1, 2]) }
}

test('fixed enum names, detached boolean/object/Hide and clear/reset return the same builder and retain no caller font alias', () => {
  const { studio, facade, builder } = setup()
  try {
    assert.equal(facade.Enum.ChartSelectModeEnum, api.ChartLegendSelectModes)
    assert.equal(facade.Enum.ChartLegendPositionEnum, api.ChartLegendPositions)
    assert.deepEqual(facade.Enum.ChartSelectModeEnum, { Single: 'single', Multiple: 'multiple', Close: 'close' })
    const font = { family: 'Georgia' }
    assert.equal(builder.setLegend({ position: facade.Enum.ChartLegendPositionEnum.Top, selectMode: facade.Enum.ChartSelectModeEnum.Multiple, wrap: true, font }), builder)
    font.family = 'Arial'; assert.equal(builder.build().legend.font.family, 'Georgia')
    builder.setLegend({ color: '#123456', fontSize: 15, bold: true, italic: false, family: 'Cambria' })
    assert.deepEqual(builder.build().legend.font, { color: '#123456', fontSize: 15, bold: true, italic: false, family: 'Cambria' })
    assert.equal(Object.hasOwn(builder.build().legend, 'fontSize'), false)
    builder.setLegend(false); assert.equal(builder.build().legend.visible, false)
    builder.setLegend(true); assert.equal(builder.build().legend.visible, true)
    builder.setLegend({ position: facade.Enum.ChartLegendPositionEnum.Hide }); assert.equal(builder.build().legend.position, 'top'); assert.equal(builder.build().legend.visible, false)
    const before = builder.build()
    assert.throws(() => builder.setLegend({ selectMode: 'toggle' }), TypeError)
    assert.deepEqual(builder.build(), before)
    assert.equal(builder.clearLegend(), builder); assert.equal(builder.build().legend, undefined)
  } finally { studio.dispose() }
})

test('live configuration uses document content history and survives internal snapshot reopening; selected/page view state is absent', async () => {
  const { studio, workbook, sheet, builder } = setup()
  let reopened
  try {
    const chart = await sheet.insertChart(builder.build()), initial = chart.getInfo()
    assert.equal(await chart.setLegend({ selectMode: 'single', wrap: true, position: 'top', fontSize: 14, family: 'Cambria' }), chart)
    const authored = chart.getInfo()
    assert.deepEqual(authored.legend.font, { fontSize: 14, family: 'Cambria' })
    assert.equal(Object.hasOwn(authored.legend, 'fontSize'), false)
    workbook.undo(); assert.deepEqual(chart.getInfo(), initial)
    workbook.redo(); assert.deepEqual(chart.getInfo(), authored)
    const saved = JSON.parse(JSON.stringify(workbook.save()))
    assert.equal(JSON.stringify(saved.charts).includes('legendSelection'), false)
    assert.equal(JSON.stringify(saved.charts).includes('scrollDataIndex'), false)
    reopened = new api.OfficeStudio({ workbook: saved })
    const again = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChart(chart.getChartId())
    assert.deepEqual(again.getInfo().legend, authored.legend)
    await chart.clearLegend()
    assert.deepEqual(chart.getInfo().legend, { visible: true, position: 'bottom' })
    workbook.undo(); assert.deepEqual(chart.getInfo().legend, authored.legend)
    workbook.redo(); assert.deepEqual(chart.getInfo().legend, { visible: true, position: 'bottom' })
  } finally { reopened?.dispose(); studio.dispose() }
})
