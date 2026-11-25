import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// Draft only; the later run requires the normal OfficeStudio/WASM platform.
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

test('fixed title/alignment enum names, object/font detachment, string compatibility and independent clears share the builder', () => {
  const { studio, facade, builder } = setup()
  try {
    assert.equal(facade.Enum.ChartTitlePositionEnum, api.ChartTitlePositions)
    assert.equal(facade.Enum.ChartLabelAlignEnum, api.ChartTextAlignments)
    const spec = { text: 'Main', fontSize: 18, family: 'Georgia', position: facade.Enum.ChartTitlePositionEnum.Left, alignment: facade.Enum.ChartLabelAlignEnum.Right }
    assert.equal(builder.setTitle(spec), builder); spec.family = 'Arial'
    builder.setSubtitle({ text: 'Sub', fontSize: 12, alignment: 'left' })
    builder.setTitle('Legacy replacement')
    assert.equal(builder.build().title, 'Legacy replacement'); assert.equal(builder.build().options.titleStyle.family, 'Georgia')
    const before = builder.build()
    assert.throws(() => builder.setTitle({ position: 'middle' }), TypeError); assert.deepEqual(builder.build(), before)
    assert.equal(builder.clearTitle(), builder); assert.equal(builder.build().title, undefined); assert.equal(builder.build().options.titleStyle, undefined)
    assert.equal(builder.build().subtitle, 'Sub'); assert.equal(builder.build().options.subtitleStyle.alignment, 'left')
    assert.equal(builder.clearSubtitle(), builder); assert.equal(builder.build().subtitle, undefined); assert.equal(builder.build().options.subtitleStyle, undefined)
  } finally { studio.dispose() }
})

test('live titles use one content commit, undo/redo and internal JSON reopening preserve authored position/alignment/font and clear semantics', async () => {
  const { studio, workbook, sheet, builder } = setup()
  let reopened
  try {
    const chart = await sheet.insertChart(builder.setTitle('Legacy').setSubtitle('Sub').build()), initial = chart.getInfo()
    await chart.setTitle({ text: 'Moved', position: 'bottom', alignment: 'right', fontSize: 19, color: '#123456' })
    const authored = chart.getInfo()
    assert.equal(authored.title, 'Moved'); assert.equal(authored.options.titleStyle.position, 'bottom')
    workbook.undo(); assert.deepEqual(chart.getInfo(), initial); workbook.redo(); assert.deepEqual(chart.getInfo(), authored)
    const saved = JSON.parse(JSON.stringify(workbook.save()))
    reopened = new api.OfficeStudio({ workbook: saved })
    const again = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChart(chart.getChartId())
    assert.deepEqual(again.getInfo().options.titleStyle, authored.options.titleStyle); assert.equal(again.getInfo().title, 'Moved')
    await chart.clearTitle(); const cleared = chart.getInfo()
    assert.equal(cleared.title, undefined); assert.equal(cleared.options.titleStyle, undefined); assert.equal(cleared.subtitle, 'Sub')
    workbook.undo(); assert.deepEqual(chart.getInfo(), authored); workbook.redo(); assert.deepEqual(chart.getInfo(), cleared)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('invalid live title configuration rejects before content revision/history change and keeps the prior view author', async () => {
  const { studio, workbook, sheet, builder } = setup()
  try {
    const chart = await sheet.insertChart(builder.setTitle('Before').build()), before = chart.getInfo(), revision = studio.getWorkbookRevision(workbook.getId())
    await assert.rejects(chart.setTitle({ visible: 'yes' }), TypeError)
    await assert.rejects(chart.setSubtitle({ fontSize: -1 }), RangeError)
    assert.deepEqual(chart.getInfo(), before); assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
  } finally { studio.dispose() }
})
