import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { waterfallLayoutFixture } from './fixtures/chart-waterfall-fixtures.mjs'

// P18 integration drafts. The production Facade/service/decoder consume fixed
// transport tuples; no Rust/WASM, server, tests or browser have been started.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })

function transport() {
  const revisions = new Map(), expected = [], calls = []
  return { calls, expect(name) { expected.push(name) },
    setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    submit(source, options) {
      const name = expected.shift(); assert.ok(name, 'enqueue a fixed result before computing a new revision')
      const fixture = waterfallLayoutFixture(name), task = structuredClone(source,
        { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] }), job = task.params.jobs[0]
      assert.equal(task.kind, 'chartStatistics'); assert.equal(task.params.rowCount, 2)
      assert.equal(job.kind, 'waterfallLayout'); assert.equal(job.stackType, fixture.stackType); assert.equal(job.useSubtotal, fixture.useSubtotal)
      assert.deepEqual(task.params.series.map(series => series.field), [1, 2])
      const values = new Float64Array([fixture.cumulative, ...fixture.points.flatMap(point => [point.start, point.end, point.value]),
        ...fixture.connectors.flatMap(link => [link.from.value, link.to.value])])
      const indices = new Uint32Array([fixture.points.length, fixture.connectors.length,
        ...fixture.points.flatMap(point => [point.seriesPosition, point.dataPointIndex, point.categoryIndex, point.synthetic ? 3 : point.role === 'total' ? 1 : 0]),
        ...fixture.connectors.flatMap(link => [link.from.categoryIndex, link.to.categoryIndex])])
      const anchor = task.params.series[job.seriesIndex], requestId = `waterfall-history-${calls.length}`
      calls.push({ name, task, options })
      return { requestId, cancel() {}, promise: Promise.resolve({ output: { kind: 'chartStatistics', values, indices,
        segments: [{ kind: 'waterfallLayout', jobIndex: 0, field: anchor.field, seriesPosition: anchor.seriesPosition, slot: job.slot,
          valueOffset: 0, valueLength: values.length, indexOffset: 0, indexLength: indices.length, validSampleCount: 4, available: true }] } }) }
    },
  }
}
function gradient() { return api.chartLinearGradient({ start: { x: 0, y: 0 }, end: { x: 1, y: 1 },
  stops: [{ offset: 0, color: 'red', opacity: 0.4 }, { offset: 1, color: 'blue' }] }) }
function base(studio) {
  const workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C3').setValues([['Label', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]])
  return { workbook, sheet }
}
async function setup() {
  const compute = transport(), studio = new api.OfficeStudio({ computeClient: compute, currentUserId: 'owner' }), { workbook, sheet } = base(studio)
  const chart = await sheet.insertChart(sheet.newChart('waterfall').setSource('A1:C3').setCategoryField(0).setValueFields([1, 2])
    .setPalette(['#112233', '#223344', '#334455', '#445566']).setTitle('Flow')
    .setOptions({ waterfall: { pointRoles: [], connector: false } }).build())
  return { studio, compute, workbook, sheet, chart }
}
async function read(compute, chart, name) {
  compute.expect(name)
  const [layout, commands] = await Promise.all([chart.getComputedLayout(), chart.getComputedRenderCommands()])
  return { layout, commands }
}
const json = value => JSON.parse(JSON.stringify(value))

test('waterfall Builder role and layout operations chain, detach patches and restore independent authorship', () => {
  const studio = new api.OfficeStudio(), { sheet } = base(studio)
  try {
    const color = gradient(), builder = sheet.newChart('waterfall').setSource('A1:C3').setValueFields([1, 2])
    assert.equal(builder.setPositiveStyle(1, { color, border: { color: 'black', width: 2 }, fillOpacity: 0.6 }), builder)
    assert.equal(builder.setNegativeStyle({ name: 'Second' }, { color: 'red', fillOpacity: 0.4 }), builder)
    assert.equal(builder.setSubtotalStyle(2, { color: 'purple', dataPoints: { 2: { color: 'pink' } } }), builder)
    assert.equal(builder.setConnector({ color: 'black', width: 3, dashType: 'dotted' }), builder)
    assert.equal(builder.setStackType('arrangement'), builder); assert.equal(builder.setUseSubtotal(false), builder)
    const authored = builder.build()
    color.stops[0].color = 'lime'
    assert.deepEqual(builder.build(), authored)
    builder.setPositiveStyle(1, { border: { opacity: 0.2 } })
    assert.deepEqual(builder.build().options.series[1].waterfallStyles.positive.border, { color: 'black', width: 2, opacity: 0.2 })
    builder.setPositiveStyle(1, { color: null })
    assert.equal(builder.build().options.series[1].waterfallStyles.positive.color, undefined)
    assert.equal(builder.build().options.series[1].waterfallStyles.positive.border.width, 2)
    const description = builder.describe({ fields: { 0: 'Label', 1: 'First', 2: 'Second' } })
    assert.equal(description.series[2].waterfallStyles.subtotal.dataPoints[2].color, 'pink')
    description.series[2].waterfallStyles.subtotal.dataPoints[2].color = 'blue'
    assert.equal(builder.build().options.series[2].waterfallStyles.subtotal.dataPoints[2].color, 'pink')
    builder.clearPositiveStyle(1).clearNegativeStyle({ name: 'Second' }).clearSubtotalStyle(2).clearConnector().resetStackType().resetUseSubtotal()
    const reset = builder.build()
    assert.equal(reset.options.series[1].waterfallStyles.positive, undefined)
    assert.equal(reset.options.namedSeries[0].style.waterfallStyles.negative, undefined)
    assert.equal(reset.options.series[2].waterfallStyles.subtotal, undefined)
    assert.equal(reset.options.waterfall.connector, undefined); assert.equal(reset.options.waterfall.stackType, undefined); assert.equal(reset.options.waterfall.useSubtotal, undefined)
    builder.setOptions({ waterfall: { pointRoles: [{ seriesIndex: 2, dataPointIndex: 0, role: 'change' }, { seriesIndex: 2, dataPointIndex: 0, role: 'total' }] } })
    assert.equal(builder.build().options.waterfall.pointRoles.length, 2)
    builder.setStackType('stacked')
    assert.equal(builder.build().options.waterfall.pointRoles, undefined)
    const before = builder.build()
    for (const operation of [() => builder.setStackType('grouped'), () => builder.setUseSubtotal('true'),
      () => builder.setConnector({ width: -1 }), () => builder.setPositiveStyle(1, { waterfallStyles: {} }),
      () => builder.setOptions({ waterfall: { pointRoles: [{ seriesIndex: 2, dataPointIndex: 0, role: 'total' }] } })]) assert.throws(operation)
    assert.deepEqual(builder.build(), before)
    assert.throws(() => sheet.newChart('line').setPositiveStyle(1, { color: 'red' }), TypeError)
  } finally { studio.dispose() }
})

test('live waterfall styles, duplicate roles, stack and subtotal resets share history and internal reopen', async () => {
  const { studio, compute, workbook, chart } = await setup()
  let reopened
  try {
    const raw = chart.getData(), color = gradient()
    await chart.setPositiveStyle(1, { color, name: 'Revenue', label: { visible: true, contentType: 'series' } })
    await chart.setNegativeStyle({ name: 'Second' }, { color: 'blue', border: { color: 'black', width: 3 } })
    await chart.setSubtotalStyle(2, { color: 'purple', name: 'Second final', dataPoints: { 2: { color: 'pink' } } })
    await chart.setConnector({ color: 'black', width: 2, dashType: 'solid' })
    const pointRoles = [{ seriesIndex: 2, dataPointIndex: 0, role: 'change' }, { seriesIndex: 2, dataPointIndex: 0, role: 'total' }]
    await chart.setOptions({ waterfall: { pointRoles } })
    pointRoles[1].role = 'change'; color.stops[0].color = 'lime'
    assert.equal(chart.getOptions().waterfall.pointRoles[1].role, 'total')
    assert.equal(chart.getOptions().series[1].waterfallStyles.positive.color.stops[0].color, 'red')
    const initial = await read(compute, chart, 'arrangement-total')
    assert.deepEqual(initial.layout.waterfall.points.filter(point => point.synthetic).map(point => point.value), [2, 2])
    assert.equal(initial.layout.series[1].points[0].baseValue, 0)
    assert.deepEqual(chart.getLayout(), initial.layout); assert.deepEqual(chart.getRenderCommands(), initial.commands)
    const authored = chart.getOptions()
    await chart.clearPositiveStyle(1)
    assert.equal(chart.getOptions().series[1].waterfallStyles.positive, undefined)
    workbook.undo(); assert.deepEqual(chart.getOptions(), authored)
    workbook.redo(); assert.equal(chart.getOptions().series[1].waterfallStyles.positive, undefined)
    workbook.undo()
    await chart.setStackType('stacked')
    assert.equal(chart.getOptions().waterfall.pointRoles, undefined)
    const stacked = await read(compute, chart, 'stacked')
    assert.deepEqual(stacked.layout.series.map(series => series.points.length), [2, 2])
    assert.equal(stacked.layout.waterfall.points.at(-1).seriesPosition, 0xffff_ffff)
    workbook.undo()
    assert.deepEqual(chart.getOptions().waterfall.pointRoles, authored.waterfall.pointRoles)
    assert.deepEqual((await read(compute, chart, 'arrangement-total')).layout, initial.layout)
    workbook.redo()
    assert.deepEqual((await read(compute, chart, 'stacked')).layout, stacked.layout)
    await chart.setUseSubtotal(false)
    assert.equal((await read(compute, chart, 'stacked-no-subtotal')).layout.waterfall.points.length, 4)
    await chart.resetUseSubtotal(); await chart.resetStackType(); await chart.clearConnector()
    const final = await read(compute, chart, 'arrangement')
    assert.equal(chart.getOptions().waterfall.stackType, undefined); assert.equal(chart.getOptions().waterfall.useSubtotal, undefined)
    assert.equal(chart.getOptions().waterfall.connector, undefined)
    assert.deepEqual(final.layout.waterfall.points.filter(point => point.synthetic).map(point => point.value), [9, 9])
    assert.deepEqual(chart.getData(), raw)
    const options = chart.getOptions(), description = chart.getDescription(), snapshot = json(workbook.save()), serialized = JSON.stringify(snapshot.charts)
    for (const transient of ['waterfallLayout', 'synthetic', 'seriesPosition', 'connectors', 'cumulative', 'lease']) assert.equal(serialized.includes(`"${transient}"`), false)
    const restoredCompute = transport(); reopened = new api.OfficeStudio({ workbook: snapshot, computeClient: restoredCompute, currentUserId: 'owner' })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(chart.getChartId())
    assert.deepEqual(json(restored.getOptions()), json(options)); assert.deepEqual(json(restored.getDescription()), json(description))
    assert.deepEqual(restored.getData(), raw)
    assert.deepEqual(await read(restoredCompute, restored, 'arrangement'), final)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('invalid and wrong-type waterfall live operations publish no content or history', async () => {
  const { studio, workbook, chart } = await setup()
  try {
    const contents = JSON.stringify(workbook.save()), history = studio.editor.history.captureSavepoint(workbook.getId())
    for (const operation of [() => chart.setPositiveStyle(1, { fillOpacity: 2 }), () => chart.setNegativeStyle(1, { color: 'not-a-color' }),
      () => chart.setSubtotalStyle(1, { dataPoints: { '-1': { color: 'red' } } }), () => chart.setConnector({ opacity: 2 }),
      () => chart.setStackType('grouped'), () => chart.setUseSubtotal(1),
      () => chart.setOptions({ waterfall: { stackType: 'stacked', pointRoles: [{ seriesIndex: 1, dataPointIndex: 0, role: 'total' }] } })]) {
      await assert.rejects(async () => { await operation() })
      assert.equal(JSON.stringify(workbook.save()), contents)
      assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), history)
    }
    await chart.update({ type: 'column' })
    const inactive = JSON.stringify(workbook.save()), afterType = studio.editor.history.captureSavepoint(workbook.getId())
    for (const operation of [() => chart.setPositiveStyle(1, { color: 'red' }), () => chart.clearPositiveStyle(1),
      () => chart.setNegativeStyle(1, { color: 'red' }), () => chart.clearNegativeStyle(1),
      () => chart.setSubtotalStyle(1, { color: 'red' }), () => chart.clearSubtotalStyle(1), () => chart.setConnector(true),
      () => chart.clearConnector(), () => chart.setStackType('stacked'), () => chart.resetStackType(),
      () => chart.setUseSubtotal(false), () => chart.resetUseSubtotal()]) await assert.rejects(async () => { await operation() }, TypeError)
    assert.equal(JSON.stringify(workbook.save()), inactive)
    assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), afterType)
  } finally { studio.dispose() }
})

test('worksheet edit protection blocks every waterfall operation and direct chart commands without history', async () => {
  const { studio, workbook, sheet, chart } = await setup()
  let visitor
  try {
    studio.executeCommandSync('sheet.command.add-worksheet-protection', { unitId: workbook.getId(), sheetId: sheet.getSheetId(),
      rule: { id: 'chart-protection', permissionId: 'chart-permission', unitType: 'worksheet', ownerUserId: 'owner',
        editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [] } })
    const compute = transport(); visitor = new api.OfficeStudio({ workbook: json(workbook.save()), currentUserId: 'visitor', computeClient: compute })
    const targetWorkbook = api.FOffice.newAPI(visitor).getActiveWorkbook(), targetSheet = targetWorkbook.getActiveSheet(), target = targetSheet.getChartById(chart.getChartId())
    assert.equal(target.getType(), 'waterfall')
    const contents = JSON.stringify(targetWorkbook.save()), history = visitor.editor.history.captureSavepoint(targetWorkbook.getId())
    for (const operation of [() => target.setPositiveStyle(1, { color: 'red' }), () => target.clearPositiveStyle(1),
      () => target.setNegativeStyle(1, { color: 'red' }), () => target.clearNegativeStyle(1),
      () => target.setSubtotalStyle(1, { color: 'red' }), () => target.clearSubtotalStyle(1), () => target.setConnector(true),
      () => target.clearConnector(), () => target.setStackType('stacked'), () => target.resetStackType(),
      () => target.setUseSubtotal(false), () => target.resetUseSubtotal(),
      () => visitor.executeCommand('sheet.command.update-chart', { unitId: targetWorkbook.getId(), chartId: chart.getChartId(), patch: { title: 'Blocked direct edit' } }),
      () => target.remove(), () => targetSheet.insertChart(targetSheet.newChart('waterfall').setSource('A1:C3').setValueFields([1, 2]).build())]) {
      await assert.rejects(async () => { await operation() }, /编辑权限/)
      assert.equal(JSON.stringify(targetWorkbook.save()), contents)
      assert.deepEqual(visitor.editor.history.captureSavepoint(targetWorkbook.getId()), history)
    }
    assert.equal(compute.calls.length, 0)
  } finally { visitor?.dispose(); studio.dispose() }
})
