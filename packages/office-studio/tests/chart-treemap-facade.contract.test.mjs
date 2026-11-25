import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { TREEMAP_TABLE } from './fixtures/chart-treemap-fixtures.mjs'
import { chartHierarchyGoldenTransport as transport } from './fixtures/chart-hierarchy-facade-transport.mjs'

// P26 Facade/service drafts with independent materialized transport tuples.
// No server, test, compiler, Worker, WASM or browser has been started.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function base(studio) {
  const workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C8').setValues(TREEMAP_TABLE)
  return { workbook, sheet }
}
function newBuilder(sheet) { return sheet.newChart('treemap').setSource('A1:C8').setHierarchyFields([0, 1]).setValueField(2).setAbsolutePosition(20, 30).setSize(560, 400) }

test('all six Treemap Builder methods chain, detach inputs and reset each authored group independently', () => {
  const studio = new api.OfficeStudio(), { sheet } = base(studio)
  try {
    const fields = [0, 1], background = { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }
    const builder = newBuilder(sheet)
    assert.equal(builder.setHierarchyFields(fields), builder); assert.equal(builder.setValueField(2), builder)
    assert.equal(builder.setParentLabelLayout('overlapping'), builder)
    assert.equal(builder.setTreemapLabel({ visible: true, contentType: 10, fontSize: 18, backgroundColor: background }), builder)
    const before = builder.build(); fields[0] = 2; background.stops[0].color = 'lime'
    assert.deepEqual(builder.build(), before)
    assert.equal(builder.resetParentLabelLayout(), builder); assert.equal(builder.build().options.treemap.parentLabelLayout, undefined)
    assert.equal(builder.build().options.treemap.label.fontSize, 18)
    assert.equal(builder.resetTreemapLabel(), builder); assert.equal(builder.build().options.treemap.label, undefined)
    assert.deepEqual(builder.build().options.mapping.hierarchyIndexes, [0, 1]); assert.deepEqual(builder.build().valueFields, [2])
    const current = builder.build()
    for (const operation of [() => builder.setHierarchyFields([]), () => builder.setHierarchyFields([0, 0]), () => builder.setHierarchyFields(Array(2)),
      () => builder.setValueField(-1), () => builder.setParentLabelLayout('floating'), () => builder.setTreemapLabel({ position: 'left' })]) assert.throws(operation)
    assert.deepEqual(builder.build(), current)
    assert.throws(() => sheet.newChart('line').setHierarchyFields([0, 1]), TypeError)
    assert.throws(() => sheet.newChart('line').setValueField(2), TypeError)
  } finally { studio.dispose() }
})

test('live labels, background union swaps and resets use real commands, detached snapshots, per-document history and internal reopen', async () => {
  const compute = transport(), studio = new api.OfficeStudio({ computeClient: compute }), { workbook, sheet } = base(studio)
  let reopened
  try {
    const chart = await sheet.insertChart(newBuilder(sheet).setParentLabelLayout('overlapping').setTreemapLabel({ visible: true, contentType: 10, fontSize: 12 }).build())
    const raw = await chart.getComputedData(), layout = await chart.getComputedLayout(), originalOptions = chart.getOptions()
    assert.deepEqual(raw.hierarchy.map(node => [node.name, node.value, node.selfValue, node.paletteIndex]), [['North', 13, -5, 1], ['South', 9, 3, 0]])
    assert.deepEqual(layout.treemap.filter(rect => rect.depth === 0).map(rect => rect.value), [18, 9])
    assert.equal(layout.treemap.filter(rect => rect.renderNode.isTreemapResidual).length, 1)
    assert.equal((await chart.setHierarchyFields([0, 1])), chart); assert.equal((await chart.setValueField(2)), chart)
    const color = { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }
    assert.equal((await chart.setTreemapLabel({ backgroundColor: color })), chart)
    const pattern = { type: 'pattern', preset: 'pct25', foregroundColor: 'black', backgroundColor: 'white' }
    await chart.setTreemapLabel({ backgroundColor: pattern, fontSize: 18 }); pattern.foregroundColor = 'red'
    const authored = chart.getOptions()
    assert.equal(authored.treemap.label.backgroundColor.foregroundColor, 'black'); assert.equal(Object.hasOwn(authored.treemap.label.backgroundColor, 'stops'), false)
    await chart.resetParentLabelLayout(); assert.equal(chart.getOptions().treemap.parentLabelLayout, undefined)
    assert.equal(chart.getOptions().treemap.label.fontSize, 18)
    const beforeReset = chart.getOptions(); await chart.resetTreemapLabel(); assert.equal(chart.getOptions().treemap.label, undefined)
    workbook.undo(); assert.deepEqual(chart.getOptions(), beforeReset)
    workbook.redo(); assert.equal(chart.getOptions().treemap.label, undefined)
    await chart.setTreemapLabel({ visible: false, contentType: 10 }); await chart.setParentLabelLayout('banner')
    const saved = workbook.save(), detachedData = await chart.getComputedData()
    detachedData.hierarchy[0].selfValue = 100; assert.equal((await chart.getComputedData()).hierarchy[0].selfValue, -5)
    reopened = api.FOffice.newAPI(studio).createWorkbook({ ...saved, id: 'treemap-reopened-draft' }, { makeCurrent: false })
    const copy = reopened.getActiveSheet().getCharts()[0]
    assert.deepEqual(copy.getOptions(), chart.getOptions()); assert.deepEqual(json(await copy.getComputedData()), json(raw))
    assert.ok((await copy.getComputedRenderCommands()).some(command => command.kind === 'text' && command.text === 'North'))
    assert.ok(compute.calls.length > 0)
    assert.ok(compute.calls.some(call => call.options.documentId !== compute.calls[0].options.documentId))
    assert.equal(originalOptions.treemap.label.fontSize, 12)
  } finally { reopened?.dispose(); studio.dispose() }
})
