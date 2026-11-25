import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { chartResourceFixture, chartResourceGoldenTransport, withChartResourcePlatform } from './fixtures/chart-resource-platform-fixture.mjs'

// Frontend host contract drafts. No server, compiler, runner, browser or WASM ran.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const registry = workbook => workbook.metadata?.chartImageResources ?? {}
function setup(table = [['Word', 'Weight'], ['A', 1]]) {
  const compute = chartResourceGoldenTransport(), studio = new api.OfficeStudio({ computeClient: compute })
  const workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange(`A1:${table[0].length === 4 ? 'D' : 'B'}${table.length}`).setValues(table)
  return { compute, studio, workbook, sheet }
}
function wordBuilder(sheet) { return sheet.newChart('wordCloud').setSource('A1:B2').setSize(560, 400).setValueFields([1]) }

test('WordCloud Builder detaches resources and clears pending payload on ID replacement/reset', () => {
  const { studio, sheet } = setup()
  try {
    const value = wordBuilder(sheet), image = chartResourceFixture()
    assert.equal(value.setMaskImageResource(image), value)
    const detached = value.build(); image.source = './changed.png'
    assert.equal(detached.options.wordCloud.maskImageResource.source, chartResourceFixture().source)
    assert.equal(value.build().options.wordCloud.maskImageResource.source, chartResourceFixture().source)
    assert.equal(value.setShape('heart').setRepeat(true).resetShape().resetRepeat(), value)
    assert.equal(value.setMaskImage('Circle'), value); assert.equal(value.build().options.wordCloud.maskImageResource, undefined)
    value.setMaskImageResource(chartResourceFixture()); value.clearMaskImage()
    assert.equal(value.build().options.wordCloud.maskImageId, undefined)
    assert.equal(value.build().options.wordCloud.maskImageResource, undefined)
    assert.throws(() => sheet.newChart('line').setMaskImageResource(chartResourceFixture()), /词云/)
  } finally { studio.dispose() }
})

test('WordCloud live resource Apply, clear, undo/redo and JSON reopen publish reference and bytes together', async () => {
  await withChartResourcePlatform(async () => {
    const { studio, workbook, sheet, compute } = setup(); let reopened
    try {
      const chart = await sheet.insertChart(wordBuilder(sheet).build())
      const before = json(workbook.save()), beforeHistory = studio.editor.history.captureSavepoint(workbook.getId())
      assert.equal(await chart.setMaskImageResource(chartResourceFixture()), chart)
      const applied = json(workbook.save()), appliedHistory = studio.editor.history.captureSavepoint(workbook.getId())
      assert.equal(chart.getOptions().wordCloud.maskImageId, chartResourceFixture().id)
      assert.equal(chart.getOptions().wordCloud.maskImageResource, undefined)
      assert.deepEqual(registry(applied)[chartResourceFixture().id], chartResourceFixture())
      workbook.undo(); assert.deepEqual(json(workbook.save()).charts, before.charts)
      assert.deepEqual(registry(workbook.save()), registry(before))
      assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), beforeHistory)
      workbook.redo(); assert.deepEqual(json(workbook.save()).charts, applied.charts)
      assert.deepEqual(registry(workbook.save()), registry(applied))
      assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), appliedHistory)
      assert.equal((await chart.getComputedData()).wordCloud.words[0].text, 'A')
      reopened = api.FOffice.newAPI(studio).createWorkbook({ ...applied, id: 'cloud-resource-reopened-draft' }, { makeCurrent: false })
      const restored = reopened.getActiveSheet().getCharts()[0]
      assert.deepEqual(restored.getOptions(), chart.getOptions())
      assert.equal((await restored.getComputedData()).wordCloud.words[0].text, 'A')
      await chart.clearMaskImage(); assert.equal(registry(workbook.save())[chartResourceFixture().id], undefined)
      workbook.undo(); assert.deepEqual(registry(workbook.save())[chartResourceFixture().id], chartResourceFixture())
      assert.ok(compute.calls.length > 0)
    } finally { reopened?.dispose(); studio.dispose() }
  })
})

test('WordCloud conflicting or missing resource ID leaves the chart and history unchanged', async () => {
  await withChartResourcePlatform(async () => {
    const { studio, workbook, sheet } = setup()
    try {
      const chart = await sheet.insertChart(wordBuilder(sheet).setMaskImageResource(chartResourceFixture()).build())
      const before = json(workbook.save()), history = studio.editor.history.captureSavepoint(workbook.getId())
      await assert.rejects(chart.setMaskImageResource({ ...chartResourceFixture(), source: 'data:image/png;base64,AAAAAAAAAAA=' }), /ID 已被/)
      await assert.rejects(chart.setMaskImage('missing-resource'), error => error.code === 'COMPUTE_INVALID_INPUT')
      assert.deepEqual(json(workbook.save()).charts, before.charts)
      assert.deepEqual(registry(workbook.save()), registry(before))
      assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), history)
    } finally { studio.dispose() }
  })
})

test('shared WordCloud resources survive removal until the final authored reference and return with undo', async () => {
  await withChartResourcePlatform(async () => {
    const { studio, workbook, sheet } = setup()
    try {
      const first = await sheet.insertChart(wordBuilder(sheet).setMaskImageResource(chartResourceFixture()).build())
      const second = await sheet.insertChart(wordBuilder(sheet).setMaskImage(chartResourceFixture().id).build())
      await first.remove(); assert.deepEqual(registry(workbook.save())[chartResourceFixture().id], chartResourceFixture())
      await second.remove(); assert.equal(registry(workbook.save())[chartResourceFixture().id], undefined)
      workbook.undo(); assert.deepEqual(registry(workbook.save())[chartResourceFixture().id], chartResourceFixture())
      assert.equal(workbook.getActiveSheet().getCharts()[0].getOptions().wordCloud.maskImageId, chartResourceFixture().id)
    } finally { studio.dispose() }
  })
})

test('a source revision during pending WordCloud fetch cannot publish an obsolete candidate', async () => {
  let resolveFetch, started, delay = false
  const awaitingFetch = new Promise(resolve => { started = resolve })
  await withChartResourcePlatform(async () => {
    const { studio, workbook, sheet } = setup()
    try {
      const chart = await sheet.insertChart(wordBuilder(sheet).build()), before = json(chart.getInfo())
      delay = true
      const editing = chart.setMaskImageResource({ ...chartResourceFixture('chart-image:late'), sourceType: 'url', source: './late.png' })
      const rejection = assert.rejects(editing, error => ['COMPUTE_STALE_RESULT', 'COMPUTE_CANCELLED'].includes(error.code))
      await awaitingFetch; sheet.getRange('B2').setValue(2)
      resolveFetch(new Response(new Uint8Array(8), { headers: { 'content-type': 'image/png' } }))
      await rejection
      assert.deepEqual(json(chart.getInfo()), before)
      assert.equal(registry(workbook.save())['chart-image:late'], undefined)
    } finally { studio.dispose() }
  }, () => delay ? new Promise(resolve => { resolveFetch = resolve; started() })
    : Promise.resolve(new Response(new Uint8Array(8), { headers: { 'content-type': 'image/png' } })))
})

test('Relation automatic matrix projection shares real live commands, history and internal reopen', async () => {
  const { studio, workbook, sheet, compute } = setup([['Name', 'Group', ' A ', 'B'], [' A ', 'g1', 0, 2], ['B', 'g2', 3, 0]])
  let reopened
  try {
    const chart = await sheet.insertChart(sheet.newChart('relation').setSource('A1:D3').setSize(560, 400).setUseValueAsSymbolSize(false).build())
    const data = await chart.getComputedData()
    assert.equal(data.graph.sourceKind, 'coWordMatrix'); assert.equal(data.graph.links.length, 4)
    assert.deepEqual(data.graph.nodes.map(node => [node.name, node.value]), [[' A ', 2], ['B', 3]])
    const before = json(chart.getInfo()), history = studio.editor.history.captureSavepoint(workbook.getId())
    await chart.update({ options: { relation: { layout: 'force', force: { gravity: 0.2, repulsion: 7000 }, emphasisEnabled: false, nodeShape: 'diamond' } } })
    const after = json(chart.getInfo())
    assert.ok(compute.calls.some(call => call.task.params.jobs[0].layout === 'force' && call.task.params.jobs[0].gravity === 0.2))
    workbook.undo(); assert.deepEqual(json(chart.getInfo()), before)
    assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), history)
    workbook.redo(); assert.deepEqual(json(chart.getInfo()), after)
    const saved = json(workbook.save())
    assert.equal(JSON.stringify(saved.charts).includes('"relation":{"width"'), false)
    reopened = api.FOffice.newAPI(studio).createWorkbook({ ...saved, id: 'relation-matrix-reopened-draft' }, { makeCurrent: false })
    const restored = reopened.getActiveSheet().getCharts()[0]
    assert.equal((await restored.getComputedData()).graph.sourceKind, 'coWordMatrix')
    assert.deepEqual(restored.getOptions(), chart.getOptions())
  } finally { reopened?.dispose(); studio.dispose() }
})
