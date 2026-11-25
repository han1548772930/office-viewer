import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook } from '../src/domain/model.ts'
import { consumeChartPreparationSync } from '../src/features/charts/chart-compute-lifecycle.ts'
import { chartRelationSourceFormat, chartRelationGraphSteps } from '../src/features/charts/chart-relation-data.ts'
import { projectChartRecords } from '../src/features/charts/chart-record-projection.ts'
import { chartRelationMappingControls } from '../src/features/charts/chart-relation-settings-controls.ts'
import { CHART_IMAGE_MAX_BYTES, CHART_IMAGE_RESOURCE_KEY, normalizeChartImageResource,
  publishChartImageResourceCandidate, readChartImageResources, resolveChartImageResource, pruneChartImageResources,
  readChartImageFile } from '../src/features/charts/chart-image-resources.ts'
import { cloneChartOptions, mergeChartOptions } from '../src/features/charts/chart-options.ts'
import { drawChartWordCloudMaskImage } from '../src/features/charts/chart-word-cloud-mask.ts'
import { chartWordCloudData, chartWordCloudFontSize } from '../src/features/charts/chart-word-cloud.ts'

// Frontend contract drafts only. No runner, compiler, browser, server, WASM or benchmark ran.
const consume = consumeChartPreparationSync
const matrix = () => ({ fields: ['Name', 'Group', ' A ', 'B'], records: [[' A ', 'g1', 0, 2], ['B', 'g2', 3, 0]] })
const binding = (options = {}) => ({ type: 'relation', categoryField: 0, valueFields: [], options })
const resource = (id = 'chart-image:one') => ({ id, sourceType: 'dataUrl', source: 'data:image/png;base64,iVBORw0KGgo=', mimeType: 'image/png', sizeBytes: 8 })
const cloud = (value = resource()) => ({ id: 'cloud-1', type: 'wordCloud', options: { wordCloud: { maskImageId: value.id, maskImageResource: value } } })
async function withGlobals(values, work) {
  const original = new Map(Object.keys(values).map(key => [key, Object.getOwnPropertyDescriptor(globalThis, key)]))
  for (const [key, value] of Object.entries(values)) Object.defineProperty(globalThis, key, { configurable: true, writable: true, value })
  try { return await work() } finally {
    for (const [key, descriptor] of original) if (descriptor) Object.defineProperty(globalThis, key, descriptor); else delete globalThis[key]
  }
}
function canvasOwner() {
  const draws = []
  return { canvas: { width: 4, height: 4 }, context: { drawImage(...args) { draws.push(args) } }, release() {}, draws }
}

test('automatic matrix recognition keeps square label identity and explicit edge/format precedence', () => {
  const source = matrix(), before = structuredClone(source)
  assert.equal(chartRelationSourceFormat(binding(), source), 'coWordMatrix')
  assert.equal(chartRelationSourceFormat(binding({ mapping: { edge: { sourceIndex: 0, targetIndex: 1 } } }), source), 'edgeList')
  assert.equal(chartRelationSourceFormat(binding({ mapping: { relationFormat: 'edgeList' } }), source), 'edgeList')
  const mismatch = matrix(); mismatch.records[0][0] = 'A'
  assert.equal(chartRelationSourceFormat(binding(), mismatch), 'edgeList')
  const short = matrix(); short.records[1].pop()
  assert.equal(chartRelationSourceFormat(binding(), short), 'edgeList')
  const duplicate = matrix(); duplicate.fields[3] = ' A '; duplicate.records[1][0] = ' A '
  assert.equal(chartRelationSourceFormat(binding(), duplicate), 'edgeList')
  const graph = consume(chartRelationGraphSteps(source, binding()))
  assert.deepEqual(graph.nodes.map(node => [node.name, node.category, node.value]), [[' A ', 'g1', 2], ['B', 'g2', 3]])
  assert.deepEqual(graph.links.map(link => [link.source, link.target, link.value]), [[' A ', ' A ', 0], [' A ', 'B', 2], ['B', ' A ', 3], ['B', 'B', 0]])
  assert.equal(graph.sourceKind, 'coWordMatrix'); assert.deepEqual(source, before)
})

test('matrix projection and settings use the same recognition and respect hidden primary columns', () => {
  const source = matrix(), chart = binding()
  const data = projectChartRecords(source, chart)
  assert.equal(data.graph.sourceKind, 'coWordMatrix'); assert.deepEqual(data.categories, [' A ', 'B'])
  const controls = chartRelationMappingControls(4, source), format = controls[0]
  assert.equal(format.read(chart), 'auto')
  assert.ok(controls.filter(control => control.key.startsWith('mapping.edge')).every(control => control.visible(chart) === false))
  format.write(chart, 'edgeList')
  assert.ok(controls.filter(control => control.key.startsWith('mapping.edge')).every(control => control.visible(chart) === true))
  format.write(chart, 'auto'); assert.equal(chart.options.mapping.relationFormat, undefined)
  const hidden = projectChartRecords({ ...source, visibleFields: [0, 1, 2] }, binding())
  assert.deepEqual(hidden.categories, []); assert.deepEqual(hidden.series, [])
})

test('relation edge data preserves zeros, repeated directed links and self incident counting', () => {
  const source = { fields: ['Source', 'Target', 'Weight'], records: [[' A ', 'B', 2], [' A ', 'B', 3], ['B', ' A ', 0], [' A ', ' A ', 4]] }
  const graph = consume(chartRelationGraphSteps(source, binding()))
  assert.equal(graph.valid, true); assert.equal(graph.links.length, 4)
  assert.deepEqual(graph.nodes.map(node => [node.name, node.value]), [[' A ', 9], ['B', 5]])
  const bad = { ...source, records: [...source.records, ['B', ' A ', '2']] }
  const rejected = consume(chartRelationGraphSteps(bad, binding()))
  assert.equal(rejected.valid, false); assert.deepEqual(rejected.nodes, []); assert.match(rejected.diagnostic, /有限非负/)
})

test('resource validation separates locators, persistent data URLs, declared bytes and option text budgets', () => {
  assert.deepEqual(normalizeChartImageResource(resource()), resource())
  for (const patch of [{ sourceType: 'url', source: 'opaque-id' }, { sourceType: 'url', source: 'blob:temporary' },
    { mimeType: 'text/html' }, { sizeBytes: 7 }, { sizeBytes: CHART_IMAGE_MAX_BYTES + 1 }, { source: 'data:image/jpeg;base64,iVBORw0KGgo=' }]) {
    assert.throws(() => normalizeChartImageResource({ ...resource(), ...patch }))
  }
  assert.equal(normalizeChartImageResource({ ...resource(), sourceType: 'url', source: './mask.png' }).source, './mask.png')
  const payload = 'AAAA'.repeat(300_000), large = { ...resource(), source: `data:image/png;base64,${payload}`, sizeBytes: 900_000 }
  const options = cloneChartOptions({ wordCloud: { maskImageId: large.id, maskImageResource: large } })
  assert.equal(options.wordCloud.maskImageResource.sizeBytes, 900_000)
  assert.throws(() => cloneChartOptions({ wordCloud: { maskImageId: 'different', maskImageResource: resource() } }), /不一致/)
  const cleared = mergeChartOptions(options, { wordCloud: { maskImageId: undefined, maskImageResource: undefined } })
  assert.equal(cleared.wordCloud.maskImageResource, undefined); assert.equal(cleared.wordCloud.maskImageId, undefined)
})

test('resource candidate publication, shared references, pruning and JSON reopen retain one authoritative registry', () => {
  const workbook = createBlankWorkbook(), chart = cloud(), second = cloud()
  second.id = 'cloud-2'; workbook.charts = [chart, second]
  const pending = structuredClone(workbook)
  assert.equal(workbook.metadata?.[CHART_IMAGE_RESOURCE_KEY], undefined)
  publishChartImageResourceCandidate(pending, pending.charts[0])
  assert.equal(pending.charts[0].options.wordCloud.maskImageResource, undefined)
  assert.deepEqual(readChartImageResources(pending)[resource().id], resource())
  assert.deepEqual(resolveChartImageResource(pending, pending.charts[0], resource().id), resource())
  const reopened = JSON.parse(JSON.stringify(pending))
  assert.deepEqual(readChartImageResources(reopened), readChartImageResources(pending))
  reopened.charts.splice(0, 1); pruneChartImageResources(reopened)
  assert.ok(readChartImageResources(reopened)[resource().id])
  reopened.charts = []; pruneChartImageResources(reopened)
  assert.equal(reopened.metadata[CHART_IMAGE_RESOURCE_KEY], undefined)
  assert.deepEqual(workbook.charts[0].options.wordCloud.maskImageResource, resource())
})

test('conflicting IDs and malformed registries fail without partial candidate publication', () => {
  const workbook = createBlankWorkbook(), chart = cloud(); workbook.charts = [chart]
  publishChartImageResourceCandidate(workbook, chart)
  const candidate = structuredClone(workbook), conflict = { ...resource(), source: 'data:image/png;base64,AAAAAAAAAAA=' }
  candidate.charts[0].options.wordCloud.maskImageResource = conflict
  const before = structuredClone(candidate)
  assert.throws(() => publishChartImageResourceCandidate(candidate, candidate.charts[0]), /ID 已被/)
  assert.deepEqual(candidate, before)
  assert.throws(() => readChartImageResources({ metadata: { [CHART_IMAGE_RESOURCE_KEY]: { different: resource() } } }), /ID/)
})

test('mask loader resolves resource IDs before fetching and releases successful/stale ImageBitmaps', async () => {
  let closed = 0, fetched = [], stale = false
  const bitmap = () => ({ width: 2, height: 2, close() { closed++ } })
  await withGlobals({ fetch: async url => { fetched.push(url); return new Response(new Uint8Array(8), { headers: { 'content-type': 'image/png' } }) },
    createImageBitmap: async () => bitmap() }, async () => {
    const owner = canvasOwner()
    await drawChartWordCloudMaskImage(owner, resource().id, { resolveImageResource: () => resource() })
    assert.deepEqual(fetched, [resource().source]); assert.equal(owner.draws.length, 1); assert.equal(closed, 1)
    await assert.rejects(drawChartWordCloudMaskImage(canvasOwner(), 'missing-resource', {}), error => error.code === 'COMPUTE_INVALID_INPUT')
    assert.equal(fetched.length, 1)
    await withGlobals({ createImageBitmap: async () => { stale = true; return bitmap() } }, async () => {
      await assert.rejects(drawChartWordCloudMaskImage(canvasOwner(), resource().id, {
        resolveImageResource: () => resource(), assertCurrent() { if (stale) throw new Error('stale revision') },
      }))
    })
    assert.equal(closed, 2)
  })
})

test('mask loader enforces streaming bytes/MIME and aborts pending decode with late bitmap release', async () => {
  const owner = canvasOwner()
  let cancelled = 0
  const stream = () => new ReadableStream({ start(controller) { controller.enqueue(new Uint8Array(200)); }, cancel() { cancelled++ } })
  await withGlobals({ fetch: async () => new Response(stream(), { headers: { 'content-type': 'image/png' } }) }, async () => {
    await assert.rejects(drawChartWordCloudMaskImage(owner, './mask.png', { memoryBudgetBytes: 256 }), error => error.code === 'COMPUTE_MEMORY_LIMIT')
    assert.equal(cancelled, 1)
  })
  await withGlobals({ fetch: async () => new Response(new Uint8Array(8), { headers: { 'content-type': 'text/html' } }) }, async () => {
    await assert.rejects(drawChartWordCloudMaskImage(owner, './mask.png', {}), error => error.code === 'COMPUTE_INVALID_INPUT')
  })
  const controller = new AbortController(); let finishDecode, beganDecode, closed = 0
  const decoding = new Promise(resolve => { beganDecode = resolve })
  await withGlobals({ fetch: async () => new Response(new Uint8Array(8), { headers: { 'content-type': 'image/png' } }),
    createImageBitmap: () => { beganDecode(); return new Promise(resolve => { finishDecode = resolve }) } }, async () => {
    const loading = drawChartWordCloudMaskImage(owner, './mask.png', { signal: controller.signal })
    await decoding; controller.abort()
    await assert.rejects(loading, error => error.code === 'COMPUTE_CANCELLED')
    finishDecode({ width: 2, height: 2, close() { closed++ } }); await Promise.resolve(); await Promise.resolve()
    assert.equal(closed, 1); assert.equal(owner.draws.length, 0)
  })
})

test('FileReader cancellation and synchronous read failures remove owned callbacks', async () => {
  let current
  class Reader {
    constructor() { current = this }
    readAsDataURL() {}
    abort() { this.onabort?.() }
  }
  const file = { name: 'mask.png', type: 'image/png', size: 8 }, controller = new AbortController()
  await withGlobals({ FileReader: Reader }, async () => {
    const pending = readChartImageFile(file, controller.signal); controller.abort()
    await assert.rejects(pending, error => error.code === 'COMPUTE_CANCELLED')
    assert.equal(current.onload, null); assert.equal(current.onerror, null); assert.equal(current.onabort, null)
    await withGlobals({ FileReader: class extends Reader { readAsDataURL() { throw new Error('read failed') } } }, async () => {
      await assert.rejects(readChartImageFile(file), /read failed/)
      assert.equal(current.onload, null); assert.equal(current.onerror, null); assert.equal(current.onabort, null)
    })
  })
})

test('word-cloud all-zero reduction, repeats and font mapping preserve original tooltip values', () => {
  const data = { categories: ['A', 'B', 'C', 'D', 'E', 'F'], series: [{ values: [0, 0, 0, 0, 0, 0] }] }
  const reduced = chartWordCloudData(data)
  assert.equal(reduced.length, 5); assert.deepEqual(reduced.map(word => [word.value, word.realValue]), Array.from({ length: 5 }, () => [1, 0]))
  const repeated = chartWordCloudData(data, { repeat: true })
  assert.equal(repeated.length, 200); assert.equal(repeated[199].repeated, true); assert.equal(repeated[199].realValue, 0)
  assert.equal(chartWordCloudFontSize(1, [1, 9], [16, 56], 3), 16)
  assert.equal(chartWordCloudFontSize(4, [1, 9], [16, 56], 3), 36)
  assert.equal(chartWordCloudFontSize(9, [1, 9], [16, 56], 3), 56)
})
