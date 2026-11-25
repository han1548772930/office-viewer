import assert from 'node:assert/strict'
import { test, vi } from 'vitest'
import { CHART_IMAGE_EXPORT_MAX_PIXELS, chartImageExportEnvironment, chartImageExportFilename, exportChartCanvasImage } from '../src/features/charts/chart-image-export.ts'

// Written only. Mock encoding isolates ownership/cancellation; real PNG bytes are checked in the browser draft.
const source = () => ({ width: 32, height: 16 })
function environment() {
  const paints = [], downloads = [], encoded = new Blob(['png-encoding-fixture'], { type: 'image/png' })
  let callback
  const canvas = { width: 0, height: 0, getContext: () => ({ drawImage: (...args) => paints.push(args) }), toBlob: value => { callback = value } }
  return { paints, downloads, canvas, encoded, finish: value => callback(value === undefined ? encoded : value),
    value: { createCanvas: () => canvas, download: (blob, filename) => downloads.push({ blob, filename }) } }
}

test('export snapshots the actual supplied bitmap once before awaiting encoding and releases its owned surface', async () => {
  const env = environment(), painted = source(), exporting = exportChartCanvasImage(painted, 'object:42', { environment: env.value })
  assert.deepEqual(env.paints, [[painted, 0, 0]]); assert.equal(env.canvas.width, 32); assert.equal(env.canvas.height, 16)
  painted.width = 128; assert.equal(env.canvas.width, 32)
  env.finish(); assert.equal(await exporting, env.encoded)
  assert.deepEqual(env.downloads, [{ blob: env.encoded, filename: 'object_42.png' }]); assert.equal(env.canvas.width, 0); assert.equal(env.canvas.height, 0)
  assert.equal(chartImageExportFilename('chart<>:"/\\|?*'), 'chart_________.png')
})

test('preflight abort and pixel budget reject before allocating or copying any surface', async () => {
  const env = environment(), aborted = new AbortController(); aborted.abort()
  await assert.rejects(exportChartCanvasImage(source(), 'object', { environment: env.value, signal: aborted.signal }), error => error.name === 'AbortError')
  await assert.rejects(exportChartCanvasImage({ width: CHART_IMAGE_EXPORT_MAX_PIXELS + 1, height: 1 }, 'object', { environment: env.value }), RangeError)
  await assert.rejects(exportChartCanvasImage({ width: 0, height: 1 }, 'object', { environment: env.value }), RangeError)
  assert.equal(env.paints.length, 0); assert.equal(env.downloads.length, 0); assert.equal(env.canvas.width, 0)
})

test('an adapter cannot claim the borrowed renderer canvas as its owned encoding surface', async () => {
  const borrowed = source(), downloads = []
  await assert.rejects(exportChartCanvasImage(borrowed, 'object', { environment: { createCanvas: () => borrowed, download: blob => downloads.push(blob) } }), /独立画布/)
  assert.equal(borrowed.width, 32); assert.equal(borrowed.height, 16); assert.deepEqual(downloads, [])
})

test('abort during encoding releases the surface and a late callback cannot download', async () => {
  const env = environment(), abort = new AbortController()
  const exporting = exportChartCanvasImage(source(), 'object', { environment: env.value, signal: abort.signal })
  const rejected = assert.rejects(exporting, error => error.name === 'AbortError'); abort.abort(); await rejected
  assert.equal(env.canvas.width, 0); env.finish(); await Promise.resolve(); assert.equal(env.downloads.length, 0)
})

test('null/wrong-format encoding and a stale author reject without handing a file to the browser', async () => {
  for (const output of [null, new Blob(['wrong'], { type: 'image/jpeg' }), new Blob([], { type: 'image/png' })]) {
    const env = environment(), exporting = exportChartCanvasImage(source(), 'object', { environment: env.value })
    const rejected = assert.rejects(exporting, /PNG 编码失败/); env.finish(output); await rejected
    assert.equal(env.canvas.width, 0); assert.equal(env.downloads.length, 0)
  }
  let current = true
  const env = environment(), exporting = exportChartCanvasImage(source(), 'object', { environment: env.value,
    assertCurrent() { if (!current) throw new Error('author changed') } })
  current = false; const rejected = assert.rejects(exporting, /author changed/); env.finish(); await rejected
  assert.equal(env.downloads.length, 0); assert.equal(env.canvas.height, 0)
})

test('an encoder that never settles times out and ignores its later callback', async () => {
  vi.useFakeTimers()
  try {
    const env = environment(), exporting = exportChartCanvasImage(source(), 'object', { environment: env.value })
    const rejected = assert.rejects(exporting, /编码超时/)
    await vi.advanceTimersByTimeAsync(15_000); await rejected; env.finish()
    assert.equal(env.canvas.width, 0); assert.equal(env.downloads.length, 0)
  } finally { vi.useRealTimers() }
})

test('the browser download uses a connected anchor and revokes its URL after consumption or immediately on failure', () => {
  vi.useFakeTimers()
  const revoked = [], clicks = [], children = new Set()
  let fail = false
  const anchor = { style: {}, remove() { children.delete(this) }, click() {
    assert.equal(children.has(this), true); clicks.push({ href: this.href, filename: this.download }); if (fail) throw new Error('click failed')
  } }
  vi.stubGlobal('document', { createElement: () => anchor, body: { append: node => children.add(node) } })
  vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:export-owned'); vi.spyOn(URL, 'revokeObjectURL').mockImplementation(url => revoked.push(url))
  try {
    const blob = new Blob(['fixture'], { type: 'image/png' })
    chartImageExportEnvironment.download(blob, 'object.png')
    assert.deepEqual(clicks, [{ href: 'blob:export-owned', filename: 'object.png' }]); assert.equal(children.size, 0); assert.deepEqual(revoked, [])
    vi.advanceTimersByTime(1000); assert.deepEqual(revoked, ['blob:export-owned'])
    fail = true; assert.throws(() => chartImageExportEnvironment.download(blob, 'object.png'), /click failed/)
    assert.deepEqual(revoked, ['blob:export-owned', 'blob:export-owned']); assert.equal(children.size, 0)
  } finally { vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers() }
})
