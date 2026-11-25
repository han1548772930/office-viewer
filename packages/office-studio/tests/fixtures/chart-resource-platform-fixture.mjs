import assert from 'node:assert/strict'

// Browser adapters are mocked only to isolate host history/resource contracts.
// The PNG below is a signature-only fixture, never an accepted image sample.
export const chartResourceFixture = id => ({ id: id ?? 'chart-image:fixture', sourceType: 'dataUrl',
  source: 'data:image/png;base64,iVBORw0KGgo=', mimeType: 'image/png', sizeBytes: 8, name: 'mask.png' })
export async function withChartResourcePlatform(work, fetchImage = async () => new Response(new Uint8Array(8), { headers: { 'content-type': 'image/png' } })) {
  class Canvas {
    constructor(width, height) {
      this.width = width; this.height = height
      const context = { measureText: () => ({ width: 8, actualBoundingBoxLeft: 2, actualBoundingBoxRight: 2,
        actualBoundingBoxAscent: 3, actualBoundingBoxDescent: 1 }),
        getImageData: () => {
          const data = new Uint8ClampedArray(this.width * this.height * 4)
          for (let at = 3; at < data.length; at += 4) data[at] = 255
          return { data }
        } }
      for (const method of ['save', 'restore', 'translate', 'scale', 'rotate', 'beginPath', 'closePath', 'ellipse',
        'moveTo', 'lineTo', 'bezierCurveTo', 'fill', 'fillRect', 'strokeText', 'fillText', 'drawImage']) context[method] = () => {}
      this.getContext = () => context
    }
  }
  const replacements = { OffscreenCanvas: Canvas, createImageBitmap: async () => ({ width: 2, height: 2, close() {} }), fetch: fetchImage }
  const original = new Map(Object.keys(replacements).map(key => [key, Object.getOwnPropertyDescriptor(globalThis, key)]))
  for (const [key, value] of Object.entries(replacements)) Object.defineProperty(globalThis, key, { configurable: true, writable: true, value })
  try { return await work() } finally {
    for (const [key, descriptor] of original) if (descriptor) Object.defineProperty(globalThis, key, descriptor); else delete globalThis[key]
  }
}

/** Materialized tuples for the one-word/two-node fixtures. No production
 * layout or converter is called, so host tests cannot self-confirm the solver. */
export function chartResourceGoldenTransport() {
  const revisions = new Map(), calls = []
  return { calls, setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    submit(source, options) {
      const task = structuredClone(source, { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] })
      const job = task.params.jobs[0], field = task.params.series[job.seriesIndex].field
      assert.equal(task.kind, 'chartStatistics'); assert.equal(task.params.jobs.length, 1)
      assert.equal(source.values.byteLength, 0); assert.equal(options.ownership, 'transfer')
      let values, indices
      if (job.kind === 'wordCloud') {
        assert.equal(job.glyphs.length, 1); assert.ok(job.width > 30); assert.ok(job.height > 30)
        values = new Float64Array([15, 15, 1]); indices = new Uint32Array([0])
      } else {
        assert.equal(job.kind, 'relationLayout'); assert.equal(job.nodeCount, 2)
        assert.equal(job.useValueAsSymbolSize, false); assert.equal(job.symbolSize, 10)
        values = new Float64Array([10, 20, 10, 40, 60, 10]); indices = new Uint32Array([0, 1])
      }
      calls.push({ task, options })
      return { requestId: `chart-resource-${calls.length}`, cancel() {}, promise: Promise.resolve({ output: {
        kind: 'chartStatistics', values, indices, segments: [{ kind: job.kind, jobIndex: 0, field,
          seriesPosition: task.params.series[job.seriesIndex].seriesPosition, slot: 0, valueOffset: 0, valueLength: values.length,
          indexOffset: 0, indexLength: indices.length, validSampleCount: indices.length, available: true }] } }) }
    } }
}
