import assert from 'node:assert/strict'

// Independent materialized golden tuples for the shared signed A1:C8 sample.
// This fake transports data; it never calls either hierarchy implementation.
export function chartHierarchyGoldenTransport() {
  const revisions = new Map(), calls = []
  return { calls, setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    submit(source, options) {
      const task = structuredClone(source, { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] })
      assert.equal(task.kind, 'chartStatistics'); assert.equal(task.params.rowCount, 7)
      assert.deepEqual(task.params.series, [{ field: 2, seriesPosition: 0, column: 0 }])
      const job = task.params.jobs[0]
      assert.equal(job.kind, 'hierarchy'); assert.deepEqual(job.hierarchyColumns, [1, 2]); assert.deepEqual(job.labelColumns, [3, 4])
      assert.equal(job.keyCount, 6); assert.equal(job.labelCount, 6); assert.equal(source.values.byteLength, 0)
      assert.deepEqual(Array.from(task.values.slice(0, 7)), [8, 12, 6, 3, -2, -4, -5])
      assert.equal(options.ownership, 'transfer'); assert.ok(options.signal)
      calls.push({ task, options })
      const values = new Float64Array([13, -5, 12, 12, 6, 6, 9, 3, 6, 6])
      const indices = new Uint32Array([5, 2, 0xffffffff, 2, 2, 2, 0, 3, 3, 3, 0, 4, 4, 4, 0xffffffff, 0, 0, 0, 3, 1, 1, 1, 3, 0])
      return { requestId: 'hierarchy-golden-' + calls.length, cancel() {}, promise: Promise.resolve({ output: { kind: 'chartStatistics', values, indices,
        segments: [{ kind: 'hierarchy', jobIndex: 0, field: 2, seriesPosition: 0, slot: 0, valueOffset: 0, valueLength: values.length,
          indexOffset: 0, indexLength: indices.length, validSampleCount: 7, available: true }] } }) }
    } }
}
