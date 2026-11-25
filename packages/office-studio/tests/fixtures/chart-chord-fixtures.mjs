import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'

// P28 source fixtures and independently authored portable numerical goldens.
// Imports declare helpers only; no implementation/check/server is executed.
export const CHORD_TABLE = [['Source', 'Target', 'Weight'], ['B', 'A', 2], ['A', 'B', 3], ['B', 'B', 9]]
export const CHORD_PALETTE = ['rgba(255,0,0,0.8)', 'rgba(0,0,255,0.6)', '#00ff00', '#ff00ff']
export const CHORD_CASES = JSON.parse(readFileSync(new URL('../../../../tests/fixtures/chart-chord.json', import.meta.url), 'utf8')).cases
export const goldenValue = value => Array.isArray(value) ? value[0] / value[1] * Math.PI : value
export function chordModel(options = {}) {
  return { id: 'chord-draft', sheetId: 'sheet', revision: 0, type: 'chord', categoryField: 0, valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 340 }, palette: CHORD_PALETTE,
    legend: { visible: false, position: 'bottom' }, xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' },
    options: { mapping: { chord: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } }, ...options } }
}
export function chordRecords(records = CHORD_TABLE.slice(1), formattedRecords) {
  return { fields: ['Source', 'Target', 'Weight'], records, ...(formattedRecords ? { formattedRecords } : {}) }
}
export function chordGoldenTask(fixture) {
  const rows = fixture.rows.length, values = new Float64Array(rows * 3), states = new Uint8Array(values.length)
  for (let row = 0; row < rows; row++) for (let column = 0; column < 3; column++) {
    const value = fixture.rows[row][column]
    if (value !== null) { values[column * rows + row] = value; states[column * rows + row] = 1 }
  }
  return { kind: 'chartStatistics', values, states, categoryIds: new Uint32Array(rows), params: {
    rowCount: rows, columnCount: 3, categoryCount: rows ? 1 : 0, series: [{ field: 2, seriesPosition: 0, column: 0 }],
    jobs: [{ kind: 'chord', seriesIndex: 0, slot: 0, sourceColumn: 1, targetColumn: 2, labelCount: fixture.labelCount, ...fixture.layout }] } }
}
export function chordGoldenOutput(fixture) {
  const values = Float64Array.from(fixture.values, goldenValue), indices = Uint32Array.from(fixture.indices)
  return { kind: 'chartStatistics', values, indices, segments: [{ kind: 'chord', jobIndex: 0, field: 2, seriesPosition: 0, slot: 0,
    valueOffset: 0, valueLength: values.length, indexOffset: 0, indexLength: indices.length,
    validSampleCount: fixture.validSampleCount, available: fixture.available }] }
}
export function chordGoldenPlan(fixture) {
  return { task: chordGoldenTask(fixture), labels: Array.from({ length: fixture.labelCount }, (_, index) => String.fromCharCode(65 + index)),
    mapping: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 }, valueFieldName: 'Weight', formatter: String, retainedBytes: 1024 }
}
/** Transport for a real Facade source A1:C4, with independently known two-node
 * aggregate 5, 11 after changing C2 to 8, or count 2. Equal arcs have the closed
 * form pi minus authored padding. This never calls either numerical kernel. */
export function chartChordGoldenTransport() {
  const revisions = new Map(), calls = []
  return { calls, setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    submit(source, options) {
      const task = structuredClone(source, { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] })
      assert.equal(task.kind, 'chartStatistics'); assert.equal(task.params.rowCount, 3)
      const job = task.params.jobs[0]
      assert.equal(job.kind, 'chord'); assert.equal(job.sourceColumn, 1); assert.equal(job.targetColumn, 2); assert.equal(job.labelCount, 2)
      assert.equal(source.values.byteLength, 0); assert.equal(options.ownership, 'transfer'); assert.ok(options.signal)
      const samples = Array.from(task.values.slice(0, 3)), count = samples[0] === 1 && samples[1] === 1
      assert.deepEqual(samples, count ? [1, 1, 0] : samples[0] === 8 ? [8, 3, 0] : [2, 3, 0])
      assert.deepEqual(Array.from(task.states), [1,1,0,1,1,0,1,1,0])
      assert.deepEqual(Array.from(task.values.slice(3)), [0,1,0,1,0,0])
      assert.ok(job.padAngle >= 0 && job.padAngle < 90); assert.ok(job.minAngle >= 0 && job.minAngle < 90)
      const start = -(job.startAngle % 360) / 180 * Math.PI, direction = job.clockwise ? 1 : -1
      const padding = job.padAngle / 180 * Math.PI, firstEnd = start + direction * (Math.PI - padding)
      const secondStart = start + direction * Math.PI, secondEnd = secondStart + direction * (Math.PI - padding)
      const value = count ? 2 : samples[0] === 8 ? 11 : 5, field = task.params.series[0].field
      const values = new Float64Array([value,start,firstEnd,value,secondStart,secondEnd,value,start,firstEnd,secondStart,secondEnd]), indices = new Uint32Array([2,1,0,1,0,1])
      calls.push({ task, options })
      return { requestId: 'chord-golden-' + calls.length, cancel() {}, promise: Promise.resolve({ output: { kind: 'chartStatistics', values, indices,
        segments: [{ kind: 'chord', jobIndex: 0, field, seriesPosition: 0, slot: 0, valueOffset: 0, valueLength: values.length,
          indexOffset: 0, indexLength: indices.length, validSampleCount: 2, available: true }] } }) }
    } }
}
