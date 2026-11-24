import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { ComputeError } from '../src/compute/errors.ts'

// Deferred cross-module transaction drafts. The numerical Worker is verified
// separately; this transport holds independently specified histogram tuples.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
function deferred() {
  let resolve, reject
  const promise = new Promise((yes, no) => { resolve = yes; reject = no })
  return { promise, resolve, reject }
}
function transport() {
  const revisions = new Map(), calls = [], waiters = new Map()
  return { calls,
    setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    waitForCall(index) {
      if (calls[index]) return Promise.resolve(calls[index])
      const pending = deferred(); waiters.set(index, pending); return pending.promise
    },
    submit(source, options) {
      const task = structuredClone(source, { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] })
      const pending = deferred(), index = calls.length
      const abort = () => pending.reject(new ComputeError('COMPUTE_CANCELLED', 'Draft transport acknowledged cancellation'))
      options.signal.addEventListener('abort', abort, { once: true })
      const cleanup = () => options.signal.removeEventListener('abort', abort)
      pending.promise.then(cleanup, cleanup)
      const call = { task, options, complete(negative = false) {
        pending.resolve({ output: { kind: 'chartStatistics',
          values: new Float64Array(negative ? [0, -1, -1, 0, 0, 1, 1, 0] : [0, 1, 1, 2, 2, 3, 3, 0]),
          indices: new Uint32Array([0, 2, 1, 1, 1, 1, 2, 1]),
          segments: [{ kind: 'histogram', jobIndex: 0, field: 1, seriesPosition: 0, slot: 0,
            valueOffset: 0, valueLength: 8, indexOffset: 0, indexLength: 8, validSampleCount: 5, available: true }] } })
      } }
      calls.push(call); waiters.get(index)?.resolve(call); waiters.delete(index)
      return { requestId: 'candidate-' + index, promise: pending.promise, cancel: abort }
    },
  }
}
async function setup(negative = false) {
  const compute = transport(), studio = new api.OfficeStudio({ computeClient: compute })
  const workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:B6').setValues([['Label', 'Value'], ...[0, 1, 2, 3, 4].map(value => [String(value), negative ? value - 2 : value])])
  const chart = await sheet.insertChart(sheet.newChart('histogram').setSource('A1:B6').setValueFields([1])
    .setTitle('Published').setOption('histogram', { binning: { mode: 'count', count: 2 },
      underflowThreshold: negative ? -1 : 1, overflowThreshold: negative ? 1 : 3 }).build())
  return { compute, studio, workbook, sheet, chart }
}

test('date chart update publishes once after final statistics and belongs to ordinary undo/redo history', async () => {
  const { compute, studio, workbook, chart } = await setup(), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.id === 'sheet.command.update-chart') events.push(event) })
  try {
    const id = workbook.getId(), before = studio.getWorkbookRevision(id)
    const update = chart.update({ title: 'Validated', options: { mapping: { useDateAxis: true } } })
    const call = await compute.waitForCall(0)
    assert.equal(chart.getInfo().title, 'Published')
    assert.equal(studio.getWorkbookRevision(id), before)
    assert.equal(events.length, 0)
    call.complete(); assert.equal(await update, true)
    assert.equal(chart.getInfo().title, 'Validated')
    assert.equal(studio.getWorkbookRevision(id), before + 1)
    assert.equal(events.length, 1)
    workbook.undo(); assert.equal(chart.getInfo().title, 'Published')
    assert.equal(chart.getInfo().options.mapping?.useDateAxis, undefined)
    workbook.redo(); assert.equal(chart.getInfo().title, 'Validated')
  } finally { studio.dispose() }
})

test('an invalid derived date never publishes candidate options, content revision or history', async () => {
  const { compute, studio, workbook, chart } = await setup(true), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.id === 'sheet.command.update-chart') events.push(event) })
  try {
    const id = workbook.getId(), revision = studio.getWorkbookRevision(id), saved = structuredClone(chart.getInfo())
    const history = studio.editor.history.captureSavepoint(id)
    const update = chart.update({ title: 'Rejected', options: { mapping: { useDateAxis: true } } })
    const rejected = assert.rejects(update, /日期|date|分类/)
    ;(await compute.waitForCall(0)).complete(true); await rejected
    assert.deepEqual(chart.getInfo(), saved)
    assert.equal(studio.getWorkbookRevision(id), revision)
    assert.deepEqual(studio.editor.history.captureSavepoint(id), history)
    assert.equal(events.length, 0)
  } finally { studio.dispose() }
})

test('an intervening source edit rejects a pending candidate instead of overwriting the newer workbook', async () => {
  const { compute, studio, workbook, sheet, chart } = await setup()
  try {
    const update = chart.update({ title: 'Stale', options: { mapping: { useDateAxis: true } } })
    const rejected = assert.rejects(update, error => /COMPUTE_STALE_RESULT|COMPUTE_CANCELLED/.test(error.code ?? '') || /已更改|失效/.test(error.message))
    const call = await compute.waitForCall(0)
    sheet.getRange('B2').setValue(99)
    await rejected; call.complete()
    assert.equal(chart.getInfo().title, 'Published')
    assert.equal(sheet.getRange('B2').getValue(), 99)
    assert.equal(chart.getInfo().options.mapping?.useDateAxis, undefined)
    workbook.undo(); assert.equal(sheet.getRange('B2').getValue(), 0)
  } finally { studio.dispose() }
})

test('sync execution refuses asynchronous chart mutations before events or content changes', async () => {
  const { studio, workbook, chart } = await setup(), events = []
  studio.events.addEvent('BeforeCommandExecute', event => events.push(event.id))
  try {
    const saved = structuredClone(chart.getInfo()), revision = studio.getWorkbookRevision(workbook.getId())
    assert.throws(() => studio.executeCommandSync('sheet.command.update-chart',
      { unitId: workbook.getId(), chartId: chart.getInfo().id, patch: { title: 'Sync' } }), /异步/)
    assert.deepEqual(chart.getInfo(), saved)
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
    assert.equal(events.length, 0)
  } finally { studio.dispose() }
})

test('candidate validators cannot silently rewrite the prepared workbook', async () => {
  const { studio, workbook } = await setup()
  try {
    const id = workbook.getId(), saved = structuredClone(studio.getWorkbookModel(id)), revision = studio.getWorkbookRevision(id)
    await assert.rejects(() => studio.commitWorkbookWithValidation(id, 'Candidate', candidate => { candidate.name = 'Planned' },
      async candidate => { candidate.name = 'Validator mutation' }), /校验修改/)
    assert.deepEqual(studio.getWorkbookModel(id), saved)
    assert.equal(studio.getWorkbookRevision(id), revision)
  } finally { studio.dispose() }
})
