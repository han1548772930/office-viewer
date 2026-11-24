import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { ComputeError } from '../src/compute/errors.ts'

// Drafts for unified verification. The real chart service/Facade use a
// controlled Worker transport here; no second waterfall engine is provided.
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
  const revisions = new Map(), calls = [], nextCalls = new Map()
  return { calls,
    setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    waitForCall(index) {
      if (calls[index]) return Promise.resolve(calls[index])
      const pending = deferred(); nextCalls.set(index, pending); return pending.promise
    },
    submit(source, options) {
      const task = structuredClone(source, { transfer: [source.values.buffer, source.states.buffer, source.categoryIds.buffer] })
      const pending = deferred(), requestId = 'facade-draft-' + calls.length
      const abort = () => pending.reject(new ComputeError('COMPUTE_CANCELLED', 'Draft transport acknowledged cancellation'))
      options.signal.addEventListener('abort', abort, { once: true })
      const cleanup = () => options.signal.removeEventListener('abort', abort)
      pending.promise.then(cleanup, cleanup)
      const call = { task, source, options, complete() {
        const values = new Float64Array([3, 0, 5, 5, 5, 3, -2, 5, 5])
        const indices = new Uint32Array([2, 1, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1])
        pending.resolve({ output: { kind: 'chartStatistics', values, indices,
          segments: [{ kind: 'waterfallLayout', jobIndex: 0, field: 1, seriesPosition: 0, slot: 0,
            valueOffset: 0, valueLength: values.length, indexOffset: 0, indexLength: indices.length,
            validSampleCount: 2, available: true }] } })
      } }
      const index = calls.length; calls.push(call); nextCalls.get(index)?.resolve(call); nextCalls.delete(index)
      return { requestId, promise: pending.promise, cancel: abort }
    },
  }
}
async function setup() {
  const compute = transport(), studio = new api.OfficeStudio({ computeClient: compute })
  const workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:B3').setValues([['Label', 'Flow'], ['A', 5], ['B', -2]])
  const chart = await sheet.insertChart(sheet.newChart('waterfall').setSource('A1:B3').setValueFields([1])
    .setTitle('Flow').setUseSubtotal(false).setOption('waterfall', { pointRoles: [] }).build())
  return { compute, studio, workbook, sheet, chart }
}

test('async reads share the chart service, populate sync reads and detach results from the cached arrays', async () => {
  const { compute, studio, workbook, chart } = await setup()
  try {
    assert.throws(() => chart.getLayout(), error => error.code === 'COMPUTE_NOT_READY')
    assert.throws(() => chart.getRenderCommands(), error => error.code === 'COMPUTE_NOT_READY')
    await chart.update({ title: 'Before read' })
    const info = chart.getInfo(), revision = studio.getWorkbookRevision(workbook.getId())
    const dataRead = chart.getComputedData(), layoutRead = chart.getComputedLayout(), commandRead = chart.getComputedRenderCommands()
    const call = await compute.waitForCall(0)
    assert.equal(compute.calls.length, 1)
    assert.equal(call.source.values.byteLength, 0)
    assert.equal(call.options.documentId, workbook.getId())
    assert.equal(call.options.modelRevision, revision)
    call.complete()
    const [data, layout, commands] = await Promise.all([dataRead, layoutRead, commandRead])
    assert.equal(data.series[0].values[0], 5)
    assert.deepEqual(chart.getLayout(), layout)
    assert.deepEqual(chart.getRenderCommands(), commands)
    data.series[0].values[0] = 999; layout.plot.width = 1; commands.length = 0
    assert.equal((await chart.getComputedData()).series[0].values[0], 5)
    assert.ok(chart.getLayout().plot.width > 1)
    assert.ok(chart.getRenderCommands().length > 0)
    assert.equal(compute.calls.length, 1)
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
    assert.deepEqual(chart.getInfo(), info)
    workbook.undo()
    assert.equal(chart.getInfo().title, 'Flow')
  } finally { studio.dispose() }
})

test('aborting after the shared service resolves still prevents the Facade from returning a cached result', async () => {
  const { compute, studio, chart } = await setup(), controller = new AbortController()
  const service = studio.chartStatisticsService, request = service.request.bind(service)
  try {
    service.request = (...args) => {
      const pending = request(...args)
      return { ...pending, promise: pending.promise.then(snapshot => { controller.abort(); return snapshot }) }
    }
    const read = chart.getComputedLayout({ signal: controller.signal })
    const rejection = assert.rejects(read, error => error.code === 'COMPUTE_CANCELLED')
    ;(await compute.waitForCall(0)).complete()
    await rejection
    service.request = request
    assert.ok(chart.getLayout().series[0].points.length > 0)
    assert.equal(compute.calls.length, 1)
  } finally { service.request = request; studio.dispose() }
})

test('a source edit between service completion and the Facade continuation rejects old geometry', async () => {
  const { compute, studio, sheet, chart } = await setup()
  const service = studio.chartStatisticsService, request = service.request.bind(service)
  try {
    service.request = (...args) => {
      const pending = request(...args)
      return { ...pending, promise: pending.promise.then(snapshot => { sheet.getRange('B2').setValue(6); return snapshot }) }
    }
    const read = chart.getComputedRenderCommands(), rejection = assert.rejects(read, error => error.code === 'COMPUTE_STALE_RESULT')
    ;(await compute.waitForCall(0)).complete()
    await rejection
    assert.equal(sheet.getRange('B2').getValue(), 6)
    assert.throws(() => chart.getLayout(), error => error.code === 'COMPUTE_NOT_READY')
  } finally { service.request = request; studio.dispose() }
})

test('source range permissions are rechecked even when an incoming ACL update keeps the content revision', async () => {
  const { compute, studio, workbook, sheet, chart } = await setup()
  const service = studio.chartStatisticsService, request = service.request.bind(service)
  const revision = studio.getWorkbookRevision(workbook.getId())
  try {
    service.request = (...args) => {
      const pending = request(...args)
      return { ...pending, promise: pending.promise.then(snapshot => {
        const source = studio.getWorkbookModel(workbook.getId()).sheets.find(candidate => candidate.id === sheet.getSheetId())
        source.rangeProtectionRules = [{ id: 'incoming-range', permissionId: 'incoming-permission', ownerUserId: 'other-user',
          ranges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 1 }],
          editState: 'onlyMe', viewState: 'noOneElseCanView', allowedUserIds: [] }]
        return snapshot
      }) }
    }
    const read = chart.getComputedData(), rejection = assert.rejects(read, error => error.code === 'COMPUTE_INVALID_INPUT' && /源区域/.test(error.message))
    ;(await compute.waitForCall(0)).complete()
    await rejection
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
    for (const method of ['getData', 'getDescription', 'getLayout', 'getRenderCommands']) {
      assert.throws(() => chart[method](), error => error.code === 'COMPUTE_INVALID_INPUT' && /源区域/.test(error.message))
    }
  } finally { service.request = request; studio.dispose() }
})
