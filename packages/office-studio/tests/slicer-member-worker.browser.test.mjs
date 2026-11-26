import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'vite'

// Real production Worker/Rust/WASM drafts. Requires a newly built local
// protocol-2 artifact with create_slicer_members; never skips to a JS double.
// No server, browser, build, test or benchmark has run for these drafts.
let server, browser, context, page, url
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'tests/benchmarks/vite.config.ts'), server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No real slicer Worker fixture port')
  url = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/slicer-member-worker.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0; context = await browser.newContext(); page = await context.newPage()
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(url); await page.waitForFunction(() => document.querySelector('#status')?.textContent === 'ready')
})
afterEach(async () => {
  try { await page?.evaluate(() => window.slicerMemberWorkerFixture?.dispose()); assert.deepEqual(errors, []) }
  finally { await context?.close() }
})
afterAll(async () => { await browser?.close(); await server?.close() })
const run = (method, ...args) => page.evaluate(({ method, args }) => window.slicerMemberWorkerFixture[method](...args), { method, args })
const ascending = [0, 10, 1, 13, 2, 12, 8, 7, 3, 5, 4, 6]
const available = [10, 1, 13, 12, 7, 3, 4]
// compute_version returns the crate's semver, not a module-name prefix.
const assertWasm = result => assert.match(result.engineVersion, /^\d+\.\d+\.\d+(?:[-+].*)?$/)

test('real WASM returns fixed typed identities, duplicate OR, source representatives and only a bounded copied window', async () => {
  const result = await run('fixed', { start: 5, count: 4 })
  assertWasm(result); assert.deepEqual(result.output.indices, ascending.slice(5, 9)); assert.equal(result.output.total, 12)
  assert.deepEqual(result.output.hasData, [1, 0, 1, 1]); assert.equal(result.callerDetached, true)
  assert.equal(result.output.dateCount, 3); assert.equal(result.output.minDate, -2.25); assert.equal(result.output.maxDate, 5.5)
  assert.equal(result.metrics.boundary.inputCopies, 7); assert.equal(result.metrics.boundary.outputCopies, 2)
  assert.equal(result.metrics.boundary.outputBytes, 20); assert.equal(result.metrics.boundary.rustBufferClones, 0)
  assert.equal(result.metrics.boundary.jsonConversions, 0); assert.equal(result.metrics.transport.inputTransfers, 7)
  assert.equal(result.metrics.transport.outputTransfers, 2); assert.equal(result.metrics.ownershipTransfers, 7)
  assert.equal(result.context.kind, 'slicerMembers'); assert.equal(result.context.baseRevision, 0)
})

for (const crossFilter of ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData']) for (const hideItemsWithNoData of [false, true]) {
  test(`real Rust ${crossFilter}/hide=${hideItemsWithNoData} projects before slicing and retains true totals`, async () => {
    const result = await run('fixed', { crossFilter, hideItemsWithNoData })
    const expected = crossFilter === 'none' ? ascending : hideItemsWithNoData ? available
      : crossFilter === 'showItemsWithDataAtTop' ? [10, 1, 13, 12, 7, 3, 4, 0, 2, 8, 5, 6] : ascending
    assertWasm(result); assert.deepEqual(result.output.indices, expected); assert.equal(result.output.total, expected.length)
    if (crossFilter === 'none' || hideItemsWithNoData) assert.ok(result.output.hasData.every(value => value === 1))
    const tail = await run('fixed', { crossFilter, hideItemsWithNoData, start: expected.length + 5, count: 200 })
    assert.equal(tail.output.total, expected.length); assert.deepEqual(tail.output.indices, [])
    assert.equal(tail.metrics.boundary.outputCopies, 0); assert.equal(tail.metrics.transport.outputTransfers, 2)
  })
}

for (const locale of ['en-US', 'zh-CN', 'de-DE']) for (const sortOrder of ['ascending', 'descending']) {
  test(`real ${locale}/${sortOrder} Intl/calendar boundary preserves aliases, boolean order and UTF-16 identities`, async () => {
    for (const customListSort of [false, true]) {
      const result = await run('text', locale, customListSort, sortOrder)
      assertWasm(result); assert.deepEqual(result.output, result.baseline)
      assert.equal(result.output.total, new Set(result.values).size + 2)
      const booleanRows = result.output.indices.filter(row => row >= result.values.length)
      assert.deepEqual(booleanRows, sortOrder === 'ascending' ? [12, 11] : [11, 12])
      if (customListSort) {
        const januaryRows = [1, 2].filter(row => result.output.indices.includes(row))
        assert.deepEqual(result.output.indices.filter(row => januaryRows.includes(row)), januaryRows)
        assert.ok(sortOrder === 'ascending' ? result.output.indices.indexOf(1) < result.output.indices.indexOf(0)
          : result.output.indices.indexOf(1) > result.output.indices.indexOf(0))
      }
    }
  })
}

for (const kind of ['table', 'pivot']) for (const crossFilter of ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData']) for (const hideItemsWithNoData of [false, true]) {
  test(`real ${kind}/${crossFilter}/hide=${hideItemsWithNoData} service combines native/report availability without changing selection or startItem`, async () => {
    const result = await run('sourceWindow', kind, { crossFilter, hideItemsWithNoData }, { start: 1, count: 2 })
    assertWasm(result); assert.deepEqual(result.window, result.baseline); assert.equal(result.unchanged, true)
    assert.equal(result.window.total, crossFilter !== 'none' && hideItemsWithNoData ? 3 : 7)
    assert.equal(result.metrics.length, 1); assert.equal(result.metrics[0].compute.boundary.rustBufferClones, 0)
    assert.equal(result.metrics[0].compute.transport.outputBytes, result.window.items.length * 5)
  })
}

test('real metadata includes cached typed Date bounds, grouped identity and fractional serials without guessing ISO text or numbers', async () => {
  const result = await run('metadata')
  assert.deepEqual(result.metadata.timeline, { minDate: -2.25, maxDate: 100, interactive: true })
  assert.deepEqual(result.metadata.dateGroup, { type: 3, dateType: 'month' }); assert.equal(result.metadata.fieldId, 'date-source')
  assert.equal(result.unchanged, true); assert.equal(result.metrics[0].mode, 'metadata')
  assert.equal(result.metrics[0].compute.boundary.inputCopies, 5); assert.equal(result.metrics[0].compute.boundary.outputCopies, 0)
  assert.equal(result.metrics[0].compute.transport.outputTransfers, 2); assert.equal(result.metrics[0].compute.transport.outputBytes, 0)
})

test('real empty items/metadata outputs keep live independent empty buffers and no-date Timeline stays noninteractive', async () => {
  for (const mode of ['items', 'canvas', 'metadata']) {
    const result = await run('empty', mode); assertWasm(result)
    assert.deepEqual(result.output, { kind: 'slicerMembers', total: 0, indices: [], hasData: [], dateCount: 0, minDate: null, maxDate: null })
    assert.equal(result.callerDetached, true); assert.equal(result.metrics.boundary.inputCopies, 1)
    assert.equal(result.metrics.boundary.inputBytes, 4); assert.equal(result.metrics.boundary.outputCopies, 0)
  }
  const result = await run('metadata', true)
  assert.deepEqual(result.metadata.timeline, { interactive: false }); assert.equal(result.unchanged, true)
})

for (const action of ['cancel', 'canvas', 'metadata', 'revision', 'close', 'source', 'permission', 'service']) {
  test(`real ${action} invalidation while awaiting Worker rejects the owner and leaves shared client reusable`, async () => {
    const result = await run('lifecycle', action)
    const expected = { cancel: 'SLICER_RPC_ABORTED', canvas: 'SLICER_RPC_ABORTED', metadata: 'SLICER_RPC_ABORTED', revision: 'SLICER_SOURCE_INVALIDATED', close: 'SLICER_RPC_UNIT_DISPOSED',
      source: 'SLICER_SOURCE_INVALIDATED', permission: 'SLICER_SOURCE_INVALIDATED', service: 'SLICER_RPC_DISPOSED' }[action]
    assert.equal(result.triggered, true); assert.equal(result.rejected.code, expected); assert.equal(result.workerCount, 1)
    assertWasm(result); assert.deepEqual(result.recovery.indices, ascending); assert.ok(result.rejectionMs >= 0)
  })
}

test('real cooperative cancellation acknowledges release before a reused Worker completes another member query', async () => {
  const result = await run('directCancellation')
  assertWasm(result); assert.equal(result.triggered, true); assert.equal(result.rejected.code, 'COMPUTE_CANCELLED')
  assert.equal(result.workerCount, 1); assert.deepEqual(result.recovered.indices, ascending)
  assert.ok(result.acknowledgementMs >= 0); assert.ok(result.recoveryMs >= result.acknowledgementMs)
})

test('a delivered previous real Worker result cannot settle the current real WASM request', async () => {
  const result = await run('lateResult')
  assertWasm(result); assert.equal(result.replayed, true); assert.equal(result.workerCount, 1)
  assert.notEqual(result.requestId, result.previousRequestId); assert.deepEqual(result.output, result.baseline)
})

test('real output copies survive owner/client disposal and transfer independently after Rust result release', async () => {
  const result = await run('release')
  assertWasm(result); assert.equal(result.independent, true); assert.equal(result.callerDetached, true)
  assert.deepEqual(result.after, result.before); assert.deepEqual(result.after.indices, ascending)
})

test('real Canvas windows globally fold captions and preserve selected/partial/hasData across page boundaries', async () => {
  const first = await run('canvas', {}, { start:0,count:3 }), second = await run('canvas', {}, { start:3,count:200 })
  assertWasm(first); assertWasm(second); assert.equal(first.window.total,7); assert.equal(second.window.total,7)
  const items = [...first.window.items,...second.window.items]
  assert.deepEqual(items.map(item => item.label),['','0','1','-2.25','5.5','#N/A','true'])
  assert.equal(new Set(items.map(item => item.label)).size,7)
  assert.ok(items.every(item => item.selected === true && item.partial === false))
  assert.deepEqual(items.map(item => item.hasData),[false,true,true,true,true,true,true])
  const sortedLabels = ['','0','1','-2.25','5.5','#N/A','true'], dataLabels = sortedLabels.slice(1)
  for (const crossFilter of ['none','showItemsWithDataAtTop','showItemsWithNoData']) for (const hideItemsWithNoData of [false,true]) {
    const projected = await run('canvas',{ crossFilter,hideItemsWithNoData }); assertWasm(projected)
    const labels = crossFilter === 'none' ? sortedLabels : hideItemsWithNoData ? dataLabels
      : crossFilter === 'showItemsWithDataAtTop' ? [...dataLabels,''] : sortedLabels
    assert.equal(projected.window.total,labels.length); assert.deepEqual(projected.window.items.map(item => item.label),labels)
    assert.equal(new Set(projected.window.items.map(item => item.label)).size,labels.length)
    assert.ok(projected.window.items.every(item => item.selected === true && item.partial === false))
    assert.deepEqual(projected.window.items.map(item => item.hasData),labels.map(label => crossFilter === 'none' || label !== ''))
  }
})

test('real Canvas keeps ISO Date captions and all-source half selection while formal typed members stay distinct', async () => {
  const result = await run('canvasDateCaptions')
  assertWasm(result); assert.equal(result.display.total,2); assert.equal(result.formal.total,3)
  assert.deepEqual(result.display.items.map(({ label,selected,partial,hasData }) => ({ label,selected,partial,hasData })),[
    { label:'100',selected:false,partial:true,hasData:false },{ label:'1900-04-09',selected:true,partial:false,hasData:true }])
  assert.ok(result.formal.items.every(item => !Object.hasOwn(item,'label') && !Object.hasOwn(item,'selected')))
  const hidden = await run('canvasDateCaptions',true)
  assert.equal(hidden.display.total,1); assert.equal(hidden.display.items[0].label,'1900-04-09')
  assert.equal(hidden.display.items[0].selected,true); assert.equal(hidden.display.items[0].partial,false)
})
