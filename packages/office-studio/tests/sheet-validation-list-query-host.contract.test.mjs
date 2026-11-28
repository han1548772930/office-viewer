import assert from 'node:assert/strict'
import { afterEach, test, vi } from 'vitest'
import { effectScope, shallowReactive, watch } from 'vue'
import { OfficeComputeClient } from '../src/compute/client.ts'
import { ComputeError } from '../src/compute/errors.ts'
import { estimateTaskMemory } from '../src/compute/tasks.ts'
import { useSheetValidationListQuery } from '../src/components/spreadsheet/use-sheet-validation-list-query.ts'
import { createSheetValidationListNavigation, sheetValidationListNavigationIndex } from '../src/components/spreadsheet/sheet-validation-list-navigation.ts'

// Lifecycle/reentrant-publication drafts. Adapters test the controller contract;
// they do not establish Rust matching, real Listy DOM or browser/IME acceptance.
afterEach(() => vi.restoreAllMocks())
const rows = labels => Object.freeze(labels.map(label => Object.freeze({ value: label, label })))
const flush = async () => { for (let index = 0; index < 12; index++) await Promise.resolve() }
const deferred = () => { let resolve, reject; const promise = new Promise((yes, no) => { resolve = yes; reject = no }); return { promise, resolve, reject } }
function adapter() {
  const requests = []
  const compute = { submit: (input, settings) => {
    const job = deferred(), record = { input, settings, job, settled: false }
    requests.push(record)
    return { requestId: `query-${requests.length}`, cancel() {}, promise: job.promise }
  } }
  function resolve(record, indices = []) {
    if (record.settled) return
    record.settled = true
    const memory = estimateTaskMemory(record.input), output = new Uint32Array(indices)
    record.job.resolve({ type: 'result', requestId: 'fixture-result', documentId: record.settings.documentId, modelRevision: record.settings.modelRevision,
      engineVersion: 'host-adapter-no-wasm', output: { kind: 'validationListQuery', indices: output, total: indices.length },
      metrics: { elapsedMs: 1, ...memory, boundary: { inputCopies: 1 + Number(record.input.textBytes.length > 0) + Number(record.input.queryBytes.length > 0),
        inputBytes: memory.inputBytes, outputCopies: Number(output.length > 0), outputBytes: output.byteLength, rustBufferClones: 0, controlDtoConversions: 2, jsonConversions: 0 },
      transport: { inputTransfers: 3, inputBytes: memory.inputBytes, outputTransfers: 1, outputBytes: output.byteLength, structuredResultMessages: 1 } } })
  }
  const reject = (record, error) => { record.settled = true; record.job.reject(error) }
  return { compute, requests, resolve, reject }
}
function host(extra = {}) {
  const state = shallowReactive({ owner: {}, options: rows(['Alpha', 'Beta', 'Gamma']), query: '', unitId: 'book', baseRevision: 4, active: true, allowed: true })
  const hooks = {}, scope = effectScope(), transport = adapter()
  const control = scope.run(() => useSheetValidationListQuery({ minOptions: 1, ...extra,
    ...(extra.owned ? {} : { compute: transport.compute }), state: () => {
      hooks.state?.()
      return { ...state, active: state.active && state.allowed, current: () => { hooks.current?.(); return state.allowed } }
    } }))
  return { state, hooks, scope, transport, control,
    latest: () => transport.requests.at(-1), resolve: indices => transport.resolve(transport.requests.at(-1), indices),
    dispose: async () => { scope.stop(); for (const record of transport.requests) transport.resolve(record); await flush() } }
}

test('a query change cancels its predecessor, reuses packed source text and publishes only the current complete window', async () => {
  const h = host()
  try {
    await flush(); const first = h.latest(); assert.equal(h.control.busy.value, true)
    h.state.query = 'beta'; await flush(); const second = h.latest()
    assert.notEqual(second, first); assert.equal(first.settings.signal.aborted, true); assert.equal(second.settings.signal.aborted, false)
    assert.equal(first.input.textBytes, second.input.textBytes); assert.equal(first.input.offsets, second.input.offsets)
    h.transport.resolve(first, [0, 1, 2]); await flush(); assert.deepEqual(h.control.matches.value, []); assert.equal(h.control.busy.value, true)
    h.transport.resolve(second, [1]); await flush()
    assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta']); assert.equal(h.control.busy.value, false); assert.equal(h.control.error.value, '')
  } finally { await h.dispose() }
})

test('the same immutable source/query keeps result identity across revisions and a changed source receives fresh packed bytes', async () => {
  const h = host()
  try {
    await flush(); const first = h.latest(); h.resolve([0, 1, 2]); await flush(); const matches = h.control.matches.value
    h.state.baseRevision++; await flush(); assert.equal(h.transport.requests.length, 1); assert.equal(h.control.matches.value, matches)
    h.state.options = rows(['New']); await flush(); const next = h.latest()
    assert.notEqual(next.input.textBytes, first.input.textBytes); assert.equal(next.settings.modelRevision, 5)
    h.resolve([0]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.label), ['New'])
  } finally { await h.dispose() }
})

test.each(['owner', 'unit', 'revision', 'active'])('a %s change away and back cannot revive the old task even when all final fields match', async change => {
  const h = host()
  try {
    await flush(); const first = h.latest()
    if (change === 'owner') { const original = h.state.owner; h.state.owner = {}; h.state.owner = original }
    if (change === 'unit') { h.state.unitId = 'other'; h.state.unitId = 'book' }
    if (change === 'revision') { h.state.baseRevision = 5; h.state.baseRevision = 4 }
    if (change === 'active') { h.state.active = false; h.state.active = true }
    await flush(); const next = h.latest(); assert.notEqual(next, first); assert.equal(first.settings.signal.aborted, true)
    h.transport.resolve(first, [0]); await flush(); assert.deepEqual(h.control.matches.value, []); assert.equal(h.control.busy.value, true)
    h.transport.resolve(next, [1]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta'])
  } finally { await h.dispose() }
})

test('closing or losing View releases derived rows and packed source; reopening prepares a new owner without old callbacks', async () => {
  const h = host()
  try {
    await flush(); const first = h.latest(); h.resolve([0, 1, 2]); await flush()
    h.state.allowed = false; await flush(); assert.deepEqual(h.control.matches.value, []); assert.equal(h.control.busy.value, false); assert.equal(h.control.metrics.value, undefined)
    h.state.allowed = true; await flush(); const next = h.latest()
    assert.notEqual(next.input.textBytes, first.input.textBytes); h.resolve([1]); await flush()
    h.state.active = false; await flush(); assert.deepEqual(h.control.matches.value, []); assert.equal(h.control.error.value, '')
  } finally { await h.dispose() }
})

test('a recoverable failure preserves an error code and retry resolves with a fresh current result', async () => {
  const h = host()
  try {
    await flush(); const first = h.latest(); h.transport.reject(first, new ComputeError('COMPUTE_WORKER_FAILED', 'failed')); await flush()
    assert.equal(h.control.busy.value, false); assert.equal(h.control.errorCode.value, 'COMPUTE_WORKER_FAILED'); assert.deepEqual(h.control.matches.value, [])
    const retry = h.control.refresh(); await flush(); const next = h.latest(); assert.equal(next.input.textBytes, first.input.textBytes)
    h.resolve([0]); await retry; assert.deepEqual(h.control.matches.value.map(item => item.value), ['Alpha']); assert.equal(h.control.errorCode.value, undefined)
  } finally { await h.dispose() }
})

test('a success watcher starting another query cannot let the old finally erase the successor source, busy state or result', async () => {
  const h = host(); let switched = false
  try {
    await flush(); const first = h.latest()
    h.scope.run(() => watch(h.control.matches, values => {
      if (!switched && values.length) { switched = true; h.state.query = 'beta' }
    }, { flush: 'sync' }))
    h.transport.resolve(first, [0]); await flush(); const next = h.latest()
    assert.notEqual(next, first); assert.equal(next.settings.signal.aborted, false); assert.equal(h.control.busy.value, true); assert.deepEqual(h.control.matches.value, [])
    assert.equal(next.input.textBytes, first.input.textBytes)
    h.resolve([1]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta']); assert.equal(h.control.busy.value, false)
  } finally { await h.dispose() }
})

test('failure-publication reentry resolves the interrupted refresh instead of rejecting from catch or clearing the new query', async () => {
  const h = host(); let switched = false
  try {
    await flush(); const refresh = h.control.refresh(); await flush(); const failed = h.latest()
    h.scope.run(() => watch(h.control.error, value => { if (value && !switched) { switched = true; h.state.query = 'beta' } }, { flush: 'sync' }))
    h.transport.reject(failed, new ComputeError('COMPUTE_WORKER_FAILED', 'failed')); await refresh; await flush()
    const next = h.latest(); assert.notEqual(next, failed); assert.equal(next.settings.signal.aborted, false); assert.equal(h.control.error.value, ''); assert.equal(h.control.busy.value, true)
    h.resolve([1]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta'])
  } finally { await h.dispose() }
})

test('a capability getter reentering specifically during final cleanup cannot clear a later pending controller', async () => {
  const h = host(); let afterPublication = 0
  try {
    await flush(); const first = h.latest()
    h.hooks.current = () => {
      if (h.control.matches.value.length && h.control.busy.value && ++afterPublication === 2) h.state.query = 'beta'
    }
    h.transport.resolve(first, [0]); await flush(); const next = h.latest()
    assert.notEqual(next, first); assert.equal(next.settings.signal.aborted, false); assert.equal(h.control.busy.value, true)
    h.resolve([1]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta']); assert.equal(h.control.busy.value, false)
  } finally { await h.dispose() }
})

test('state getter reentry during preflight cannot publish or retain the interrupted source over a successor', async () => {
  const h = host()
  try {
    await flush(); const first = h.latest(), original = h.state.options
    h.hooks.state = () => { delete h.hooks.state; h.state.options = rows(['Interim']); h.state.options = original }
    await h.control.refresh(); await flush(); const next = h.latest()
    assert.notEqual(next, first); assert.equal(next.input.params.rowCount, 3); assert.equal(next.settings.signal.aborted, false)
    h.transport.resolve(first, [0]); await flush(); assert.deepEqual(h.control.matches.value, [])
    h.resolve([1]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta'])
  } finally { await h.dispose() }
})

test('a progress watcher changing the query cancels the old callback and never restores its rows', async () => {
  const h = host(); let switched = false
  try {
    await flush(); const first = h.latest()
    h.scope.run(() => watch(h.control.progress, value => { if (value && !switched) { switched = true; h.state.query = 'beta' } }, { flush: 'sync' }))
    assert.throws(() => first.settings.onProgress({ type: 'progress', requestId: 'progress', documentId: 'book', modelRevision: 4,
      phase: 'computing', processed: 1, total: 3, ratio: 1 / 3, unit: 'rows' }), error => error.code === 'COMPUTE_STALE_RESULT')
    await flush(); assert.equal(first.settings.signal.aborted, true); assert.equal(h.control.progress.value, undefined)
    h.transport.resolve(first, [0]); h.resolve([1]); await flush(); assert.deepEqual(h.control.matches.value.map(item => item.value), ['Beta'])
  } finally { await h.dispose() }
})

test('small queries and failures before admission do not create an owned client; scope disposal releases the one lazy client', async () => {
  const transport = adapter(), submit = vi.spyOn(OfficeComputeClient.prototype, 'submit').mockImplementation(transport.compute.submit)
  const dispose = vi.spyOn(OfficeComputeClient.prototype, 'dispose'), small = host({ owned: true, minOptions: 1024 }), limited = host({ owned: true, memoryBudgetBytes: 1 })
  let large
  try {
    await flush(); assert.equal(submit.mock.calls.length, 0); assert.equal(limited.control.errorCode.value, 'COMPUTE_MEMORY_LIMIT')
    assert.equal(small.control.matches.value.length, 3)
    large = host({ owned: true }); await flush(); assert.equal(submit.mock.calls.length, 1)
    const request = transport.requests[0]; large.scope.stop(); assert.equal(request.settings.signal.aborted, true); assert.equal(dispose.mock.calls.length, 1)
    transport.resolve(request, [0]); await flush(); assert.deepEqual(large.control.matches.value, []); assert.equal(large.control.busy.value, false)
  } finally { await small.dispose(); await limited.dispose(); await large?.dispose(); for (const record of transport.requests) transport.resolve(record) }
})

test.each([{ minOptions: 0 }, { minTextUnits: -1 }, { minOptions: 1.5 }])('invalid provisional path threshold %j cannot start a task', config => {
  const scope = effectScope()
  try { assert.throws(() => scope.run(() => useSheetValidationListQuery({ ...config, state: () => ({}) })), RangeError) }
  finally { scope.stop() }
})

function navigationHost() {
  const state = { owner: {}, items: rows(Array.from({ length: 1000 }, (_, index) => `Item ${index}`)), query: '', unitId: 'book', baseRevision: 4,
    ready: true, current: () => true }, frames = [], activated = [], scrolled = [], focused = []
  let mounted = -1, hooks = {}, navigation
  navigation = createSheetValidationListNavigation({ state: () => { hooks.state?.(); return { ...state } },
    activate: index => { activated.push(index); hooks.activate?.() }, scrollTo: (item, align) => { scrolled.push([item.value, align]); hooks.scroll?.() },
    nextFrame: () => { const frame = deferred(); frames.push(frame); return frame.promise },
    focusAt: (index, current) => { hooks.lookup?.(); if (index !== mounted || !current()) return false; focused.push(index); return true } })
  return { state, frames, activated, scrolled, focused, hooks, navigation, mount: index => { mounted = index } }
}

test.each([['ArrowDown', -1, 0], ['ArrowUp', -1, 999], ['Home', 200, 0], ['End', 200, 999], ['ArrowDown', 999, 999], ['ArrowUp', 0, 0],
  ['ArrowDown', 11, 12], ['ArrowUp', 11, 10]])('%s uses the full logical result rather than only mounted checkboxes', (key, index, expected) => {
  assert.equal(sheetValidationListNavigationIndex(key, index, 1000), expected)
})

test('End scrolls to the full last key, waits for dynamic row measurement and focuses only its mounted Ant input', async () => {
  const h = navigationHost(), job = h.navigation.focus(999, 'bottom')
  assert.deepEqual(h.scrolled, [['Item 999', 'bottom']]); h.frames[0].resolve(); await flush()
  assert.deepEqual(h.focused, []); h.mount(999); h.frames[1].resolve(); assert.equal(await job, true); assert.deepEqual(h.focused, [999])
  h.navigation.dispose()
})

test.each(['owner', 'items', 'query', 'revision', 'unit', 'ime', 'dispose'])('%s invalidation across a virtual frame cannot revive an old focus request after values return', async change => {
  const h = navigationHost(), before = { ...h.state }, job = h.navigation.focus(999)
  if (change === 'owner') h.state.owner = {}
  if (change === 'items') h.state.items = rows(['Other'])
  if (change === 'query') h.state.query = 'other'
  if (change === 'revision') h.state.baseRevision = 5
  if (change === 'unit') h.state.unitId = 'other'
  if (change === 'ime') h.state.ready = false
  if (change === 'dispose') h.navigation.dispose()
  else { h.navigation.invalidate(); Object.assign(h.state, before); h.navigation.invalidate() }
  h.mount(999); h.frames[0].resolve(); assert.equal(await job, false); assert.deepEqual(h.focused, []); h.navigation.dispose()
})

test.each(['state', 'capability', 'activate', 'scroll', 'lookup'])('reentry during %s is checked before the next scroll/frame/focus boundary', async boundary => {
  const h = navigationHost()
  h.hooks[boundary] = () => { delete h.hooks[boundary]; h.state.query = 'other'; h.navigation.invalidate() }
  if (boundary === 'capability') h.state.current = () => { h.hooks.capability?.(); return true }
  const job = h.navigation.focus(999)
  if (h.frames.length) { h.mount(999); h.frames[0].resolve() }
  assert.equal(await job, false); assert.deepEqual(h.focused, [])
  if (boundary === 'state' || boundary === 'capability' || boundary === 'activate') assert.deepEqual(h.scrolled, [])
  h.navigation.dispose()
})

test('rapid navigation keeps only the latest logical target and has a bounded missing-row retry', async () => {
  const h = navigationHost(), first = h.navigation.focus(12), latest = h.navigation.focus(999, 'bottom')
  h.mount(999); h.frames[0].resolve(); h.frames[1].resolve(); assert.equal(await first, false); assert.equal(await latest, true); assert.deepEqual(h.focused, [999])
  h.mount(-1); const missing = h.navigation.focus(1)
  for (let index = 2; index < 6; index++) { h.frames[index].resolve(); await flush() }
  assert.equal(await missing, false); assert.deepEqual(h.focused, [999]); h.navigation.dispose()
})
