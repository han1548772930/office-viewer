import assert from 'node:assert/strict'
import { afterEach, test, vi } from 'vitest'
import { getRecordValueKey } from '../src/domain/record-value.ts'
import { SlicerCanvasDataCache, sameSlicerCanvasDataOwner } from '../src/features/slicer/slicer-canvas-data.ts'

// Deferred boundary/ownership contracts. Controlled loader promises test the cache,
// not membership algorithms or WASM; real service/browser drafts are separate.
const caches = new Set()
afterEach(() => { for (const cache of caches) cache.dispose(); caches.clear(); vi.useRealTimers() })
const flush = async () => { for (let index = 0; index < 8; index++) await Promise.resolve() }
const member = (index, caption = String(index)) => {
  const value = { type: 'number', value: index }
  return { key: getRecordValueKey(value), value, label: caption, hasData: true, selected: true, partial: false }
}
const window = (start, count, total = 40, prefix = '') => ({ total,
  items: Array.from({ length: Math.max(0, Math.min(count, total - start)) }, (_, index) => member(start + index, prefix + (start + index))) })
const metadata = (name = 'Date') => ({ fieldId: 'date-source', name, ready: true,
  dateGroup: { type: 3, dateType: 'month', excelLike: { unit: 'month', autoStart: true, autoEnd: false } },
  timeline: { minDate: -2.25, maxDate: 45000.75, interactive: true } })
function pendingLoader() {
  const calls = []
  const add = (kind, start, count, signal) => new Promise((resolve, reject) => calls.push({ kind, start, count, signal, resolve, reject }))
  return { calls, members: (start, count, signal) => add('members', start, count, signal),
    metadata: signal => add('metadata', undefined, undefined, signal), assertCurrent() {} }
}
function fixture(options = {}) {
  const live = { instance: {}, model: {}, unitId: 'unit', sheetId: 'host', revision: 7, historyGeneration: 2, userId: 'owner', locale: 'en-US' }
  function owner() {
    const captured = { ...live }
    return { ...captured, assertCurrent() {
      if (Object.keys(captured).some(key => captured[key] !== live[key])) throw new Error('owner invalidated')
    } }
  }
  const cache = new SlicerCanvasDataCache({ pageSize: 2, maxCachedMembers: 8, ...options })
  caches.add(cache); cache.reset(owner())
  return { cache, live, owner }
}

test.each(['instance', 'model', 'unitId', 'sheetId', 'revision', 'historyGeneration', 'userId', 'locale'])(
  'changing %s rejects the old result and immediately admits a new owner', async key => {
    const f = fixture({ maxConcurrent: 1 }), loader = pendingLoader()
    f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
    const old = loader.calls[0]
    f.live[key] = typeof f.live[key] === 'number' ? f.live[key] + 1 : typeof f.live[key] === 'string' ? 'changed' : {}
    f.cache.reset(f.owner()); f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
    assert.equal(old.signal.aborted, true); assert.equal(loader.calls.length, 2)
    old.resolve(window(0, 2, 10, 'old')); await flush()
    assert.equal(f.cache.state('s').pending, true); assert.equal(f.cache.read('s', 0, 2), undefined)
    loader.calls[1].resolve(window(0, 2, 10, 'new')); await flush()
    assert.deepEqual(f.cache.read('s', 0, 2).items.map(item => item.label), ['new0', 'new1'])
  })

test('owner equality includes the complete tuple and reset snapshots a caller-owned object', async () => {
  const f = fixture({ maxConcurrent: 1 }), loader = pendingLoader(), captured = { ...f.live, assertCurrent() {} }
  assert.equal(sameSlicerCanvasDataOwner(captured, { ...captured }), true)
  assert.equal(sameSlicerCanvasDataOwner(captured, { ...captured, historyGeneration: 3 }), false)
  f.cache.reset(captured); f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
  captured.historyGeneration++
  f.cache.reset(captured); f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
  assert.equal(loader.calls[0].signal.aborted, true); assert.equal(loader.calls.length, 2)
  loader.calls[0].resolve(window(0, 2, 10, 'old')); loader.calls[1].resolve(window(0, 2, 10, 'new')); await flush()
  assert.equal(f.cache.read('s', 0, 2).items[0].label, 'new0')
})

test('same-id loader replacement cancels an abort-ignoring promise without blocking the next loader', async () => {
  const f = fixture({ maxConcurrent: 1 }), old = pendingLoader(), current = pendingLoader()
  f.cache.bind('s', 'slicer', old); f.cache.request('s', 6, 2)
  f.cache.bind('s', 'slicer', current); f.cache.request('s', 6, 2)
  assert.equal(old.calls[0].signal.aborted, true); assert.equal(current.calls.length, 1)
  current.calls[0].resolve(window(6, 2, 20, 'new')); await flush()
  old.calls[0].resolve(window(6, 2, 20, 'old')); await flush()
  assert.deepEqual(f.cache.read('s', 6, 2).items.map(item => item.label), ['new6', 'new7'])
  assert.equal(f.cache.state('s').error, undefined)
})

test('repeated stable bind/request does not replace in-flight work or generate an update loop', async () => {
  let changes = 0
  const f = fixture({ onChange: () => changes++ }), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
  assert.equal(loader.calls.length, 1); assert.equal(loader.calls[0].signal.aborted, false)
  loader.calls[0].resolve(window(0, 2)); await flush()
  const settled = changes
  for (let index = 0; index < 20; index++) { f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2) }
  assert.equal(changes, settled); assert.equal(loader.calls.length, 1)
})

test('retry retains an explicit scroll window and a stale finally cannot clear the replacement page token', async () => {
  const f = fixture({ maxConcurrent: 1 }), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 8, 2)
  f.cache.retry('s')
  assert.deepEqual(loader.calls.map(call => call.start), [8, 8]); assert.equal(loader.calls[0].signal.aborted, true)
  loader.calls[0].resolve(window(8, 2, 20, 'old')); await flush()
  assert.equal(f.cache.state('s').pending, true)
  loader.calls[1].resolve(window(8, 2, 20, 'new')); await flush()
  assert.equal(f.cache.state('s').pending, false); assert.equal(f.cache.read('s', 8, 2).items[0].label, 'new8')
})

test('a replacement scroll cancels obsolete running and queued pages and loads only the new window', async () => {
  const f = fixture({ maxConcurrent: 1 }), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 6); f.cache.request('s', 10, 2)
  assert.deepEqual(loader.calls.map(call => call.start), [0, 10]); assert.equal(loader.calls[0].signal.aborted, true)
  loader.calls[1].resolve(window(10, 2)); loader.calls[0].resolve(window(0, 2)); await flush()
  assert.deepEqual(loader.calls.map(call => call.start), [0, 10])
  assert.equal(f.cache.read('s', 0, 2), undefined); assert.equal(f.cache.read('s', 10, 2).items[0].label, '10')
})

test('overlapping scroll preserves useful queued pages while removing the page outside the window', async () => {
  const f = fixture({ maxConcurrent: 1 }), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 6); f.cache.request('s', 2, 4)
  assert.deepEqual(loader.calls.map(call => call.start), [0, 2])
  loader.calls[1].resolve(window(2, 2)); await flush()
  assert.deepEqual(loader.calls.map(call => call.start), [0, 2, 4])
  loader.calls[2].resolve(window(4, 2)); await flush()
  assert.deepEqual(f.cache.read('s', 2, 4).items.map(item => item.label), ['2', '3', '4', '5'])
})

test('retain removal releases a slot and never allows a removed control to publish', async () => {
  const f = fixture({ maxConcurrent: 1 }), first = pendingLoader(), second = pendingLoader()
  f.cache.bind('a', 'slicer', first); f.cache.bind('b', 'slicer', second)
  f.cache.request('a', 0, 2); f.cache.request('b', 0, 2); assert.equal(second.calls.length, 0)
  f.cache.retain(new Set(['b'])); assert.equal(first.calls[0].signal.aborted, true); assert.equal(second.calls.length, 1)
  first.calls[0].resolve(window(0, 2)); second.calls[0].resolve(window(0, 2)); await flush()
  assert.deepEqual(f.cache.state('a'), { pending: false }); assert.equal(f.cache.read('b', 0, 2).items.length, 2)
})

test('scope disposal clears timers, cancels metadata and rejects late pages and rebind attempts', async () => {
  vi.useFakeTimers()
  const f = fixture(), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.bind('t', 'timeline', loader)
  f.cache.request('s', 0, 2); f.cache.request('t')
  f.cache.dispose(); f.cache.dispose()
  assert.equal(vi.getTimerCount(), 0); assert.equal(loader.calls.every(call => call.signal.aborted), true)
  for (const call of loader.calls) call.resolve(call.kind === 'metadata' ? metadata() : window(0, 2))
  await flush(); f.cache.reset(f.owner()); f.cache.bind('s', 'slicer', loader); f.cache.request('s')
  assert.deepEqual(f.cache.state('s'), { pending: false }); assert.equal(loader.calls.length, 2)
})

test('an owner guard failure on read cancels and drops data even before the host explicitly resets', async () => {
  const f = fixture(), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
  f.live.historyGeneration++
  assert.deepEqual(f.cache.state('s'), { pending: false }); assert.equal(loader.calls[0].signal.aborted, true)
  loader.calls[0].resolve(window(0, 2)); await flush(); assert.equal(f.cache.read('s', 0, 2), undefined)
})

test('revoked read permission cancels queries, clears cached captions and requires an explicit retry', async () => {
  const f = fixture(), loader = pendingLoader(); let allowed = true
  loader.assertCurrent = () => { if (!allowed) throw Object.assign(new Error('read permission revoked'), { code: 'permission-denied' }) }
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 2)
  allowed = false; assert.equal(f.cache.state('s').errorCode, 'permission-denied')
  assert.equal(loader.calls[0].signal.aborted, true); assert.equal(f.cache.read('s', 0, 2), undefined)
  loader.calls[0].resolve(window(0, 2)); await flush()
  allowed = true; f.cache.request('s', 0, 2); assert.equal(loader.calls.length, 1)
  f.cache.retry('s'); loader.calls[1].resolve(window(0, 2)); await flush()
  assert.equal(f.cache.state('s').error, undefined); assert.equal(f.cache.read('s', 0, 2).items.length, 2)
})

test('a failed sibling page cancels other work and cannot expose a partial successful window', async () => {
  const f = fixture(), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 4)
  loader.calls[0].resolve(window(0, 2)); await flush()
  loader.calls[1].reject(new Error('query crashed')); await flush()
  assert.equal(f.cache.state('s').error, 'query crashed'); assert.equal(f.cache.state('s').total, undefined)
  assert.equal(f.cache.read('s', 0, 2), undefined)
  f.cache.retry('s'); assert.deepEqual(loader.calls.slice(2).map(call => call.start), [0, 2])
})

test.each(['total', 'caption'])('inconsistent %s between pages fails the entire entry', async kind => {
  const f = fixture(), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 0, 4)
  loader.calls[0].resolve(window(0, 2, 10)); await flush()
  const result = window(2, 2, kind === 'total' ? 11 : 10)
  if (kind === 'caption') result.items[0].label = '0'
  loader.calls[1].resolve(result); await flush()
  assert.equal(f.cache.state('s').errorCode, 'SLICER_CANVAS_INVALID_RESULT'); assert.equal(f.cache.read('s', 0, 2), undefined)
})

test('LRU removes the least recently read page outside the new window', async () => {
  const f = fixture({ maxCachedMembers: 4 }), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader)
  f.cache.request('s', 0, 2); loader.calls[0].resolve(window(0, 2)); await flush()
  f.cache.request('s', 2, 2); loader.calls[1].resolve(window(2, 2)); await flush()
  assert.ok(f.cache.read('s', 0, 2))
  f.cache.request('s', 4, 2); loader.calls[2].resolve(window(4, 2)); await flush()
  assert.ok(f.cache.read('s', 0, 2)); assert.equal(f.cache.read('s', 2, 2), undefined); assert.ok(f.cache.read('s', 4, 2))
})

test('an older visible window in another control remains pinned when the first control scrolls', async () => {
  const f = fixture({ maxCachedMembers: 4 }), a = pendingLoader(), b = pendingLoader()
  f.cache.bind('a', 'slicer', a); f.cache.bind('b', 'slicer', b)
  f.cache.request('a', 0, 2); f.cache.request('b', 0, 2)
  b.calls[0].resolve(window(0, 2)); a.calls[0].resolve(window(0, 2)); await flush()
  f.cache.request('a', 2, 2); a.calls[1].resolve(window(2, 2)); await flush()
  assert.ok(f.cache.read('b', 0, 2)); assert.ok(f.cache.read('a', 2, 2)); assert.equal(f.cache.read('a', 0, 2), undefined)
})

test('too many simultaneous pinned windows fail visibly and retry succeeds after another control is removed', async () => {
  const f = fixture({ maxCachedMembers: 2 }), a = pendingLoader(), b = pendingLoader()
  f.cache.bind('a', 'slicer', a); f.cache.bind('b', 'slicer', b)
  f.cache.request('a', 0, 2); a.calls[0].resolve(window(0, 2)); await flush()
  f.cache.request('b', 0, 2); b.calls[0].resolve(window(0, 2)); await flush()
  assert.equal(f.cache.state('b').errorCode, 'SLICER_CANVAS_MEMORY_LIMIT'); assert.ok(f.cache.read('a', 0, 2))
  f.cache.retain(new Set(['b'])); f.cache.retry('b'); b.calls[1].resolve(window(0, 2)); await flush()
  assert.ok(f.cache.read('b', 0, 2))
})

test('metadata is included in the shared payload budget and released on removal', async () => {
  const f = fixture({ memoryBudgetBytes: 1400 }), loader = pendingLoader()
  f.cache.bind('t1', 'timeline', loader); f.cache.bind('t2', 'timeline', loader); f.cache.bind('s', 'slicer', loader)
  f.cache.request('t1'); f.cache.request('t2')
  loader.calls[0].resolve(metadata()); loader.calls[1].resolve(metadata()); await flush()
  assert.ok(f.cache.state('t1').metadata); assert.ok(f.cache.state('t2').metadata)
  f.cache.request('s', 0, 2); loader.calls[2].resolve(window(0, 2)); await flush()
  assert.equal(f.cache.state('s').errorCode, 'SLICER_CANVAS_MEMORY_LIMIT')
  f.cache.retain(new Set(['s'])); f.cache.retry('s'); loader.calls[3].resolve(window(0, 2)); await flush()
  assert.ok(f.cache.read('s', 0, 2))
})

test('oversized captions and metadata cannot bypass the cache memory admission', async () => {
  for (const kind of ['slicer', 'timeline']) {
    const f = fixture({ memoryBudgetBytes: 500 }), loader = pendingLoader()
    f.cache.bind('s', kind, loader); f.cache.request('s', 0, 2)
    loader.calls[0].resolve(kind === 'timeline' ? metadata('x'.repeat(1000)) : window(0, 2, 10, 'x'.repeat(1000)))
    await flush(); assert.equal(f.cache.state('s').errorCode, 'SLICER_CANVAS_MEMORY_LIMIT')
    assert.equal(f.cache.state('s').metadata, undefined); assert.equal(f.cache.state('s').total, undefined)
  }
})

test('a timeline without a reliable range state fails rather than being silently treated as empty', async () => {
  const f = fixture(), loader = pendingLoader()
  f.cache.bind('t', 'timeline', loader); f.cache.request('t')
  loader.calls[0].resolve({ fieldId: 'date', name: 'Date', ready: true }); await flush()
  assert.equal(f.cache.state('t').errorCode, 'SLICER_CANVAS_INVALID_RESULT')
})

test('callers cannot mutate retained member values, arrays or nested timeline metadata', async () => {
  const f = fixture(), loader = pendingLoader(), raw = window(0, 2), meta = metadata()
  f.cache.bind('s', 'slicer', loader); f.cache.bind('t', 'timeline', loader)
  f.cache.request('s', 0, 2); f.cache.request('t')
  loader.calls[0].resolve(raw); loader.calls[1].resolve(meta); await flush()
  raw.items[0].label = 'caller'; raw.items[1].value.value = 99; meta.timeline.minDate = 100
  const read = f.cache.read('s', 0, 2), state = f.cache.state('t')
  read.items[0].label = 'reader'; read.items[1].value.value = 100; read.items.pop()
  state.metadata.timeline.maxDate = 0; state.metadata.dateGroup.excelLike.autoStart = false
  assert.deepEqual(f.cache.read('s', 0, 2).items.map(item => [item.label, item.value.value]), [['0', 0], ['1', 1]])
  assert.equal(f.cache.state('t').metadata.timeline.minDate, -2.25)
  assert.equal(f.cache.state('t').metadata.timeline.maxDate, 45000.75)
  assert.equal(f.cache.state('t').metadata.dateGroup.excelLike.autoStart, true)
})

test('queue pressure fails one entry without cancelling another control and without automatic retry', async () => {
  const f = fixture({ maxConcurrent: 1, maxQueuedTasks: 1 }), a = pendingLoader(), b = pendingLoader()
  f.cache.bind('a', 'slicer', a); f.cache.bind('b', 'slicer', b)
  f.cache.request('a', 0, 2); f.cache.request('b', 0, 4)
  assert.equal(a.calls[0].signal.aborted, false); assert.equal(b.calls.length, 0)
  assert.equal(f.cache.state('b').errorCode, 'SLICER_CANVAS_QUEUE_FULL')
  f.cache.request('b', 0, 2); assert.equal(b.calls.length, 0)
  a.calls[0].resolve(window(0, 2)); await flush()
  f.cache.request('b', 0, 2); assert.equal(b.calls.length, 0)
})

test('queued deadlines count from admission and timed-out work immediately frees a slot', async () => {
  vi.useFakeTimers()
  const f = fixture({ maxConcurrent: 1, timeoutMs: 20 }), a = pendingLoader(), b = pendingLoader()
  f.cache.bind('a', 'slicer', a); f.cache.bind('b', 'slicer', b)
  f.cache.request('a', 0, 2); await vi.advanceTimersByTimeAsync(5); f.cache.request('b', 0, 2)
  await vi.advanceTimersByTimeAsync(15)
  assert.equal(f.cache.state('a').errorCode, 'SLICER_CANVAS_TIMEOUT'); assert.equal(a.calls[0].signal.aborted, true)
  assert.equal(b.calls.length, 1)
  b.calls[0].resolve(window(0, 2)); await flush(); assert.ok(f.cache.read('b', 0, 2))
  a.calls[0].resolve(window(0, 2)); await flush(); assert.equal(f.cache.read('a', 0, 2), undefined)
  f.cache.dispose(); assert.equal(vi.getTimerCount(), 0)
})

test('tail and empty windows are published with the actual total and a final change notification', async () => {
  let changes = 0
  const f = fixture({ onChange: () => changes++ }), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader); f.cache.request('s', 3, 2)
  loader.calls[0].resolve(window(2, 2, 5)); loader.calls[1].resolve(window(4, 2, 5)); await flush()
  assert.deepEqual(f.cache.read('s', 3, 2).items.map(item => item.label), ['3', '4'])
  f.cache.bind('empty', 'slicer', loader); f.cache.request('empty', 20, 2)
  const before = changes
  loader.calls[2].resolve(window(20, 2, 0)); await flush()
  assert.equal(changes > before, true); assert.deepEqual(f.cache.read('empty', 0, 2), { start: 0, total: 0, items: [] })
  assert.equal(f.cache.state('empty').pending, false)
})

test.each([{ pageSize: 201 }, { maxConcurrent: 33 }, { maxCachedMembers: 1, pageSize: 2 },
  { maxQueuedTasks: 0 }, { memoryBudgetBytes: 0 }, { timeoutMs: Infinity }])(
  'invalid budget %j is rejected without scheduling', options => assert.throws(() => new SlicerCanvasDataCache(options), /预算/))

test('invalid or over-budget windows cannot trigger an unbounded queue or read allocation', () => {
  const f = fixture(), loader = pendingLoader()
  f.cache.bind('s', 'slicer', loader)
  f.cache.request('s', Number.MAX_SAFE_INTEGER, 2)
  assert.equal(loader.calls.length, 0); assert.equal(f.cache.state('s').errorCode, 'SLICER_CANVAS_INVALID_WINDOW')
  assert.throws(() => f.cache.read('s', 0, Number.MAX_SAFE_INTEGER), /窗口/)
})
