import { getSlicerDisplayMembers, SlicerMemberQueryService } from '../../packages/office-studio/src/features/slicer/slicer-member-query.ts'
import { getSlicerSourceCellRange } from '../../packages/office-studio/src/features/slicer/slicer-connections.ts'
import { sheetCellRecordValue } from '../../packages/office-studio/src/features/sheet-workbench/sheet-record-value.ts'
import { readSlicerItemState } from '../../packages/office-studio/src/features/slicer/slicer-engine.ts'
import { generateSlicerWorkerRows, fingerprintSlicerWorkerRows, SLICER_WORKER_SEED, SLICER_WORKER_DATASET_VERSION,
  slicerWorkerTableWorkbook } from '../../packages/office-studio/tests/fixtures/slicer-member-worker-data.mjs'
import { observedSlicerWorkerClient, createSlicerMemberWorkerRuntime } from '../../packages/office-studio/tests/fixtures/slicer-member-worker-runtime.mjs'

// Reproducible benchmark draft only. No timings have been collected. The JS
// baseline is explicit and never injected into the production query service.
const nextTurn = () => new Promise(resolve => setTimeout(resolve, 0))
const nextFrame = () => new Promise(resolve => requestAnimationFrame(resolve))
const sourceOptions = { sortOrder: 'descending', customListSort: true,
  crossFilter: 'showItemsWithDataAtTop', hideItemsWithNoData: false }

function percentile(values, fraction) {
  if (!values.length) return null
  const sorted = [...values].sort((left, right) => left - right)
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * fraction))]
}
function summary(values) { return { count: values.length, p50Ms: percentile(values, .5), p95Ms: percentile(values, .95), maxMs: values.length ? Math.max(...values) : null } }
function heapBytes() {
  const value = performance.memory?.usedJSHeapSize
  return Number.isFinite(value) ? value : null
}
async function userAgentMemory() {
  if (!crossOriginIsolated || typeof performance.measureUserAgentSpecificMemory !== 'function') return null
  try { const memory = await performance.measureUserAgentSpecificMemory(); return { bytes: memory.bytes, breakdown: memory.breakdown } }
  catch (error) { return { unavailableReason: String(error) } }
}
function equalWindow(first, second) { return JSON.stringify(first) === JSON.stringify(second) }
function normalizeValue(mode, value) { return mode !== 'metadata' ? value : { timeline: value.timeline ?? null } }
function baselineQuery(workbook, slicer, mode, request, locale) {
  if (mode !== 'metadata') {
    const projected = getSlicerDisplayMembers(workbook, slicer, locale)
    if (mode === 'canvas') {
      const state = readSlicerItemState(workbook,slicer), checked = new Set(state.checkedItems), partial = new Set(state.partialItems), captions = new Map()
      for (const member of projected) {
        const previous = captions.get(member.label)
        if (previous) previous.hasData ||= member.hasData
        else captions.set(member.label,{ ...member,selected:checked.has(member.label),partial:partial.has(member.label) })
      }
      const folded = [...captions.values()]
      return { total:folded.length,items:folded.slice(request.start,request.start+request.count) }
    }
    return { total: projected.length, items: projected.slice(request.start, request.start + request.count)
      .map(({ key, value, hasData }) => ({ key, value, hasData })) }
  }
  const { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer)
  let minDate = null, maxDate = null
  for (let row = startRow; row <= endRow; row++) {
    const value = sheetCellRecordValue(sheet.cells[`${row}:${column}`], workbook.date1904)
    if (value.type !== 'date') continue
    minDate = minDate === null ? value.value : Math.min(minDate, value.value)
    maxDate = maxDate === null ? value.value : Math.max(maxDate, value.value)
  }
  return { timeline: minDate === null ? null : { minDate, maxDate, interactive: true } }
}

function frameProbe() {
  const intervals = []; let previous, handle, stopped = false, ready
  const first = new Promise(resolve => { ready = resolve })
  const frame = time => {
    if (stopped) return
    if (previous !== undefined) intervals.push({ start: previous, end: time, duration: time - previous })
    else ready()
    previous = time; handle = requestAnimationFrame(frame)
  }
  handle = requestAnimationFrame(frame)
  return { first, stop(start, end) {
    stopped = true; cancelAnimationFrame(handle)
    return summary(intervals.filter(interval => interval.start <= end && interval.end >= start).map(interval => interval.duration))
  } }
}

async function measure(mode, regime, rowCount, sample, expected, work, longTasks) {
  await nextTurn()
  const frames = frameProbe(); await frames.first
  const beforeHeap = heapBytes(), startedAt = performance.now()
  let result, completedAt
  try { result = await work(); completedAt = performance.now() }
  catch (error) { frames.stop(startedAt, performance.now()); throw error }
  // Include the first frame after a synchronous baseline block, and allow the
  // PerformanceObserver to receive the operation's completed long-task entry.
  await nextFrame(); await nextTurn()
  const observedFrames = frames.stop(startedAt, completedAt)
  const observedLongTasks = longTasks.filter(task => task.startTime <= completedAt && task.startTime + task.duration >= startedAt)
  const actual = normalizeValue(mode, result.value), parity = equalWindow(actual, expected)
  if (!parity) throw new Error(`Slicer benchmark parity mismatch: ${rowCount}/${mode}/${regime}/${sample}`)
  const afterHeap = heapBytes(), metrics = result.metrics
  return { mode, regime, rowCount, sample, wallMs: completedAt - startedAt, parity,
    returnedMembers: mode !== 'metadata' ? actual.items.length : 0, totalMembers: mode !== 'metadata' ? actual.total : null,
    frameIntervals: observedFrames, longTasks: { count: observedLongTasks.length,
      totalMs: observedLongTasks.reduce((total, task) => total + task.duration, 0), maxMs: observedLongTasks.length ? Math.max(...observedLongTasks.map(task => task.duration)) : 0 },
    mainJSHeapBeforeBytes: beforeHeap, mainJSHeapAfterBytes: afterHeap,
    mainJSHeapDeltaBytes: beforeHeap !== null && afterHeap !== null ? afterHeap - beforeHeap : null,
    sourcePreparationMs: metrics?.sourcePreparationMs ?? null, queryTotalMs: metrics?.totalMs ?? null,
    sourcePackedBytes: metrics?.sourcePackedBytes ?? null, engineVersion: result.engineVersion ?? null,
    compute: metrics?.compute ?? null,
    workerLinearMemoryBytes: null, workerResidentMemoryBytes: null }
}

export async function runSlicerMemberBenchmark({ sizes = [4096, 16384, 65536], iterations = 3, locale = 'en-US' } = {}) {
  if (!Array.isArray(sizes) || !sizes.length || sizes.some(size => !Number.isSafeInteger(size) || size < 1 || size > 1_000_000)
    || !Number.isSafeInteger(iterations) || iterations < 1 || iterations > 20 || typeof locale !== 'string' || !locale.trim()) throw new TypeError('Invalid slicer benchmark configuration')
  const longTasks = [], longTaskObserverAvailable = typeof PerformanceObserver !== 'undefined' && PerformanceObserver.supportedEntryTypes?.includes('longtask') === true
  let observer
  if (longTaskObserverAvailable) {
    observer = new PerformanceObserver(list => { for (const entry of list.getEntries()) longTasks.push({ startTime: entry.startTime, duration: entry.duration }) })
    observer.observe({ type: 'longtask', buffered: false })
  }
  const environment = { userAgent: navigator.userAgent, browserBrands: navigator.userAgentData?.brands ?? null,
    hardwareConcurrency: navigator.hardwareConcurrency, deviceMemoryGiB: navigator.deviceMemory ?? null,
    viewport: { width: innerWidth, height: innerHeight, devicePixelRatio }, crossOriginIsolated,
    secureContext: isSecureContext, locale, timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    longTaskObserverAvailable, mainJSHeapBeforeBytes: heapBytes(), userAgentMemoryBefore: await userAgentMemory(),
    renderingBackend: 'requestAnimationFrame probe; no office Canvas mounted' }
  const workloads = []
  try {
    for (const rowCount of sizes) {
      const generated = generateSlicerWorkerRows(rowCount), { workbook, slicer } = slicerWorkerTableWorkbook(generated.rows, generated.availability, sourceOptions)
      const request = { start: 17, count: 200 }, samples = []
      const owner = { unitId: workbook.id, baseRevision: 0, locale, assertCurrent() {} }
      const observed = observedSlicerWorkerClient(); let lastMetrics
      const reused = new SlicerMemberQueryService({ computeClient: observed.client, onMetrics: value => { lastMetrics = value } })
      const call = (service, mode) => mode === 'items' ? service.getItems(workbook, slicer, request, owner)
        : mode === 'canvas' ? service.getDisplayMemberWindow(workbook,slicer,request,owner) : service.getSourceMetadata(workbook,slicer,owner)
      try {
        // All three modes warm up the same real client/module; discarded
        // warmup observations never enter ordinary sample statistics.
        await call(reused,'items'); await call(reused,'canvas'); await call(reused,'metadata')
        for (const mode of ['items','canvas','metadata']) {
          const expected = baselineQuery(workbook, slicer, mode, request, locale)
          for (let sample = 0; sample < iterations; sample++) {
            const regimes = ['main-thread-js', 'new-worker-wasm', 'reused-worker-wasm']
            const order = [...regimes.slice(sample % regimes.length), ...regimes.slice(0, sample % regimes.length)]
            for (const regime of order) {
              samples.push(await measure(mode, regime, rowCount, sample, expected, async () => {
                if (regime === 'main-thread-js') return { value: baselineQuery(workbook, slicer, mode, request, locale) }
                if (regime === 'reused-worker-wasm') {
                  lastMetrics = undefined; const value = await call(reused, mode)
                  return { value, metrics: lastMetrics, engineVersion: observed.events.findLast(event => event.data.type === 'result')?.data.engineVersion }
                }
                const fresh = observedSlicerWorkerClient(); let metrics
                const service = new SlicerMemberQueryService({ computeClient: fresh.client, onMetrics: value => { metrics = value } })
                try {
                  const value = await call(service, mode)
                  return { value, metrics, engineVersion: fresh.events.findLast(event => event.data.type === 'result')?.data.engineVersion }
                } finally { service.dispose(); fresh.client.dispose() }
              }, longTasks))
            }
          }
        }
      } finally { reused.dispose(); observed.client.dispose() }
      workloads.push({ rowCount, seed: generated.seed, datasetVersion: generated.version,
        fingerprint: fingerprintSlicerWorkerRows(generated.rows, generated.availability), request, sourceOptions,
        reusedWorkerCount: observed.workers.length, samples,
        summaries: ['items','canvas','metadata'].flatMap(mode => ['main-thread-js', 'new-worker-wasm', 'reused-worker-wasm'].map(regime => ({
          mode, regime, wall: summary(samples.filter(sample => sample.mode === mode && sample.regime === regime).map(sample => sample.wallMs)) }))) })
    }
    const cancellationRuntime = createSlicerMemberWorkerRuntime(); let cancellation
    try { cancellation = await cancellationRuntime.directCancellation() }
    finally { cancellationRuntime.dispose() }
    return { schemaVersion: 1, generatedAt: new Date().toISOString(), environment: { ...environment,
      mainJSHeapAfterBytes: heapBytes(), userAgentMemoryAfter: await userAgentMemory() },
      methodology: { datasetVersion: SLICER_WORKER_DATASET_VERSION, seed: SLICER_WORKER_SEED, iterations,
        baseline: 'Items: existing complete main-thread typed scan/Map/dedup/Intl/sort/availability projection, then bounded slice. Canvas: that projection plus the original all-source checked/partial scan and global caption fold, then slice. Metadata: explicit same-typed-Date source scan.',
        worker: 'Real service includes source read, availability predicate, dictionary packing, queue/startup, real Worker Intl ranks, Rust index/Date scan, output transfer and bounded model projection.',
        sourceModelConstruction: 'Outside sample timing; every regime receives the same unchanged workbook and native Region availability filter.',
        coldDefinition: 'Fresh client and Worker per query. Browser/Vite/network caches may be warm; this is not guaranteed disk-cold startup.',
        warmDefinition: 'One real Worker/module reused after discarded items, Canvas and metadata warmup.',
        transportAccounting: 'Source packed bytes, all seven input transfers/copies and only requested indices/hasData output copies/transfers are recorded by real compute metrics.',
        memoryLimitations: ['Main-thread usedJSHeapSize is optional and GC-sensitive; it does not measure Worker resident/WASM memory.',
          'User-agent-specific memory is optional and requires secure cross-origin isolation; default fixture may return null.',
          'estimatedPeakBytes is conservative admission accounting, not measured peak memory; source model heap is excluded.',
          'Worker linear/resident memory, peak memory and sustained large Office Canvas workloads require separate validation.'],
        frameLimitations: 'Frame intervals and long tasks measure this query fixture only. Office first paint, scroll, edit and zoom are unmeasured.',
        cancellation: 'Triggered by the first real Worker progress packet; rejection, real cancelled acknowledgement and next query completion are separate clocks.' },
      workloads, cancellation, officeCanvasMetrics: { firstPaintMs: null, scroll: null, edit: null, zoom: null,
        status: '未验收；本夹具未挂载工作台 Canvas' } }
  } finally { observer?.disconnect() }
}

window.slicerMemberBenchmark = { run: runSlicerMemberBenchmark }
document.querySelector('#status').textContent = 'ready'
