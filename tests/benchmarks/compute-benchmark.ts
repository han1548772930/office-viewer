import { OfficeComputeClient } from '../../packages/office-studio/src/compute/client.ts'
import type { ComputeMetrics } from '../../packages/office-studio/src/compute/protocol.ts'
import { NumericIndexService } from '../../packages/office-studio/src/features/sheet-compute/numeric-index-service.ts'

type NumericColumn = Array<number | null>
type OperationName = 'sort' | 'filter'
type Regime = 'typescript' | 'new-worker' | 'reused-worker'

interface BenchmarkConfig { sizes: number[]; iterations: number }
interface LongTaskSummary { count: number; totalMs: number; maxMs: number }
interface TimingSample {
  operation: OperationName
  regime: Regime
  rowCount: number
  sample: number
  wallMs: number
  inputBytes: number
  responseBytes: number
  boundaryBytes: number
  workerComputeMs: number | null
  workerQueueAndStartupMs: number | null
  clientTotalMs: number | null
  estimatedPeakBytes: number | null
  browserHeapBeforeBytes: number | null
  browserHeapAfterBytes: number | null
  browserHeapDeltaBytes: number | null
  longTasks: LongTaskSummary
  parity: boolean
}

interface DatasetResult {
  rowCount: number
  keyCount: number
  defaultMinRows: number
  defaultRoute: { sortUsesWorker: boolean; filterUsesWorker: boolean }
  cold: TimingSample[]
  warm: TimingSample[]
  typescript: TimingSample[]
  summary: Array<{ operation: OperationName; regime: Regime; p50WallMs: number; p95WallMs: number }>
}

interface BrowserBenchmarkResult {
  schemaVersion: 1
  generatedAt: string
  environment: {
    userAgent: string
    browserBrands: Array<{ brand: string; version: string }> | null
    hardwareConcurrency: number
    deviceMemoryGiB: number | null
    viewport: { width: number; height: number; devicePixelRatio: number }
    usedJSHeapSizeBeforeBytes: number | null
    usedJSHeapSizeAfterBytes: number | null
    longTaskObserverAvailable: boolean
  }
  routing: { configuredMinRows: number; benchmarkMinRowsOverride: number; thresholdStatus: string }
  methodology: {
    inputGenerator: string
    keyCount: number
    sortDirections: string[]
    nullsLast: boolean
    filterPredicates: string[]
    iterations: number
    coldDefinition: string
    warmDefinition: string
    memoryLimitations: string[]
  }
  workloads: DatasetResult[]
}

declare global {
  interface Window {
    computeBenchmark: { run(config: BenchmarkConfig): Promise<BrowserBenchmarkResult> }
    __officeComputeLongTasks?: Array<{ startTime: number; duration: number }>
  }
}

const status = document.querySelector<HTMLElement>('#status')
if (!status) throw new Error('Benchmark status element is missing')
status.textContent = 'ready'

const longTaskObserverAvailable = typeof PerformanceObserver !== 'undefined'
  && PerformanceObserver.supportedEntryTypes?.includes('longtask') === true
if (longTaskObserverAvailable) {
  window.__officeComputeLongTasks = []
  new PerformanceObserver((list) => {
    for (const entry of list.getEntries()) window.__officeComputeLongTasks?.push({ startTime: entry.startTime, duration: entry.duration })
  }).observe({ type: 'longtask', buffered: true })
}

window.computeBenchmark = { run: runBenchmark }

async function runBenchmark(config: BenchmarkConfig): Promise<BrowserBenchmarkResult> {
  if (!Array.isArray(config.sizes) || !config.sizes.length || config.sizes.some((size) => !Number.isSafeInteger(size) || size < 1)) throw new TypeError('sizes must contain positive safe integers')
  if (!Number.isSafeInteger(config.iterations) || config.iterations < 1) throw new TypeError('iterations must be a positive safe integer')
  const environmentBefore = readBrowserEnvironment()
  const workloads: DatasetResult[] = []
  for (const rowCount of config.sizes) workloads.push(await benchmarkDataset(rowCount, config.iterations))
  return {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    environment: { ...environmentBefore, usedJSHeapSizeAfterBytes: readHeapBytes(), longTaskObserverAvailable },
    routing: {
      configuredMinRows: 4096,
      benchmarkMinRowsOverride: 0,
      thresholdStatus: '4096 is a provisional routing threshold; benchmark results must be reviewed across supported devices before calibration.',
    },
    methodology: {
      inputGenerator: 'Deterministic two-column values: column 0=((row*17)%8191)-4096 except row%127=0 is null; column 1=((row*7919)%12289)-6144 except row%199=0 is null.',
      keyCount: 2,
      sortDirections: ['ascending', 'descending'],
      nullsLast: true,
      filterPredicates: ['column 0 between -1000 and 1000', 'column 1 greater than or equal to 0', 'match=all'],
      iterations: config.iterations,
      coldDefinition: 'A newly constructed OfficeComputeClient and Worker for each single task. Browser/network/module caches may already be warm; this measures fresh Worker initialization, not a guaranteed disk-cold load.',
      warmDefinition: 'A single client/Worker is initialized and both task types are run once as discarded warmup; subsequent samples reuse that Worker and WASM module.',
      memoryLimitations: [
        'Worker metrics estimatedPeakBytes is a schema-based estimate including JS input, a WASM input copy, index allocation, and fixed runtime allowance; it is not measured resident WASM memory.',
        'Chrome usedJSHeapSize is sampled before and after each operation when available; the delta includes garbage-collection and browser scheduling noise.',
        'Typed input and result bytes are exact payload byte counts; JS Array baseline storage is implementation-dependent and is not equated to this byte count.',
      ],
    },
    workloads,
  }
}

async function benchmarkDataset(rowCount: number, iterations: number): Promise<DatasetResult> {
  const columns = createColumns(rowCount)
  const expectedSort = sortInTypeScript(columns)
  const expectedFilter = filterInTypeScript(columns)
  const baselineSort = await measureTypeScript('sort', rowCount, iterations, () => sortInTypeScript(columns), expectedSort)
  const baselineFilter = await measureTypeScript('filter', rowCount, iterations, () => filterInTypeScript(columns), expectedFilter)

  const defaultClient = new OfficeComputeClient()
  const defaultService = new NumericIndexService(defaultClient)
  defaultClient.setDocumentRevision(`bench-route-${rowCount}`, 1)
  const routeContext = { documentId: `bench-route-${rowCount}`, modelRevision: 1 }
  const defaultSortProbe = await defaultService.sortRows({ rowCount, columns, directions: ['ascending', 'descending'], nullsLast: true }, routeContext)
  const defaultFilterProbe = await defaultService.filterRows({ rowCount, columns, predicates: [{ operator: 'between', lower: -1000, upper: 1000 }, { operator: 'ge', lower: 0 }], match: 'all' }, routeContext)
  defaultClient.dispose()

  const cold: TimingSample[] = []
  for (const operation of ['sort', 'filter'] as const) {
    for (let sample = 0; sample < Math.max(1, Math.min(iterations, 3)); sample += 1) {
      cold.push(await measureWorker(operation, 'new-worker', rowCount, sample, columns, expectedSort, expectedFilter))
    }
  }

  const client = new OfficeComputeClient()
  const service = new NumericIndexService(client, { minRows: 0 })
  const documentId = `bench-reused-${rowCount}`
  client.setDocumentRevision(documentId, 1)
  const context = { documentId, modelRevision: 1 }
  try {
    await runWorkerOperation(service, 'sort', rowCount, columns, context)
    await runWorkerOperation(service, 'filter', rowCount, columns, context)
    const warm: TimingSample[] = []
    for (let sample = 0; sample < iterations; sample += 1) {
      warm.push(await measureWorkerWithClient(service, 'sort', 'reused-worker', rowCount, sample, columns, context, expectedSort, expectedFilter))
      warm.push(await measureWorkerWithClient(service, 'filter', 'reused-worker', rowCount, sample, columns, context, expectedSort, expectedFilter))
    }
    const typescript = [...baselineSort, ...baselineFilter]
    return {
      rowCount,
      keyCount: 2,
      defaultMinRows: 4096,
      defaultRoute: { sortUsesWorker: defaultSortProbe !== null, filterUsesWorker: defaultFilterProbe !== null },
      cold,
      warm,
      typescript,
      summary: summarize([...typescript, ...cold, ...warm]),
    }
  } finally {
    client.dispose()
  }
}

async function measureTypeScript(operation: OperationName, rowCount: number, iterations: number, compute: () => Uint32Array, expected: Uint32Array): Promise<TimingSample[]> {
  compute() // warm-up JIT outside measured samples
  const samples: TimingSample[] = []
  for (let sample = 0; sample < iterations; sample += 1) {
    const heapBefore = readHeapBytes()
    const longTaskStart = performance.now()
    const longTaskIndex = window.__officeComputeLongTasks?.length ?? 0
    const start = performance.now()
    const indices = compute()
    const end = performance.now()
    assertSameIndices(indices, expected, `${operation} TypeScript baseline`)
    await yieldForPerformanceObserver()
    const heapAfter = readHeapBytes()
    samples.push({
      operation, regime: 'typescript', rowCount, sample, wallMs: end - start,
      inputBytes: rowCount * 2 * 8,
      responseBytes: indices.byteLength,
      boundaryBytes: 0,
      workerComputeMs: null, workerQueueAndStartupMs: null, clientTotalMs: null, estimatedPeakBytes: null,
      browserHeapBeforeBytes: heapBefore, browserHeapAfterBytes: heapAfter,
      browserHeapDeltaBytes: heapDelta(heapBefore, heapAfter),
      longTasks: collectLongTasks(longTaskIndex, longTaskStart, end + 1), parity: true,
    })
  }
  return samples
}

async function measureWorker(operation: OperationName, regime: 'new-worker', rowCount: number, sample: number,
  columns: NumericColumn[], expectedSort: Uint32Array, expectedFilter: Uint32Array): Promise<TimingSample> {
  const client = new OfficeComputeClient()
  const service = new NumericIndexService(client, { minRows: 0 })
  const documentId = `bench-${regime}-${operation}-${rowCount}-${sample}`
  client.setDocumentRevision(documentId, 1)
  try {
    return await measureWorkerWithClient(service, operation, regime, rowCount, sample, columns,
      { documentId, modelRevision: 1 }, expectedSort, expectedFilter)
  } finally {
    client.dispose()
  }
}

async function measureWorkerWithClient(service: NumericIndexService, operation: OperationName, regime: 'reused-worker' | 'new-worker',
  rowCount: number, sample: number, columns: NumericColumn[], context: { documentId: string; modelRevision: number },
  expectedSort: Uint32Array, expectedFilter: Uint32Array): Promise<TimingSample> {
  const heapBefore = readHeapBytes()
  const longTaskIndex = window.__officeComputeLongTasks?.length ?? 0
  const start = performance.now()
  const result = await runWorkerOperation(service, operation, rowCount, columns, context)
  const end = performance.now()
  await yieldForPerformanceObserver()
  const heapAfter = readHeapBytes()
  const expected = operation === 'sort' ? expectedSort : expectedFilter
  assertSameIndices(result.indices, expected, `${operation} Worker/WASM`)
  const responseBytes = result.indices.byteLength
  const clientTotalMs = result.metrics.totalMs ?? null
  const workerComputeMs = result.metrics.elapsedMs
  const queueMs = result.metrics.queueMs ?? null
  return {
    operation, regime, rowCount, sample, wallMs: end - start,
    inputBytes: result.metrics.inputBytes,
    responseBytes,
    boundaryBytes: result.metrics.inputBytes + responseBytes,
    workerComputeMs,
    workerQueueAndStartupMs: queueMs,
    clientTotalMs,
    estimatedPeakBytes: result.metrics.estimatedPeakBytes,
    browserHeapBeforeBytes: heapBefore, browserHeapAfterBytes: heapAfter,
    browserHeapDeltaBytes: heapDelta(heapBefore, heapAfter),
    longTasks: collectLongTasks(longTaskIndex, start, end + 1),
    parity: true,
  }
}

async function runWorkerOperation(service: NumericIndexService, operation: OperationName, rowCount: number, columns: NumericColumn[], context: { documentId: string; modelRevision: number }) {
  const result = operation === 'sort'
    ? await service.sortRows({ rowCount, columns, directions: ['ascending', 'descending'], nullsLast: true }, context)
    : await service.filterRows({ rowCount, columns, predicates: [{ operator: 'between', lower: -1000, upper: 1000 }, { operator: 'ge', lower: 0 }], match: 'all' }, context)
  if (!result) throw new Error(`NumericIndexService unexpectedly selected the TypeScript fallback at ${rowCount} rows`)
  return result
}

function createColumns(rowCount: number): NumericColumn[] {
  const first: NumericColumn = new Array(rowCount)
  const second: NumericColumn = new Array(rowCount)
  for (let row = 0; row < rowCount; row += 1) {
    first[row] = row % 127 === 0 ? null : ((row * 17) % 8191) - 4096
    second[row] = row % 199 === 0 ? null : ((row * 7919) % 12289) - 6144
  }
  return [first, second]
}

function sortInTypeScript(columns: NumericColumn[]): Uint32Array {
  const indices = Array.from({ length: columns[0]!.length }, (_, row) => row)
  indices.sort((left, right) => {
    for (let key = 0; key < columns.length; key += 1) {
      const l = columns[key]![left]!
      const r = columns[key]![right]!
      const leftBlank = l === null
      const rightBlank = r === null
      if (leftBlank || rightBlank) {
        if (leftBlank !== rightBlank) return leftBlank ? 1 : -1
      } else {
        const order = l! < r! ? -1 : l! > r! ? 1 : 0
        if (order !== 0) return key === 1 ? -order : order
      }
    }
    return left - right
  })
  return Uint32Array.from(indices)
}

function filterInTypeScript(columns: NumericColumn[]): Uint32Array {
  const first = columns[0]!
  const second = columns[1]!
  const matched: number[] = []
  for (let row = 0; row < first.length; row += 1) {
    const value = first[row]!
    const secondary = second[row]!
    if (value !== null && secondary !== null && value >= -1000 && value <= 1000 && secondary >= 0) matched.push(row)
  }
  return Uint32Array.from(matched)
}

function summarize(samples: TimingSample[]): DatasetResult['summary'] {
  const groups = new Map<string, number[]>()
  for (const sample of samples) {
    const key = `${sample.operation}:${sample.regime}`
    groups.set(key, [...(groups.get(key) ?? []), sample.wallMs])
  }
  return [...groups].map(([key, timings]) => {
    const [operation, regime] = key.split(':') as [OperationName, Regime]
    return { operation, regime, p50WallMs: percentile(timings, 0.5), p95WallMs: percentile(timings, 0.95) }
  })
}

function percentile(values: number[], quantile: number): number {
  const sorted = [...values].sort((a, b) => a - b)
  return sorted[Math.max(0, Math.ceil(quantile * sorted.length) - 1)] ?? 0
}

function assertSameIndices(actual: Uint32Array, expected: Uint32Array, label: string): void {
  if (actual.length !== expected.length) throw new Error(`${label} result length mismatch: ${actual.length} != ${expected.length}`)
  for (let index = 0; index < actual.length; index += 1) if (actual[index] !== expected[index]) throw new Error(`${label} parity mismatch at index ${index}: ${actual[index]} != ${expected[index]}`)
}

function readHeapBytes(): number | null {
  const memory = (performance as Performance & { memory?: { usedJSHeapSize: number } }).memory
  return memory && Number.isFinite(memory.usedJSHeapSize) ? memory.usedJSHeapSize : null
}

function heapDelta(before: number | null, after: number | null): number | null { return before === null || after === null ? null : after - before }

function collectLongTasks(index: number, start: number, end: number): LongTaskSummary {
  const tasks = (window.__officeComputeLongTasks ?? []).slice(index).filter((task) => task.startTime >= start && task.startTime <= end)
  return { count: tasks.length, totalMs: tasks.reduce((sum, task) => sum + task.duration, 0), maxMs: tasks.reduce((max, task) => Math.max(max, task.duration), 0) }
}

function yieldForPerformanceObserver(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0))
}

function readBrowserEnvironment(): BrowserBenchmarkResult['environment'] {
  const uaData = (navigator as Navigator & { userAgentData?: { brands?: Array<{ brand: string; version: string }> } }).userAgentData
  return {
    userAgent: navigator.userAgent,
    browserBrands: uaData?.brands ?? null,
    hardwareConcurrency: navigator.hardwareConcurrency,
    deviceMemoryGiB: (navigator as Navigator & { deviceMemory?: number }).deviceMemory ?? null,
    viewport: { width: window.innerWidth, height: window.innerHeight, devicePixelRatio: window.devicePixelRatio },
    usedJSHeapSizeBeforeBytes: readHeapBytes(), usedJSHeapSizeAfterBytes: null, longTaskObserverAvailable,
  }
}
