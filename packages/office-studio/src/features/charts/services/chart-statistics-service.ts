import type { ChartModel, WorkbookModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { OfficeComputeClient } from '../../../compute/client.ts'
import type { ComputeMetrics, ComputeProgress } from '../../../compute/protocol.ts'
import { COMPUTE_LIMITS } from '../../../compute/protocol.ts'
import { ComputeError, normalizeComputeError } from '../../../compute/errors.ts'
import { estimateChartTaskMemory } from '../../../compute/chart-tasks.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../../sheet-protection/range-protection'
import type { ChartComputedStatistics } from '../chart-derived'
import { prepareChartComputeInputs } from '../chart-input-preparation'
import { projectChartRecords } from '../chart-record-projection'
import { readChartRecordsAsync } from '../chart-source-compute'
import { decodeChartStatistics, prepareChartStatisticsTask } from '../chart-statistics-compute'
import { estimateChartDataMemory, estimateChartDecodedMemory } from '../chart-compute-memory'
import { assertChartPreparationCurrent, assertChartPreparationMemory, chartPreparationBudget, type ChartComputePreparationContext } from '../chart-compute-lifecycle'

export interface ChartStatisticsRequestContext extends ChartComputePreparationContext {
  documentId: string
  modelRevision: number
  userId?: string
  priority?: 'interactive' | 'background'
  timeoutMs?: number
  onProgress?: (progress: ComputeProgress) => void
}

export interface ChartStatisticsSnapshot {
  documentId: string
  modelRevision: number
  chartId: string
  chartRevision: number
  signature: string
  statistics: ChartComputedStatistics
  metrics?: ComputeMetrics
  preparationMs: number
  decodingMs: number
  estimatedCacheBytes: number
}

export interface ChartStatisticsRequest {
  promise: Promise<ChartStatisticsSnapshot>
  cancel(): void
}

export interface ChartStatisticsServiceOptions {
  memoryBudgetBytes?: number
  cacheBudgetBytes?: number
  maxCachedCharts?: number
  maxPendingCharts?: number
}

export interface ChartStatisticsInvalidation {
  documentId: string
  chartId?: string
  reason: 'revision' | 'configuration' | 'closed' | 'disposed'
}

interface Consumer {
  progress?: (progress: ComputeProgress) => void
  deadlineAt: number
  fail(error: ComputeError): void
}
interface Entry {
  documentId: string; modelRevision: number; chartId: string; signature: string; userId: string
  controller: AbortController; consumers: Set<Consumer>; promise: Promise<ChartStatisticsSnapshot>
  snapshot?: ChartStatisticsSnapshot; failure?: ComputeError
  startedAt: number; memoryBudget: number; priority: 'interactive' | 'background'
  resolve(snapshot: ChartStatisticsSnapshot): void
  reject(error: ComputeError): void
}
interface ScheduledChart {
  workbook: WorkbookModel; chart: ChartModel; entry: Entry
}

/** The Studio's shared compute client owns Worker/WASM scheduling. This service
 * owns prepared chart data and request/cache lifetimes, never document content
 * or history. One chart's complete preparation/compute/decode lifecycle runs at
 * a time to bound resident JS data. Queued charts retain source references only;
 * cancellation removes those references immediately. The shared compute client
 * still schedules other calculation domains independently. */
export class ChartStatisticsService {
  private readonly client: OfficeComputeClient
  private readonly memoryBudget: number
  private readonly cacheBudget: number
  private readonly maxCached: number
  private readonly maxPending: number
  private readonly revisions = new Map<string, number>()
  private readonly entries = new Map<string, Map<string, Entry>>()
  private readonly listeners = new Set<(event: ChartStatisticsInvalidation) => void>()
  private queue: ScheduledChart[] = []
  private active?: ScheduledChart
  private pumpScheduled = false
  private disposed = false
  /** Observation failures are available for diagnostics and cannot stop invalidation. */
  readonly listenerErrors: unknown[] = []

  constructor(client: OfficeComputeClient, options: ChartStatisticsServiceOptions = {}) {
    this.client = client
    this.memoryBudget = chartPreparationBudget({ memoryBudgetBytes: options.memoryBudgetBytes })
    this.cacheBudget = options.cacheBudgetBytes ?? Math.min(64 * 1024 * 1024, this.memoryBudget)
    this.maxCached = options.maxCachedCharts ?? 64
    this.maxPending = options.maxPendingCharts ?? COMPUTE_LIMITS.maxPendingTasks
    if (!Number.isSafeInteger(this.cacheBudget) || this.cacheBudget <= 0 || this.cacheBudget > this.memoryBudget
      || !Number.isSafeInteger(this.maxCached) || this.maxCached < 1 || !Number.isSafeInteger(this.maxPending) || this.maxPending < 1) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '图表缓存或任务队列预算无效')
    }
  }

  subscribe(listener: (event: ChartStatisticsInvalidation) => void): { dispose(): void } {
    this.assertAlive(); this.listeners.add(listener)
    return { dispose: () => this.listeners.delete(listener) }
  }

  setDocumentRevision(documentId: string, revision: number): void {
    this.assertAlive()
    this.client.setDocumentRevision(documentId, revision)
    const previous = this.revisions.get(documentId)
    this.revisions.set(documentId, revision)
    if (previous !== undefined && previous !== revision) this.invalidateUnit(documentId, 'revision')
  }

  forgetDocument(documentId: string): void {
    this.assertAlive()
    this.revisions.delete(documentId)
    this.invalidateUnit(documentId, 'closed')
  }

  isCurrent(context: Pick<ChartStatisticsRequestContext, 'documentId' | 'modelRevision'>): boolean {
    return !this.disposed && this.revisions.get(context.documentId) === context.modelRevision && this.client.isCurrent(context)
  }

  getCached(chart: ChartModel, context: ChartStatisticsRequestContext): ChartStatisticsSnapshot | undefined {
    const budget = Math.min(this.memoryBudget, chartPreparationBudget(context))
    assertChartPreparationCurrent(context)
    if (!this.isCurrent(context)) return undefined
    const entry = this.entries.get(context.documentId)?.get(chart.id)
    return entry?.modelRevision === context.modelRevision && !entry.failure && entry.userId === (context.userId ?? 'local-user')
      && entry.signature === chartStatisticsSignature(chart) && (entry.snapshot?.estimatedCacheBytes ?? 0) <= budget ? entry.snapshot : undefined
  }

  request(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext): ChartStatisticsRequest {
    this.assertAlive()
    const startedAt = performance.now(), timeoutMs = context.timeoutMs ?? 30_000
    const memoryBudget = Math.min(this.memoryBudget, chartPreparationBudget(context))
    assertChartPreparationCurrent(context)
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0 || timeoutMs > 2_147_483_647
      || context.priority !== undefined && context.priority !== 'interactive' && context.priority !== 'background') {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '图表计算超时或优先级配置无效')
    }
    if (!this.isCurrent(context) || workbook.id !== context.documentId || !workbook.charts?.some(candidate => candidate.id === chart.id
      && chartStatisticsSignature(candidate) === chartStatisticsSignature(chart))) {
      throw new ComputeError('COMPUTE_STALE_RESULT', '图表来源、文档或配置已失效')
    }
    const signature = chartStatisticsSignature(chart), userId = context.userId ?? 'local-user'
    assertChartReadPermission(workbook, chart, userId)
    let unit = this.entries.get(context.documentId)
    if (!unit) { unit = new Map(); this.entries.set(context.documentId, unit) }
    let entry = unit.get(chart.id)
    const compatible = (candidate: Entry): boolean => !candidate.failure && candidate.modelRevision === context.modelRevision
      && candidate.signature === signature && candidate.userId === userId
    if (entry && !compatible(entry)) {
      // Remove only the outgoing slot before abort/observation callbacks can
      // reenter. A callback's new request must never be overwritten below.
      unit.delete(chart.id)
      this.rejectEntry(entry, new ComputeError('COMPUTE_STALE_RESULT', '图表配置或来源已更新'))
      this.emit({ documentId: context.documentId, chartId: chart.id, reason: 'configuration' })
      this.assertAlive()
      assertChartPreparationCurrent(context)
      if (!this.isCurrent(context) || chartStatisticsSignature(chart) !== signature) throw new ComputeError('COMPUTE_STALE_RESULT', '图表配置观察期间来源已更新')
      unit = this.entries.get(context.documentId)
      if (!unit) { unit = new Map(); this.entries.set(context.documentId, unit) }
      entry = unit.get(chart.id)
      if (entry && !compatible(entry)) throw new ComputeError('COMPUTE_STALE_RESULT', '新的图表请求已替代当前配置')
    }
    if (!entry) {
      if (this.pendingCount() >= this.maxPending) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表数据准备队列已满')
      const frozenChart = cloneModel(chart), controller = new AbortController()
      let resolve!: Entry['resolve'], reject!: Entry['reject']
      const promise = new Promise<ChartStatisticsSnapshot>((yes, no) => { resolve = yes; reject = no })
      const active: Entry = { documentId: context.documentId, modelRevision: context.modelRevision, chartId: chart.id, signature, userId,
        controller, consumers: new Set(), promise, resolve, reject, startedAt, memoryBudget, priority: context.priority ?? 'background' }
      entry = active
      unit.set(chart.id, active)
      const scheduled: ScheduledChart = { workbook, chart: frozenChart, entry: active }
      const firstBackground = this.queue.findIndex(job => job.entry.priority === 'background')
      if (active.priority === 'interactive' && firstBackground >= 0) this.queue.splice(firstBackground, 0, scheduled)
      else this.queue.push(scheduled)
      this.scheduleNext()
    }
    if (context.priority === 'interactive') this.promoteInteractive(entry)
    return this.attach(entry, context, startedAt, timeoutMs)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    for (const documentId of [...this.entries.keys()]) this.invalidateUnit(documentId, 'disposed')
    this.queue = []
    this.revisions.clear(); this.listeners.clear()
  }

  private scheduleNext(): void {
    if (this.active || this.pumpScheduled || this.disposed || !this.queue.length) return
    this.pumpScheduled = true
    queueMicrotask(() => { this.pumpScheduled = false; this.runNext() })
  }

  private promoteInteractive(entry: Entry): void {
    if (entry.priority === 'interactive') return
    entry.priority = 'interactive'
    const index = this.queue.findIndex(job => job.entry === entry)
    if (index < 0) return
    const [job] = this.queue.splice(index, 1)
    if (!job) return
    const firstBackground = this.queue.findIndex(candidate => candidate.entry.priority === 'background')
    if (firstBackground >= 0) this.queue.splice(firstBackground, 0, job)
    else this.queue.push(job)
  }

  private runNext(): void {
    if (this.active || this.disposed) return
    const job = this.queue.shift()
    if (!job) return
    if (job.entry.failure) { this.scheduleNext(); return }
    this.active = job
    const entry = job.entry
    void this.prepare(job.workbook, job.chart, entry).then(snapshot => {
      this.assertEntryCurrent(entry)
      entry.snapshot = snapshot
      this.trimCache(entry)
      entry.resolve(snapshot)
    }).catch(error => this.rejectEntry(entry, normalizeComputeError(error))).finally(() => {
      if (this.active === job) this.active = undefined
      this.scheduleNext()
    })
  }

  private async prepare(workbook: WorkbookModel, chart: ChartModel, entry: Entry): Promise<ChartStatisticsSnapshot> {
    this.assertEntryCurrent(entry)
    const available = entry.memoryBudget - this.cacheBytes()
    if (available <= 0) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表缓存占用了数据准备预算')
    const preparation: ChartComputePreparationContext = { signal: entry.controller.signal, memoryBudgetBytes: available,
      assertCurrent: () => { this.assertEntryCurrent(entry); assertChartReadPermission(workbook, chart, entry.userId) } }
    let records = await readChartRecordsAsync(workbook, chart, preparation)
    // Source + projection/normalization/aggregation references are charged
    // before allocating projected numeric arrays. Strings are checked by the reader.
    assertChartPreparationMemory(records.records.length * (records.fields.length * 64 + 192) + 8 * 1024 * 1024, preparation)
    let data = projectChartRecords(records, chart)
    records = { fields: [], records: [] }
    data = await prepareChartComputeInputs(data, chart, preparation)
    const dataBytes = await estimateChartDataMemory(data, preparation)
    assertChartPreparationMemory(dataBytes, preparation)
    // A no-statistics chart needs no extra bulk allocation. Keep its control
    // budget valid at an exact admission boundary; every real task is charged
    // against dataBytes + its complete peak before it reaches the Worker.
    const taskContext: ChartComputePreparationContext = { ...preparation, memoryBudgetBytes: Math.max(1, available - dataBytes) }
    const plan = await prepareChartStatisticsTask(data, chart, taskContext)
    const preparationMs = performance.now() - entry.startedAt
    let statistics: ChartComputedStatistics = { preparedData: data }, metrics: ComputeMetrics | undefined, decodingMs = 0, decodedBytes = 0
    if (plan.task) {
      assertChartPreparationMemory(dataBytes + estimateChartTaskMemory(plan.task.params).estimatedPeakBytes, preparation)
      const result = await this.client.submit(plan.task, { documentId: entry.documentId, modelRevision: entry.modelRevision,
        // A later view may outlive the original consumer. Per-consumer deadlines
        // include preparation/queueing and abort this task when the last view
        // leaves; the shared Worker must not inherit the first view's timeout.
        ownership: 'transfer', priority: entry.priority, signal: entry.controller.signal, timeoutMs: 2_147_483_647,
        onProgress: progress => this.publishProgress(entry, progress) }).promise
      this.assertEntryCurrent(entry)
      if (result.output.kind !== 'chartStatistics') throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表任务返回了其他类型的结果')
      decodedBytes = estimateChartDecodedMemory(result.output)
      assertChartPreparationMemory(dataBytes + result.output.values.byteLength + result.output.indices.byteLength + decodedBytes, preparation)
      const decodingStart = performance.now()
      statistics = await decodeChartStatistics(plan, result.output, taskContext)
      decodingMs = performance.now() - decodingStart; metrics = result.metrics
    }
    this.assertEntryCurrent(entry)
    return { documentId: entry.documentId, modelRevision: entry.modelRevision, chartId: entry.chartId, chartRevision: chart.revision,
      signature: entry.signature, statistics, metrics, preparationMs, decodingMs, estimatedCacheBytes: dataBytes + decodedBytes }
  }

  private attach(entry: Entry, context: ChartStatisticsRequestContext, startedAt: number, timeoutMs: number): ChartStatisticsRequest {
    let settled = false, removeAbort = () => {}, rejectPromise: (error: ComputeError) => void = () => {}
    let deadline: ReturnType<typeof setTimeout> | undefined
    const consumer: Consumer = { progress: context.onProgress, deadlineAt: startedAt + timeoutMs, fail: error => {
      if (settled) return
      settled = true; detach(); rejectPromise(error)
    } }
    const detach = () => {
      entry.consumers.delete(consumer); removeAbort(); clearTimeout(deadline)
      if (!entry.snapshot && !entry.consumers.size) this.rejectEntry(entry, new ComputeError('COMPUTE_CANCELLED', '图表计算已取消'))
    }
    const cancel = () => {
      if (settled) return
      settled = true; detach(); rejectPromise(new ComputeError('COMPUTE_CANCELLED', '图表计算请求已取消'))
    }
    const promise = new Promise<ChartStatisticsSnapshot>((resolve, reject) => {
      rejectPromise = reject; entry.consumers.add(consumer)
      // Subscribe before attaching cancellation. Even an abort during attachment
      // has a rejection observer for the shared entry's promise.
      entry.promise.then(snapshot => {
        if (settled) return
        try {
          if (performance.now() >= consumer.deadlineAt) throw new ComputeError('COMPUTE_TIMEOUT', '图表计算请求超时')
          if (!this.isCurrent(context) || entry.failure) throw entry.failure ?? new ComputeError('COMPUTE_STALE_RESULT', '图表结果已过期')
          assertChartPreparationMemory(snapshot.estimatedCacheBytes, context)
          context.assertCurrent?.()
          if (!this.isCurrent(context) || entry.failure) throw entry.failure ?? new ComputeError('COMPUTE_STALE_RESULT', '图表观察期间来源已更新')
          if (performance.now() >= consumer.deadlineAt) throw new ComputeError('COMPUTE_TIMEOUT', '图表计算请求超时')
          settled = true; detach(); resolve(snapshot)
        } catch (error) { settled = true; detach(); reject(normalizeComputeError(error)) }
      }, error => { if (!settled) { settled = true; detach(); reject(normalizeComputeError(error)) } })
      if (entry.failure) { consumer.fail(entry.failure); return }
      if (context.signal) {
        context.signal.addEventListener('abort', cancel, { once: true })
        removeAbort = () => context.signal?.removeEventListener('abort', cancel)
        if (context.signal.aborted) { cancel(); return }
      }
      if (!entry.snapshot) deadline = setTimeout(() => consumer.fail(new ComputeError('COMPUTE_TIMEOUT', '图表计算请求超时')),
        Math.max(0, consumer.deadlineAt - performance.now()))
    })
    return { promise, cancel }
  }

  private assertEntryCurrent(entry: Entry): void {
    if (entry.failure) throw entry.failure
    this.expireConsumers(entry)
    if (entry.failure) throw entry.failure
    if (entry.controller.signal.aborted) throw new ComputeError('COMPUTE_CANCELLED', '图表计算已取消')
    if (!this.isCurrent(entry) || this.entries.get(entry.documentId)?.get(entry.chartId) !== entry) throw new ComputeError('COMPUTE_STALE_RESULT', '图表来源版本已失效')
  }

  private rejectEntry(entry: Entry, error: ComputeError): void {
    if (entry.failure) return
    entry.failure = error
    if (this.entries.get(entry.documentId)?.get(entry.chartId) === entry) this.entries.get(entry.documentId)!.delete(entry.chartId)
    this.queue = this.queue.filter(job => job.entry !== entry)
    entry.controller.abort()
    entry.reject(error)
    for (const consumer of [...entry.consumers]) consumer.fail(error)
    this.scheduleNext()
  }

  private invalidateUnit(documentId: string, reason: ChartStatisticsInvalidation['reason']): void {
    const unit = this.entries.get(documentId)
    this.entries.delete(documentId)
    for (const entry of unit?.values() ?? []) this.rejectEntry(entry, new ComputeError(reason === 'disposed' ? 'COMPUTE_DISPOSED' : 'COMPUTE_STALE_RESULT', '图表来源已关闭或更新'))
    this.emit({ documentId, reason })
  }

  private pendingCount(): number {
    return this.queue.length + (this.active && !this.active.entry.failure && !this.active.entry.snapshot ? 1 : 0)
  }
  private cacheBytes(): number {
    let bytes = 0
    for (const unit of this.entries.values()) for (const entry of unit.values()) bytes += entry.snapshot?.estimatedCacheBytes ?? 0
    return bytes
  }
  private trimCache(latest: Entry): void {
    const cached: Entry[] = []
    for (const unit of this.entries.values()) for (const entry of unit.values()) if (entry.snapshot) cached.push(entry)
    let bytes = this.cacheBytes(), count = cached.length
    for (const entry of [...cached.filter(entry => entry !== latest), latest]) {
      if (bytes <= this.cacheBudget && count <= this.maxCached) break
      this.entries.get(entry.documentId)?.delete(entry.chartId)
      bytes -= entry.snapshot!.estimatedCacheBytes; count--
    }
  }
  private publishProgress(entry: Entry, progress: ComputeProgress): void {
    this.expireConsumers(entry)
    for (const consumer of [...entry.consumers]) {
      if (entry.failure) break
      if (!entry.consumers.has(consumer)) continue
      try { consumer.progress?.(progress) }
      catch (error) { this.recordListenerError(error) }
    }
  }
  private expireConsumers(entry: Entry): void {
    const now = performance.now()
    for (const consumer of [...entry.consumers]) {
      if (now >= consumer.deadlineAt) consumer.fail(new ComputeError('COMPUTE_TIMEOUT', '图表计算请求超时'))
    }
  }
  private emit(event: ChartStatisticsInvalidation): void {
    for (const listener of [...this.listeners]) {
      try { listener(event) }
      catch (error) { this.recordListenerError(error) }
    }
  }
  private recordListenerError(error: unknown): void { this.listenerErrors.push(error); if (this.listenerErrors.length > 16) this.listenerErrors.shift() }
  private assertAlive(): void { if (this.disposed) throw new ComputeError('COMPUTE_DISPOSED', '图表统计服务已释放') }
}

export function chartStatisticsSignature(chart: ChartModel): string {
  return JSON.stringify([chart.sheetId, chart.type, chart.source, chart.pivotTableId, chart.categoryField, chart.valueFields, chart.seriesToCategory, chart.options])
}

export function assertChartReadPermission(workbook: WorkbookModel, chart: ChartModel, userId: string): void {
  const host = workbook.sheets.find(sheet => sheet.id === chart.sheetId)
  if (!host || !canViewWorksheetProtection(host, userId)) throw new ComputeError('COMPUTE_INVALID_INPUT', '没有查看图表所在工作表的权限')
  if (chart.pivotTableId) {
    const pivot = workbook.pivotTables?.find(candidate => candidate.id === chart.pivotTableId)
    const pivotSource = workbook.sheets.find(sheet => sheet.id === pivot?.source.subUnitId)
    const pivotOutput = workbook.sheets.find(sheet => sheet.id === pivot?.anchor.subUnitId)
    if (!pivot || !pivotSource || !pivotOutput || !canViewWorksheetProtection(pivotSource, userId) || !canViewWorksheetProtection(pivotOutput, userId)
      || !canViewProtectedRanges(pivotSource, [pivot.source.range], userId)
      || pivot.outputRange && !canViewProtectedRanges(pivotOutput, [pivot.outputRange], userId)) throw new ComputeError('COMPUTE_INVALID_INPUT', '没有查看透视图源数据或输出的权限')
  } else {
    const source = workbook.sheets.find(sheet => sheet.id === (chart.source.sheetId ?? chart.sheetId))
    if (!source || !canViewWorksheetProtection(source, userId) || !canViewProtectedRanges(source, [chart.source.range], userId)) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '没有查看图表源区域的权限')
    }
  }
}
