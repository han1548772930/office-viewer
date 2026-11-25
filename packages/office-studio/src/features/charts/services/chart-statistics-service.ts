import type { ChartModel, WorkbookModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { OfficeComputeClient } from '../../../compute/client.ts'
import type { ComputeMetrics, ComputeProgress } from '../../../compute/protocol.ts'
import { COMPUTE_LIMITS } from '../../../compute/protocol.ts'
import { ComputeError, normalizeComputeError } from '../../../compute/errors.ts'
import { estimateChartTaskMemory } from '../../../compute/chart-tasks.ts'
import { canViewProtectedRanges, canViewWorksheetProtection } from '../../sheet-protection/range-protection'
import type { ChartComputedStatistics } from '../chart-derived'
import { updateChartDefinition, type ChartData, type ChartUpdatePatch } from '../chart-engine'
import { initializeChartWaterfallOptionsAsync } from '../chart-waterfall'
import { validateChartSourceDataAsync } from '../chart-source-validation'
import { prepareChartComputeInputs } from '../chart-input-preparation'
import { projectChartRecordsAsync } from '../chart-record-projection'
import { readChartRecordsAsync } from '../chart-source-compute'
import { decodeChartStatistics, prepareChartStatisticsTask } from '../chart-statistics-compute'
import { estimateChartDataMemory, estimateChartDecodedMemory, estimateChartFinalLayoutMemory } from '../chart-compute-memory'
import { assertChartPreparationCurrent, assertChartPreparationMemory, chartPreparationBudget, type ChartComputePreparationContext } from '../chart-compute-lifecycle'
import { layoutChartData } from '../chart-layout'
import { prepareChartWordCloudLayout } from '../chart-word-cloud-layout'
import { prepareChartRelationLayout } from '../chart-relation-layout'
import { assertChartSourcePresent, chartSourceNotFoundDiagnostic, ChartSourceNotFoundError } from '../chart-source-diagnostic'
import { resolveChartImageResource } from '../chart-image-resources'

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
  /** Only detached authoring owns this normalized model; never a view cache. */
  authoredChart?: ChartModel
}

export interface ChartStatisticsRequest {
  promise: Promise<ChartStatisticsSnapshot>
  cancel(): void
}

export interface ChartSourceDataRequest {
  promise: Promise<ChartData>
  cancel(): void
}

export interface ChartAuthoringPlan {
  previous?: ChartModel
  patch?: ChartUpdatePatch
  initializeDefaults?: boolean
}

export interface ChartAuthoringRequest {
  promise: Promise<ChartModel>
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
  assertCurrent?: () => void
  deadlineAt: number
  fail(error: ComputeError): void
}
interface Entry {
  documentId: string; modelRevision: number; chartId: string; signature: string; userId: string
  /** Detached edits use a separate slot, sharing the same bounded scheduler. */
  candidateId?: number
  /** Settings read the authored source, without statistical or scene allocation. */
  sourceOnly?: boolean
  /** Settings statistics do not admit a final publishable date scene. */
  settingsOnly?: boolean
  authoring?: ChartAuthoringPlan
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
  private readonly candidates = new Map<number, Entry>()
  private nextCandidateId = 0
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

  getCached(chart: ChartModel, context: ChartStatisticsRequestContext, workbook?: WorkbookModel): ChartStatisticsSnapshot | undefined {
    if (workbook) this.assertSourcePresent(workbook, chart, context.documentId)
    const budget = Math.min(this.memoryBudget, chartPreparationBudget(context))
    assertChartPreparationCurrent(context)
    if (!this.isCurrent(context)) return undefined
    const entry = this.entries.get(context.documentId)?.get(chart.id)
    return entry?.modelRevision === context.modelRevision && !entry.failure && entry.userId === (context.userId ?? 'local-user')
      && entry.signature === chartStatisticsSignature(chart) && (entry.snapshot?.estimatedCacheBytes ?? 0) <= budget ? entry.snapshot : undefined
  }

  request(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext): ChartStatisticsRequest {
    this.assertAlive()
    this.assertSourcePresent(workbook, chart, context.documentId)
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

  /** Compute a detached candidate without replacing a published chart's cache
   * or notifying its views. No candidate snapshot survives its last consumer. */
  requestCandidate(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext): ChartStatisticsRequest {
    return this.requestDetached(workbook, chart, context, false)
  }

  /** Read a draft which need not be published. It shares admission, queueing,
   * cancellation and version lifetimes with ordinary chart calculations. */
  requestSourceData(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext): ChartSourceDataRequest {
    const request = this.requestDetached(workbook, chart, context, true)
    return { cancel: request.cancel, promise: request.promise.then(snapshot => {
      if (!snapshot.statistics.preparedData) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表数据准备没有返回源数据')
      return snapshot.statistics.preparedData
    }) }
  }

  /** Read derived identities for an unpublished settings draft. Statistics
   * use the same Rust engine as views; no published cache is replaced. */
  requestSettingsStatistics(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext): ChartStatisticsRequest {
    return this.requestDetached(workbook, chart, context, false, { settingsOnly: true })
  }

  /** Creation/type conversion reads, projects and validates within one bounded
   * lifecycle. A caller publishes only the returned normalized small model. */
  requestAuthoredCandidate(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext,
    authoring: ChartAuthoringPlan = {}): ChartAuthoringRequest {
    const request = this.requestDetached(workbook, chart, context, false, { authoring })
    return { cancel: request.cancel, promise: request.promise.then(snapshot => {
      if (!snapshot.authoredChart) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表候选准备没有返回模型')
      return snapshot.authoredChart
    }) }
  }

  private requestDetached(workbook: WorkbookModel, chart: ChartModel, context: ChartStatisticsRequestContext, sourceOnly: boolean,
    extra: Pick<Entry, 'settingsOnly' | 'authoring'> = {}): ChartStatisticsRequest {
    this.assertAlive()
    this.assertSourcePresent(workbook, chart, context.documentId)
    const startedAt = performance.now(), timeoutMs = context.timeoutMs ?? 30_000
    const memoryBudget = Math.min(this.memoryBudget, chartPreparationBudget(context))
    assertChartPreparationCurrent(context)
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0 || timeoutMs > 2_147_483_647
      || context.priority !== undefined && context.priority !== 'interactive' && context.priority !== 'background') {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '图表候选计算超时或优先级配置无效')
    }
    if (!this.isCurrent(context) || workbook.id !== context.documentId || !sourceOnly && !extra.settingsOnly && !workbook.charts?.some(candidate => candidate.id === chart.id
      && chartStatisticsSignature(candidate) === chartStatisticsSignature(chart))) throw new ComputeError('COMPUTE_STALE_RESULT', '图表候选不属于当前文档版本')
    const userId = context.userId ?? 'local-user'
    assertChartReadPermission(workbook, chart, userId)
    if (this.pendingCount() >= this.maxPending) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表候选数据准备队列已满')
    if (!Number.isSafeInteger(this.nextCandidateId + 1)) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表候选会话身份超出容量')
    const candidateId = ++this.nextCandidateId, controller = new AbortController(), frozenChart = cloneModel(chart)
    let resolve!: Entry['resolve'], reject!: Entry['reject']
    const promise = new Promise<ChartStatisticsSnapshot>((yes, no) => { resolve = yes; reject = no })
    const entry: Entry = { documentId: context.documentId, modelRevision: context.modelRevision, chartId: chart.id,
      candidateId, sourceOnly, ...cloneModel(extra), signature: chartStatisticsSignature(chart), userId, controller, consumers: new Set(), promise, resolve, reject,
      startedAt, memoryBudget, priority: context.priority ?? 'interactive' }
    this.candidates.set(candidateId, entry)
    const scheduled = { workbook, chart: frozenChart, entry }, firstBackground = this.queue.findIndex(job => job.entry.priority === 'background')
    if (entry.priority === 'interactive' && firstBackground >= 0) this.queue.splice(firstBackground, 0, scheduled)
    else this.queue.push(scheduled)
    this.scheduleNext()
    return this.attach(entry, context, startedAt, timeoutMs)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    for (const documentId of new Set([...this.entries.keys(), ...Array.from(this.candidates.values(), entry => entry.documentId)])) this.invalidateUnit(documentId, 'disposed')
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
      if (entry.candidateId === undefined) this.trimCache(entry)
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
    let hierarchyMetrics: ComputeMetrics | undefined
    const preparation: ChartComputePreparationContext = { signal: entry.controller.signal, memoryBudgetBytes: available,
      assertCurrent: () => { this.assertEntryCurrent(entry); assertChartReadPermission(workbook, chart, entry.userId) },
      resolveImageResource: id => resolveChartImageResource(workbook, chart, id),
      runChartTask: async task => {
        this.assertEntryCurrent(entry)
        const result = await this.client.submit(task, { documentId: entry.documentId, modelRevision: entry.modelRevision,
          ownership: 'transfer', priority: entry.priority, signal: entry.controller.signal, timeoutMs: 2_147_483_647,
          onProgress: progress => this.publishProgress(entry, progress) }).promise
        this.assertEntryCurrent(entry)
        assertChartReadPermission(workbook, chart, entry.userId)
        if (result.output.kind !== 'chartStatistics') throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表布局任务返回了其他类型的结果')
        hierarchyMetrics = result.metrics
        return result.output
      } }
    let records = await readChartRecordsAsync(workbook, chart, preparation)
    if (entry.authoring?.previous) {
      chart = updateChartDefinition(workbook, entry.authoring.previous, entry.authoring.patch ?? {}, records)
      this.assertEntryCurrent(entry)
      assertChartReadPermission(workbook, chart, entry.userId)
    }
    // Source + projection/normalization/aggregation references are charged
    // before allocating projected numeric arrays. Strings are checked by the reader.
    assertChartPreparationMemory(records.records.length * (records.fields.length * 64 + 192) + 8 * 1024 * 1024, preparation)
    let data = await projectChartRecordsAsync(records, chart, preparation)
    records = { fields: [], records: [] }
    if (entry.authoring) {
      if (entry.authoring.initializeDefaults) chart.options = await initializeChartWaterfallOptionsAsync(chart, data.series, preparation)
      await validateChartSourceDataAsync(chart, data, preparation)
    }
    data = await prepareChartComputeInputs(data, chart, preparation)
    let dataBytes = await estimateChartDataMemory(data, preparation)
    assertChartPreparationMemory(dataBytes, preparation)
    if (chart.type === 'wordCloud' && !entry.sourceOnly) {
      data = await prepareChartWordCloudLayout(data, chart, { ...preparation, memoryBudgetBytes: Math.max(1, available - dataBytes) })
      dataBytes = await estimateChartDataMemory(data, preparation)
      assertChartPreparationMemory(dataBytes, preparation)
    }
    if (chart.type === 'relation' && !entry.sourceOnly) {
      data = await prepareChartRelationLayout(data, chart, { ...preparation, memoryBudgetBytes: Math.max(1, available - dataBytes) })
      dataBytes = await estimateChartDataMemory(data, preparation)
      assertChartPreparationMemory(dataBytes, preparation)
    }
    if (entry.sourceOnly || entry.authoring && !chart.options?.mapping?.useDateAxis) {
      this.assertEntryCurrent(entry)
      return { documentId: entry.documentId, modelRevision: entry.modelRevision, chartId: entry.chartId, chartRevision: chart.revision,
        signature: chartStatisticsSignature(chart), statistics: { preparedData: data }, preparationMs: performance.now() - entry.startedAt,
        metrics: hierarchyMetrics,
        ...(entry.authoring ? { authoredChart: chart } : {}),
        decodingMs: 0, estimatedCacheBytes: dataBytes }
    }
    // A no-statistics chart needs no extra bulk allocation. Keep its control
    // budget valid at an exact admission boundary; every real task is charged
    // against dataBytes + its complete peak before it reaches the Worker.
    const taskContext: ChartComputePreparationContext = { ...preparation, memoryBudgetBytes: Math.max(1, available - dataBytes) }
    const plan = await prepareChartStatisticsTask(data, chart, taskContext)
    const preparationMs = performance.now() - entry.startedAt
    let statistics: ChartComputedStatistics = { preparedData: data }, metrics: ComputeMetrics | undefined = hierarchyMetrics, decodingMs = 0, decodedBytes = 0
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
    if (chart.options?.mapping?.useDateAxis && !entry.settingsOnly) {
      // Calendar ticks depend on final Rust bins and on trend/X-error extents.
      // Validate that actual scene before exposing a computed snapshot, reusing
      // decoded statistics rather than running a second statistical calculation.
      assertChartPreparationMemory(dataBytes + decodedBytes + estimateChartFinalLayoutMemory(data, chart, statistics), preparation)
      layoutChartData(data, chart, statistics)
      this.assertEntryCurrent(entry)
    }
    return { documentId: entry.documentId, modelRevision: entry.modelRevision, chartId: entry.chartId, chartRevision: chart.revision,
      signature: chartStatisticsSignature(chart), statistics, metrics, preparationMs, decodingMs, estimatedCacheBytes: dataBytes + decodedBytes,
      ...(entry.authoring ? { authoredChart: chart } : {}) }
  }

  private attach(entry: Entry, context: ChartStatisticsRequestContext, startedAt: number, timeoutMs: number): ChartStatisticsRequest {
    let settled = false, removeAbort = () => {}, rejectPromise: (error: ComputeError) => void = () => {}
    let deadline: ReturnType<typeof setTimeout> | undefined
    const consumer: Consumer = { progress: context.onProgress, assertCurrent: context.assertCurrent, deadlineAt: startedAt + timeoutMs, fail: error => {
      if (settled) return
      settled = true; detach(); rejectPromise(error)
    } }
    const detach = () => {
      entry.consumers.delete(consumer); removeAbort(); clearTimeout(deadline)
      if (!entry.snapshot && !entry.consumers.size) this.rejectEntry(entry, new ComputeError('COMPUTE_CANCELLED', '图表计算已取消'))
      if (!entry.consumers.size && entry.candidateId !== undefined) this.candidates.delete(entry.candidateId)
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
    const retained = entry.candidateId === undefined ? this.entries.get(entry.documentId)?.get(entry.chartId) : this.candidates.get(entry.candidateId)
    if (!this.isCurrent(entry) || retained !== entry) throw new ComputeError('COMPUTE_STALE_RESULT', '图表来源版本已失效')
  }

  private assertSourcePresent(workbook: WorkbookModel, chart: ChartModel, documentId: string): void {
    const diagnostic = chartSourceNotFoundDiagnostic(workbook, chart)
    if (!diagnostic) return
    const error = new ChartSourceNotFoundError(diagnostic), cached = this.entries.get(documentId)?.get(chart.id)
    if (cached) { cached.snapshot = undefined; this.rejectEntry(cached, error) }
    for (const entry of [...this.candidates.values()]) if (entry.documentId === documentId && entry.chartId === chart.id) {
      entry.snapshot = undefined; this.rejectEntry(entry, error)
    }
    if (cached) this.emit({ documentId, chartId: chart.id, reason: 'configuration' })
    throw error
  }

  private rejectEntry(entry: Entry, error: ComputeError): void {
    if (entry.failure) return
    entry.failure = error
    if (entry.candidateId !== undefined) this.candidates.delete(entry.candidateId)
    else if (this.entries.get(entry.documentId)?.get(entry.chartId) === entry) this.entries.get(entry.documentId)!.delete(entry.chartId)
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
    for (const entry of [...this.candidates.values()].filter(entry => entry.documentId === documentId)) this.rejectEntry(entry,
      new ComputeError(reason === 'disposed' ? 'COMPUTE_DISPOSED' : 'COMPUTE_STALE_RESULT', '图表候选来源已关闭或更新'))
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
      else {
        try { consumer.assertCurrent?.() }
        catch (error) { consumer.fail(normalizeComputeError(error)) }
      }
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
  return JSON.stringify([chart.sheetId, chart.type, chart.source, chart.pivotTableId, chart.categoryField, chart.valueFields, chart.seriesToCategory,
    chart.options, chart.xAxis, chart.yAxis])
}

export function assertChartReadPermission(workbook: WorkbookModel, chart: ChartModel, userId: string): void {
  assertChartSourcePresent(workbook, chart)
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
