import { ComputeError, isErrorData, normalizeComputeError } from './errors.ts'
import { progressUnits, readProgress, readResult } from './guards.ts'
import { COMPUTE_LIMITS, COMPUTE_PROTOCOL_VERSION, isRecord } from './protocol.ts'
import type { ComputeContext, ComputeProgress, ComputeResult, ComputeTask } from './protocol.ts'
import { cloneTask, estimateTaskMemory, taskViews, transferTaskBuffers, validateTask } from './tasks.ts'
import { createWasmTaskControl, matchesWasmTaskControl } from '../platform/wasm/task-protocol.ts'
import type { WasmTaskControl } from '../platform/wasm/task-protocol.ts'

export type ComputeWorkerPort = Pick<Worker, 'postMessage' | 'terminate' | 'onmessage' | 'onerror'> & Partial<Pick<Worker, 'onmessageerror'>>
export interface ComputeClientOptions {
  workerFactory?: () => ComputeWorkerPort
  wasmModuleUrl?: string
  /** Includes queued inputs, active inputs, WASM copies and result allocation. */
  memoryBudgetBytes?: number
  maxPendingTasks?: number
  /** Kill/recreate a nonresponsive Worker after cooperative cancellation. */
  cancellationGraceMs?: number
  initializationTimeoutMs?: number
}
export interface SubmitComputeOptions {
  documentId: string; modelRevision: number
  /** transfer detaches complete, exclusive caller buffers during submit. */
  ownership: 'copy' | 'transfer'
  /** Interactive work runs before queued background jobs; active work finishes. */
  priority?: 'interactive' | 'background'
  signal?: AbortSignal
  timeoutMs?: number
  onProgress?: (progress: ComputeProgress) => void
}
export interface ComputeJob { requestId: string; promise: Promise<ComputeResult>; cancel(): void }
interface PendingJob extends ComputeContext {
  control: WasmTaskControl
  task: ComputeTask; options: SubmitComputeOptions
  memory: ReturnType<typeof estimateTaskMemory>
  inputLength: number; progressUnits: ReturnType<typeof progressUnits>
  submittedAt: number; dispatchedAt?: number
  resolve(value: ComputeResult): void; reject(error: ComputeError): void
  settled: boolean; deadline?: ReturnType<typeof setTimeout>
  removeAbort?: () => void
}

function validDuration(value: number): boolean { return Number.isFinite(value) && value > 0 && value <= 2_147_483_647 }

/** One bounded calculation at a time. Feature commands must recheck the returned
 * document revision immediately before applying a result as one mutation. */
export class OfficeComputeClient {
  private readonly options: ComputeClientOptions
  private readonly memoryBudget: number
  private readonly maxPending: number
  private worker: ComputeWorkerPort | null = null
  private ready = false
  private disposed = false
  private sequence = 0
  private queue: PendingJob[] = []
  private active: PendingJob | null = null
  private revisions = new Map<string, number>()
  private initTimer?: ReturnType<typeof setTimeout>
  private cancellationTimer?: ReturnType<typeof setTimeout>

  constructor(options: ComputeClientOptions = {}) {
    this.options = { ...options }
    this.memoryBudget = options.memoryBudgetBytes ?? COMPUTE_LIMITS.defaultMemoryBudget
    this.maxPending = options.maxPendingTasks ?? COMPUTE_LIMITS.maxPendingTasks
    if (!Number.isSafeInteger(this.memoryBudget) || this.memoryBudget <= 0 || this.memoryBudget > 0xffff_ffff
      || !Number.isSafeInteger(this.maxPending) || this.maxPending <= 0
      || !validDuration(options.initializationTimeoutMs ?? 15_000) || !validDuration(options.cancellationGraceMs ?? 250)) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '计算服务内存、队列及超时配置不合法')
    }
  }

  setDocumentRevision(documentId: string, modelRevision: number): void {
    this.validateContext(documentId, modelRevision)
    if (this.disposed) throw new ComputeError('COMPUTE_DISPOSED', '计算服务已释放')
    const previous = this.revisions.get(documentId)
    if (previous !== undefined && modelRevision < previous) throw new ComputeError('COMPUTE_INVALID_INPUT', '文档版本不能倒退')
    this.revisions.set(documentId, modelRevision)
    for (const job of [...this.queue, ...(this.active ? [this.active] : [])]) {
      if (job.documentId === documentId && job.modelRevision !== modelRevision) {
        this.cancelJob(job.requestId, new ComputeError('COMPUTE_STALE_RESULT', '模型已更新，旧计算任务已失效'))
      }
    }
  }

  isCurrent(context: Pick<ComputeContext, 'documentId' | 'modelRevision'>): boolean {
    return !this.disposed && this.revisions.get(context.documentId) === context.modelRevision
  }

  forgetDocument(documentId: string): void {
    for (const job of [...this.queue, ...(this.active ? [this.active] : [])]) {
      if (job.documentId === documentId) this.cancelJob(job.requestId, new ComputeError('COMPUTE_STALE_RESULT', '文档已关闭'))
    }
    this.revisions.delete(documentId)
  }

  submit(task: ComputeTask, options: SubmitComputeOptions): ComputeJob {
    const requestId = `compute-${++this.sequence}`
    const submittedAt = performance.now()
    let resolve!: (result: ComputeResult) => void
    let reject!: (error: ComputeError) => void
    const promise = new Promise<ComputeResult>((yes, no) => { resolve = yes; reject = no })
    const handle = { requestId, promise, cancel: () => this.cancelJob(requestId) }
    try {
      if (this.disposed) throw new ComputeError('COMPUTE_DISPOSED', '计算服务已释放')
      if (!options) throw new ComputeError('COMPUTE_INVALID_INPUT', '必须提供计算上下文')
      this.validateContext(options.documentId, options.modelRevision)
      if (options.ownership !== 'copy' && options.ownership !== 'transfer') throw new ComputeError('COMPUTE_INVALID_INPUT', '必须明确缓冲区所有权')
      if (options.priority !== undefined && options.priority !== 'interactive' && options.priority !== 'background') throw new ComputeError('COMPUTE_INVALID_INPUT', '计算优先级不合法')
      if (options.signal && (typeof options.signal.aborted !== 'boolean' || typeof options.signal.addEventListener !== 'function' || typeof options.signal.removeEventListener !== 'function')) {
        throw new ComputeError('COMPUTE_INVALID_INPUT', '取消信号不合法')
      }
      if (options.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '计算已取消')
      if (options.timeoutMs !== undefined && !validDuration(options.timeoutMs)) throw new ComputeError('COMPUTE_INVALID_INPUT', '超时时间超出有效范围')
      validateTask(task, this.memoryBudget)
      const memory = estimateTaskMemory(task)
      const pending = [...this.queue, ...(this.active ? [this.active] : [])]
      // Retained inputs plus the largest concurrent WASM/output allocation.
      const inputBytes = pending.reduce((sum, job) => sum + job.memory.inputBytes, memory.inputBytes)
      const maxExtra = pending.reduce((max, job) => Math.max(max, job.memory.estimatedPeakBytes - job.memory.inputBytes), memory.estimatedPeakBytes - memory.inputBytes)
      if (pending.length >= this.maxPending || inputBytes + maxExtra > this.memoryBudget) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '计算队列已达到任务或内存预算')
      const current = this.revisions.get(options.documentId)
      if (current !== undefined && options.modelRevision < current) throw new ComputeError('COMPUTE_STALE_RESULT', '不能提交旧模型版本的计算')
      // Take ownership immediately, even while the job is waiting for the Worker.
      const ownedTask = options.ownership === 'copy' ? cloneTask(task) : structuredClone(task, { transfer: transferTaskBuffers(task) })
      const control = createWasmTaskControl(ownedTask.kind, requestId, options.documentId, options.modelRevision, {
        maxInputBytes: this.memoryBudget, maxOutputBytes: this.memoryBudget, maxWorkingBytes: this.memoryBudget,
      })
      const job: PendingJob = { requestId, documentId: options.documentId, modelRevision: options.modelRevision, control, task: ownedTask,
        options: { ...options }, resolve, reject, settled: false, memory, inputLength: ownedTask.kind === 'validationListQuery' ? ownedTask.params.rowCount
          : ownedTask.kind === 'checkboxCells' ? ownedTask.params.existingCount : ownedTask.values.length,
        progressUnits: progressUnits(ownedTask), submittedAt }
      this.setDocumentRevision(options.documentId, options.modelRevision)
      if (options.signal) {
        const signal = options.signal
        const abort = () => this.cancelJob(requestId)
        signal.addEventListener('abort', abort, { once: true })
        job.removeAbort = () => signal.removeEventListener('abort', abort)
      }
      const background = job.options.priority === 'background'
      const firstBackground = this.queue.findIndex((queued) => queued.options.priority === 'background')
      if (!background && firstBackground >= 0) this.queue.splice(firstBackground, 0, job)
      else this.queue.push(job)
      // Deadlines include preparing, queueing and module initialization.
      const remainingMs = (options.timeoutMs ?? 30_000) - (performance.now() - submittedAt)
      if (remainingMs <= 0) this.cancelJob(requestId, new ComputeError('COMPUTE_TIMEOUT', '计算任务超时'))
      else job.deadline = setTimeout(() => this.cancelJob(requestId, new ComputeError('COMPUTE_TIMEOUT', '计算任务超时')), remainingMs)
      if (!job.settled) { this.ensureWorker(); this.dispatch() }
    } catch (error) { reject(normalizeComputeError(error)) }
    return handle
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.failWorker(new ComputeError('COMPUTE_DISPOSED', '计算服务已释放'))
    this.revisions.clear()
  }

  private validateContext(documentId: string, revision: number): void {
    if (typeof documentId !== 'string' || !documentId.trim() || !Number.isSafeInteger(revision) || revision < 0) throw new ComputeError('COMPUTE_INVALID_INPUT', '必须提供文档 ID 和非负安全整数模型版本')
  }

  private ensureWorker(): void {
    if (this.worker || this.disposed || this.queue.length === 0) return
    try {
      const worker = this.options.workerFactory?.() ?? new Worker(new URL('./compute.worker.ts', import.meta.url), { type: 'module', name: 'office-compute' })
      this.worker = worker
      worker.onmessage = (event: MessageEvent<unknown>) => { if (this.worker === worker) this.onMessage(event.data) }
      worker.onerror = (event: ErrorEvent) => { if (this.worker === worker) this.failWorker(new ComputeError('COMPUTE_WORKER_FAILED', '计算 Worker 运行失败', event.message)) }
      worker.onmessageerror = () => { if (this.worker === worker) this.failWorker(new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '计算 Worker 消息无法传输')) }
      const moduleUrl = this.options.wasmModuleUrl ?? new URL(`${import.meta.env?.BASE_URL ?? '/'}compute/office_compute.js`, location.href).href
      this.initTimer = setTimeout(() => {
        if (this.worker === worker) this.failWorker(new ComputeError('COMPUTE_WASM_LOAD_FAILED', '计算模块初始化超时'))
      }, this.options.initializationTimeoutMs ?? 15_000)
      worker.postMessage({ type: 'init', protocolVersion: COMPUTE_PROTOCOL_VERSION, moduleUrl })
    } catch (error) { this.failWorker(normalizeComputeError(error, 'COMPUTE_WORKER_FAILED')) }
  }

  private dispatch(): void {
    if (!this.ready || !this.worker || this.active || this.disposed) return
    const job = this.queue.shift()
    if (!job) return
    if (job.settled || !this.isCurrent(job)) { this.rejectJob(job, new ComputeError('COMPUTE_STALE_RESULT', '计算版本已失效')); this.dispatch(); return }
    this.active = job
    job.dispatchedAt = performance.now()
    try {
      this.worker.postMessage({ type: 'run', context: job.control, task: job.task }, transferTaskBuffers(job.task))
    } catch (error) { this.failWorker(normalizeComputeError(error, 'COMPUTE_WORKER_FAILED')) }
  }

  private onMessage(data: unknown): void {
    const protocolFailure = () => this.failWorker(new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '计算 Worker 消息结构或版本不合法'))
    if (!isRecord(data)) { protocolFailure(); return }
    if (data.type === 'ready') {
      if (this.ready || data.protocolVersion !== COMPUTE_PROTOCOL_VERSION || typeof data.engineVersion !== 'string' || !data.engineVersion.trim()) { protocolFailure(); return }
      clearTimeout(this.initTimer); this.ready = true; this.dispatch(); return
    }
    if (data.type === 'init-error') { if (isErrorData(data.error)) this.failWorker(normalizeComputeError(data.error)); else protocolFailure(); return }
    const job = this.active
    if (!job) return
    // Late events from an already completed request cannot complete the active job.
    if (isRecord(data.context) && typeof data.context.requestId === 'string' && data.context.requestId !== job.requestId) return
    if (!matchesWasmTaskControl(data.context, job.control)) { protocolFailure(); return }
    if (data.type === 'progress') {
      const progress = readProgress(data, job, job.progressUnits)
      if (!progress) { protocolFailure(); return }
      if (!job.settled && this.isCurrent(job)) {
        try { job.options.onProgress?.(progress) }
        catch (error) { this.cancelJob(job.requestId, normalizeComputeError(error)) }
      }
      return
    }
    if (data.type === 'cancelled') { this.rejectJob(job, new ComputeError('COMPUTE_CANCELLED', '计算已取消')); this.finishActive(); return }
    if (data.type === 'error') {
      if (!isErrorData(data.error)) { protocolFailure(); return }
      const error = normalizeComputeError(data.error)
      if (error.code === 'COMPUTE_WASM_TRAP') this.failWorker(error)
      else { this.rejectJob(job, error); this.finishActive() }
      return
    }
    if (data.type !== 'result') { protocolFailure(); return }
    const result = readResult(data, job, job.control, job.task, job.memory, job.inputLength)
    if (!result) { protocolFailure(); return }
    if (!this.isCurrent(job)) this.rejectJob(job, new ComputeError('COMPUTE_STALE_RESULT', '旧计算结果已丢弃'))
    else if (!job.settled) {
      result.metrics.queueMs = (job.dispatchedAt ?? performance.now()) - job.submittedAt
      result.metrics.totalMs = performance.now() - job.submittedAt
      result.metrics.preparedInputCopies = job.options.ownership === 'copy' ? taskViews(job.task).length : 0
      result.metrics.preparedInputBytes = job.options.ownership === 'copy' ? job.memory.inputBytes : 0
      result.metrics.ownershipTransfers = job.options.ownership === 'transfer' ? taskViews(job.task).length : 0
      job.settled = true; this.cleanupJob(job); job.resolve(result)
    }
    this.finishActive()
  }

  private cancelJob(requestId: string, error = new ComputeError('COMPUTE_CANCELLED', '计算已取消')): void {
    const queued = this.queue.findIndex((job) => job.requestId === requestId)
    if (queued >= 0) { const [job] = this.queue.splice(queued, 1); if (job) this.rejectJob(job, error); return }
    const job = this.active
    if (!job || job.requestId !== requestId || job.settled) return
    this.rejectJob(job, error)
    const reset = () => {
      if (this.active !== job) return
      this.stopWorker(); this.active = null
      this.ensureWorker(); this.dispatch()
    }
    this.cancellationTimer = setTimeout(reset, this.options.cancellationGraceMs ?? 250)
    try { this.worker?.postMessage({ type: 'cancel', context: job.control }) }
    catch { reset() }
  }

  private finishActive(): void {
    clearTimeout(this.cancellationTimer); this.active = null; this.dispatch()
  }
  private cleanupJob(job: PendingJob): void { clearTimeout(job.deadline); job.removeAbort?.() }
  private rejectJob(job: PendingJob, error: ComputeError): void {
    if (!job.settled) { job.settled = true; this.cleanupJob(job); job.reject(error) }
  }
  private stopWorker(): void {
    clearTimeout(this.initTimer); clearTimeout(this.cancellationTimer)
    const worker = this.worker
    this.worker = null; this.ready = false
    if (worker) { worker.onmessage = null; worker.onerror = null; worker.onmessageerror = null; worker.terminate() }
  }
  private failWorker(error: ComputeError): void {
    if (this.active) this.rejectJob(this.active, error)
    for (const job of this.queue) this.rejectJob(job, error)
    this.queue = []; this.active = null; this.stopWorker()
  }
}
