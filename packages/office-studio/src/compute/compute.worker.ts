import { ComputeError, normalizeComputeError } from './errors.ts'
import { progressUnits } from './guards.ts'
import { COMPUTE_PROTOCOL_VERSION, isContext, isRecord } from './protocol.ts'
import type { ComputeContext, ComputeOutput, ComputeTask, ComputeWorkerEvent, ComputeWorkerRequest, NumericStatistics } from './protocol.ts'
import { estimateTaskMemory, validateTask } from './tasks.ts'
import { loadComputeWasm } from './wasm.worker.ts'
import type { ComputeWasmModule, WasmStepper } from './wasm.worker.ts'

interface WorkerScope {
  postMessage(message: ComputeWorkerEvent, transfer?: Transferable[]): void
  addEventListener(type: 'message', callback: (event: MessageEvent<unknown>) => void): void
}
const scope: WorkerScope = {
  postMessage: (message, transfer = []) => globalThis.postMessage(message, { transfer }),
  addEventListener: (type, callback) => globalThis.addEventListener(type, callback),
}
let wasm: ComputeWasmModule | null = null
let active: string | null = null
const cancelled = new Set<string>()
const yieldToMessages = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

function checkCancellation(context: ComputeContext): void {
  if (cancelled.has(context.requestId)) throw new ComputeError('COMPUTE_CANCELLED', '计算已取消')
}

function statistics(values: Float64Array): NumericStatistics {
  if (values.length !== 13) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '统计结果槽位版本不匹配')
  const count = (index: number) => values[index] ?? 0
  const finite = (index: number) => Number.isFinite(values[index]) ? values[index]! : null
  return { numericCount: count(0), nonEmptyCount: count(1), emptyCount: count(2), errorCount: count(3), invalidNumericCount: count(4), sum: finite(5), mean: finite(6), min: finite(7), max: finite(8), populationVariance: finite(9), sampleVariance: finite(10), sampleStandardDeviation: finite(11), overflowed: count(12) !== 0 }
}

async function calculate(module: ComputeWasmModule, task: ComputeTask, context: ComputeContext): Promise<ComputeOutput> {
  let lastProgress = -Infinity
  const { total, unit } = progressUnits(task)
  const progress = (ratio: number, force = false) => {
    const now = performance.now()
    if (force || now - lastProgress >= 32) {
      scope.postMessage({ type: 'progress', ...context, phase: ratio >= 1 ? 'completed' : 'computing', ratio, total, unit, processed: Math.round(total * ratio) })
      lastProgress = now
    }
  }
  progress(0, true)
  // Yield before a potentially large linear-memory allocation so a cancellation
  // posted during module initialization can be observed before copying inputs.
  await yieldToMessages(); checkCancellation(context)
  if (task.kind === 'aggregate') {
    const accumulator = new module.NumericAggregator()
    try {
      const chunkSize = 65_536
      for (let offset = 0; offset < task.values.length; offset += chunkSize) {
        checkCancellation(context)
        const end = Math.min(offset + chunkSize, task.values.length)
        accumulator.push(task.values.subarray(offset, end), task.states.subarray(offset, end))
        progress(end / task.values.length)
        await yieldToMessages()
      }
      checkCancellation(context); progress(1, true)
      return { kind: 'aggregate', statistics: statistics(accumulator.result()) }
    } finally { accumulator.free() }
  }
  let stepper: WasmStepper
  if (task.kind === 'sort') stepper = new module.NumericIndexSort(task.values, task.states, task.rowCount, task.directions, task.nullsLast)
  else stepper = new module.NumericIndexFilter(task.values, task.states, task.rowCount, task.operators, task.lower, task.upper, task.match === 'all')
  try {
    let done = false
    while (!done) {
      checkCancellation(context)
      done = stepper.step(task.kind === 'sort' ? 32_768 : 16_384)
      progress(stepper.progress())
      await yieldToMessages()
    }
    checkCancellation(context); progress(1, true)
    return { kind: task.kind, indices: stepper.finish() }
  } finally { stepper.free() }
}

async function handle(request: ComputeWorkerRequest): Promise<void> {
  if (request.type === 'init') {
    try {
      if (request.protocolVersion !== COMPUTE_PROTOCOL_VERSION) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '计算 Worker 协议版本不兼容')
      wasm = await loadComputeWasm(request.moduleUrl)
      scope.postMessage({ type: 'ready', protocolVersion: COMPUTE_PROTOCOL_VERSION, engineVersion: wasm.compute_version() })
    } catch (error) { scope.postMessage({ type: 'init-error', error: normalizeComputeError(error).toJSON() }) }
    return
  }
  if (request.type === 'cancel') { if (active === request.requestId) cancelled.add(request.requestId); return }
  const context: ComputeContext = { requestId: request.requestId, documentId: request.documentId, modelRevision: request.modelRevision }
  if (active) { scope.postMessage({ type: 'error', ...context, error: { code: 'COMPUTE_NOT_READY', message: '计算 Worker 正在执行其他任务' } }); return }
  active = context.requestId
  const started = performance.now()
  try {
    if (!wasm) throw new ComputeError('COMPUTE_NOT_READY', '计算 Worker 尚未就绪')
    validateTask(request.task, request.memoryBudgetBytes)
    const output = await calculate(wasm, request.task, context)
    const transfer: ArrayBuffer[] = []
    if (output.kind !== 'aggregate') {
      if (!(output.indices.buffer instanceof ArrayBuffer)) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '计算结果必须是独立 ArrayBuffer')
      transfer.push(output.indices.buffer)
    }
    scope.postMessage({ type: 'result', ...context, engineVersion: wasm.compute_version(), output, metrics: { elapsedMs: performance.now() - started, ...estimateTaskMemory(request.task) } }, transfer)
  } catch (error) {
    const failure = normalizeComputeError(error)
    if (failure.code === 'COMPUTE_CANCELLED') scope.postMessage({ type: 'cancelled', ...context })
    else scope.postMessage({ type: 'error', ...context, error: failure.toJSON() })
  } finally { cancelled.delete(context.requestId); active = null }
}

scope.addEventListener('message', (event) => {
  const request = event.data
  if (!isRecord(request)) {
    scope.postMessage({ type: 'init-error', error: { code: 'COMPUTE_PROTOCOL_MISMATCH', message: '计算请求必须是结构化对象' } }); return
  }
  if (request.type === 'init' && typeof request.moduleUrl === 'string' && typeof request.protocolVersion === 'number') {
    void handle({ type: 'init', moduleUrl: request.moduleUrl, protocolVersion: request.protocolVersion })
  } else if (request.type === 'cancel' && typeof request.requestId === 'string') {
    void handle({ type: 'cancel', requestId: request.requestId })
  } else if (request.type === 'run' && isContext(request)) {
    const context: ComputeContext = { requestId: request.requestId, documentId: request.documentId, modelRevision: request.modelRevision }
    try {
      if (typeof request.memoryBudgetBytes !== 'number') throw new ComputeError('COMPUTE_INVALID_INPUT', '必须提供计算内存预算')
      validateTask(request.task, request.memoryBudgetBytes)
      void handle({ type: 'run', ...context, task: request.task, memoryBudgetBytes: request.memoryBudgetBytes })
    } catch (error) { scope.postMessage({ type: 'error', ...context, error: normalizeComputeError(error).toJSON() }) }
  } else {
    scope.postMessage({ type: 'init-error', error: { code: 'COMPUTE_PROTOCOL_MISMATCH', message: '计算请求协议不合法' } })
  }
})
