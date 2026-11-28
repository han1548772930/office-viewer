import { ComputeError, normalizeComputeError } from './errors.ts'
import { isNumericStatistics, progressUnits } from './guards.ts'
import { COMPUTE_PROTOCOL_VERSION, isRecord } from './protocol.ts'
import type { ComputeOutput, ComputeTask, ComputeWorkerEvent, ComputeWorkerRequest } from './protocol.ts'
import { estimateTaskMemory, taskViews, validateTask } from './tasks.ts'
import { loadComputeWasm } from './wasm.worker.ts'
import type { LoadedComputeWasm, WasmAggregate, WasmChartResult, WasmIndexResult, WasmSlicerMemberResult, WasmCheckboxCellsResult, WasmStepper } from './wasm.worker.ts'
import { readChartStatisticsOutput } from './chart-results.ts'
import { readSlicerMembersOutput } from './slicer-member-results.ts'
import { readCheckboxCellsOutput } from './checkbox-cell-results.ts'
import { readValidationListQueryOutput } from './validation-list-query-results.ts'
import { prepareSlicerTextRanks } from './slicer-member.worker.ts'
import {
  matchesWasmTaskControl, readBoundaryCopyStats, requireIndependentOutput, requireMatchingWasmResult,
  requireWasmTaskControl, WasmProtocolError,
} from '../platform/wasm/task-protocol.ts'
import type { BoundaryCopyStats, WasmTaskControl } from '../platform/wasm/task-protocol.ts'

interface WorkerScope {
  postMessage(message: ComputeWorkerEvent, transfer?: Transferable[]): void
  addEventListener(type: 'message', callback: (event: MessageEvent<unknown>) => void): void
}
const scope: WorkerScope = {
  postMessage: (message, transfer = []) => globalThis.postMessage(message, { transfer }),
  addEventListener: (type, callback) => globalThis.addEventListener(type, callback),
}
let wasm: LoadedComputeWasm | null = null
let initializing = false
let active: WasmTaskControl | null = null
const cancelled = new Set<string>()

function checkCancellation(context: WasmTaskControl): void {
  if (cancelled.has(context.requestId)) throw new ComputeError('COMPUTE_CANCELLED', '计算已取消')
}
function isAggregate(value: unknown): value is WasmAggregate {
  return isRecord(value) && ['step', 'progress', 'result', 'free'].every((key) => typeof value[key] === 'function')
}
function isStepper(value: unknown): value is WasmStepper<unknown> {
  return isRecord(value) && ['step', 'progress', 'finish', 'free'].every((key) => typeof value[key] === 'function')
}
function isIndexResult(value: unknown): value is WasmIndexResult {
  return isRecord(value) && ['copy_indices', 'metadata', 'free'].every((key) => typeof value[key] === 'function')
}
function isChartResult(value: unknown): value is WasmChartResult {
  return isRecord(value) && ['copy_values', 'copy_indices', 'metadata', 'free'].every((key) => typeof value[key] === 'function')
}
function isSlicerMemberResult(value: unknown): value is WasmSlicerMemberResult {
  return isRecord(value) && ['copy_indices', 'copy_has_data', 'metadata', 'free'].every((key) => typeof value[key] === 'function')
}
function isCheckboxCellsResult(value: unknown): value is WasmCheckboxCellsResult {
  return isRecord(value) && ['copy_runs', 'metadata', 'free'].every(key => typeof value[key] === 'function')
}
interface Calculated { output: ComputeOutput; boundary: BoundaryCopyStats; buffers: ArrayBuffer[] }

async function calculate(loaded: LoadedComputeWasm, task: ComputeTask, context: WasmTaskControl): Promise<Calculated> {
  const module = loaded.module
  let lastProgress = -Infinity
  const { total, unit } = progressUnits(task)
  const progress = (ratio: number, force = false) => {
    if (!Number.isFinite(ratio) || ratio < 0 || ratio > 1) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', 'WASM 进度超出有效范围')
    const now = performance.now()
    if (force || now - lastProgress >= 32) {
      scope.postMessage({ type: 'progress', context, phase: ratio >= 1 ? 'completed' : 'computing',
        ratio, total, unit, processed: Math.round(total * ratio) })
      lastProgress = now
    }
  }
  progress(0, true)
  // This calls the shared Rust Future/Promise + web-sys macro-task adapter.
  // Large creation/copy steps remain cancellable by the client's kill/recreate.
  await module.compute_yield()
  checkCancellation(context)
  if (task.kind === 'validationListQuery') {
    const candidate: unknown = module.create_validation_list_query(task.offsets, task.textBytes, task.queryBytes, { context, params: task.params })
    if (!isStepper(candidate)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '列表查询缺少协议 2 步进接口')
    try {
      let done = false
      while (!done) { checkCancellation(context); done = candidate.step(16_384); progress(candidate.progress()); await module.compute_yield() }
      checkCancellation(context)
      const result: unknown = candidate.finish()
      if (!isIndexResult(result)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '列表查询缺少结果所有权接口')
      try {
        const indices = result.copy_indices()
        if (!(indices instanceof Uint32Array)) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '列表查询必须返回 Uint32Array')
        const buffer = requireIndependentOutput(indices, loaded.memory), metadata = requireMatchingWasmResult(result.metadata(), context)
        if (metadata.indexLength !== indices.length || metadata.byteLength !== indices.byteLength) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '列表查询与元数据长度不匹配')
        const output = readValidationListQueryOutput({ kind: 'validationListQuery', indices, total: metadata.total }, task.params)
        if (!output) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '列表查询窗口或计数无效')
        progress(1, true)
        return { output, boundary: readBoundaryCopyStats(metadata.transfer), buffers: [buffer] }
      } finally { result.free() }
    } finally { candidate.free() }
  }
  if (task.kind === 'checkboxCells') {
    const candidate: unknown = module.create_checkbox_cells(task.ranges, task.merges, task.positions, task.actions, { context, params: task.params })
    if (!isStepper(candidate)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '复选框计划缺少协议 2 步进接口')
    try {
      let done = false
      while (!done) { checkCancellation(context); done = candidate.step(16_384); progress(candidate.progress()); await module.compute_yield() }
      checkCancellation(context)
      const result: unknown = candidate.finish()
      if (!isCheckboxCellsResult(result)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '复选框计划缺少结果所有权接口')
      try {
        const runs = result.copy_runs()
        if (!(runs instanceof Uint32Array)) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '复选框计划必须返回 Uint32Array')
        const buffer = requireIndependentOutput(runs, loaded.memory), metadata = requireMatchingWasmResult(result.metadata(), context)
        if (metadata.runLength !== runs.length || metadata.byteLength !== runs.byteLength) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '复选框计划与元数据长度不匹配')
        const output = readCheckboxCellsOutput({ kind: 'checkboxCells', runs, changedCells: metadata.changedCells }, task.params)
        if (!output) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '复选框输出范围或变化计数无效')
        progress(1, true)
        return { output, boundary: readBoundaryCopyStats(metadata.transfer), buffers: [buffer] }
      } finally { result.free() }
    } finally { candidate.free() }
  }
  if (task.kind === 'slicerMembers') {
    await prepareSlicerTextRanks(task, async () => { checkCancellation(context); await module.compute_yield(); checkCancellation(context) })
    checkCancellation(context)
    const candidate: unknown = module.create_slicer_members(task.values, task.states, task.textIds, task.availability,
      task.textOffsets, task.textBytes, task.textRanks, { context, params: task.params })
    if (!isStepper(candidate)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '成员查询缺少协议 2 步进接口')
    try {
      let done = false
      while (!done) {
        checkCancellation(context)
        done = candidate.step(16_384)
        progress(candidate.progress())
        await module.compute_yield()
      }
      checkCancellation(context)
      const result: unknown = candidate.finish()
      if (!isSlicerMemberResult(result)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '成员查询缺少协议 2 结果所有权接口')
      try {
        const indices = result.copy_indices(), hasData = result.copy_has_data()
        if (!(indices instanceof Uint32Array) || !(hasData instanceof Uint8Array)) {
          throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '成员查询必须返回源行及可用性 TypedArray')
        }
        const indexBuffer = requireIndependentOutput(indices, loaded.memory), dataBuffer = requireIndependentOutput(hasData, loaded.memory)
        const metadata = requireMatchingWasmResult(result.metadata(), context)
        if (metadata.indexLength !== indices.length || metadata.byteLength !== indices.byteLength + hasData.byteLength) {
          throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '成员窗口与 WASM 元数据长度不匹配')
        }
        const output = readSlicerMembersOutput({ kind: 'slicerMembers', indices, hasData, total: metadata.total,
          dateCount: metadata.dateCount, minDate: metadata.minDate, maxDate: metadata.maxDate }, task.params)
        if (!output) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '成员窗口或日期边界与请求不匹配')
        progress(1, true)
        return { output, boundary: readBoundaryCopyStats(metadata.transfer), buffers: [indexBuffer, dataBuffer] }
      } finally { result.free() }
    } finally { candidate.free() }
  }
  if (task.kind === 'aggregate') {
    const accumulator: unknown = module.create_numeric_aggregate(task.values, task.states, context)
    if (!isAggregate(accumulator)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '聚合任务缺少协议 2 步进接口')
    try {
      let done = false
      while (!done) {
        checkCancellation(context)
        done = accumulator.step(65_536)
        progress(accumulator.progress())
        await module.compute_yield()
      }
      checkCancellation(context)
      const result = requireMatchingWasmResult(accumulator.result(), context)
      if (!isNumericStatistics(result.statistics, task.values.length)) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '聚合结果不是协议 2 统计 DTO')
      progress(1, true)
      return { output: { kind: 'aggregate', statistics: result.statistics }, boundary: readBoundaryCopyStats(result.transfer), buffers: [] }
    } finally { accumulator.free() }
  }
  const candidate: unknown = task.kind === 'chartStatistics'
    ? module.create_chart_statistics(task.values, task.states, task.categoryIds, { context, params: task.params })
    : task.kind === 'sort'
    ? module.create_numeric_sort(task.values, task.states, { context, rowCount: task.rowCount,
      directions: Array.from(task.directions), nullsLast: task.nullsLast })
    : module.create_numeric_filter(task.values, task.states, { context, rowCount: task.rowCount,
      operators: Array.from(task.operators), lower: Array.from(task.lower), upper: Array.from(task.upper), matchAll: task.match === 'all' })
  if (!isStepper(candidate)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '索引任务缺少协议 2 步进接口')
  try {
    let done = false
    while (!done) {
      checkCancellation(context)
      done = candidate.step(task.kind === 'chartStatistics' ? 1 : task.kind === 'sort' ? 32_768 : 16_384)
      progress(candidate.progress())
      await module.compute_yield()
    }
    checkCancellation(context)
    const result: unknown = candidate.finish()
    if (task.kind === 'chartStatistics') {
      if (!isChartResult(result)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '图表结果缺少协议 2 所有权接口')
      try {
        const values = result.copy_values(), indices = result.copy_indices()
        if (!(values instanceof Float64Array) || !(indices instanceof Uint32Array)) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表结果必须为数值及索引 TypedArray')
        const valueBuffer = requireIndependentOutput(values, loaded.memory), indexBuffer = requireIndependentOutput(indices, loaded.memory)
        const metadata = requireMatchingWasmResult(result.metadata(), context)
        if (metadata.valueLength !== values.length || metadata.indexLength !== indices.length
          || metadata.byteLength !== values.byteLength + indices.byteLength) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表结果长度与元数据不匹配')
        const output = readChartStatisticsOutput({ kind: 'chartStatistics', values, indices, segments: metadata.segments }, task.params)
        if (!output) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '图表分段结果与请求不匹配')
        progress(1, true)
        return { output, boundary: readBoundaryCopyStats(metadata.transfer), buffers: [valueBuffer, indexBuffer] }
      } finally { result.free() }
    }
    if (!isIndexResult(result)) throw new ComputeError('COMPUTE_WASM_API_MISSING', '索引结果缺少协议 2 所有权接口')
    try {
      const indices = result.copy_indices()
      if (!(indices instanceof Uint32Array)) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '索引结果必须是 Uint32Array')
      const buffer = requireIndependentOutput(indices, loaded.memory)
      const metadata = requireMatchingWasmResult(result.metadata(), context)
      if (metadata.byteLength !== indices.byteLength) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '索引结果长度与元数据不匹配')
      progress(1, true)
      return { output: { kind: task.kind, indices }, boundary: readBoundaryCopyStats(metadata.transfer), buffers: [buffer] }
    } finally { result.free() }
  } finally { candidate.free() }
}

async function handle(request: ComputeWorkerRequest): Promise<void> {
  if (request.type === 'init') {
    try {
      if (request.protocolVersion !== COMPUTE_PROTOCOL_VERSION) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '计算 Worker 需要协议 2')
      if (initializing || wasm || active) throw new ComputeError('COMPUTE_NOT_READY', '计算 Worker 已初始化或正在使用')
      initializing = true
      wasm = await loadComputeWasm(request.moduleUrl)
      scope.postMessage({ type: 'ready', protocolVersion: COMPUTE_PROTOCOL_VERSION, engineVersion: wasm.module.compute_version() })
    } catch (error) { scope.postMessage({ type: 'init-error', error: normalizeComputeError(error).toJSON() }) }
    finally { initializing = false }
    return
  }
  if (request.type === 'cancel') {
    if (active && matchesWasmTaskControl(request.context, active)) cancelled.add(active.requestId)
    return
  }
  const context = request.context
  if (active) {
    scope.postMessage({ type: 'error', context, error: { code: 'COMPUTE_NOT_READY', message: '计算 Worker 正在执行其他任务' } })
    return
  }
  active = context
  const started = performance.now()
  try {
    if (!wasm) throw new ComputeError('COMPUTE_NOT_READY', '计算 Worker 尚未就绪')
    validateTask(request.task, context.budget.maxWorkingBytes)
    const memory = estimateTaskMemory(request.task)
    const { output, boundary, buffers } = await calculate(wasm, request.task, context)
    checkCancellation(context)
    const outputBytes = output.kind === 'aggregate' ? 0 : output.kind === 'checkboxCells' ? output.runs.byteLength : output.kind === 'chartStatistics'
      ? output.values.byteLength + output.indices.byteLength : output.kind === 'slicerMembers'
      ? output.indices.byteLength + output.hasData.byteLength : output.indices.byteLength
    scope.postMessage({ type: 'result', context, engineVersion: wasm.module.compute_version(), output, metrics: {
      elapsedMs: performance.now() - started, ...memory, boundary,
      transport: { inputTransfers: taskViews(request.task).length, inputBytes: memory.inputBytes,
        outputTransfers: buffers.length, outputBytes,
        structuredResultMessages: output.kind === 'aggregate' || output.kind === 'chartStatistics' || output.kind === 'slicerMembers' || output.kind === 'checkboxCells' ? 1 : 0 },
    } }, buffers)
  } catch (error) {
    const failure = normalizeComputeError(error)
    if (failure.code === 'COMPUTE_CANCELLED') scope.postMessage({ type: 'cancelled', context })
    else scope.postMessage({ type: 'error', context, error: failure.toJSON() })
  } finally { cancelled.delete(context.requestId); active = null }
}

scope.addEventListener('message', (event) => {
  const data = event.data
  let context: WasmTaskControl | undefined
  try {
    if (!isRecord(data)) throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '计算请求必须是结构化对象')
    if (data.type === 'init' && typeof data.moduleUrl === 'string' && data.moduleUrl.trim() && typeof data.protocolVersion === 'number') {
      void handle({ type: 'init', moduleUrl: data.moduleUrl, protocolVersion: data.protocolVersion }); return
    }
    if (data.type !== 'run' && data.type !== 'cancel') throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '计算请求必须使用协议 2')
    if (!isRecord(data.context) || (data.context.kind !== 'aggregate' && data.context.kind !== 'sort' && data.context.kind !== 'filter'
      && data.context.kind !== 'chartStatistics' && data.context.kind !== 'slicerMembers' && data.context.kind !== 'checkboxCells')) {
      throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '计算请求的计算域无效')
    }
    context = requireWasmTaskControl(data.context, data.context.kind)
    if (data.type === 'cancel') { void handle({ type: 'cancel', context }); return }
    validateTask(data.task, context.budget.maxWorkingBytes)
    if (data.task.kind !== context.kind) throw new WasmProtocolError('WASM_PROTOCOL_MISMATCH', '任务上下文与计算参数不匹配')
    void handle({ type: 'run', context, task: data.task })
  } catch (error) {
    const failure = normalizeComputeError(error).toJSON()
    scope.postMessage(context ? { type: 'error', context, error: failure } : { type: 'init-error', error: failure })
  }
})
