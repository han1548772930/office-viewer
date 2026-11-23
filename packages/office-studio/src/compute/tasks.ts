import { ComputeError } from './errors.ts'
import { COMPUTE_LIMITS, isRecord } from './protocol.ts'
import type { ComputeTask } from './protocol.ts'

export function taskViews(task: ComputeTask): Array<Float64Array | Uint8Array> {
  if (task.kind === 'sort') return [task.values, task.states, task.directions]
  if (task.kind === 'filter') return [task.values, task.states, task.operators, task.lower, task.upper]
  return [task.values, task.states]
}

export function estimateTaskMemory(task: ComputeTask): { inputBytes: number; estimatedPeakBytes: number } {
  const inputBytes = taskViews(task).reduce((sum, view) => sum + view.byteLength, 0)
  // JS input + linear-memory input copy + merge/filter index buffers + a returned
  // JS index array, with a fixed allowance for module/runtime and binding copies.
  const indexBytes = task.kind === 'aggregate' ? 0 : task.rowCount * (task.kind === 'sort' ? 12 : 8)
  const wasmInput = inputBytes
  return { inputBytes, estimatedPeakBytes: inputBytes + wasmInput + indexBytes + 8 * 1024 * 1024 }
}

export function validateTask(task: unknown, memoryBudgetBytes: number = COMPUTE_LIMITS.defaultMemoryBudget): asserts task is ComputeTask {
  if (!isRecord(task) || (task.kind !== 'aggregate' && task.kind !== 'sort' && task.kind !== 'filter')
    || !(task.values instanceof Float64Array) || !(task.states instanceof Uint8Array)) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '计算任务必须提供数值及状态 TypedArray')
  }
  const views: Array<Float64Array | Uint8Array> = [task.values, task.states]
  if (task.kind !== 'aggregate') {
    const keys = task.kind === 'sort' ? task.directions : task.operators
    if (!(keys instanceof Uint8Array) || typeof task.rowCount !== 'number' || !Number.isSafeInteger(task.rowCount) || task.rowCount < 0
      || task.rowCount > COMPUTE_LIMITS.maxRows || keys.length === 0 || keys.length > COMPUTE_LIMITS.maxKeys
      || task.rowCount * keys.length !== task.values.length) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '列主序表格维度或关键字不合法')
    }
    views.push(keys)
    if (task.kind === 'filter' && (!(task.lower instanceof Float64Array) || !(task.upper instanceof Float64Array)
      || task.lower.length !== keys.length || task.upper.length !== keys.length || (task.match !== 'all' && task.match !== 'any'))) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '筛选操作数或组合方式不合法')
    }
    if (task.kind === 'filter' && task.lower instanceof Float64Array && task.upper instanceof Float64Array) views.push(task.lower, task.upper)
    if (task.kind === 'sort' && typeof task.nullsLast !== 'boolean') throw new ComputeError('COMPUTE_INVALID_INPUT', '必须明确空值排序位置')
  }
  for (const view of views) {
    if (!(view.buffer instanceof ArrayBuffer)) throw new ComputeError('COMPUTE_INVALID_INPUT', '输入必须使用可转移的 ArrayBuffer；当前不依赖 SharedArrayBuffer')
    try { new Uint8Array(view.buffer) } catch { throw new ComputeError('COMPUTE_INVALID_INPUT', '输入缓冲区已被转移') }
  }
  if (task.values.length !== task.states.length) throw new ComputeError('COMPUTE_INVALID_INPUT', '数值与状态数组长度不一致')
  if (task.values.length > COMPUTE_LIMITS.maxElements) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '输入元素超过计算内核上限')
  const inputBytes = views.reduce((sum, view) => sum + view.byteLength, 0)
  const indexBytes = task.kind === 'aggregate' ? 0 : Number(task.rowCount) * (task.kind === 'sort' ? 12 : 8)
  const wasmInput = inputBytes
  if (!Number.isSafeInteger(memoryBudgetBytes) || memoryBudgetBytes <= 0 || memoryBudgetBytes > 0xffff_ffff
    || inputBytes + wasmInput + indexBytes + 8 * 1024 * 1024 > memoryBudgetBytes) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '任务预计内存超过配置预算')
  }
}

export function cloneTask(task: ComputeTask): ComputeTask {
  const common = { values: task.values.slice(), states: task.states.slice() }
  if (task.kind === 'sort') return { kind: 'sort', ...common, rowCount: task.rowCount, nullsLast: task.nullsLast, directions: task.directions.slice() }
  if (task.kind === 'filter') return { kind: 'filter', ...common, rowCount: task.rowCount, match: task.match, operators: task.operators.slice(), lower: task.lower.slice(), upper: task.upper.slice() }
  return { kind: 'aggregate', ...common }
}

export function transferTaskBuffers(task: ComputeTask): ArrayBuffer[] {
  const buffers = new Set<ArrayBuffer>()
  for (const view of taskViews(task)) {
    const buffer = view.buffer
    if (!(buffer instanceof ArrayBuffer) || view.byteOffset !== 0 || view.byteLength !== buffer.byteLength) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '转移输入要求各 TypedArray 独占完整 ArrayBuffer；切片请使用 copy 模式')
    }
    if (buffers.has(buffer)) throw new ComputeError('COMPUTE_INVALID_INPUT', '转移输入不能在多个列数组间共享缓冲区')
    buffers.add(buffer)
  }
  return Array.from(buffers)
}

