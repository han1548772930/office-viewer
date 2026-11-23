import { isRecord } from './protocol.ts'
import type { ComputeContext, ComputeMetrics, ComputeOutput, ComputeProgress, ComputeResult, ComputeTask, NumericStatistics } from './protocol.ts'
import { taskViews } from './tasks.ts'
import { isUnsignedCount, isWasmTaskMetrics, matchesWasmTaskControl } from '../platform/wasm/task-protocol.ts'
import type { WasmTaskControl } from '../platform/wasm/task-protocol.ts'

function isFiniteOrNull(value: unknown): value is number | null {
  return value === null || (typeof value === 'number' && Number.isFinite(value))
}
export function isNumericStatistics(value: unknown, length: number): value is NumericStatistics {
  if (!isRecord(value) || !isUnsignedCount(value.numericCount) || !isUnsignedCount(value.nonEmptyCount) || !isUnsignedCount(value.emptyCount)
    || !isUnsignedCount(value.errorCount) || !isUnsignedCount(value.invalidNumericCount) || typeof value.overflowed !== 'boolean') return false
  if (value.nonEmptyCount + value.emptyCount !== length
    || value.numericCount + value.errorCount + value.invalidNumericCount > value.nonEmptyCount) return false
  return ['sum', 'mean', 'min', 'max', 'populationVariance', 'sampleVariance', 'sampleStandardDeviation'].every((key) => isFiniteOrNull(value[key]))
}
function isMetrics(value: unknown, expected: Pick<ComputeMetrics, 'inputBytes' | 'estimatedPeakBytes'>): value is ComputeMetrics {
  return isWasmTaskMetrics(value) && isRecord(value) && value.inputBytes === expected.inputBytes
    && value.estimatedPeakBytes === expected.estimatedPeakBytes
}
export function progressUnits(task: ComputeTask): Pick<ComputeProgress, 'total' | 'unit'> {
  if (task.kind === 'aggregate') return { total: task.values.length, unit: 'values' }
  if (task.kind === 'filter') return { total: task.rowCount, unit: 'rows' }
  return { total: task.rowCount < 2 ? 0 : task.rowCount * Math.ceil(Math.log2(task.rowCount)), unit: 'merge-writes' }
}
export function readProgress(data: Record<string, unknown>, context: ComputeContext, expected: Pick<ComputeProgress, 'total' | 'unit'>): ComputeProgress | null {
  if (data.type !== 'progress' || typeof data.ratio !== 'number' || !Number.isFinite(data.ratio) || data.ratio < 0 || data.ratio > 1
    || !isUnsignedCount(data.processed) || data.total !== expected.total || data.unit !== expected.unit
    || data.processed > expected.total || (data.phase !== 'initializing' && data.phase !== 'computing' && data.phase !== 'completed')) return null
  return { type: 'progress', ...context, phase: data.phase, ratio: data.ratio, processed: data.processed, ...expected }
}
export function readResult(data: Record<string, unknown>, context: ComputeContext, control: WasmTaskControl, task: ComputeTask,
  memory: Pick<ComputeMetrics, 'inputBytes' | 'estimatedPeakBytes'>, inputLength: number): ComputeResult | null {
  if (data.type !== 'result' || !matchesWasmTaskControl(data.context, control)
    || typeof data.engineVersion !== 'string' || !data.engineVersion.trim() || !isMetrics(data.metrics, memory)
    || !isRecord(data.output) || data.output.kind !== task.kind) return null
  const boundary = data.metrics.boundary
  const transport = data.metrics.transport
  if (boundary.inputBytes !== inputLength * 9 || boundary.inputCopies !== (inputLength ? 2 : 0)
    || boundary.jsonConversions !== 0 || boundary.rustBufferClones !== 0
    || transport.inputBytes !== memory.inputBytes || transport.inputTransfers !== taskViews(task).length) return null
  let output: ComputeOutput
  let outputBytes = 0
  if (data.output.kind === 'aggregate') {
    if (!isNumericStatistics(data.output.statistics, inputLength)) return null
    output = { kind: 'aggregate', statistics: data.output.statistics }
  } else {
    const indices = data.output.indices
    if ((data.output.kind !== 'sort' && data.output.kind !== 'filter') || task.kind === 'aggregate' || !(indices instanceof Uint32Array)
      || !(indices.buffer instanceof ArrayBuffer) || indices.byteOffset !== 0 || indices.byteLength !== indices.buffer.byteLength
      || (task.kind === 'sort' ? indices.length !== task.rowCount : indices.length > task.rowCount)) return null
    // Correctness and range membership are tested in the Rust kernel. The main
    // thread validates ownership and dimensions without rescanning all indices.
    output = { kind: data.output.kind, indices }
    outputBytes = indices.byteLength
  }
  if (boundary.outputBytes !== outputBytes || boundary.outputCopies !== (outputBytes ? 1 : 0)
    || transport.outputBytes !== outputBytes || transport.outputTransfers !== (task.kind === 'aggregate' ? 0 : 1)
    || transport.structuredResultMessages !== (task.kind === 'aggregate' ? 1 : 0)) return null
  return { type: 'result', ...context, context: control, engineVersion: data.engineVersion, output,
    metrics: { elapsedMs: data.metrics.elapsedMs, inputBytes: memory.inputBytes, estimatedPeakBytes: memory.estimatedPeakBytes,
      boundary, transport } }
}
