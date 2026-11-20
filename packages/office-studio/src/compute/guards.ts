import { isRecord } from './protocol.ts'
import type { ComputeContext, ComputeMetrics, ComputeOutput, ComputeProgress, ComputeResult, ComputeTask, NumericStatistics } from './protocol.ts'

function isCount(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}
function isFiniteOrNull(value: unknown): value is number | null {
  return value === null || (typeof value === 'number' && Number.isFinite(value))
}
function isStatistics(value: unknown, length: number): value is NumericStatistics {
  if (!isRecord(value) || !isCount(value.numericCount) || !isCount(value.nonEmptyCount) || !isCount(value.emptyCount)
    || !isCount(value.errorCount) || !isCount(value.invalidNumericCount) || typeof value.overflowed !== 'boolean') return false
  if (value.nonEmptyCount + value.emptyCount !== length
    || value.numericCount + value.errorCount + value.invalidNumericCount > value.nonEmptyCount) return false
  return ['sum', 'mean', 'min', 'max', 'populationVariance', 'sampleVariance', 'sampleStandardDeviation'].every((key) => isFiniteOrNull(value[key]))
}
function isMetrics(value: unknown, expected: Pick<ComputeMetrics, 'inputBytes' | 'estimatedPeakBytes'>): value is ComputeMetrics {
  return isRecord(value) && typeof value.elapsedMs === 'number' && Number.isFinite(value.elapsedMs) && value.elapsedMs >= 0
    && value.inputBytes === expected.inputBytes && value.estimatedPeakBytes === expected.estimatedPeakBytes
}

export function progressUnits(task: ComputeTask): Pick<ComputeProgress, 'total' | 'unit'> {
  if (task.kind === 'aggregate') return { total: task.values.length, unit: 'values' }
  if (task.kind === 'filter') return { total: task.rowCount, unit: 'rows' }
  return { total: task.rowCount < 2 ? 0 : task.rowCount * Math.ceil(Math.log2(task.rowCount)), unit: 'merge-writes' }
}

export function readProgress(data: Record<string, unknown>, context: ComputeContext, expected: Pick<ComputeProgress, 'total' | 'unit'>): ComputeProgress | null {
  if (data.type !== 'progress' || typeof data.ratio !== 'number' || !Number.isFinite(data.ratio) || data.ratio < 0 || data.ratio > 1
    || !isCount(data.processed) || data.total !== expected.total || data.unit !== expected.unit
    || data.processed > expected.total || (data.phase !== 'initializing' && data.phase !== 'computing' && data.phase !== 'completed')) return null
  return { type: 'progress', ...context, phase: data.phase, ratio: data.ratio, processed: data.processed, ...expected }
}

export function readResult(data: Record<string, unknown>, context: ComputeContext, task: ComputeTask,
  memory: Pick<ComputeMetrics, 'inputBytes' | 'estimatedPeakBytes'>, inputLength: number): ComputeResult | null {
  if (data.type !== 'result' || typeof data.engineVersion !== 'string' || !data.engineVersion.trim() || !isMetrics(data.metrics, memory)
    || !isRecord(data.output) || data.output.kind !== task.kind) return null
  let output: ComputeOutput
  if (data.output.kind === 'aggregate') {
    if (!isStatistics(data.output.statistics, inputLength)) return null
    output = { kind: 'aggregate', statistics: data.output.statistics }
  } else {
    const indices = data.output.indices
    if ((data.output.kind !== 'sort' && data.output.kind !== 'filter') || task.kind === 'aggregate' || !(indices instanceof Uint32Array)
      || !(indices.buffer instanceof ArrayBuffer) || indices.byteOffset !== 0 || indices.byteLength !== indices.buffer.byteLength
      || (task.kind === 'sort' ? indices.length !== task.rowCount : indices.length > task.rowCount)) return null
    // Index content is computed and tested inside Rust. Traversing millions of
    // indices again here would bring the heavy work back to the main thread.
    output = { kind: data.output.kind, indices }
  }
  return { type: 'result', ...context, engineVersion: data.engineVersion, output,
    metrics: { elapsedMs: data.metrics.elapsedMs, inputBytes: memory.inputBytes, estimatedPeakBytes: memory.estimatedPeakBytes } }
}
