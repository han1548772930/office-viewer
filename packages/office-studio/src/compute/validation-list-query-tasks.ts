import { ComputeError } from './errors.ts'
import { COMPUTE_LIMITS, isRecord } from './protocol.ts'
import type { ValidationListQueryParams, ValidationListQueryTask } from './validation-list-query-protocol.ts'

const parameterKeys = ['rowCount', 'textByteLength', 'queryByteLength', 'start', 'count'] as const
const unsigned = (value: unknown): value is number => typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 && value <= 0xffff_ffff

/** Shared admission estimate: JS/linear-memory inputs, KMP prefix table,
 * retained window and independent output. Source option/model heaps are separate.
 */
export function estimateValidationListQueryMemory(params: ValidationListQueryParams): { inputBytes: number; outputBytes: number; estimatedPeakBytes: number } {
  const inputBytes = (params.rowCount + 1) * 4 + params.textByteLength + params.queryByteLength
  const outputBytes = Math.min(Math.max(0, params.rowCount - params.start), params.count) * 4
  const estimatedPeakBytes = inputBytes * 2 + params.queryByteLength * 2 + outputBytes * 2 + 8 * 1024 * 1024
  return { inputBytes, outputBytes, estimatedPeakBytes }
}

export function validateValidationListQueryTask(task: Record<string, unknown>): asserts task is Record<string, unknown> & ValidationListQueryTask {
  const params = task.params
  if (!isRecord(params) || Object.keys(params).length !== parameterKeys.length
    || Object.keys(params).some(key => !parameterKeys.some(known => known === key))
    || !unsigned(params.rowCount) || params.rowCount > COMPUTE_LIMITS.maxRows
    || !unsigned(params.textByteLength) || params.textByteLength % 2 !== 0 || params.textByteLength > COMPUTE_LIMITS.maxElements * 2
    || !unsigned(params.queryByteLength) || params.queryByteLength % 2 !== 0 || params.queryByteLength > COMPUTE_LIMITS.maxElements * 2
    || !unsigned(params.start) || !unsigned(params.count) || params.count === 0 || params.count > COMPUTE_LIMITS.maxRows
    || !(task.offsets instanceof Uint32Array) || task.offsets.length !== params.rowCount + 1
    || !(task.textBytes instanceof Uint8Array) || task.textBytes.length !== params.textByteLength
    || !(task.queryBytes instanceof Uint8Array) || task.queryBytes.length !== params.queryByteLength) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '列表查询参数与 UTF-16 字典维度不一致')
  }
}
