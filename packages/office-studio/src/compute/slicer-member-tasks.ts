import { ComputeError } from './errors.ts'
import { COMPUTE_LIMITS, isRecord } from './protocol.ts'
import type { SlicerMembersParams, SlicerMembersTask } from './slicer-member-protocol.ts'

const parameterKeys = ['mode', 'rowCount', 'dictionaryCount', 'textByteLength', 'start', 'count', 'locale',
  'sortOrder', 'customListSort', 'crossFilter', 'hideItemsWithNoData']
const unsigned = (value: unknown): value is number => typeof value === 'number' && Number.isSafeInteger(value) && value >= 0

/** Conservative, shared with Rust; includes input ownership, dictionary/Intl
 * preparation, hash entries, merge scratch and independent bounded outputs.
 * This is admission accounting, not a measurement of the source model heap.
 */
export function estimateSlicerMembersMemory(params: SlicerMembersParams): { inputBytes: number; outputBytes: number; estimatedPeakBytes: number } {
  const inputBytes = params.rowCount * (params.mode === 'canvas' ? 18 : 14) + (params.dictionaryCount * 2 + 1) * 4 + params.textByteLength
  const outputBytes = params.mode !== 'metadata' ? Math.min(params.count, params.rowCount) * 5 : 0
  const estimatedPeakBytes = inputBytes * 2 + params.rowCount * 128 + params.dictionaryCount * 160
    + params.textByteLength * 2 + outputBytes * 2 + 8 * 1024 * 1024
  return { inputBytes, outputBytes, estimatedPeakBytes }
}

export function validateSlicerMembersTask(task: Record<string, unknown>): asserts task is Record<string, unknown> & SlicerMembersTask {
  const params = task.params
  if (!isRecord(params) || Object.keys(params).length !== parameterKeys.length
    || Object.keys(params).some(key => !parameterKeys.includes(key))
    || (params.mode !== 'items' && params.mode !== 'canvas' && params.mode !== 'metadata')
    || !unsigned(params.rowCount) || params.rowCount > COMPUTE_LIMITS.maxRows
    || !unsigned(params.dictionaryCount) || params.dictionaryCount > params.rowCount * (params.mode === 'canvas' ? 2 : 1)
    || !unsigned(params.textByteLength) || params.textByteLength > 0xffff_ffff || params.textByteLength % 2 !== 0
    || !unsigned(params.start) || !unsigned(params.count) || params.count === 0 || params.count > COMPUTE_LIMITS.maxRows
    || typeof params.locale !== 'string' || !params.locale.trim() || params.locale.length > 128
    || (params.sortOrder !== 'ascending' && params.sortOrder !== 'descending')
    || typeof params.customListSort !== 'boolean' || typeof params.hideItemsWithNoData !== 'boolean'
    || typeof params.crossFilter !== 'string' || !['none', 'showItemsWithDataAtTop', 'showItemsWithNoData'].includes(params.crossFilter)) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器 typed 成员查询参数不合法')
  }
  if (!(task.values instanceof Float64Array) || !(task.states instanceof Uint8Array)
    || !(task.textIds instanceof Uint32Array) || !(task.availability instanceof Uint8Array)
    || !(task.textOffsets instanceof Uint32Array) || !(task.textBytes instanceof Uint8Array) || !(task.textRanks instanceof Uint32Array)
    || task.values.length !== params.rowCount || task.states.length !== params.rowCount
    || task.textIds.length !== params.rowCount * (params.mode === 'canvas' ? 2 : 1) || task.availability.length !== params.rowCount
    || task.textOffsets.length !== params.dictionaryCount + 1 || task.textBytes.length !== params.textByteLength
    || task.textRanks.length !== params.dictionaryCount) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器 typed 成员列和字典维度不一致')
  }
  if (params.mode === 'metadata' && (params.dictionaryCount !== 0 || params.textByteLength !== 0)) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '日期 metadata 不携带文本字典')
  }
}
