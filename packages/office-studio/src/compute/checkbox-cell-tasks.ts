import { ComputeError } from './errors.ts'
import { COMPUTE_LIMITS, isRecord } from './protocol.ts'
import type { CheckboxCellsParams, CheckboxCellsTask } from './checkbox-cell-protocol.ts'
const keys = ['rowCount', 'columnCount', 'rangeCount', 'mergeCount', 'existingCount', 'coveredCells', 'maxOutputRuns']
const unsigned = (value: unknown): value is number => typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 && value <= 0xffff_ffff

/** Admission accounting shared with Rust. Includes independent input/output copies, sparse lookup, merge fragments and row-run lookup. */
export function estimateCheckboxCellsMemory(params: CheckboxCellsParams) {
  const inputBytes = (params.rangeCount + params.mergeCount) * 16 + params.existingCount * 9
  const outputBytes = params.maxOutputRuns * 20
  const estimatedPeakBytes = inputBytes * 2 + params.existingCount * 96 + params.maxOutputRuns * 160
    + (params.mergeCount + params.rangeCount) * 32 + 8 * 1024 * 1024
  return { inputBytes, outputBytes, estimatedPeakBytes }
}
export function validateCheckboxCellsTask(task: Record<string, unknown>): asserts task is Record<string, unknown> & CheckboxCellsTask {
  const p = task.params
  if (!isRecord(p) || Object.keys(p).length !== keys.length || Object.keys(p).some(key => !keys.includes(key)) || !keys.every(key => unsigned(p[key]))
    || Number(p.rowCount) < 1 || Number(p.rowCount) > 1_048_576 || Number(p.columnCount) < 1 || Number(p.columnCount) > 16_384
    || Number(p.rangeCount) > 4_096 || Number(p.mergeCount) > COMPUTE_LIMITS.maxElements || Number(p.existingCount) > COMPUTE_LIMITS.maxElements
    || Number(p.coveredCells) > COMPUTE_LIMITS.maxElements || Number(p.existingCount) > Number(p.coveredCells)
    || Number(p.maxOutputRuns) > Number(p.coveredCells) || Number(p.coveredCells) > 0 && Number(p.maxOutputRuns) === 0
    || !(task.ranges instanceof Uint32Array) || !(task.merges instanceof Uint32Array) || !(task.positions instanceof Uint32Array) || !(task.actions instanceof Uint8Array)
    || task.ranges.length !== Number(p.rangeCount) * 4 || task.merges.length !== Number(p.mergeCount) * 4
    || task.positions.length !== Number(p.existingCount) * 2 || task.actions.length !== p.existingCount) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '复选框批量范围、稀疏位置或控制参数不合法')
  }
}
