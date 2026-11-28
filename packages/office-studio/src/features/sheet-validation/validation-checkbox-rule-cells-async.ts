import type { CellRange, SheetDataValidationRule, SpreadsheetCell, SpreadsheetSheet, WorkbookModel } from '../../domain/model.ts'
import type { OfficeComputeClient } from '../../compute/client.ts'
import type { ComputeMetrics, ComputeProgress } from '../../compute/protocol.ts'
import type { CheckboxCellsTask } from '../../compute/checkbox-cell-protocol.ts'
import { COMPUTE_LIMITS } from '../../compute/protocol.ts'
import { estimateCheckboxCellsMemory } from '../../compute/checkbox-cell-tasks.ts'
import { ComputeError } from '../../compute/errors.ts'
import { sheetRangeContainsCell, sheetRangesIntersect } from '../sheet-workbench/sheet-range-command.ts'
import { createSheetCheckboxRuleCell, planSheetCheckboxRuleCells, prepareSheetCheckboxRuleCells,
  sheetCheckboxRuleCellAction, type SheetCheckboxRuleCellPlan } from './validation-checkbox-rule-cells.ts'

export interface SheetCheckboxRuleCellProgress {
  phase: 'enumerating' | 'classifying' | 'computing' | 'materializing'; processed: number; total: number | null
  compute?: ComputeProgress
}
export interface SheetDataValidationAsyncOptions {
  signal?: AbortSignal; onProgress?(progress: SheetCheckboxRuleCellProgress): void
  /** Admission estimate for detached cells and content history; this does not measure the existing model/browser heap. */
  maxMaterializationBytes?: number
}
export interface SheetCheckboxRuleCellsAsyncContext extends SheetDataValidationAsyncOptions {
  compute: Pick<OfficeComputeClient, 'submit'>; modelRevision: number; assertCurrent(): void
}
export interface SheetCheckboxRuleCellBulkPlan extends SheetCheckboxRuleCellPlan {
  metrics?: { compute: ComputeMetrics; materializationEstimatedBytes: number; preparedBulkBytes: number }
}
/** Provisional bounded policy. The threshold is not a measured performance crossover; compare both paths during unified validation. */
export const SHEET_CHECKBOX_SYNC_CELL_LIMIT = 4_096
const hostYield = () => new Promise<void>(resolve => globalThis.setTimeout(resolve, 0))
const encodedRanges = (ranges: readonly CellRange[]) => new Uint32Array(ranges.flatMap(r => [r.startRow, r.endRow, r.startCol, r.endCol]))

export async function planSheetCheckboxRuleCellsAsync(workbook: WorkbookModel, sheet: SpreadsheetSheet, rule: SheetDataValidationRule,
  previous: SheetDataValidationRule | undefined, userId: string, region: string, context: SheetCheckboxRuleCellsAsyncContext): Promise<SheetCheckboxRuleCellBulkPlan> {
  const materializationBudget = context.maxMaterializationBytes ?? COMPUTE_LIMITS.defaultMemoryBudget
  if (!Number.isSafeInteger(materializationBudget) || materializationBudget < 1 || materializationBudget > 0xffff_ffff) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '复选框物化内存预算必须是有效的正整数')
  }
  const check = () => { if (context.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '复选框批量规则操作已取消'); context.assertCurrent() }
  let lastProgress = -Infinity
  const report = (progress: SheetCheckboxRuleCellProgress, force = false) => {
    if (!context.onProgress || !force && performance.now() - lastProgress < 32) return
    lastProgress = performance.now()
    try { context.onProgress(progress) } catch { /* Progress observers do not own the content transaction. */ }
    check()
  }
  const yieldChecked = async () => { check(); await hostYield(); check() }
  check()
  const prepared = prepareSheetCheckboxRuleCells(workbook, sheet, rule, previous, userId)
  check()
  if (!prepared) return { cells: [], ranges: [] }
  const coveredCells = prepared.ranges.reduce((sum, r) => sum + (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1), 0)
  if (coveredCells <= SHEET_CHECKBOX_SYNC_CELL_LIMIT) {
    const result = planSheetCheckboxRuleCells(workbook, sheet, rule, previous, userId, region)
    const bytes = result.cells.reduce((sum, cell) => sum + (JSON.stringify(cell).length * 2 + 256) * 4, 0)
    if (bytes > materializationBudget) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '复选框单元格和历史预计内存超过预算')
    check(); return result
  }
  if (coveredCells > COMPUTE_LIMITS.maxElements || prepared.ranges.length > 4_096) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '复选框批量范围超过本次计算任务的有界上限')
  }
  const position = (key: string) => {
    const match = /^(\d+):(\d+)$/.exec(key)
    if (!match) return undefined
    const row = Number(match[1]), column = Number(match[2])
    return prepared.ranges.some(r => sheetRangeContainsCell(r, row, column)) ? { row, column } : undefined
  }
  let scanned = 0, existingCount = 0
  // Enumerate keys before touching target scalar getters. The preparation checks complete View ranges first.
  for (const key in sheet.cells) {
    if (!Object.hasOwn(sheet.cells, key)) continue
    if (position(key)) existingCount++
    if (++scanned % 4_096 === 0) { report({ phase: 'enumerating', processed: scanned, total: null }); await yieldChecked() }
  }
  const rowCount = sheet.rowCount ?? 1_048_576, columnCount = sheet.columnCount ?? 16_384
  const merges = sheet.mergedRanges.filter(merge => prepared.ranges.some(r => sheetRangesIntersect(r, merge)))
    .map(merge => ({ ...merge, endRow: Math.min(merge.endRow, rowCount - 1), endCol: Math.min(merge.endCol, columnCount - 1) }))
  const params = { rowCount, columnCount, rangeCount: prepared.ranges.length, mergeCount: merges.length, existingCount, coveredCells,
    maxOutputRuns: Math.min(coveredCells, 65_536) }
  const memory = estimateCheckboxCellsMemory(params)
  if (memory.estimatedPeakBytes > COMPUTE_LIMITS.defaultMemoryBudget) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '复选框稀疏输入和计划预计内存超过预算')
  const task: CheckboxCellsTask = { kind: 'checkboxCells', ranges: encodedRanges(prepared.ranges), merges: encodedRanges(merges),
    positions: new Uint32Array(existingCount * 2), actions: new Uint8Array(existingCount), params }
  let index = 0; scanned = 0
  for (const key in sheet.cells) {
    if (!Object.hasOwn(sheet.cells, key)) continue
    const p = position(key)
    if (p) {
      task.positions[index * 2] = p.row; task.positions[index * 2 + 1] = p.column
      task.actions[index] = sheetCheckboxRuleCellAction(prepared, sheet.cells[key], p.row, p.column); index++
    }
    if (++scanned % 4_096 === 0) { report({ phase: 'classifying', processed: index, total: existingCount }); await yieldChecked() }
  }
  if (index !== existingCount) throw new ComputeError('COMPUTE_STALE_RESULT', '复选框源单元格在准备期间已更改')
  report({ phase: 'classifying', processed: index, total: existingCount }, true); check()
  const job = context.compute.submit(task, { documentId: workbook.id, modelRevision: context.modelRevision, ownership: 'transfer', priority: 'interactive', signal: context.signal,
    onProgress: compute => report({ phase: 'computing', processed: compute.processed, total: compute.total, compute }) })
  const result = await job.promise
  check()
  if (result.documentId !== workbook.id || result.modelRevision !== context.modelRevision || result.output.kind !== 'checkboxCells') {
    throw new ComputeError('COMPUTE_STALE_RESULT', '复选框批量计划不属于当前内容版本')
  }
  const output = result.output, cells: SpreadsheetCell[] = [], ranges: CellRange[] = []
  let materializationEstimatedBytes = 8 * 1024 * 1024, processed = 0
  if (materializationEstimatedBytes + output.changedCells * 1_024 > materializationBudget) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '复选框物化单元格和历史预计内存超过预算')
  }
  for (let i = 0; i < output.runs.length; i += 5) {
    const sr = output.runs[i]!, er = output.runs[i + 1]!, sc = output.runs[i + 2]!, ec = output.runs[i + 3]!, action = output.runs[i + 4]!
    ranges.push({ startRow: sr, endRow: er, startCol: sc, endCol: ec })
    for (let row = sr; row <= er; row++) for (let column = sc; column <= ec; column++) {
      const cell = createSheetCheckboxRuleCell(workbook, sheet, row, column, action === 1 ? prepared.values.unchecked : prepared.values.checked, region)
      materializationEstimatedBytes += (JSON.stringify(cell).length * 2 + 256) * 4
      if (materializationEstimatedBytes > materializationBudget) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '复选框样式和元数据物化预计内存超过预算')
      cells.push(cell)
      if (++processed % 256 === 0) { report({ phase: 'materializing', processed, total: output.changedCells }); await yieldChecked() }
    }
  }
  if (processed !== output.changedCells) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '复选框物化与变化计数不一致')
  report({ phase: 'materializing', processed, total: output.changedCells }, true); check()
  return { cells, ranges, metrics: { compute: result.metrics, materializationEstimatedBytes, preparedBulkBytes: memory.inputBytes } }
}
