import { createBlankSheet, type SpreadsheetSheet, type WorkbookModel } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canEditProtectedRanges, canEditWorksheetProtection, canViewWorksheetProtection } from '../sheet-protection/range-protection.ts'
import { createUniqueSheetName, validateSheetName } from './sheet-workbench.ts'
import { rebindWorksheetScopedNames, rewriteWorkbookWorksheetReferences } from './sheet-worksheet-references.ts'
import type { CopyWorksheetCommandInput } from './sheet-worksheet-copy.ts'
import { reconcileSlicerConnectionResources } from '../slicer/slicer-connections-lifecycle.ts'
import { invalidateRemovedPivotSourceWorksheet } from '../pivot/services/pivot-source-lifecycle.ts'
import { synchronizeLinkedPivotChartReports } from '../pivot/services/pivot-chart-report-lifecycle.ts'

export const WORKSHEET_LIFECYCLE_COMMANDS = {
  insert: 'sheet.command.insert-sheet', remove: 'sheet.command.remove-sheet', rename: 'sheet.command.set-worksheet-name',
  hidden: 'sheet.command.set-worksheet-hidden', show: 'sheet.command.set-worksheet-show',
  order: 'sheet.command.set-worksheet-order', tabColor: 'sheet.command.set-tab-color',
} as const
export type WorksheetHiddenState = 0 | 1 | 2
export type WorksheetLifecycleCommandId = typeof WORKSHEET_LIFECYCLE_COMMANDS[keyof typeof WORKSHEET_LIFECYCLE_COMMANDS]
export interface WorksheetLifecycleCommandInput extends CopyWorksheetCommandInput {
  index?: number
  sheet?: Partial<SpreadsheetSheet>
  name?: string
  order?: number
  hidden?: WorksheetHiddenState
  value?: string
}
export interface WorksheetLifecyclePlan { workbook: WorkbookModel; sheetId: string; changed: boolean }

export function worksheetHiddenState(sheet: Pick<SpreadsheetSheet, 'visible' | 'metadata'>): WorksheetHiddenState {
  const state = sheet.metadata?.hiddenState
  if (state === 2 || state === 'veryHidden') return 2
  if (state === 1 || state === 'hidden' || sheet.visible === false) return 1
  return 0
}

function targetSheet(workbook: WorkbookModel, id: string, userId: string): SpreadsheetSheet {
  const sheet = workbook.sheets.find(item => item.id === id)
  if (!sheet) throw new Error('工作表不存在。')
  if (!canViewWorksheetProtection(sheet, userId) || !canEditWorksheetProtection(sheet, userId)) throw new Error('当前用户没有修改此工作表的权限。')
  return sheet
}
function result(workbook: WorkbookModel, sheetId: string, changed: boolean): WorksheetLifecyclePlan {
  if (changed) workbook.modified = true
  return { workbook, sheetId, changed }
}
function replaceActiveSheet(candidate: WorkbookModel, original: WorkbookModel, id: string, userId: string): void {
  if (candidate.activeSheetId !== id) return
  const index = original.sheets.findIndex(sheet => sheet.id === id)
  const neighbors = [...original.sheets.slice(index + 1), ...original.sheets.slice(0, index).reverse()]
  const next = neighbors.find(sheet => candidate.sheets.some(item => item.id === sheet.id && worksheetHiddenState(item) === 0)
    && canViewWorksheetProtection(sheet, userId))
  if (!next) throw new Error('没有可切换的可见工作表，操作未提交。')
  candidate.activeSheetId = next.id
}

export function planWorksheetInsert(original: WorkbookModel, options: { userId: string; index?: number; sheet?: Partial<SpreadsheetSheet>; name?: string } ): WorksheetLifecyclePlan {
  const index = options.index ?? original.sheets.length
  if (!Number.isSafeInteger(index) || index < 0 || index > original.sheets.length) throw new RangeError('工作表插入位置无效。')
  const supplied = options.sheet ? cloneModel(options.sheet) : {}
  const requestedName = options.name ?? supplied.name
  if (requestedName !== undefined && typeof requestedName !== 'string') throw new TypeError('工作表名称无效。')
  const validation = validateSheetName(requestedName ?? '工作表', [])
  if (!validation.valid) throw new Error(validation.message)
  const name = requestedName === undefined ? createUniqueSheetName(original.sheets)
    : original.sheets.some(sheet => sheet.name.toLowerCase() === requestedName.toLowerCase())
      ? createUniqueSheetName(original.sheets, requestedName) : requestedName
  const sheet = Object.assign(createBlankSheet(name), Object.fromEntries(Object.entries(supplied).filter(([, value]) => value !== undefined)))
  sheet.name = name
  if (typeof sheet.id !== 'string' || !sheet.id.trim() || original.sheets.some(item => item.id === sheet.id)) throw new Error('工作表 ID 必须唯一。')
  if (!canViewWorksheetProtection(sheet, options.userId)) throw new Error('当前用户没有查看新工作表的权限。')
  const candidate = cloneModel(original)
  candidate.sheets.splice(index, 0, sheet)
  rebindWorksheetScopedNames(candidate.metadata, original.sheets, candidate.sheets)
  for (const item of candidate.sheets) if (item.id !== sheet.id) rebindWorksheetScopedNames(item.metadata, original.sheets, candidate.sheets)
  if (!candidate.sheets.some(item => worksheetHiddenState(item) === 0)) throw new Error('工作簿必须包含可见工作表。')
  if (worksheetHiddenState(sheet) === 0) candidate.activeSheetId = sheet.id
  return result(candidate, sheet.id, true)
}

export function planWorksheetRename(original: WorkbookModel, id: string, name: string, userId: string): WorksheetLifecyclePlan {
  const source = targetSheet(original, id, userId)
  const validation = validateSheetName(name, original.sheets, id)
  if (!validation.valid) throw new Error(validation.message)
  const candidate = cloneModel(original)
  if (name === source.name) return result(candidate, id, false)
  candidate.sheets.find(sheet => sheet.id === id)!.name = name
  rewriteWorkbookWorksheetReferences(candidate, original, { sheetId: id, oldName: source.name, newName: name })
  synchronizeLinkedPivotChartReports(original, candidate)
  return result(candidate, id, true)
}

export function planWorksheetRemove(original: WorkbookModel, id: string, userId: string): WorksheetLifecyclePlan {
  const source = targetSheet(original, id, userId)
  if (original.sheets.length <= 1 || worksheetHiddenState(source) === 0 && original.sheets.filter(sheet => worksheetHiddenState(sheet) === 0).length <= 1) {
    throw new Error('工作簿必须保留至少一个可见工作表。')
  }
  const full = { startRow: 0, endRow: (source.rowCount ?? 1_048_576) - 1, startCol: 0, endCol: (source.columnCount ?? 16_384) - 1 }
  if (!canEditProtectedRanges(source, [full], userId)) throw new Error('工作表包含当前用户无权删除的保护区域。')
  const candidate = cloneModel(original)
  candidate.sheets = candidate.sheets.filter(sheet => sheet.id !== id)
  replaceActiveSheet(candidate, original, id, userId)
  const removedTables = new Set((candidate.tables ?? []).filter(table => table.sheetId === id).map(table => table.id))
  const removedTableNames = new Set((candidate.tables ?? []).filter(table => table.sheetId === id).map(table => table.name.toLowerCase()))
  const removedPivots = new Set((candidate.pivotTables ?? []).filter(pivot => pivot.anchor.subUnitId === id).map(pivot => pivot.id))
  if (candidate.tables) candidate.tables = candidate.tables.filter(table => table.sheetId !== id)
  if (candidate.pivotTables) candidate.pivotTables = candidate.pivotTables.filter(pivot => !removedPivots.has(pivot.id))
  if (candidate.charts) candidate.charts = candidate.charts.filter(chart => chart.sheetId !== id && (!chart.pivotTableId || !removedPivots.has(chart.pivotTableId)))
  if (candidate.slicers) candidate.slicers = candidate.slicers.filter(slicer => slicer.hostSheetId !== id
    && (!slicer.tableId || !removedTables.has(slicer.tableId)))
  rewriteWorkbookWorksheetReferences(candidate, original, { sheetId: id, oldName: source.name }, removedTableNames)
  invalidateRemovedPivotSourceWorksheet(original, candidate, id)
  reconcileSlicerConnectionResources(original, candidate)
  synchronizeLinkedPivotChartReports(original, candidate)
  return result(candidate, id, true)
}

export function planWorksheetHiddenState(original: WorkbookModel, id: string, hidden: WorksheetHiddenState, userId: string): WorksheetLifecyclePlan {
  if (hidden !== 0 && hidden !== 1 && hidden !== 2) throw new RangeError('工作表隐藏状态无效。')
  const source = targetSheet(original, id, userId), state = worksheetHiddenState(source), candidate = cloneModel(original)
  if (state === hidden) return result(candidate, id, false)
  if (state === 0 && hidden !== 0 && original.sheets.filter(sheet => worksheetHiddenState(sheet) === 0).length <= 1) {
    throw new Error('不能隐藏唯一可见的工作表。')
  }
  const sheet = candidate.sheets.find(item => item.id === id)!
  sheet.visible = hidden === 0
  sheet.metadata = { ...sheet.metadata, hiddenState: hidden }
  if (hidden === 0) candidate.activeSheetId = id
  else replaceActiveSheet(candidate, original, id, userId)
  return result(candidate, id, true)
}

export function planWorksheetOrder(original: WorkbookModel, id: string, order: number, userId: string): WorksheetLifecyclePlan {
  targetSheet(original, id, userId)
  if (!Number.isSafeInteger(order) || order < 0 || order >= original.sheets.length) throw new RangeError('工作表顺序无效。')
  const candidate = cloneModel(original), from = candidate.sheets.findIndex(sheet => sheet.id === id)
  if (from === order) return result(candidate, id, false)
  const [sheet] = candidate.sheets.splice(from, 1)
  candidate.sheets.splice(order, 0, sheet!)
  rebindWorksheetScopedNames(candidate.metadata, original.sheets, candidate.sheets)
  for (const item of candidate.sheets) rebindWorksheetScopedNames(item.metadata, original.sheets, candidate.sheets)
  return result(candidate, id, true)
}

export function planWorksheetTabColor(original: WorkbookModel, id: string, value: string, userId: string): WorksheetLifecyclePlan {
  const source = targetSheet(original, id, userId)
  if (typeof value !== 'string' || value.length > 128 || /[\u0000-\u001f]/.test(value)) throw new TypeError('标签颜色无效。')
  const color = value.trim() || undefined, candidate = cloneModel(original)
  if (source.tabColor === color) return result(candidate, id, false)
  candidate.sheets.find(sheet => sheet.id === id)!.tabColor = color
  return result(candidate, id, true)
}

export function planWorksheetLifecycle(original: WorkbookModel, command: WorksheetLifecycleCommandId,
  input: WorksheetLifecycleCommandInput, userId: string): WorksheetLifecyclePlan {
  const id = input.sheetId ?? input.subUnitId ?? original.activeSheetId
  switch (command) {
    case WORKSHEET_LIFECYCLE_COMMANDS.insert: return planWorksheetInsert(original, { userId, index: input.index, sheet: input.sheet, name: input.name })
    case WORKSHEET_LIFECYCLE_COMMANDS.remove: return planWorksheetRemove(original, id, userId)
    case WORKSHEET_LIFECYCLE_COMMANDS.rename:
      if (typeof input.name !== 'string') throw new TypeError('工作表名称无效。')
      return planWorksheetRename(original, id, input.name, userId)
    case WORKSHEET_LIFECYCLE_COMMANDS.hidden:
      if (input.hidden === undefined && worksheetHiddenState(targetSheet(original, id, userId)) !== 0) return result(cloneModel(original), id, false)
      return planWorksheetHiddenState(original, id, input.hidden ?? 1, userId)
    case WORKSHEET_LIFECYCLE_COMMANDS.show: return planWorksheetHiddenState(original, id, 0, userId)
    case WORKSHEET_LIFECYCLE_COMMANDS.order:
      if (input.order === undefined) throw new TypeError('工作表顺序无效。')
      return planWorksheetOrder(original, id, input.order, userId)
    case WORKSHEET_LIFECYCLE_COMMANDS.tabColor:
      if (typeof input.value !== 'string') throw new TypeError('标签颜色无效。')
      return planWorksheetTabColor(original, id, input.value, userId)
    default: throw new Error('工作表命令不存在。')
  }
}
