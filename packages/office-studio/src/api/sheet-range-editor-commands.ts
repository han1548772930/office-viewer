import type { OfficeStudio } from './office-studio.ts'
import type { CellRange } from '../domain/model.ts'
import { applySheetClearCommand, clearModeForCommand, SHEET_CLEAR_COMMAND_IDS, type SheetClearCommandInput } from '../features/sheet-workbench/sheet-clear-command.ts'
import { applySheetMergeCommand, mergeActionForCommand, previewSheetMergeCommand, SHEET_MERGE_COMMAND_IDS,
  type SheetMergeCommandInput, type SheetMergeMode } from '../features/sheet-workbench/sheet-merge-command.ts'
import { sheetRangeCommandInputRanges, type SheetRangeCommandInput } from '../features/sheet-workbench/sheet-range-command.ts'
import { assertSheetRangeChangedValuesValid } from '../features/sheet-workbench/sheet-range-editor-validation.ts'

export function resolveSheetRangeCommandTarget(api: OfficeStudio, input?: SheetRangeCommandInput) {
  const explicitRanges = sheetRangeCommandInputRanges(input)
  const unitId = input?.unitId ?? api.activeWorkbookId, workbook = api.getWorkbookModel(unitId)
  if (!workbook) throw new Error('找不到目标工作簿。')
  const selected = api.getSelection()
  const sheetId = input?.sheetId ?? input?.subUnitId ?? (selected?.unitId === unitId ? selected.sheetId : undefined) ?? workbook.activeSheetId
  const sheet = workbook.sheets.find(candidate => candidate.id === sheetId)
  if (!sheet) throw new Error('找不到目标工作表。')
  const ranges: CellRange[] = explicitRanges ?? (selected?.unitId === unitId && selected.sheetId === sheetId ? api.getSelections().map(range => ({
    startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
  })) : [])
  return { unitId, workbook, sheetId, sheet, ranges }
}
function changedValueEvent(api: OfficeStudio, unitId: string, sheetId: string, ranges: readonly CellRange[]): void {
  api.events.emit('SheetValueChanged', { workbookId: unitId, sheetId, ranges: ranges.map(range => ({
    startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol, unitId, sheetId,
  })) })
}

/** UI, keyboard, context menus and Facade all reach these registered transactions. */
export function registerSheetRangeEditorCommands(api: OfficeStudio): void {
  for (const id of SHEET_CLEAR_COMMAND_IDS) api.registerCommand<SheetClearCommandInput | undefined, boolean>(id, input => {
    const mode = clearModeForCommand(id, input)
    if (!mode || mode === 'none') return false
    const target = resolveSheetRangeCommandTarget(api, input)
    if (!target.ranges.length) return false
    const changed = api.commitWorkbook(target.unitId, mode === 'all' ? '清除全部' : mode === 'format' ? '清除格式' : '清除内容', model => {
      const sheet = model.sheets.find(candidate => candidate.id === target.sheetId)!
      const filteredRows = [...(sheet.autoFilter?.filteredRows ?? []), ...(model.tables ?? []).filter(table => table.sheetId === sheet.id).flatMap(table => table.autoFilter?.filteredRows ?? [])]
      applySheetClearCommand(sheet, target.ranges, mode, api.currentUserId, filteredRows, model)
      if (mode !== 'format') assertSheetRangeChangedValuesValid(target.sheet, model, target.sheetId)
    })
    if (changed && mode !== 'format') changedValueEvent(api, target.unitId, target.sheetId, target.ranges)
    return changed
  })
  for (const id of SHEET_MERGE_COMMAND_IDS) api.registerCommand<SheetMergeCommandInput | SheetMergeMode | undefined, boolean>(id, params => {
    const input = typeof params === 'string' ? { mode: params } : params, action = mergeActionForCommand(id, input)
    if (!action) return false
    const target = resolveSheetRangeCommandTarget(api, input)
    if (!target.ranges.length) return false
    if (input?.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)) throw new TypeError('工作簿版本无效。')
    if (input?.expectedRevision !== undefined && api.getWorkbookRevision(target.unitId) !== input.expectedRevision) throw new Error('工作簿已更改，请重新选择区域后合并。')
    const preview = previewSheetMergeCommand(target.sheet, target.ranges, action, api.currentUserId)
    if (preview.requiresConfirmation) {
      api.events.emit('SheetMergeConfirmationRequested', {
        requestId: crypto.randomUUID(), workbookId: target.unitId, sheetId: target.sheetId, commandId: id,
        baseRevision: api.getWorkbookRevision(target.unitId), discardedContentCount: preview.discardedContentCount,
        input: { ...input, unitId: target.unitId, sheetId: target.sheetId, ranges: target.ranges.map(range => ({ ...range })) },
      })
      return false
    }
    if (!preview.added.length && !preview.removed.length) return false
    const changed = api.commitWorkbook(target.unitId, action.kind === 'unmerge' ? '取消合并' : '合并单元格', model => {
      applySheetMergeCommand(model.sheets.find(candidate => candidate.id === target.sheetId)!, target.ranges, action, api.currentUserId)
      if (action.kind === 'merge') assertSheetRangeChangedValuesValid(target.sheet, model, target.sheetId)
    })
    if (changed && action.kind === 'merge') changedValueEvent(api, target.unitId, target.sheetId, target.ranges)
    return changed
  })
}
