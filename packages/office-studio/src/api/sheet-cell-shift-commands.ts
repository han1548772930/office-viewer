import type { OfficeStudio } from './office-studio.ts'
import type { ApiSheetSelection, Disposable } from './events.ts'
import { cloneModel } from '../domain/serialization.ts'
import { toExcelWorkbookView } from '../services/excelViewAdapter.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { assertSheetCellShiftConfirmationAllowed, planWorkbookCellShiftCommand, SHEET_CELL_SHIFT_COMMAND_IDS, SHEET_CELL_SHIFT_CONFIRM_COMMANDS,
  sheetCellShiftCommandInput, sheetCellShiftCommandRange, sheetCellShiftDirectionForCommand, sheetCellShiftEdit,
  type SheetCellShiftCommandInput } from '../features/sheet-workbench/sheet-cell-shift-command.ts'

export function registerSheetCellShiftCommands(api: OfficeStudio): Disposable {
  const registrations = SHEET_CELL_SHIFT_COMMAND_IDS.map(commandId => api.registerCommand<unknown, boolean>(commandId, params => {
    const input = sheetCellShiftCommandInput(params), direction = sheetCellShiftDirectionForCommand(commandId)
    if (!direction) return false
    const target = resolveSheetRangeCommandTarget(api, input), baseRevision = api.getWorkbookRevision(target.unitId)
    const historyGeneration = api.editor.history.captureSavepoint(target.unitId).generation, userId = api.currentUserId
    const assertCurrent = () => {
      if (api.isDisposed || api.getWorkbookModel(target.unitId) !== target.workbook || api.getWorkbookRevision(target.unitId) !== baseRevision
        || api.editor.history.captureSavepoint(target.unitId).generation !== historyGeneration || api.currentUserId !== userId) throw new Error('工作簿内容或历史已更改，请重新执行单元格移位。')
      if (input.expectedRevision !== undefined && input.expectedRevision !== baseRevision
        || input.expectedHistoryGeneration !== undefined && input.expectedHistoryGeneration !== historyGeneration) throw new Error('工作簿内容或历史已更改，请重新选择单元格移位范围。')
      if (input.host && (input.host.unitId !== target.unitId || input.host.sheetId !== target.sheetId || api.activeWorkbookId !== input.host.unitId
        || api.editor.state.activeKind !== 'workbook' || target.workbook.activeSheetId !== input.host.sheetId)) throw new Error('单元格移位宿主已切换，请重新选择范围。')
    }
    assertCurrent()
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    api.assertCanEditWorksheet(target.unitId, target.sheetId)
    const currentSelection = api.getSelectionState(target.unitId, target.sheetId)
    const range = sheetCellShiftCommandRange(input, target.ranges, currentSelection?.activeRangeIndex)
    if (!range) return false
    if (commandId === SHEET_CELL_SHIFT_CONFIRM_COMMANDS[direction]) assertSheetCellShiftConfirmationAllowed(target.workbook, target.sheetId, sheetCellShiftEdit(range, direction))
    const plan = planWorkbookCellShiftCommand(target.workbook, target.sheetId, range, direction, userId, input.dismantlePartialMerges, target.ranges)
    for (const [sheetId, ranges] of Object.entries(plan.affected)) {
      api.assertCanViewWorksheet(target.unitId, sheetId); api.assertCanEditWorksheet(target.unitId, sheetId)
      for (const range of ranges) api.assertCanEditRange(target.unitId, sheetId, range)
    }
    assertCurrent()
    if (plan.requiresConfirmation) {
      api.publishObserverEvent('SheetCellShiftConfirmationRequested', () => ({ requestId: crypto.randomUUID(), workbookId: target.unitId,
        sheetId: target.sheetId, commandId, baseRevision, mergeRanges: cloneModel(plan.mergeRanges), input: {
          ...input, unitId: target.unitId, sheetId: target.sheetId, range: cloneModel(range), ranges: [cloneModel(range)],
          expectedRevision: baseRevision, expectedHistoryGeneration: historyGeneration,
        } }))
      return false
    }
    const after: ApiSheetSelection = { ranges: [{ unitId: target.unitId, sheetId: target.sheetId, startRow: range.startRow, endRow: range.endRow,
      startColumn: range.startCol, endColumn: range.endCol, rangeType: 0 }], activeRangeIndex: 0,
      primary: { row: range.startRow, column: range.startCol } }
    const changed = api.commitWorkbook(target.unitId, plan.edit.kind === 'insert' ? '插入并移动单元格' : '删除并移动单元格', model => {
      assertCurrent(); Object.assign(model, plan.workbook)
    }, { preserveStyleInheritance: true, expectedRevision: baseRevision, expectedHistoryGeneration: historyGeneration,
      ...(currentSelection ? { selections: { sheetId: target.sheetId, before: currentSelection, after } } : {}),
      prepareView: model => { toExcelWorkbookView(model) },
    })
    if (changed) {
      api.publishObserverEvent('SheetStructureChanged', () => ({ workbookId: target.unitId, sheetId: target.sheetId, edits: [cloneModel(plan.edit)] }))
      for (const [sheetId, ranges] of Object.entries(plan.affected)) api.publishObserverEvent('SheetValueChanged', () => ({ workbookId: target.unitId, sheetId,
        ranges: ranges.map(range => ({ unitId: target.unitId, sheetId, startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol })) }))
    }
    return changed
  }))
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}

export { SHEET_CELL_SHIFT_COMMANDS, SHEET_CELL_SHIFT_CONFIRM_COMMANDS, SHEET_CELL_SHIFT_COMMAND_IDS, SheetCellShiftDimensions,
  type SheetCellShiftCommandInput, type SheetCellShiftConfirmationRequest, type SheetCellShiftDimension } from '../features/sheet-workbench/sheet-cell-shift-command.ts'
