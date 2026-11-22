import type { OfficeStudio } from './office-studio.ts'
import type { ApiSheetSelection, Disposable } from './events.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { planSheetAxisVisibilityCommand, sheetAxisVisibilityActionForCommand, sheetAxisVisibilityCommandInput,
  SHEET_AXIS_VISIBILITY_COMMAND_IDS } from '../features/sheet-workbench/sheet-axis-visibility.ts'

export function registerSheetAxisVisibilityCommands(api: OfficeStudio): Disposable {
  const commands = SHEET_AXIS_VISIBILITY_COMMAND_IDS.map(id => api.registerCommand<unknown, boolean>(id, params => {
    const action = sheetAxisVisibilityActionForCommand(id), input = sheetAxisVisibilityCommandInput(params)
    if (!action) return false
    const target = resolveSheetRangeCommandTarget(api, input)
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    if (input.expectedRevision !== undefined && api.getWorkbookRevision(target.unitId) !== input.expectedRevision) throw new Error('工作簿已更改，请重新选择行列。')
    const filtered = (target.workbook.tables ?? []).filter(table => table.sheetId === target.sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])
    const plan = planSheetAxisVisibilityCommand(target.sheet, target.ranges, action, input, api.currentUserId, filtered)
    if (!plan.changed) return false
    const before = api.getSelectionState(), active = before?.ranges[before.activeRangeIndex], last = plan.selections.at(-1)
    const after: ApiSheetSelection = { ranges: plan.selections.map(range => ({ unitId: target.unitId, sheetId: target.sheetId, startRow: range.startRow,
      endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol, rangeType: action.axis === 'row' ? 1 : 2 })),
      activeRangeIndex: plan.selections.length - 1, ...(last ? { primary: { row: last.startRow, column: last.startCol } } : {}) }
    const changed = api.commitWorkbook(target.unitId, `${action.visible ? '显示' : '隐藏'}${action.axis === 'row' ? '行' : '列'}`, model => {
      Object.assign(model.sheets.find(sheet => sheet.id === target.sheetId)!, plan.changes)
    }, { ...(active?.unitId === target.unitId && active.sheetId === target.sheetId && before ? { selections: { sheetId: target.sheetId, before, after } } : {}) })
    if (changed) api.events.emit('SheetVisibilityChanged', { workbookId: target.unitId, sheetId: target.sheetId, axis: action.axis,
      visible: action.visible, ranges: plan.ranges.map(range => ({ unitId: target.unitId, sheetId: target.sheetId,
        startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol })) })
    return changed
  }))
  return { dispose: () => commands.forEach(command => command.dispose()) }
}
