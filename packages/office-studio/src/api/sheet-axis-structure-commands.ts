import type { OfficeStudio } from './office-studio.ts'
import type { ApiSheetSelection, Disposable } from './events.ts'
import { resolveSheetRangeCommandTarget } from './sheet-range-editor-commands.ts'
import { planWorkbookAxisStructureCommand, SHEET_AXIS_STRUCTURE_COMMAND_IDS, sheetAxisStructureActionForCommand,
  sheetAxisStructureCommandInput, sheetAxisStructureEdits } from '../features/sheet-workbench/sheet-axis-structure-command.ts'

export function registerSheetAxisStructureCommands(api: OfficeStudio): Disposable {
  const registrations = SHEET_AXIS_STRUCTURE_COMMAND_IDS.map(id => api.registerCommand<unknown, boolean>(id, params => {
    const input = sheetAxisStructureCommandInput(params), action = sheetAxisStructureActionForCommand(id)
    if (!action) return false
    const target = resolveSheetRangeCommandTarget(api, input)
    api.assertCanViewWorksheet(target.unitId, target.sheetId)
    if (input.expectedRevision !== undefined && api.getWorkbookRevision(target.unitId) !== input.expectedRevision) throw new Error('工作簿已更改，请重新选择行列。')
    const filteredRows = [...(target.sheet.autoFilter?.filteredRows ?? []), ...(target.workbook.tables ?? [])
      .filter(table => table.sheetId === target.sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])]
    const edits = sheetAxisStructureEdits(target.sheet, target.ranges, action, input, filteredRows)
    if (!edits.length) return false
    const after = input.direction === 'after' || input.direction === 'down' || input.direction === 'right'
      || input.direction === 1 || input.direction === 2 || input.direction === undefined && action.position === 'after'
    const plan = planWorkbookAxisStructureCommand(target.workbook, target.sheetId, edits, api.currentUserId, after ? 'after' : 'before')
    const currentSelection = api.getSelectionState(), active = currentSelection?.ranges[currentSelection.activeRangeIndex]
    const selected = active?.unitId === target.unitId && active.sheetId === target.sheetId
    const selection: ApiSheetSelection = { ranges: [{ unitId: target.unitId, sheetId: target.sheetId, startRow: plan.selection.startRow, endRow: plan.selection.endRow,
      startColumn: plan.selection.startCol, endColumn: plan.selection.endCol, rangeType: action.axis === 'row' ? 1 : 2 }], activeRangeIndex: 0,
      primary: { row: plan.selection.startRow, column: plan.selection.startCol } }
    const label = `${action.kind === 'insert' ? '插入' : action.kind === 'delete' ? '删除' : action.kind === 'move' ? '移动' : '复制'}${action.axis === 'row' ? '行' : '列'}`
    const changed = api.commitWorkbook(target.unitId, label, model => Object.assign(model, plan.workbook), {
      ...(selected && currentSelection ? { selections: { sheetId: target.sheetId, before: currentSelection, after: selection } } : {}),
    })
    if (changed) {
      api.events.emit('SheetStructureChanged', { workbookId: target.unitId, sheetId: target.sheetId, edits: plan.edits })
      api.events.emit('SheetValueChanged', { workbookId: target.unitId, sheetId: target.sheetId, ranges: plan.affected.map(range => ({
        unitId: target.unitId, sheetId: target.sheetId, startRow: range.startRow, endRow: range.endRow, startColumn: range.startCol, endColumn: range.endCol,
      })) })
    }
    return changed
  }))
  return { dispose: () => registrations.forEach(registration => registration.dispose()) }
}
