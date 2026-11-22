import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange } from '../../domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { sheetRangeCommandInputRanges } from '../../features/sheet-workbench/sheet-range-command.ts'
import type { SheetSelectionCell } from '../../features/sheet-workbench/sheet-selection.ts'
import { planSheetAxisVisibilityCommand, sheetAxisVisibilityActionForCommand, sheetAxisVisibilityCommandInput,
  type SheetAxisVisibilityCommandInput } from '../../features/sheet-workbench/sheet-axis-visibility.ts'

interface SheetAxisVisibilityUiContext {
  studio(): OfficeStudio | undefined; workbook(): ExcelWorkbookView; sheetId(): string; ranges(): CellRange[]
  userId(): string; region(): string; active(): boolean; finishEdit(): boolean
  prepareView(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void; selection(ranges: CellRange[], primary?: SheetSelectionCell): void; changed(): void
}
export function useSheetAxisVisibility(context: SheetAxisVisibilityUiContext) {
  function runAxisVisibilityCommand(id: string, params?: unknown): boolean {
    if (!context.active() || !context.finishEdit()) return false
    const action = sheetAxisVisibilityActionForCommand(id)
    if (!action) return false
    const input = sheetAxisVisibilityCommandInput(params), view = context.workbook(), studio = context.studio()
    const sheetId = input.sheetId ?? input.subUnitId ?? context.sheetId(), ranges = sheetRangeCommandInputRanges(input) ?? context.ranges()
    const request: SheetAxisVisibilityCommandInput = { ...input, unitId: input.unitId ?? view.id, sheetId, ranges }
    if (studio) {
      request.unitId ??= studio.activeWorkbookId
      const changed = studio.executeCommandSync<SheetAxisVisibilityCommandInput, boolean>(id, request)
      if (!changed) return false
      if (request.unitId === view.id) {
        const model = studio.getWorkbookModel(request.unitId)
        if (model) context.applyView(toExcelWorkbookView(model))
        const selection = studio.getSelectionState(), active = selection?.ranges[selection.activeRangeIndex]
        if (sheetId === context.sheetId() && active?.unitId === request.unitId && active.sheetId === sheetId && selection) context.selection(selection.ranges.map(range => ({
          startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn,
        })), selection.primary)
        context.changed()
      }
      return true
    }
    if (input.unitId !== undefined && input.unitId !== view.id) throw new Error('找不到目标工作簿。')
    const model = applyExcelWorkbookView(createBlankWorkbook(), view, context.region()), sheet = model.sheets.find(item => item.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
    const filtered = (model.tables ?? []).filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])
    const plan = planSheetAxisVisibilityCommand(sheet, ranges, action, request, context.userId(), filtered)
    if (!plan.changed) return false
    Object.assign(sheet, plan.changes)
    const next = context.prepareView(toExcelWorkbookView(model), view)
    if (view.id === undefined) delete next.id
    context.applyView(next)
    const last = plan.selections.at(-1)
    if (sheetId === context.sheetId()) context.selection(plan.selections, last ? { row: last.startRow, column: last.startCol } : undefined)
    context.changed()
    return true
  }
  return { runAxisVisibilityCommand }
}
