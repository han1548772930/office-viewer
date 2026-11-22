import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange } from '../../domain/model.ts'
import { createBlankWorkbook } from '../../domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { sheetRangeCommandInputRanges } from '../../features/sheet-workbench/sheet-range-command.ts'
import { planWorkbookAxisStructureCommand, sheetAxisStructureActionForCommand, sheetAxisStructureCommandInput,
  sheetAxisStructureEdits, type SheetAxisStructureCommandInput } from '../../features/sheet-workbench/sheet-axis-structure-command.ts'

interface SheetAxisStructureUiContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  ranges(): CellRange[]
  userId(): string
  region(): string
  active(): boolean
  finishEdit(): boolean
  prepareView(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void
  selection(range: CellRange): void
  changed(): void
}

/** Both hosts use the same transaction planner; the studio owns document history. */
export function useSheetAxisStructure(context: SheetAxisStructureUiContext) {
  function runAxisStructureCommand(commandId: string, params?: unknown): boolean {
    if (!context.active() || !context.finishEdit()) return false
    const action = sheetAxisStructureActionForCommand(commandId)
    if (!action) return false
    const input = sheetAxisStructureCommandInput(params), view = context.workbook(), studio = context.studio()
    const sheetId = input.sheetId ?? input.subUnitId ?? context.sheetId()
    const ranges = sheetRangeCommandInputRanges(input) ?? context.ranges()
    const request: SheetAxisStructureCommandInput = { ...input, unitId: input.unitId ?? view.id, sheetId, ranges }
    if (studio) {
      request.unitId ??= studio.activeWorkbookId
      const changed = studio.executeCommandSync<SheetAxisStructureCommandInput, boolean>(commandId, request)
      if (!changed) return false
      if (request.unitId === view.id) {
        const model = studio.getWorkbookModel(request.unitId)
        if (model) context.applyView(toExcelWorkbookView(model))
        const selection = studio.getSelection()
        if (sheetId === context.sheetId() && selection?.unitId === request.unitId && selection.sheetId === sheetId) {
          context.selection({ startRow: selection.startRow, endRow: selection.endRow,
            startCol: selection.startColumn, endCol: selection.endColumn })
        }
        context.changed()
      }
      return true
    }
    if (input.unitId !== undefined && input.unitId !== view.id) throw new Error('找不到目标工作簿。')
    const model = applyExcelWorkbookView(createBlankWorkbook(), view, context.region())
    const sheet = model.sheets.find(candidate => candidate.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
    const filteredRows = [...(sheet.autoFilter?.filteredRows ?? []), ...(model.tables ?? [])
      .filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])]
    const edits = sheetAxisStructureEdits(sheet, ranges, action, request, filteredRows)
    if (!edits.length) return false
    const after = input.direction === 'after' || input.direction === 'down' || input.direction === 'right'
      || input.direction === 1 || input.direction === 2 || input.direction === undefined && action.position === 'after'
    const plan = planWorkbookAxisStructureCommand(model, sheetId, edits, context.userId(), after ? 'after' : 'before')
    const next = context.prepareView(toExcelWorkbookView(plan.workbook), view)
    // A detached standalone host has no studio unit ID; conversion must keep that identity.
    if (view.id === undefined) delete next.id
    context.applyView(next)
    if (sheetId === context.sheetId()) context.selection(plan.selection)
    context.changed()
    return true
  }
  return { runAxisStructureCommand }
}
