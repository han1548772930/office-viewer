import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange } from '../../domain/model.ts'
import { applyExcelWorkbookView, toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { sheetRangeCommandInputRanges } from '../../features/sheet-workbench/sheet-range-command.ts'
import type { SheetSelectionCell } from '../../features/sheet-workbench/sheet-selection.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { captureSheetAxisCommandHost, type SheetAxisHostChangeOptions } from './sheet-axis-command-host.ts'
import { assertSheetAxisHasVisiblePosition, planSheetAxisVisibilityCommand, sheetAxisVisibilityActionForCommand, sheetAxisVisibilityCommandInput,
  type SheetAxisVisibilityCommandInput } from '../../features/sheet-workbench/sheet-axis-visibility.ts'

interface SheetAxisVisibilityUiContext {
  studio(): OfficeStudio | undefined; workbook(): ExcelWorkbookView; sheetId(): string; ranges(): CellRange[]
  userId(): string; region(): string; active(): boolean; finishEdit(): boolean
  prepareView(view: ExcelWorkbookView, previous: ExcelWorkbookView): ExcelWorkbookView
  applyView(view: ExcelWorkbookView): void; selection(ranges: CellRange[], primary?: SheetSelectionCell): void
  changed(options?: SheetAxisHostChangeOptions): void
}
export function useSheetAxisVisibility(context: SheetAxisVisibilityUiContext) {
  function runAxisVisibilityCommand(id: string, params?: unknown): boolean {
    const action = sheetAxisVisibilityActionForCommand(id)
    if (!action || !context.active()) return false
    const input = sheetAxisVisibilityCommandInput(params)
    if (!context.finishEdit() || !context.active()) return false
    const host = captureSheetAxisCommandHost(context), { view, studio } = host
    if (!host.sameTarget()) return false
    const sheetId = input.sheetId ?? input.subUnitId ?? context.sheetId(), ranges = sheetRangeCommandInputRanges(input) ?? context.ranges()
    const request: SheetAxisVisibilityCommandInput = { ...input, unitId: input.unitId ?? view.id, sheetId, ranges }
    if (studio && view.id) {
      const changed = studio.executeCommandSync<SheetAxisVisibilityCommandInput, boolean>(id, request)
      if (!changed) return false
      if (request.unitId === view.id && host.sameTarget()) {
        const model = studio.getWorkbookModel(request.unitId)
        if (!model) return true
        const revision = studio.getWorkbookRevision(request.unitId)
        const next = toExcelWorkbookView(model)
        if (!host.sameTarget() || studio.getWorkbookRevision(request.unitId) !== revision) return true
        // The studio already published its complete selection. The host projects it while applying this view.
        context.applyView(next)
        if (host.sameTarget() && studio.getWorkbookRevision(request.unitId) === revision) {
          context.changed({ syncSelection: false, current: host.publishedCurrent(revision) })
        }
      }
      return true
    }
    if (input.unitId !== undefined && input.unitId !== view.id) throw new Error('找不到目标工作簿。')
    const sourceSignature = JSON.stringify(view)
    const model = applyExcelWorkbookView(createBlankWorkbook(), view, context.region()), sheet = model.sheets.find(item => item.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
    const filtered = (model.tables ?? []).filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])
    const plan = planSheetAxisVisibilityCommand(sheet, ranges, action, request, context.userId(), filtered)
    if (!plan.changed) return false
    Object.assign(sheet, plan.changes)
    const next = context.prepareView(toExcelWorkbookView(model), cloneModel(view))
    if (!host.sourceCurrent() || JSON.stringify(context.workbook()) !== sourceSignature) return false
    if (!action.visible) {
      const target = next.sheets.find(item => item.id === sheetId)
      if (!target) throw new Error('隐藏的目标工作表已移除。')
      const filtered = (next.tables ?? []).filter(table => table.sheetId === sheetId).flatMap(table => table.autoFilter?.filteredRows ?? [])
      assertSheetAxisHasVisiblePosition(target, action.axis, filtered)
    }
    if (view.id === undefined) delete next.id
    else next.id = view.id
    const preparedSignature = JSON.stringify(next)
    context.applyView(next)
    if (!host.sameTarget() || JSON.stringify(context.workbook()) !== preparedSignature) return true
    const current = host.publishedCurrent()
    const last = plan.selections.at(-1)
    if (sheetId === context.sheetId()) context.selection(plan.selections, last ? { row: last.startRow, column: last.startCol } : undefined)
    if (current()) context.changed({ current })
    return true
  }
  return { runAxisVisibilityCommand }
}
