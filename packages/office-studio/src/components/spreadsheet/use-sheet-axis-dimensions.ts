import { nextTick, ref, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../api/office-studio.ts'
import type { CellRange } from '../../domain/model.ts'
import { sheetColumnWidth, sheetRowHeight } from '../../domain/sheet-axis.ts'
import { toExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import { applySheetViewAxisDimensionCommand, SHEET_AXIS_DIMENSION_COMMANDS, sheetAxisDimensionActionForCommand,
  sheetAxisDimensionCommandInput, sheetAxisDimensionTargetRanges, type SheetAxisDimensionCommandInput } from '../../features/sheet-workbench/sheet-axis-dimensions.ts'
import { canEditWorksheetProtection, canViewWorksheetProtection } from '../../features/sheet-protection/range-protection.ts'
import { measureSpreadsheetText } from '../../rendering/spreadsheet/text-metrics.ts'

interface SheetAxisUiContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  ranges(): CellRange[]
  userId(): string
  region(): string
  active(): boolean
  finishEdit(): boolean
  applyView(view: ExcelWorkbookView): void
  changed(): void
  focus(): void
}
export interface SheetAxisSizeRequest {
  axis: 'row' | 'column'
  value: number
  sheetId: string
  count: number
}
interface PendingAxisSize {
  request: SheetAxisSizeRequest
  input: SheetAxisDimensionCommandInput
  unitId?: string
  signature: string
  studio?: OfficeStudio
}

export function useSheetAxisDimensions(context: SheetAxisUiContext) {
  const pending = shallowRef<PendingAxisSize | null>(null), axisSizeError = ref('')
  function cancelAxisSize(): void { pending.value = null; axisSizeError.value = '' }
  watch([context.studio, () => context.workbook().id ?? context.workbook().name, context.sheetId, context.active], cancelAxisSize, { flush: 'sync' })
  function afterAxisSizeClose(): void { void nextTick(() => { if (context.active() && !pending.value) context.focus() }) }
  function runAxisDimensionCommand(commandId: string, params?: unknown): boolean {
    if (!context.active() || !context.finishEdit()) return false
    const input = sheetAxisDimensionCommandInput(params), action = sheetAxisDimensionActionForCommand(commandId)
    if (!action) return false
    const workbook = context.workbook(), sheetId = input.sheetId ?? input.subUnitId ?? context.sheetId()
    const ranges = input.ranges ?? input.selections ?? (input.range ? [input.range] : context.ranges())
    const request: SheetAxisDimensionCommandInput = { ...input, ranges, sheetId, unitId: input.unitId ?? workbook.id }
    const studio = context.studio(), unitId = request.unitId
    let changed = false
    if (studio && unitId) {
      changed = studio.executeCommandSync<SheetAxisDimensionCommandInput, boolean>(commandId, request)
      if (changed) {
        const model = studio.getWorkbookModel(unitId)
        if (model) context.applyView(toExcelWorkbookView(model))
      }
    } else {
      const candidate: ExcelWorkbookView = JSON.parse(JSON.stringify(workbook))
      const sheet = candidate.sheets.find(item => item.id === sheetId)
      if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
      const normalized = sheetAxisDimensionCommandInput(request)
      changed = applySheetViewAxisDimensionCommand(sheet, normalized.ranges?.map(range => 'startCol' in range ? range
        : { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }) ?? [],
        action, normalized, context.userId(), { measure: measureSpreadsheetText, region: context.region() })
      if (changed) context.applyView(candidate)
    }
    if (changed) context.changed()
    return changed
  }
  function openAxisSizeDialog(axis: 'row' | 'column', params?: unknown): void {
    if (!context.active() || !context.finishEdit()) return
    const input = sheetAxisDimensionCommandInput(params), workbook = context.workbook(), sheetId = input.sheetId ?? input.subUnitId ?? context.sheetId()
    const sheet = workbook.sheets.find(item => item.id === sheetId)
    if (!sheet || !canViewWorksheetProtection(sheet, context.userId())) throw new Error('找不到可访问的目标工作表。')
    if (!canEditWorksheetProtection(sheet, context.userId())) throw new Error('此工作表受保护，当前用户不能调整行列尺寸。')
    const normalized = sheetAxisDimensionCommandInput({ ...input, ranges: input.ranges ?? input.selections ?? (input.range ? [input.range] : context.ranges()) })
    const anchorIndex = input.index ?? (axis === 'row' ? input.anchorRow : input.anchorCol)
    const ranges = sheetAxisDimensionTargetRanges(sheet, axis, normalized.ranges?.map(range => 'startCol' in range ? range
      : { startRow: range.startRow, endRow: range.endRow, startCol: range.startColumn, endCol: range.endColumn }) ?? [], anchorIndex)
    if (!ranges.length) return
    const anchor = anchorIndex ?? (axis === 'row' ? ranges[0]!.startRow : ranges[0]!.startCol)
    const studio = context.studio(), unitId = workbook.id
    pending.value = { request: { axis, sheetId, value: axis === 'row' ? sheetRowHeight(sheet, anchor) : sheetColumnWidth(sheet, anchor),
      count: ranges.reduce((sum, range) => sum + (axis === 'row' ? range.endRow - range.startRow + 1 : range.endCol - range.startCol + 1), 0) },
      studio, unitId, signature: studio ? '' : JSON.stringify(workbook), input: { unitId, sheetId, ranges,
        expectedRevision: studio && unitId ? studio.getWorkbookRevision(unitId) : undefined } }
    axisSizeError.value = ''
  }
  function confirmAxisSize(value: number): void {
    const state = pending.value
    if (!state || !context.active()) return
    try {
      if (!context.finishEdit()) return
      if (state.studio !== context.studio() || state.unitId !== context.workbook().id || state.request.sheetId !== context.sheetId()) throw new Error('目标工作簿或工作表已切换，请重新打开尺寸设置。')
      if (!state.studio && state.signature !== JSON.stringify(context.workbook())) throw new Error('工作簿已更改，请重新打开尺寸设置。')
      runAxisDimensionCommand(state.request.axis === 'row' ? SHEET_AXIS_DIMENSION_COMMANDS.rowHeight : SHEET_AXIS_DIMENSION_COMMANDS.columnWidth, { ...state.input, value })
      cancelAxisSize()
    } catch (error: unknown) { axisSizeError.value = error instanceof Error ? error.message : '无法设置行列尺寸。' }
  }
  return { pending, axisSizeError, openAxisSizeDialog, runAxisDimensionCommand, confirmAxisSize, cancelAxisSize, afterAxisSizeClose }
}
