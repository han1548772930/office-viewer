import type { OfficeStudio } from '../../api/office-studio.ts'
import { createBlankWorkbook, type CellRange } from '../../domain/model.ts'
import { applyExcelWorkbookView, type ExcelWorkbookView } from '../../services/excelViewAdapter.ts'
import type { SheetClipboardCommandInput } from '../../features/sheet-workbench/sheet-clipboard-types.ts'

export interface SheetClipboardHostContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  range(): CellRange | undefined
  region(): string
  active(): boolean
  editing(): boolean
}

/** Capture before browser I/O. IDs alone cannot distinguish a reopened model or a replaced studio. */
export function captureSheetClipboardHost(context: SheetClipboardHostContext, standaloneUnitId: string) {
  const owner = context.studio(), view = context.workbook(), sheetId = context.sheetId(), range = context.range()
  if (!context.active() || context.editing() || !range) return undefined
  const unitId = view.id ?? standaloneUnitId, studio = owner && view.id ? owner : undefined
  const model = studio ? studio.getWorkbookModel(unitId)
    : applyExcelWorkbookView({ ...createBlankWorkbook(), id: unitId }, view, context.region())
  const sheet = model?.sheets.find(candidate => candidate.id === sheetId)
  const viewSheet = view.sheets.find(candidate => candidate.id === sheetId)
  if (!model || !sheet || !viewSheet) return undefined
  const revision = studio?.getWorkbookRevision(unitId)
  const historyGeneration = studio?.editor.history.captureSavepoint(unitId).generation
  const signature = studio ? undefined : JSON.stringify(view)
  const focused = typeof document === 'undefined' ? undefined : document.activeElement
  const sourceRange = { ...range }, selection = JSON.stringify(range)
  const sameTarget = () => context.active() && !context.editing() && context.studio() === owner
    && (context.workbook().id ?? standaloneUnitId) === unitId && context.sheetId() === sheetId
    && (!studio || !studio.isDisposed && studio.activeWorkbookId === unitId
      && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookModel(unitId)?.activeSheetId === sheetId)
  const current = () => sameTarget() && context.workbook() === view
    && view.sheets.find(candidate => candidate.id === sheetId) === viewSheet && (studio
    ? studio.getWorkbookModel(unitId) === model && model.sheets.find(candidate => candidate.id === sheetId) === sheet
      && studio.getWorkbookRevision(unitId) === revision
      && studio.editor.history.captureSavepoint(unitId).generation === historyGeneration
    : context.workbook() === view && JSON.stringify(context.workbook()) === signature)
  if (!current()) return undefined
  const request: SheetClipboardCommandInput = {
    unitId, sheetId, range: sourceRange,
    ...(studio ? { expectedRevision: revision, expectedHistoryGeneration: historyGeneration, host: { unitId, sheetId } } : {}),
  }
  const focusUnchanged = () => typeof document === 'undefined' || document.activeElement === focused
    || document.activeElement === document.body && focused instanceof Element && !focused.isConnected
  return { owner, studio, view, viewSheet, model, sheet, unitId, sheetId, range: sourceRange, revision, historyGeneration,
    sameTarget, current, selectionUnchanged: () => JSON.stringify(context.range()) === selection, focusUnchanged, request }
}

export type SheetClipboardHost = NonNullable<ReturnType<typeof captureSheetClipboardHost>>
