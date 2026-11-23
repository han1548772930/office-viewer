import type { OfficeStudio } from '../../api/office-studio.ts'
import type { ExcelWorkbookView } from '../../services/excelViewAdapter.ts'

export interface SheetAxisHostContext {
  studio(): OfficeStudio | undefined
  workbook(): ExcelWorkbookView
  sheetId(): string
  active(): boolean
}

export interface SheetAxisHostChangeOptions {
  syncSelection?: boolean
  current?: () => boolean
}

/** A completed axis command may synchronously notify observers that navigate or edit again. */
export function captureSheetAxisCommandHost(context: SheetAxisHostContext) {
  const studio = context.studio(), view = context.workbook(), sheetId = context.sheetId(), unitId = view.id
  const sameTarget = () => context.active() && context.studio() === studio && context.workbook().id === unitId
    && context.sheetId() === sheetId && (!studio || !unitId || !studio.isDisposed && studio.activeWorkbookId === unitId
      && studio.editor.state.activeKind === 'workbook' && studio.getWorkbookModel(unitId)?.activeSheetId === sheetId)
  const sourceCurrent = () => sameTarget() && context.workbook() === view
  function publishedCurrent(expectedRevision?: number): () => boolean {
    const publishedView = context.workbook(), revision = expectedRevision ?? (studio && unitId ? studio.getWorkbookRevision(unitId) : undefined)
    const signature = studio && unitId ? undefined : JSON.stringify(publishedView)
    return () => sameTarget() && context.workbook() === publishedView
      && (!studio || !unitId || studio.getWorkbookRevision(unitId) === revision)
      && (signature === undefined || JSON.stringify(context.workbook()) === signature)
  }
  return { studio, view, sheetId, sameTarget, sourceCurrent, publishedCurrent }
}
