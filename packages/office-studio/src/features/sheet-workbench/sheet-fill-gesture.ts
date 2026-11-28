import type { GridRange } from '../../rendering/spreadsheet/layout.ts'

/** A positive acknowledgement from the workbench's existing edit commit path. */
export interface SheetFillCommitContext {
  sheetId: string
  workbookRevision?: number
}

export interface SheetFillGestureContext extends SheetFillCommitContext {
  selectedKey: string
  selectedRanges: readonly GridRange[]
  editingKey: string | null
  editSessionId?: string
}

export function snapshotSheetFillGesture(context: SheetFillGestureContext): SheetFillGestureContext {
  return { ...context, selectedRanges: context.selectedRanges.map(range => ({ ...range })) }
}

export function sheetFillGestureSelectionMatches(origin: SheetFillGestureContext, current: SheetFillGestureContext): boolean {
  return origin.sheetId === current.sheetId && origin.selectedKey === current.selectedKey
    && origin.selectedRanges.length === current.selectedRanges.length
    && origin.selectedRanges.every((range, index) => {
      const next = current.selectedRanges[index]!
      return range.startRow === next.startRow && range.endRow === next.endRow
        && range.startColumn === next.startColumn && range.endColumn === next.endColumn
        // Ordinary cell selections may acquire the explicit NORMAL=0 tag
        // when the edit commit round-trips through the shared selection API.
        && (range.rangeType ?? 0) === (next.rangeType ?? 0)
    })
}

/** The commit may advance the revision; it must finish this edit without navigating away. */
export function canResumeSheetFillGesture(
  origin: SheetFillGestureContext,
  committed: SheetFillCommitContext,
  current: SheetFillGestureContext,
): boolean {
  return !current.editingKey && !current.editSessionId
    && committed.sheetId === origin.sheetId && committed.workbookRevision === current.workbookRevision
    && sheetFillGestureSelectionMatches(origin, current)
}

/** Once dragging starts, content, selection, or another edit invalidates the gesture. */
export function sheetFillGestureStillCurrent(origin: SheetFillGestureContext, current: SheetFillGestureContext): boolean {
  return !current.editingKey && !current.editSessionId
    && origin.workbookRevision === current.workbookRevision && sheetFillGestureSelectionMatches(origin, current)
}
