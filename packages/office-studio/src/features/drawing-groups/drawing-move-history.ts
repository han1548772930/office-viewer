import type { HistorySavepoint } from '../../platform/history/history-types.ts'

export const SHEET_DRAWING_MOVE_HISTORY_WINDOW = 1_000
export interface SheetDrawingMoveHistoryContext {
  unitId: string
  sheetId: string
  userId: string
  selectionKey: string
  selectionGeneration: number
  contentRevision: number
  savepoint: HistorySavepoint
  savedRevision: number
  now: number
}

/** Only a receipt is retained here. Snapshots, redo and undo stay in the existing content history. */
export class SheetDrawingMoveHistory {
  private recent?: SheetDrawingMoveHistoryContext

  amendment(current: SheetDrawingMoveHistoryContext, canAmend: (receipt: HistorySavepoint) => boolean): HistorySavepoint | undefined {
    const previous = this.recent
    if (!previous || current.now < previous.now || current.now > previous.now + SHEET_DRAWING_MOVE_HISTORY_WINDOW
      || previous.unitId !== current.unitId || previous.sheetId !== current.sheetId || previous.userId !== current.userId
      || previous.selectionKey !== current.selectionKey || previous.selectionGeneration !== current.selectionGeneration
      || previous.contentRevision !== current.contentRevision || previous.savedRevision !== current.savedRevision
      || previous.savepoint.generation !== current.savepoint.generation || previous.savepoint.revision !== current.savepoint.revision
      || current.savedRevision === current.savepoint.revision || !canAmend(previous.savepoint)) return undefined
    return { ...previous.savepoint }
  }

  remember(current: SheetDrawingMoveHistoryContext): void {
    this.recent = { ...current, savepoint: { ...current.savepoint } }
  }

  clear(): void { this.recent = undefined }
}
