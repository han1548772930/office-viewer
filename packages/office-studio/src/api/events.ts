import type { SheetBorderFormatState } from '../features/sheet-workbench/sheet-border-format-command.ts'
import type { SheetMergeConfirmationRequest } from '../features/sheet-workbench/sheet-merge-command.ts'
import type { SheetFormatPainterConfirmationRequest, SheetFormatPainterState } from '../features/sheet-workbench/sheet-format-painter-session.ts'
import type { SheetAxisEdit } from '../domain/sheet-axis-structure.ts'
import type { SheetCellShiftEdit } from '../domain/sheet-cell-shift.ts'
import type { SheetCellShiftConfirmationRequest } from '../features/sheet-workbench/sheet-cell-shift-command.ts'
import type { SheetSelectionCell } from '../features/sheet-workbench/sheet-selection.ts'
import type { DocumentInputSession } from '../features/document/model/document-input-session.ts'
import type { FWorkbook, FWorksheet } from './facade.ts'
import type { ISparklineGroup } from '../features/sparkline/sparkline-api-types.ts'
import type { SheetSelectionRange } from '../features/sheet-workbench/sheet-selection.ts'
import type { SheetTableSelectorRequest } from '../features/sheet-tables/sheet-table-create.ts'

export interface Disposable {
  dispose(): void
}

export interface ApiEventMap {
  LifeCycleChanged: { stage: 'ready' | 'disposed' }
  CommandExecuted: { id: string; params?: unknown; unitId?: string; kind?: 'workbook' | 'document' }
  EventListenerError: { eventName: keyof ApiEventMap; error: unknown }
  BeforeCommandExecute: { id: string; params?: unknown; unitId?: string; kind?: 'workbook' | 'document'; cancel?: boolean }
  BeforeUndo: { unitId: string; kind: 'workbook' | 'document'; cancel?: boolean }
  Undo: { unitId: string; kind: 'workbook' | 'document'; label?: string }
  BeforeRedo: { unitId: string; kind: 'workbook' | 'document'; cancel?: boolean }
  Redo: { unitId: string; kind: 'workbook' | 'document'; label?: string }
  WorkbookCreated: { workbookId: string }
  WorkbookDisposed: { workbookId: string }
  DocCreated: { unitId: string; documentId: string }
  DocDisposed: { unitId: string; documentId: string }
  DocumentInputStateChanged: { unitId: string; state?: DocumentInputSession }
  DocumentZoomRatioChanged: { documentId: string; previousZoomRatio: number; zoomRatio: number }
  BeforeSheetNameChange: { workbookId: string; sheetId: string; oldName: string; newName: string; cancel?: boolean }
  SheetNameChanged: { workbookId: string; sheetId: string; oldName: string; newName: string }
  SheetValueChanged: { workbookId: string; sheetId: string; ranges: ApiRange[] }
  SheetSparklineChanged: { workbook: FWorkbook; worksheet: FWorksheet; sparklines: ISparklineGroup[] }
  SheetOutlineActionDialogRequested: { unitId: string; sheetId: string; ranges: SheetSelectionRange[]; operation: 'group' | 'remove';
    expectedRevision: number; expectedHistoryGeneration: number; expectedUserId: string }
  SheetTableSelectorRequested: SheetTableSelectorRequest
  SheetDimensionsChanged: { workbookId: string; sheetId: string; axis: 'row' | 'column'; automatic: boolean; ranges: ApiRange[] }
  SheetStructureChanged: { workbookId: string; sheetId: string; edits: Array<SheetAxisEdit | SheetCellShiftEdit> }
  SheetCellShiftConfirmationRequested: SheetCellShiftConfirmationRequest
  SheetVisibilityChanged: { workbookId: string; sheetId: string; axis: 'row' | 'column'; visible: boolean; ranges: ApiRange[] }
  SheetBorderFormatStateChanged: { workbookId: string; sheetId: string; state: SheetBorderFormatState }
  SheetMergeConfirmationRequested: SheetMergeConfirmationRequest
  SheetFormatPainterStateChanged: { state: SheetFormatPainterState }
  SheetFormatPainterConfirmationRequested: SheetFormatPainterConfirmationRequest
  SelectionChanged: { workbookId: string; sheetId: string; range: ApiRange }
  SelectionsChanged: { workbookId: string; sheetId: string; selection: ApiSheetSelection }
  BeforePivotTableChange: { unitId: string; subUnitId: string; pivotTableId: string; commandId: string; cancel?: boolean }
  PivotTableChanged: { unitId: string; subUnitId: string; pivotTableId: string; revision: number }
  PivotTableRendered: { unitId: string; subUnitId: string; pivotTableId: string; revision: number; outputRange?: { startRow: number; startCol: number; endRow: number; endCol: number } }
  BeforePivotChartChange: { unitId: string; subUnitId: string; pivotChartId: string; commandId: string; cancel?: boolean }
  PivotChartAdded: { unitId: string; subUnitId: string; pivotChartId: string; sourceKind: 'range' | 'pivotTable' }
  PivotChartChanged: { unitId: string; subUnitId: string; pivotChartId: string; revision: number; reportRevision: number }
  PivotChartRemoved: { unitId: string; subUnitId: string; pivotChartId: string }
  BeforeSlicerChange: { unitId: string; hostSheetId: string; slicerId: string; commandId: string; cancel?: boolean }
  SlicerAdded: { unitId: string; hostSheetId: string; slicerId: string; type: 'slicer' | 'timeline' }
  SlicerChanged: { unitId: string; hostSheetId: string; slicerId: string; revision: number }
  SlicerRemoved: { unitId: string; hostSheetId: string; slicerId: string }
}

export class FEventName {
  readonly LifeCycleChanged = 'LifeCycleChanged'
  readonly BeforeCommandExecute = 'BeforeCommandExecute'
  readonly CommandExecuted = 'CommandExecuted'
  readonly EventListenerError = 'EventListenerError'
  readonly BeforeUndo = 'BeforeUndo'
  readonly Undo = 'Undo'
  readonly BeforeRedo = 'BeforeRedo'
  readonly Redo = 'Redo'
  readonly WorkbookCreated = 'WorkbookCreated'
  readonly WorkbookDisposed = 'WorkbookDisposed'
  readonly DocCreated = 'DocCreated'
  readonly DocDisposed = 'DocDisposed'
  readonly DocumentInputStateChanged = 'DocumentInputStateChanged'
  readonly DocumentZoomRatioChanged = 'DocumentZoomRatioChanged'
  readonly BeforeSheetNameChange = 'BeforeSheetNameChange'
  readonly SheetNameChanged = 'SheetNameChanged'
  readonly SheetValueChanged = 'SheetValueChanged'
  readonly SheetSparklineChanged = 'SheetSparklineChanged'
  readonly SheetOutlineActionDialogRequested = 'SheetOutlineActionDialogRequested'
  readonly SheetTableSelectorRequested = 'SheetTableSelectorRequested'
  readonly SheetDimensionsChanged = 'SheetDimensionsChanged'
  readonly SheetStructureChanged = 'SheetStructureChanged'
  readonly SheetCellShiftConfirmationRequested = 'SheetCellShiftConfirmationRequested'
  readonly SheetVisibilityChanged = 'SheetVisibilityChanged'
  readonly SheetBorderFormatStateChanged = 'SheetBorderFormatStateChanged'
  readonly SheetMergeConfirmationRequested = 'SheetMergeConfirmationRequested'
  readonly SheetFormatPainterStateChanged = 'SheetFormatPainterStateChanged'
  readonly SheetFormatPainterConfirmationRequested = 'SheetFormatPainterConfirmationRequested'
  readonly SelectionChanged = 'SelectionChanged'
  readonly SelectionsChanged = 'SelectionsChanged'
  readonly BeforePivotTableChange = 'BeforePivotTableChange'
  readonly PivotTableChanged = 'PivotTableChanged'
  readonly PivotTableRendered = 'PivotTableRendered'
  readonly BeforePivotChartChange = 'BeforePivotChartChange'
  readonly PivotChartAdded = 'PivotChartAdded'
  readonly PivotChartChanged = 'PivotChartChanged'
  readonly PivotChartRemoved = 'PivotChartRemoved'
  readonly BeforeSlicerChange = 'BeforeSlicerChange'
  readonly SlicerAdded = 'SlicerAdded'
  readonly SlicerChanged = 'SlicerChanged'
  readonly SlicerRemoved = 'SlicerRemoved'
}

export interface ApiRange {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
  sheetId?: string
  unitId?: string
  rangeType?: 0 | 1 | 2 | 3
}

export interface ApiSheetSelection {
  ranges: ApiRange[]
  activeRangeIndex: number
  primary?: SheetSelectionCell
}

type Listener<T> = (event: T) => void
export interface EventDeliveryError<K extends PropertyKey = PropertyKey> { eventName: K; error: unknown }

export class EventHub<TEvents extends object> {
  private readonly listeners = new Map<keyof TEvents, Set<Listener<never>>>()
  hasListeners(event: keyof TEvents): boolean { return Boolean(this.listeners.get(event)?.size) }

  addEvent<K extends keyof TEvents>(event: K, listener: Listener<TEvents[K]>): Disposable {
    const listeners = this.listeners.get(event) ?? new Set<Listener<never>>()
    listeners.add(listener as Listener<never>)
    this.listeners.set(event, listeners)
    let disposed = false
    return {
      dispose: () => {
        if (disposed) return
        disposed = true
        listeners.delete(listener as Listener<never>)
        if (!listeners.size) this.listeners.delete(event)
      },
    }
  }

  emit<K extends keyof TEvents>(event: K, payload: TEvents[K]): void {
    const listeners = this.listeners.get(event)
    if (!listeners) return
    for (const listener of [...listeners]) (listener as Listener<TEvents[K]>)(payload)
  }

  /** Observations receive their own payload and return failures; cancellable Before events still use emit(). */
  emitSettled<K extends keyof TEvents>(event: K, createPayload: () => TEvents[K], shouldContinue?: () => boolean): EventDeliveryError<K>[] {
    const errors: EventDeliveryError<K>[] = []
    const listeners = this.listeners.get(event)
    if (!listeners) return errors
    for (const listener of [...listeners]) {
      try { if (shouldContinue && !shouldContinue()) break }
      catch (error) { errors.push({ eventName: event, error }); break }
      try { (listener as Listener<TEvents[K]>)(createPayload()) }
      catch (error) { errors.push({ eventName: event, error }) }
    }
    return errors
  }

  clear(): void {
    this.listeners.clear()
  }
}

