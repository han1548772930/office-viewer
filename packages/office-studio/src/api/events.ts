export interface Disposable {
  dispose(): void
}

export interface ApiEventMap {
  LifeCycleChanged: { stage: 'ready' | 'disposed' }
  CommandExecuted: { id: string; params?: unknown; unitId?: string; kind?: 'workbook' | 'document' }
  BeforeCommandExecute: { id: string; params?: unknown; unitId?: string; kind?: 'workbook' | 'document'; cancel?: boolean }
  BeforeUndo: { unitId: string; kind: 'workbook' | 'document'; cancel?: boolean }
  Undo: { unitId: string; kind: 'workbook' | 'document'; label?: string }
  BeforeRedo: { unitId: string; kind: 'workbook' | 'document'; cancel?: boolean }
  Redo: { unitId: string; kind: 'workbook' | 'document'; label?: string }
  WorkbookCreated: { workbookId: string }
  WorkbookDisposed: { workbookId: string }
  DocCreated: { unitId: string; documentId: string }
  DocDisposed: { unitId: string; documentId: string }
  BeforeSheetNameChange: { workbookId: string; sheetId: string; oldName: string; newName: string; cancel?: boolean }
  SheetNameChanged: { workbookId: string; sheetId: string; oldName: string; newName: string }
  SheetValueChanged: { workbookId: string; sheetId: string; ranges: ApiRange[] }
  SelectionChanged: { workbookId: string; sheetId: string; range: ApiRange }
  BeforePivotTableChange: { unitId: string; subUnitId: string; pivotTableId: string; commandId: string; cancel?: boolean }
  PivotTableChanged: { unitId: string; subUnitId: string; pivotTableId: string; revision: number }
  PivotTableRendered: { unitId: string; subUnitId: string; pivotTableId: string; revision: number; outputRange?: { startRow: number; startCol: number; endRow: number; endCol: number } }
}

export class FEventName {
  readonly LifeCycleChanged = 'LifeCycleChanged'
  readonly BeforeCommandExecute = 'BeforeCommandExecute'
  readonly CommandExecuted = 'CommandExecuted'
  readonly BeforeUndo = 'BeforeUndo'
  readonly Undo = 'Undo'
  readonly BeforeRedo = 'BeforeRedo'
  readonly Redo = 'Redo'
  readonly WorkbookCreated = 'WorkbookCreated'
  readonly WorkbookDisposed = 'WorkbookDisposed'
  readonly DocCreated = 'DocCreated'
  readonly DocDisposed = 'DocDisposed'
  readonly BeforeSheetNameChange = 'BeforeSheetNameChange'
  readonly SheetNameChanged = 'SheetNameChanged'
  readonly SheetValueChanged = 'SheetValueChanged'
  readonly SelectionChanged = 'SelectionChanged'
  readonly BeforePivotTableChange = 'BeforePivotTableChange'
  readonly PivotTableChanged = 'PivotTableChanged'
  readonly PivotTableRendered = 'PivotTableRendered'
}

export interface ApiRange {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
  sheetId?: string
  unitId?: string
}

type Listener<T> = (event: T) => void

export class EventHub<TEvents extends object> {
  private readonly listeners = new Map<keyof TEvents, Set<Listener<never>>>()

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

  clear(): void {
    this.listeners.clear()
  }
}

