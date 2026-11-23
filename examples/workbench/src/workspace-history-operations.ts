import type { WorkspaceKind } from './workspace-file-operations'

export interface WorkspaceHistoryTarget {
  kind: WorkspaceKind
  unitId: string
}

export interface WorkspaceHistoryServices {
  captureTarget: () => WorkspaceHistoryTarget | undefined
  replay: (direction: 'undo' | 'redo', target: WorkspaceHistoryTarget) => boolean
  onError: (direction: 'undo' | 'redo', error: unknown) => void
}

/** Ribbon and global shortcuts use the API history path for the current unit. */
export function createWorkspaceHistoryOperations(services: WorkspaceHistoryServices) {
  function replay(direction: 'undo' | 'redo', kind?: WorkspaceKind, unitId?: string): boolean {
    const target = services.captureTarget()
    if (!target || kind !== undefined && kind !== target.kind || unitId !== undefined && unitId !== target.unitId) return false
    try {
      return services.replay(direction, target)
    } catch (error: unknown) {
      services.onError(direction, error)
      return false
    }
  }

  return {
    undo: (kind?: WorkspaceKind, unitId?: string) => replay('undo', kind, unitId),
    redo: (kind?: WorkspaceKind, unitId?: string) => replay('redo', kind, unitId),
  }
}
