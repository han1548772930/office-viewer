import type { Disposable } from '../../api/events'

export const OPEN_FEATURE_SEARCH = 'ui.operation.open-feature-search'
export const TOGGLE_SHORTCUT_PANEL = 'base-ui.operation.toggle-shortcut-panel'

export type WorkbenchUiOperation = typeof OPEN_FEATURE_SEARCH | typeof TOGGLE_SHORTCUT_PANEL
export type WorkbenchUiOperationHandler = (id: WorkbenchUiOperation) => boolean

export interface WorkbenchUiHost {
  bind: (run: WorkbenchUiOperationHandler) => Disposable
  execute: WorkbenchUiOperationHandler
}

export function isWorkbenchUiOperation(id: string): id is WorkbenchUiOperation {
  return id === OPEN_FEATURE_SEARCH || id === TOGGLE_SHORTCUT_PANEL
}

export function workbenchUiShortcut(event: KeyboardEvent): WorkbenchUiOperation | undefined {
  if (event.defaultPrevented || event.repeat || event.isComposing || event.keyCode === 229
    || event.altKey || !(event.ctrlKey || event.metaKey)) return
  if (event.shiftKey && (event.code === 'KeyP' || event.key.toLowerCase() === 'p')) return OPEN_FEATURE_SEARCH
  if (!event.shiftKey && (event.code === 'Backslash' || event.key === '\\')) return TOGGLE_SHORTCUT_PANEL
}
