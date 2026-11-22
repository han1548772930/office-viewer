import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellWrapStrategy } from './sheet-cell-layout.ts'

export interface SheetCellInputSelection { start: number; end: number }
export interface SheetCellInputReplacement { value: string; selection: SheetCellInputSelection }
export interface SheetCellInputKey {
  key: string
  altKey?: boolean
  ctrlKey?: boolean
  metaKey?: boolean
  shiftKey?: boolean
  isComposing?: boolean
  keyCode?: number
}
export type SheetCellInputAction = 'native' | 'line-break' | 'commit' | 'next-row' | 'previous-row' | 'next-column' | 'previous-column' | 'cancel'

/** Selection offsets refer to the library input's UTF-16 value, as in setSelectionRange. */
export function replaceSheetCellInputSelection(value: string, insertion: string, selection?: SheetCellInputSelection): SheetCellInputReplacement {
  const offset = (candidate: number | undefined) => candidate !== undefined && Number.isFinite(candidate)
    ? Math.max(0, Math.min(value.length, Math.trunc(candidate))) : value.length
  const first = offset(selection?.start), last = offset(selection?.end)
  const start = Math.min(first, last), end = Math.max(first, last)
  const inserted = insertion.replace(/\r\n?/gu, '\n')
  const caret = start + inserted.length
  return { value: value.slice(0, start) + inserted + value.slice(end), selection: { start: caret, end: caret } }
}

/** Sheet editing shortcuts leave native caret, clipboard and composition behavior in the input. */
export function sheetCellInputActionForKey(event: SheetCellInputKey, composing = false): SheetCellInputAction {
  if (composing || event.isComposing || event.keyCode === 229) return 'native'
  if (event.key === 'Escape') return 'cancel'
  if (event.key === 'Enter') {
    if (event.altKey && !event.ctrlKey && !event.metaKey) return 'line-break'
    if (event.altKey) return 'native'
    if (event.ctrlKey || event.metaKey) return 'commit'
    return event.shiftKey ? 'previous-row' : 'next-row'
  }
  if (event.key === 'Tab' && !event.altKey && !event.ctrlKey && !event.metaKey) return event.shiftKey ? 'previous-column' : 'next-column'
  return 'native'
}

/** Once a draft becomes multiline its bridge remains a TextArea until that edit session ends. */
export function sheetCellNeedsMultilineEditor(cell: ExcelCellView | undefined, draft: string, retained = false): boolean {
  return retained || /[\r\n]/u.test(draft) || sheetCellWrapStrategy(cell) === 'wrap'
}
