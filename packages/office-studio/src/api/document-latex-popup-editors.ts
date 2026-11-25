import type { OfficeStudio } from './office-studio'
import type { DocumentLatexPopupTarget } from '../features/document/controllers/document-latex-popup'

export interface DocumentLatexPopupEditor {
  open(target: DocumentLatexPopupTarget): boolean
  close(force?: boolean): boolean
}
export interface DocumentLatexPopupEditorHandle { dispose(): void; isCurrent(): boolean }
const editors = new WeakMap<OfficeStudio, Map<string, DocumentLatexPopupEditor>>()

export function registerDocumentLatexPopupEditor(api: OfficeStudio, unitId: string, editor: DocumentLatexPopupEditor): DocumentLatexPopupEditorHandle {
  if (!api.getDocumentModel(unitId) || !editor || typeof editor.open !== 'function' || typeof editor.close !== 'function') throw new TypeError('公式编辑器目标无效。')
  let units = editors.get(api)
  if (!units) { units = new Map(); editors.set(api, units) }
  const previous = units.get(unitId), owner: DocumentLatexPopupEditor = { open: target => editor.open(target), close: force => editor.close(force) }
  units.set(unitId, owner)
  previous?.close(true)
  return { isCurrent: () => editors.get(api)?.get(unitId) === owner, dispose: () => {
    const current = editors.get(api)
    if (current?.get(unitId) !== owner) return
    current.delete(unitId)
    if (!current.size) editors.delete(api)
    owner.close(true)
  } }
}
export function requestDocumentLatexPopupOpen(api: OfficeStudio, target: DocumentLatexPopupTarget): boolean {
  return api.getDocumentModel(target.unitId) === target.source ? editors.get(api)?.get(target.unitId)?.open(target) ?? false : false
}
export function requestDocumentLatexPopupClose(api: OfficeStudio, unitId: string): boolean {
  return editors.get(api)?.get(unitId)?.close(false) ?? true
}
export function disposeDocumentLatexPopupEditors(api: OfficeStudio, unitId?: string): void {
  const current = editors.get(api)
  if (!current) return
  const owners = unitId === undefined ? [...current.values()] : [current.get(unitId)].filter((owner): owner is DocumentLatexPopupEditor => !!owner)
  if (unitId === undefined) editors.delete(api)
  else { current.delete(unitId); if (!current.size) editors.delete(api) }
  for (const owner of owners) owner.close(true)
}
