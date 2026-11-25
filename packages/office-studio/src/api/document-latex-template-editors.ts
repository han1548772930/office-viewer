import type { OfficeStudio } from './office-studio'
import { getDocumentLatexTemplate } from '../features/document/latex/latex-templates'

interface TemplateEditor { insert: (templateId: string) => boolean }
export interface DocumentLatexTemplateEditorHandle { dispose(): void; isCurrent(): boolean }
const editors = new WeakMap<OfficeStudio, Map<string, TemplateEditor>>()

/** Runtime draft ownership is per studio/unit and never enters a document snapshot. */
export function registerDocumentLatexTemplateEditor(api: OfficeStudio, unitId: string, insert: TemplateEditor['insert']): DocumentLatexTemplateEditorHandle {
  if (!api.getDocumentModel(unitId) || typeof insert !== 'function') throw new TypeError('公式模板编辑器目标无效。')
  let units = editors.get(api)
  if (!units) { units = new Map(); editors.set(api, units) }
  const owner: TemplateEditor = { insert }
  units.set(unitId, owner)
  return { isCurrent: () => editors.get(api)?.get(unitId) === owner, dispose: () => {
    const current = editors.get(api)
    if (current && current.get(unitId) === owner) {
      current.delete(unitId)
      if (!current.size) editors.delete(api)
    }
  } }
}
export function requestDocumentLatexTemplate(api: OfficeStudio, unitId: string, templateId: unknown): boolean {
  const entry = getDocumentLatexTemplate(templateId)
  return entry && api.getDocumentModel(unitId) ? editors.get(api)?.get(unitId)?.insert(entry.id) ?? false : false
}
export function disposeDocumentLatexTemplateEditors(api: OfficeStudio, unitId?: string): void {
  if (unitId === undefined) { editors.delete(api); return }
  const current = editors.get(api)
  current?.delete(unitId)
  if (current && !current.size) editors.delete(api)
}
