import type { RibbonItem } from '../../platform/ribbon/ribbon-schema'
import type { DocumentLatexFormulaTarget } from '../../features/document/latex/latex-formula'
import type { DocumentLatexCommand } from '../../features/document/commands/document-latex-commands'
import { DOCUMENT_LATEX_POPUP_OPERATIONS } from '../../features/document/controllers/document-latex-popup'

export const DOCUMENT_LATEX_MENU_MOUNTS = [
  { id: 'empty-layout', host: 'doc.menu.empty-paragraph', group: 'layout', order: 8 },
  { id: 'insert-quick-bottom', host: 'doc.menu.paragraph-t.insert', group: 'quickBottom', order: 4 },
  { id: 'insert-menu', host: 'doc.menu.paragraph-t.insert', group: 'insert', order: 3 },
  { id: 'edit-quick-bottom', host: 'doc.menu.paragraph-t.edit', group: 'quickBottom', order: 5 },
  { id: 'float-text-style', host: 'doc.menu.float-toolbar', group: 'doc.menu.float-text-style', order: 13 },
] as const
export type DocumentLatexMenuMount = typeof DOCUMENT_LATEX_MENU_MOUNTS[number]['id']

export function documentLatexMountedMenuItem(mount: DocumentLatexMenuMount, state: { available: boolean; reason?: string },
  translate: (key: string) => string): RibbonItem {
  const definition = DOCUMENT_LATEX_MENU_MOUNTS.find(item => item.id === mount)!
  return { id: `doc-context-latex-${mount}`, type: 'button', title: translate('document.latex.menu-formula'), order: definition.order,
    disabled: !state.available, disabledReason: state.reason, commandId: DOCUMENT_LATEX_POPUP_OPERATIONS.open }
}

/** Mobile actions and the context menu share the same checked content command. */
export function documentLatexCommandFromMenu(id: string, params: unknown): Extract<DocumentLatexCommand,
  { id: 'doc.command.remove-latex-formula' | 'doc.command.replace-latex-formula-with-text' }> | undefined {
  if (!params || typeof params !== 'object' || Array.isArray(params)) return undefined
  const rangeId: unknown = Reflect.get(params, 'rangeId')
  if (typeof rangeId !== 'string' || !rangeId.trim()) return undefined
  if (id === 'doc.command.remove-latex-formula') return { id, rangeId }
  const text: unknown = Reflect.get(params, 'text')
  return id === 'doc.command.replace-latex-formula-with-text' && typeof text === 'string' ? { id, rangeId, text } : undefined
}

export function documentLatexContextItems(item: DocumentLatexFormulaTarget, state: { available: boolean; reason?: string },
  translate: (key: string) => string): RibbonItem[] {
  const common = { type: 'button' as const, disabled: !state.available, disabledReason: state.reason }
  return [
    { ...common, id: 'doc-context-latex-edit', title: translate('document.latex.edit'),
      commandId: DOCUMENT_LATEX_POPUP_OPERATIONS.open, params: { rangeId: item.rangeId } },
    { ...common, id: 'doc-context-latex-remove', title: translate('document.latex.remove'),
      commandId: 'doc.command.remove-latex-formula', params: { rangeId: item.rangeId } },
    { ...common, id: 'doc-context-latex-text', title: translate('document.latex.replace-text'),
      commandId: 'doc.command.replace-latex-formula-with-text', params: { rangeId: item.rangeId, text: item.formula.latex } },
  ]
}
