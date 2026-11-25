import { Sigma } from 'lucide-vue-next'
import type { RibbonItem, RibbonTab } from '../../platform/ribbon/ribbon-schema'
import type { OfficeTranslator } from '../../i18n/office-i18n'
import { DOCUMENT_LATEX_CATEGORY_LABELS, DOCUMENT_LATEX_TEMPLATE_CATEGORIES, DOCUMENT_LATEX_TEMPLATE_OPERATION } from '../../features/document/latex/latex-templates'
import DocumentLatexTemplateGallery from './DocumentLatexTemplateGallery.vue'

export const DOCUMENT_LATEX_RIBBON_TAB = 'docs-latex-ui.ribbon.formula'
export const DOCUMENT_LATEX_RIBBON_GROUP = 'docs-latex-ui.ribbon.group.formula'
export interface DocumentLatexRibbonContext { visible: boolean; disabled: boolean; reason?: string }
/** Fixed contextual roots: one selector per category, in source order and both grid rows. */
export function createDocumentLatexRibbon(context: DocumentLatexRibbonContext, translate?: OfficeTranslator): RibbonTab | undefined {
  if (!context.visible) return undefined
  return { id: DOCUMENT_LATEX_RIBBON_TAB, title: translate?.('document.latex.formula-tab', '公式') ?? '公式', order: 100, contextual: true,
    groups: [{ id: DOCUMENT_LATEX_RIBBON_GROUP, order: 0, items: DOCUMENT_LATEX_TEMPLATE_CATEGORIES.map((category, order): RibbonItem => ({
      id: 'docs-latex-ui.ribbon.menu.' + category, type: 'menu', icon: Sigma,
      title: translate?.('document.latex.category.' + category, DOCUMENT_LATEX_CATEGORY_LABELS[category][0]) ?? DOCUMENT_LATEX_CATEGORY_LABELS[category][0],
      order, disabled: context.disabled, disabledReason: context.reason, showLabel: true,
      popupComponent: DocumentLatexTemplateGallery, popupCommandId: DOCUMENT_LATEX_TEMPLATE_OPERATION, popupProps: { category },
      gridLayout: { row: 1, column: order + 1, rowSpan: 2, showLabel: true },
    })) }] }
}
