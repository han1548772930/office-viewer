import type { DocumentBlock, DocumentModel, DocumentPageSetup, DocumentParagraphStyle, DocumentSectionSetup } from '../../../domain/model.ts'
import { DEFAULT_DOCUMENT_SECTION_ID, documentSectionIds } from './section-header-footer.ts'

export type DocumentParagraphStylePatch = Partial<DocumentParagraphStyle>
export type DocumentSectionSetupPatch = Partial<DocumentSectionSetup>

export const DOCUMENT_PARAGRAPH_STYLE_FIELDS = [
  'align', 'indent', 'lineHeight', 'spaceBefore', 'spaceAfter', 'indentStart', 'indentEnd',
  'indentFirstLine', 'hanging', 'spacingRule', 'wordWrap', 'snapToGrid',
] as const
const paragraphStyleFields = new Set<string>(DOCUMENT_PARAGRAPH_STYLE_FIELDS)

export function validateDocumentParagraphStylePatch(patch: DocumentParagraphStylePatch, currentSpacingRule: NonNullable<DocumentParagraphStyle['spacingRule']> = 'auto'): void {
  if (typeof patch !== 'object' || patch === null || Array.isArray(patch)) throw new TypeError('Paragraph style patch must be an object')
  for (const [key, value] of Object.entries(patch)) {
    if (!paragraphStyleFields.has(key)) throw new TypeError(`Unsupported paragraph setting: ${key}`)
    if (key === 'align') {
      if (!['left', 'center', 'right', 'justify'].includes(value)) throw new TypeError('Unsupported paragraph alignment')
      continue
    }
    if (key === 'spacingRule') {
      if (!['auto', 'atLeast', 'exact'].includes(value)) throw new TypeError('Unsupported line spacing rule')
      continue
    }
    if (key === 'wordWrap' || key === 'snapToGrid') {
      if (typeof value !== 'boolean') throw new TypeError(`Paragraph ${key === 'wordWrap' ? 'word wrap' : 'snap to grid'} must be boolean`)
      continue
    }
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) throw new RangeError(`Paragraph ${key} must be a non-negative finite number`)
    if (key === 'lineHeight') {
      if (value === 0) throw new RangeError('Paragraph line height must be positive')
      const maximum = (patch.spacingRule ?? currentSpacingRule) === 'auto' ? 5 : 100
      if (value < 1 || value > maximum) throw new RangeError(`Paragraph line spacing must be between 1 and ${maximum}`)
    } else if (value > 100) throw new RangeError(`Paragraph ${key} must be between 0 and 100`)
  }
}

export function applyDocumentParagraphStylePatch(block: DocumentBlock, patch: DocumentParagraphStylePatch): void {
  validateDocumentParagraphStylePatch(patch, block.style.spacingRule ?? 'auto')
  if (!['paragraph', 'heading', 'bullet', 'numbered', 'checklist', 'quote'].includes(block.type)) {
    throw new TypeError('Paragraph settings can only be applied to text paragraphs')
  }
  if (patch.spacingRule !== undefined && patch.lineHeight === undefined) {
    validateDocumentParagraphStylePatch({ lineHeight: block.style.lineHeight }, patch.spacingRule)
  }
  block.style = { ...block.style, ...patch }
}

export function validateDocumentSectionSetupPatch(patch: DocumentSectionSetupPatch): void {
  for (const value of [patch.marginHeader, patch.marginFooter]) {
    if (value !== undefined && (!Number.isFinite(value) || value < 0 || value > 200)) throw new RangeError('Header/footer distance must be between 0 and 200')
  }
  if (patch.orientation !== undefined && patch.orientation !== 'portrait' && patch.orientation !== 'landscape') throw new TypeError('Unsupported section orientation')
  if (patch.columnSeparator !== undefined && patch.columnSeparator !== 'none' && patch.columnSeparator !== 'line') throw new TypeError('Unsupported section column separator')
  if (patch.sectionStart !== undefined && !['unspecified', 'nextPage', 'continuous', 'nextColumn', 'evenPage', 'oddPage'].includes(patch.sectionStart)) throw new TypeError('Unsupported section start')
  for (const key of ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft', 'columnGap', 'pageNumberStart'] as const) {
    const value = patch[key]
    if (value !== undefined && (!Number.isFinite(value) || value < 0)) throw new RangeError(`Section ${key} must be a non-negative finite number`)
  }
  if (patch.width !== undefined && patch.width < 1) throw new RangeError('Section page width must be at least 1 pixel')
  if (patch.height !== undefined && patch.height < 1) throw new RangeError('Section page height must be at least 1 pixel')
  if (patch.pageNumberStart !== undefined && (!Number.isInteger(patch.pageNumberStart) || patch.pageNumberStart < 1)) throw new RangeError('Section page number start must be a positive integer')
  if (patch.columnCount !== undefined && (!Number.isInteger(patch.columnCount) || patch.columnCount < 1 || patch.columnCount > 12)) throw new RangeError('Section column count must be between 1 and 12')
  if (patch.columnGap !== undefined && patch.columnGap > 1000) throw new RangeError('Section column gap cannot exceed 1000 pixels')
}

export function applyDocumentSectionSetupPatch(model: DocumentModel, sectionId: string, patch: DocumentSectionSetupPatch): void {
  validateDocumentSectionSetupPatch(patch)
  if (!documentSectionIds(model).includes(sectionId)) throw new RangeError(`Unknown document section: ${sectionId}`)
  const section = sectionId === DEFAULT_DOCUMENT_SECTION_ID
    ? undefined
    : model.blocks.find(block => block.type === 'sectionBreak' && block.id === sectionId)
  if (sectionId !== DEFAULT_DOCUMENT_SECTION_ID && !section) throw new RangeError(`Unknown document section: ${sectionId}`)
  if (sectionId === DEFAULT_DOCUMENT_SECTION_ID) {
    const pagePatch: Partial<DocumentPageSetup> = {}
    for (const key of ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'] as const) {
      if (patch[key] !== undefined) pagePatch[key] = patch[key]
    }
    model.page = { ...model.page, ...pagePatch }
    const sectionPatch = { ...patch }
    for (const key of ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'] as const) delete sectionPatch[key]
    model.sectionSetup = { ...model.sectionSetup, ...sectionPatch }
    return
  }
  const block = section as DocumentBlock
  const { sectionStart, ...setupPatch } = patch
  block.sectionBreak = {
    ...block.sectionBreak!,
    ...(sectionStart === undefined ? {} : { type: sectionStart }),
    setup: { ...block.sectionBreak?.setup, ...setupPatch },
  }
}

export function documentSectionSetup(model: DocumentModel, sectionId: string): DocumentSectionSetup {
  if (!documentSectionIds(model).includes(sectionId)) throw new RangeError(`Unknown document section: ${sectionId}`)
  if (sectionId === DEFAULT_DOCUMENT_SECTION_ID) return { ...model.page, orientation: model.page.width > model.page.height ? 'landscape' : 'portrait', columnCount: 1, columnGap: 18, columnSeparator: 'none', pageNumberStart: 1, ...model.sectionSetup }
  const block = model.blocks.find(candidate => candidate.id === sectionId)
  return { ...model.page, orientation: model.page.width > model.page.height ? 'landscape' : 'portrait', columnCount: 1, columnGap: 18, columnSeparator: 'none', sectionStart: block?.sectionBreak?.type, pageNumberStart: 1, ...block?.sectionBreak?.setup }
}
