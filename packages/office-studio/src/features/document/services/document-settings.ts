import type { DocumentBlock, DocumentModel, DocumentPageSetup, DocumentParagraphStyle, DocumentSectionSetup } from '../../../domain/model.ts'
import { DEFAULT_DOCUMENT_SECTION_ID, documentSectionIds } from './section-header-footer.ts'
import { validateDocumentNoteSettings } from '../notes/document-note-settings'
import { markDocumentParagraphPaginationAuthored } from '../model/document-paragraph-pagination'
import { cloneModel } from '../../../domain/serialization'
import { tableForBlock } from '../model/document-model'
import { documentBlockChildContainers } from '../model/document-structural-blocks'
import { documentStoredBlockHosts } from '../model/document-block-host'
import { documentTableGrid } from '../model/document-table-grid'
import { documentTableColumnWidths } from '../model/document-table-style'

export type DocumentParagraphStylePatch = Partial<DocumentParagraphStyle>
export type DocumentSectionSetupPatch = Partial<DocumentSectionSetup>

export interface DocumentPageSetupRequest {
  pageSize: Pick<DocumentPageSetup, 'width' | 'height'>
  pageOrient: 0 | 1
  documentFlavor?: 0 | 1 | 2 | 3
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
}

export const DOCUMENT_MODERN_PAGE_MARGIN = 50 / 0.75

/** Imported documents may omit metadata; their page geometry remains authoritative. */
export function documentPageOrientation(model: Pick<DocumentModel, 'page' | 'metadata'>): 'portrait' | 'landscape' {
  const metadata = model.metadata ?? {}
  const orientation = metadata.documentPageOrientation
  if (orientation === 'portrait' || orientation === 'landscape') return orientation
  return model.page.width > model.page.height ? 'landscape' : 'portrait'
}

/** Validate public and editor DTOs before creating a candidate or touching input state. */
export function assertDocumentPageSetupRequest(value: unknown): asserts value is DocumentPageSetupRequest {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('页面设置参数无效。')
  const size = Reflect.get(value, 'pageSize')
  if (!size || typeof size !== 'object' || Array.isArray(size)) throw new TypeError('页面尺寸参数无效。')
  if (![Reflect.get(size, 'width'), Reflect.get(size, 'height')].every(dimension => typeof dimension === 'number' && Number.isFinite(dimension) && dimension >= 1)
    || !['marginTop', 'marginBottom', 'marginLeft', 'marginRight'].every(key => {
      const margin = Reflect.get(value, key)
      return typeof margin === 'number' && Number.isFinite(margin) && margin >= 0
    })) throw new RangeError('页面尺寸须为正数，页边距须为非负数。')
  const orientation = Reflect.get(value, 'pageOrient'), flavor = Reflect.get(value, 'documentFlavor')
  if (orientation !== 0 && orientation !== 1) throw new RangeError('页面方向无效。')
  if (flavor !== undefined && (typeof flavor !== 'number' || ![0, 1, 2, 3].includes(flavor))) throw new RangeError('文档模式无效。')
}

export function prepareDocumentPageSetup(model: DocumentModel, request: DocumentPageSetupRequest): DocumentModel {
  assertDocumentPageSetupRequest(request)
  const page = { ...model.page, width: request.pageSize.width, height: request.pageSize.height, marginTop: request.marginTop, marginBottom: request.marginBottom,
    marginLeft: request.marginLeft, marginRight: request.marginRight }
  if (![page.width, page.height].every(value => Number.isFinite(value) && value >= 1)
    || ![page.marginTop, page.marginBottom, page.marginLeft, page.marginRight].every(value => Number.isFinite(value) && value >= 0)) {
    throw new RangeError('页面尺寸须为正数，页边距须为非负数。')
  }
  if (page.width <= page.marginLeft + page.marginRight || page.height <= page.marginTop + page.marginBottom) throw new RangeError('页边距必须保留可用内容区。')
  const candidate = cloneModel(model)
  candidate.page = page
  if (request.documentFlavor !== undefined) candidate.metadata = { ...candidate.metadata, documentFlavor: request.documentFlavor, documentFlavorContract: 'univer' }
  const modern = candidate.metadata?.documentFlavor === 2
  const orientation = request.pageOrient === 1 ? 'landscape' : 'portrait'
  candidate.metadata = { ...candidate.metadata, documentPageOrientation: orientation }
  if (!modern) {
    candidate.sectionSetup = { ...candidate.sectionSetup, ...page, orientation }
    for (const block of candidate.blocks) if (block.type === 'sectionBreak') block.sectionBreak = {
      ...block.sectionBreak!, setup: { ...block.sectionBreak?.setup, ...page, orientation },
    }
  }
  for (const id of documentSectionIds(candidate)) documentEffectiveSectionSetup(candidate, id)
  const oldContentWidth = model.page.width - (model.metadata?.documentFlavor === 2 ? DOCUMENT_MODERN_PAGE_MARGIN * 2 : model.page.marginLeft + model.page.marginRight)
  const newContentWidth = page.width - (modern ? DOCUMENT_MODERN_PAGE_MARGIN * 2 : page.marginLeft + page.marginRight)
  if (Math.abs(oldContentWidth - newContentWidth) <= 2) return candidate
  const resizeTables = (blocks: DocumentBlock[]) => {
    for (const block of blocks) {
      const table = tableForBlock(block)
      if (table && (table.columnWidths?.length || table.style?.width !== undefined)) {
        const grid = documentTableGrid(table), widths = documentTableColumnWidths(table, grid.columnCount, oldContentWidth)
        const total = widths.reduce((sum, width) => sum + width, 0)
        if (Math.abs(total - oldContentWidth) <= 2) {
          table.columnWidths = widths.map(width => width / total * newContentWidth)
          table.style = { ...table.style, width: newContentWidth }
        }
      }
      for (const child of documentBlockChildContainers(block)) resizeTables(child)
    }
  }
  resizeTables(candidate.blocks)
  for (const entry of documentStoredBlockHosts(candidate)) resizeTables(entry.blocks)
  return candidate
}

export const DOCUMENT_PARAGRAPH_STYLE_FIELDS = [
  'align', 'indent', 'lineHeight', 'spaceBefore', 'spaceAfter', 'indentStart', 'indentEnd',
  'indentFirstLine', 'hanging', 'spacingRule', 'wordWrap', 'snapToGrid',
  'keepLines', 'keepNext', 'widowControl', 'pageBreakBefore',
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
    if (['wordWrap', 'snapToGrid', 'keepLines', 'keepNext', 'widowControl', 'pageBreakBefore'].includes(key)) {
      if (typeof value !== 'boolean') throw new TypeError(`Paragraph ${key} must be boolean`)
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
  markDocumentParagraphPaginationAuthored(block, patch)
}

export function validateDocumentSectionSetupPatch(patch: DocumentSectionSetupPatch): void {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new TypeError('Section settings patch must be an object')
  const fields = new Set(['marginHeader', 'marginFooter', 'width', 'height', 'orientation', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft', 'columnCount', 'columnGap', 'columnSeparator', 'sectionStart', 'pageNumberStart', 'noteProperties', 'suppressEndnotes'])
  for (const key of Object.keys(patch)) if (!fields.has(key)) throw new TypeError(`Unsupported section setting: ${key}`)
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
  if (patch.suppressEndnotes !== undefined && typeof patch.suppressEndnotes !== 'boolean') throw new TypeError('Section endnote suppression must be boolean')
  if (patch.noteProperties !== undefined) {
    if (!patch.noteProperties || typeof patch.noteProperties !== 'object' || Array.isArray(patch.noteProperties)) throw new TypeError('Section note properties must be an object')
    for (const key of Object.keys(patch.noteProperties)) if (key !== 'footnote' && key !== 'endnote') throw new TypeError('Unsupported section note property: ' + key)
    for (const kind of ['footnote', 'endnote'] as const) if (patch.noteProperties[kind] !== undefined) validateDocumentNoteSettings(patch.noteProperties[kind]!, kind, 'section-validation')
  }
}

export function applyDocumentSectionSetupPatch(model: DocumentModel, sectionId: string, patch: DocumentSectionSetupPatch): void {
  validateDocumentSectionSetupPatch(patch)
  if (!documentSectionIds(model).includes(sectionId)) throw new RangeError(`Unknown document section: ${sectionId}`)
  const section = sectionId === DEFAULT_DOCUMENT_SECTION_ID
    ? undefined
    : model.blocks.find(block => block.type === 'sectionBreak' && block.id === sectionId)
  if (sectionId !== DEFAULT_DOCUMENT_SECTION_ID && !section) throw new RangeError(`Unknown document section: ${sectionId}`)
  if (Object.values(patch).every(value => value === undefined)) return
  // Validate the effective geometry before publishing a mutation. The first
  // section's legacy page fields are also the defaults used by later sections.
  const prospectivePage = { ...model.page }
  const prospectiveSetup = { ...model.sectionSetup }
  if (sectionId === DEFAULT_DOCUMENT_SECTION_ID) {
    for (const key of ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'] as const) {
      if (patch[key] !== undefined) {
        prospectivePage[key] = patch[key]
        delete prospectiveSetup[key]
      }
    }
  }
  const firstSectionPatch = { ...patch }
  for (const key of ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'] as const) delete firstSectionPatch[key]
  const prospective = {
    ...model,
    page: prospectivePage,
    sectionSetup: sectionId === DEFAULT_DOCUMENT_SECTION_ID ? { ...prospectiveSetup, ...firstSectionPatch } : model.sectionSetup,
    blocks: sectionId === DEFAULT_DOCUMENT_SECTION_ID ? model.blocks : model.blocks.map(block => block !== section ? block : {
      ...block,
      sectionBreak: { ...block.sectionBreak!, setup: { ...block.sectionBreak?.setup, ...patch } },
    }),
  }
  for (const id of sectionId === DEFAULT_DOCUMENT_SECTION_ID ? documentSectionIds(model) : [sectionId]) documentEffectiveSectionSetup(prospective, id)
  if (sectionId === DEFAULT_DOCUMENT_SECTION_ID) {
    const pagePatch: Partial<DocumentPageSetup> = {}
    for (const key of ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'] as const) {
      if (patch[key] !== undefined) pagePatch[key] = patch[key]
    }
    model.page = { ...model.page, ...pagePatch }
    model.sectionSetup = { ...prospectiveSetup, ...firstSectionPatch }
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
  if (sectionId === DEFAULT_DOCUMENT_SECTION_ID) {
    const setup: DocumentSectionSetup = { ...model.page, columnCount: 1, columnGap: 18, columnSeparator: 'none', pageNumberStart: 1, ...model.sectionSetup }
    setup.orientation ??= (setup.width ?? model.page.width) > (setup.height ?? model.page.height) ? 'landscape' : 'portrait'
    return setup
  }
  const block = model.blocks.find(candidate => candidate.id === sectionId)
  const setup: DocumentSectionSetup = { ...model.page, columnCount: 1, columnGap: 18, columnSeparator: 'none', sectionStart: block?.sectionBreak?.type, pageNumberStart: 1, ...block?.sectionBreak?.setup }
  setup.orientation ??= (setup.width ?? model.page.width) > (setup.height ?? model.page.height) ? 'landscape' : 'portrait'
  return setup
}

export interface EffectiveDocumentSectionSetup {
  sectionId: string
  page: DocumentPageSetup
  orientation: 'portrait' | 'landscape'
  columnCount: number
  columnGap: number
  columnSeparator: 'none' | 'line'
  /** Only an explicitly persisted value restarts a later section's numbering. */
  pageNumberStart?: number
  marginHeader: number
  marginFooter: number
}

/** Pure nominal geometry; missing section overrides use the legacy document defaults. */
export function documentEffectiveSectionSetup(model: DocumentModel, sectionId: string): EffectiveDocumentSectionSetup {
  const source = documentSectionSetup(model, sectionId)
  validateDocumentSectionSetupPatch(source)
  const nominalPage: DocumentPageSetup = {
    width: source.width ?? model.page.width, height: source.height ?? model.page.height,
    marginTop: source.marginTop ?? model.page.marginTop, marginBottom: source.marginBottom ?? model.page.marginBottom,
    marginLeft: source.marginLeft ?? model.page.marginLeft, marginRight: source.marginRight ?? model.page.marginRight,
  }
  const metadata = model.metadata ?? {}
  const modern = metadata.documentFlavor === 2
  const page = modern ? { ...model.page, marginTop: DOCUMENT_MODERN_PAGE_MARGIN, marginBottom: DOCUMENT_MODERN_PAGE_MARGIN,
    marginLeft: DOCUMENT_MODERN_PAGE_MARGIN, marginRight: DOCUMENT_MODERN_PAGE_MARGIN } : nominalPage
  const contentWidth = page.width - page.marginLeft - page.marginRight
  const contentHeight = page.height - page.marginTop - page.marginBottom
  if (![page.width, page.height, contentWidth, contentHeight].every(value => Number.isFinite(value) && value > 0)) {
    throw new RangeError('Section page size and margins must leave a positive finite content area')
  }
  const columnCount = source.columnCount ?? 1, columnGap = source.columnGap ?? 18
  if (contentWidth - columnGap * (columnCount - 1) <= 0) throw new RangeError('Section columns and gaps must leave a positive content width')
  const explicit = sectionId === DEFAULT_DOCUMENT_SECTION_ID ? model.sectionSetup
    : model.blocks.find(block => block.id === sectionId)?.sectionBreak?.setup
  return {
    sectionId, page, orientation: modern ? (metadata.documentPageOrientation === 'landscape' ? 'landscape' : 'portrait') : source.orientation ?? (page.width > page.height ? 'landscape' : 'portrait'),
    columnCount, columnGap, columnSeparator: source.columnSeparator ?? 'none',
    ...(explicit?.pageNumberStart === undefined ? {} : { pageNumberStart: explicit.pageNumberStart }),
    marginHeader: source.marginHeader ?? 48, marginFooter: source.marginFooter ?? 48,
  }
}
