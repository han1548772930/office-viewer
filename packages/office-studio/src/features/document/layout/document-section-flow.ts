import type { DocumentModel, DocumentSectionBreakType } from '../../../domain/model'
import { documentEffectiveSectionSetup, type EffectiveDocumentSectionSetup } from '../services/document-settings'
import { DEFAULT_DOCUMENT_SECTION_ID, documentSectionIds, resolveSectionHeaderFooter } from '../services/section-header-footer'
import { PAGE_GAP, PAGE_INSET, type LayoutColumn, type LayoutFlowLocation, type LayoutPage, type LayoutSection } from './document-layout-types'

const EPSILON = 1e-6
const equal = (left: number, right: number) => Math.abs(left - right) < EPSILON

export interface DocumentFlowPageFrame extends Pick<LayoutPage, 'index' | 'x' | 'y' | 'width' | 'height' | 'content' | 'orientation' | 'sectionId' | 'pageNumber' | 'pageNumberStart' | 'sectionPageIndex' | 'isParityBlank'> {
  /** A partial first note column can defer an indivisible object to a full continuation page. */
  maximumContentHeight?: number
  columnCount?: number; columnGap?: number
}

export interface DocumentSectionFlowOptions {
  reservedPageBottom?: ReadonlyMap<number, number>
  pageFrame?: (localIndex: number) => DocumentFlowPageFrame
  referencePageFloors?: ReadonlyMap<string, number>
  beforePageOverflow?: (flow: DocumentSectionFlow) => void
  onPageStart?: (flow: DocumentSectionFlow) => void
}

export interface DocumentAuxiliaryFlowSegment {
  beforePageOverflow?: (flow: DocumentSectionFlow) => void
  onPageStart?: (flow: DocumentSectionFlow) => void
  reservedBottom?: (page: LayoutPage) => number
  prefixHeight?: (page: LayoutPage) => number
}

/** Stateful layout cursor; all state belongs to the derived layout, never the document. */
export class DocumentSectionFlow {
  readonly settings: Map<string, EffectiveDocumentSectionSetup>
  readonly pages: LayoutPage[] = []
  page!: LayoutPage
  section!: LayoutSection
  column!: LayoutColumn
  cursor = 0
  private config: EffectiveDocumentSectionSetup
  private logicalPageNumber: number
  private numberingStart: number
  private fragment = 0
  private boundary: 'page' | 'section' | undefined
  private readonly firstSectionPages = new Map<string, number>()
  private readonly oddEvenSections: Set<string>
  private readonly populatedPages = new Set<number>()
  private readonly maximumHeights = new Map<number, number>()
  private readonly frameColumns = new Map<number, { count: number; gap: number }>()
  private auxiliarySegment?: DocumentAuxiliaryFlowSegment
  private readonly columnInsets: Array<{ left: number; right: number }> = []
  private readonly columnBases = new Map<LayoutColumn, { x: number; width: number }>()

  constructor(private readonly model: DocumentModel, private readonly options: DocumentSectionFlowOptions = {}) {
    const ids = documentSectionIds(model)
    this.settings = new Map(ids.map(id => [id, documentEffectiveSectionSetup(model, id)]))
    this.oddEvenSections = new Set(ids.filter(id => resolveSectionHeaderFooter(model, id).differentOddEven))
    this.config = this.settings.get(DEFAULT_DOCUMENT_SECTION_ID)!
    this.numberingStart = this.config.pageNumberStart ?? 1
    this.logicalPageNumber = this.numberingStart
    this.firstSectionPages.set(this.config.sectionId, 0)
    this.activatePage(this.appendPage(this.logicalPageNumber, false))
  }

  get location(): LayoutFlowLocation {
    return { sectionId: this.config.sectionId, sectionFragment: this.section.fragment, column: this.column.index }
  }

  /** Decorated containers narrow the same pagination cursor, including continuation columns. */
  withColumnInsets(left: number, right: number, work: () => void): void {
    if (![left, right].every(value => Number.isFinite(value) && value >= 0)) throw new RangeError('容器内边距无效。')
    this.columnInsets.push({ left, right }); this.applyColumnInsets(this.column)
    try { work() } finally {
      this.columnInsets.pop()
      for (const column of this.columnBases.keys()) this.applyColumnInsets(column)
      if (!this.columnInsets.length) this.columnBases.clear()
    }
  }

  /** Independent column contents share parent physical frames, rather than creating document pages. */
  fork(pageFrame: NonNullable<DocumentSectionFlowOptions['pageFrame']>): DocumentSectionFlow {
    return new DocumentSectionFlow(this.model, { pageFrame })
  }

  private applyColumnInsets(column: LayoutColumn): void {
    if (!this.columnBases.has(column)) this.columnBases.set(column, { x: column.x, width: column.width })
    const base = this.columnBases.get(column)!
    let x = base.x, width = base.width
    for (const inset of this.columnInsets) {
      const sum = inset.left + inset.right, scale = sum > width - 1 ? Math.max(0, width - 1) / Math.max(1, sum) : 1
      x += inset.left * scale; width = Math.max(1, width - sum * scale)
    }
    column.x = x; column.width = width
  }

  get fullPageContentHeight(): number {
    return Math.max(0, (this.maximumHeights.get(this.page.index) ?? this.page.content.height) - (this.auxiliarySegment?.reservedBottom?.(this.page) ?? 0)
      - (this.auxiliarySegment?.prefixHeight?.(this.page) ?? 0))
  }

  referencePageFloor(runIds: readonly string[]): number {
    return runIds.reduce((page, id) => Math.max(page, this.options.referencePageFloors?.get(id) ?? 0), 0)
  }

  /** Retain reference boundaries from earlier passes while recomputing their reservations. */
  advanceToPage(index: number): void {
    while (this.page.index < index) this.nextColumn()
  }

  /** Endnotes consume full-width flow after every populated column of a section. */
  beginAuxiliarySegment(segment: DocumentAuxiliaryFlowSegment): void {
    const top = this.section.y + this.section.usedHeight
    this.auxiliarySegment = segment
    this.config = { ...this.config, columnCount: 1, columnGap: 0, columnSeparator: 'none' }
    this.activateSection(this.createSection(this.page, top))
  }

  endAuxiliarySegment(): void {
    const top = this.section.y + this.section.usedHeight
    this.auxiliarySegment = undefined
    this.config = this.settings.get(this.config.sectionId)!
    this.activateSection(this.createSection(this.page, top))
  }

  /** Append a real physical page for a note continuation beyond the body. */
  appendAuxiliaryPage(): LayoutPage { this.newPage('overflow'); return this.page }

  addSpace(height: number): void {
    this.cursor += height
    this.recordHeight()
  }

  consume(height: number): void {
    this.cursor += height
    this.populatedPages.add(this.page.index)
    this.boundary = undefined
    this.recordHeight()
  }

  private recordHeight(): void {
    this.column.usedHeight = Math.max(this.column.usedHeight, this.cursor - this.column.y)
    this.section.usedHeight = Math.max(...this.section.columns.map(column => column.usedHeight))
  }

  /** An oversized indivisible row is placed once on a fresh full-height column. */
  ensureRoom(height: number): void {
    while (this.cursor + height > this.column.y + this.column.height + EPSILON) {
      if (this.column.height <= EPSILON) { this.nextColumn(); continue }
      const prefix = this.auxiliarySegment?.prefixHeight?.(this.page) ?? 0
      const freshFullColumn = this.cursor <= this.column.y + prefix + EPSILON && equal(this.column.y, this.page.content.y)
        && equal(this.column.height - prefix, this.fullPageContentHeight)
      if (freshFullColumn && height > this.column.height - prefix) return
      this.nextColumn()
    }
  }

  nextColumn(explicit = false): void {
    if (explicit) this.section.hasExplicitColumnBreak = true
    this.column.full = true
    const next = this.section.columns[this.column.index + 1]
    if (next) { this.column = next; if (this.columnInsets.length) this.applyColumnInsets(next); this.cursor = next.y; return }
    const previousPage = this.page
    this.auxiliarySegment?.beforePageOverflow?.(this)
    if (this.page !== previousPage) return
    this.options.beforePageOverflow?.(this)
    if (this.page !== previousPage) return
    this.newPage('overflow')
  }

  pageBreak(): void {
    // A rendered page boundary at the beginning of a new section already has
    // its physical page. Consecutive explicit page breaks remain distinct.
    if (!this.populatedPages.has(this.page.index) && this.boundary === 'section') { this.boundary = 'page'; return }
    this.newPage('page')
  }

  private newPage(reason: 'page' | 'overflow'): void {
    this.logicalPageNumber++
    this.activatePage(this.appendPage(this.logicalPageNumber, false))
    this.boundary = reason === 'page' ? 'page' : undefined
    this.auxiliarySegment?.onPageStart?.(this)
    this.options.onPageStart?.(this)
  }

  private compatiblePage(config: EffectiveDocumentSectionSetup, includeMargins: boolean): boolean {
    if (!equal(this.page.width, config.page.width) || !equal(this.page.height, config.page.height) || this.page.orientation !== config.orientation) return false
    if (!includeMargins) return true
    return equal(this.page.content.x - this.page.x, config.page.marginLeft)
      && equal(this.page.content.y - this.page.y, config.page.marginTop)
      && equal(this.page.width - this.page.content.width - (this.page.content.x - this.page.x), config.page.marginRight)
      && equal(this.page.height - this.page.content.height - (this.page.content.y - this.page.y), config.page.marginBottom)
  }

  columnsChange(sectionId: string): boolean {
    const config = this.settings.get(sectionId)
    if (!config) throw new RangeError(`Unknown document section: ${sectionId}`)
    const width = (config.page.width - config.page.marginLeft - config.page.marginRight - config.columnGap * (config.columnCount - 1)) / config.columnCount
    return config.columnCount !== this.section.columns.length || !equal(config.columnGap, this.config.columnGap)
      || this.section.columns.some(column => !equal(column.width, width))
  }

  startSection(sectionId: string, type: DocumentSectionBreakType): void {
    const config = this.settings.get(sectionId)
    if (!config) throw new RangeError(`Unknown document section: ${sectionId}`)
    const sectionType = type === 'unspecified' ? 'nextPage' : type
    const previousNumber = this.logicalPageNumber, previousColumn = this.column.index, previousTop = this.section.y
    const physicalCompatible = this.compatiblePage(config, false), marginsCompatible = this.compatiblePage(config, true)
    const bottom = this.section.y + this.section.usedHeight
    const columnsFit = this.page.content.width - config.columnGap * (config.columnCount - 1) > 0
    const reusePageBoundary = !this.populatedPages.has(this.page.index) && this.boundary === 'page'
      && (sectionType === 'nextPage' || sectionType === 'continuous')
    this.config = config
    this.logicalPageNumber = config.pageNumberStart ?? previousNumber
    this.numberingStart = this.logicalPageNumber

    if (!reusePageBoundary && sectionType === 'continuous' && physicalCompatible && columnsFit && bottom < this.page.content.y + this.page.content.height - EPSILON) {
      this.firstSectionPages.set(sectionId, this.page.index)
      this.activateSection(this.createSection(this.page, bottom))
      this.boundary = undefined
      return
    }
    if (sectionType === 'nextColumn' && marginsCompatible && columnsFit && previousColumn + 1 < config.columnCount) {
      this.firstSectionPages.set(sectionId, this.page.index)
      const fragment = this.createSection(this.page, previousTop)
      fragment.columns.slice(0, previousColumn + 1).forEach(column => { column.full = true })
      this.activateSection(fragment, previousColumn + 1)
      this.boundary = undefined
      return
    }

    let pageNumber = config.pageNumberStart ?? previousNumber + (reusePageBoundary ? 0 : 1)
    if (reusePageBoundary) this.pages.pop()
    if (sectionType === 'nextPage' && config.pageNumberStart !== undefined && this.oddEvenSections.has(sectionId)
      && (this.pages.length + 1) % 2 !== config.pageNumberStart % 2) this.appendPage(previousNumber + 1, true)
    if ((sectionType === 'evenPage' || sectionType === 'oddPage') && pageNumber % 2 !== (sectionType === 'evenPage' ? 0 : 1)) {
      this.appendPage(pageNumber, true)
      pageNumber++
    }
    this.logicalPageNumber = pageNumber
    this.numberingStart = pageNumber
    this.firstSectionPages.set(sectionId, this.pages.length)
    this.activatePage(this.appendPage(pageNumber, false))
    this.boundary = 'section'
  }

  private appendPage(pageNumber: number, isParityBlank: boolean): LayoutPage {
    if (this.pages.length >= 40_000) throw new RangeError('文档排版超过单次物理页预算。')
    const setup = this.config.page, previous = this.pages.at(-1), index = this.pages.length
    const y = previous ? previous.y + previous.height + PAGE_GAP : PAGE_INSET
    const frame = this.options.pageFrame?.(index)
    const page: LayoutPage = {
      index, x: PAGE_INSET, y, width: setup.width, height: setup.height, orientation: this.config.orientation,
      content: { x: PAGE_INSET + setup.marginLeft, y: y + setup.marginTop,
        width: setup.width - setup.marginLeft - setup.marginRight, height: setup.height - setup.marginTop - setup.marginBottom },
      sectionId: this.config.sectionId, pageNumber, pageNumberStart: this.numberingStart,
      sectionPageIndex: index - (this.firstSectionPages.get(this.config.sectionId) ?? index), isParityBlank,
      ...(frame ? { ...frame, content: { ...frame.content } } : {}), sections: [],
    }
    const reserved = this.options.reservedPageBottom?.get(page.index) ?? 0
    if (!Number.isFinite(reserved) || reserved < 0 || reserved > page.content.height) throw new RangeError('脚注预留空间必须位于真实页面内容高度内。')
    this.maximumHeights.set(page.index, frame?.maximumContentHeight ?? page.content.height - reserved)
    if (frame?.columnCount !== undefined) this.frameColumns.set(page.index, { count: frame.columnCount, gap: frame.columnGap ?? 0 })
    this.pages.push(page)
    return page
  }

  private activatePage(page: LayoutPage): void {
    this.page = page
    this.activateSection(this.createSection(page, page.content.y))
  }

  private createSection(page: LayoutPage, y: number): LayoutSection {
    const frame = this.frameColumns.get(page.index), count = this.auxiliarySegment ? 1 : frame?.count ?? this.config.columnCount,
      gap = this.auxiliarySegment ? 0 : frame?.gap ?? this.config.columnGap
    const available = page.content.width - gap * (count - 1)
    if (available <= 0) throw new RangeError('Section columns and gaps must leave a positive content width')
    const width = available / count
    const reserved = this.options.reservedPageBottom?.get(page.index) ?? 0
    const height = Math.max(0, page.content.y + page.content.height - reserved - (this.auxiliarySegment?.reservedBottom?.(page) ?? 0) - y)
    const columns = Array.from({ length: count }, (_, index): LayoutColumn => ({
      index, x: page.content.x + index * (width + gap), y, width, height, usedHeight: 0, full: false,
    }))
    const section: LayoutSection = { fragment: this.fragment++, sectionId: this.config.sectionId,
      x: page.content.x, y, width: page.content.width, height, columns, usedHeight: 0,
      hasExplicitColumnBreak: false, columnSeparator: this.config.columnSeparator }
    page.sections.push(section)
    return section
  }

  private activateSection(section: LayoutSection, column = 0): void {
    this.section = section
    this.column = section.columns[column]!
    if (this.columnInsets.length) this.applyColumnInsets(this.column)
    this.cursor = this.column.y
  }
}
