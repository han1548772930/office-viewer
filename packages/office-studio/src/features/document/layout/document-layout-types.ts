import type { DocumentPageField } from '../../../domain/model'
import type { DocumentRun, DocumentTableBorder, DocumentTableBorderSide, TextTarget } from '../model/document-model'
import type { LatexLayout } from '../latex/latex-formula'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'

export interface LayoutRect { x: number; y: number; width: number; height: number }
export interface LayoutFlowLocation { sectionId: string; sectionFragment: number; column: number }
export type LayoutTabLeader = 'none' | 'dots' | 'dashes' | 'underline' | 'middleDots'
/** A transient right tab in the wrapped text; right is CSS px from that line's text origin. */
export interface LayoutRightTabStop { offset: number; right?: number; leader: LayoutTabLeader }
export interface LayoutTabGlyph { leader: LayoutTabLeader; leaderText: string; leaderWidth: number }
export interface LayoutGlyph extends LayoutRect { text: string; offset: number; end: number; run: DocumentRun; font: string; latex?: LatexLayout; tab?: LayoutTabGlyph }
export interface LayoutLine extends LayoutRect, Partial<LayoutFlowLocation> {
  host?: DocumentAuxiliaryBlockHost
  target: TextTarget; start: number; end: number; baseline: number; glyphLineHeight?: number; glyphs: LayoutGlyph[]; page: number
  clip?: LayoutRect; tableCellId?: string; repeatedHeader?: boolean
  marker?: string; markerStyle?: { x: number; font: string; color: string }
  checklist?: { blockId: string; checked: boolean }; codeBlock?: { language: string; lineNumber: number }
}
export interface LayoutColumn extends LayoutRect { index: number; usedHeight: number; full: boolean }
/** A physical page can contain several sections, including a NEXT_COLUMN fragment. */
export interface LayoutSection extends LayoutRect {
  fragment: number; sectionId: string; columns: LayoutColumn[]; usedHeight: number; hasExplicitColumnBreak: boolean; columnSeparator: 'none' | 'line'
}
export interface LayoutPage extends LayoutRect {
  index: number; content: LayoutRect; sectionId: string; pageNumber: number; pageNumberStart: number
  sectionPageIndex: number; orientation: 'portrait' | 'landscape'; sections: LayoutSection[]; isParityBlank: boolean
}
export interface LayoutCell extends LayoutRect, Partial<LayoutFlowLocation> {
  host?: DocumentAuxiliaryBlockHost
  target: TextTarget; fill: string; page: number; clip?: LayoutRect
  tableId?: string; cellId?: string; row?: number; columnIndex?: number; rowSpan?: number; colSpan?: number; depth?: number
  borders?: Record<DocumentTableBorderSide, DocumentTableBorder>; continuedBefore?: boolean; continuedAfter?: boolean; repeatedHeader?: boolean
}
export interface LayoutImage extends LayoutRect, Partial<LayoutFlowLocation> {
  host?: DocumentAuxiliaryBlockHost
  blockId: string; src: string; alt: string; page: number; clip?: LayoutRect; tableCellId?: string; repeatedHeader?: boolean
}
export interface LayoutDivider extends LayoutRect, Partial<LayoutFlowLocation> { page: number; clip?: LayoutRect; host?: DocumentAuxiliaryBlockHost }
export interface LayoutTableRowMetric { row: number; id: string; y: number; height: number; contentHeight?: number }
/** A table can have several fragments on pages/columns; repeated headers share model IDs. */
export interface LayoutTableFragment extends LayoutRect, Partial<LayoutFlowLocation> {
  host?: DocumentAuxiliaryBlockHost
  tableId: string; page: number; depth: number; sourceStart: number; sourceEnd: number
  columnWidths: number[]; rows: LayoutTableRowMetric[]; repeatedHeader: boolean; continuedBefore: boolean; continuedAfter: boolean
  clip?: LayoutRect
}
export interface LayoutPageContent {
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]; columnSeparators: LayoutDivider[]
}
export const STORY_LINE_HEIGHT = 14
export type LayoutStoryLine = Pick<LayoutLine, 'x' | 'y' | 'width' | 'height' | 'start' | 'end' | 'baseline' | 'glyphs'>
export interface LayoutPageHeaderFooter extends LayoutRect {
  page: number; sectionId: string; pageNumber: number; pageNumberStart: number; isParityBlank: boolean
  headerStory: DocumentPageField['story']; footerStory: DocumentPageField['story']; header: string; footer: string
  headerLines: string[]; footerLines: string[]; headerLayoutLines: LayoutStoryLine[]; footerLayoutLines: LayoutStoryLine[]
  headerTop: number; footerBottom: number
  headerBlockLayout?: LayoutAuxiliaryContent; footerBlockLayout?: LayoutAuxiliaryContent
}
export interface LayoutAuxiliaryContent extends LayoutRect {
  host: DocumentAuxiliaryBlockHost; sourceHost: DocumentAuxiliaryBlockHost; page: number
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]; tableFragments: LayoutTableFragment[]
}
export interface DocumentLayout {
  width: number; height: number; pages: LayoutPage[]; pageContent: LayoutPageContent[]
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]; columnSeparators: LayoutDivider[]
  pageHeaderFooters: LayoutPageHeaderFooter[]; headings: Array<{ id: string; text: string; level: number; y: number; page: number }>
  tableFragments?: LayoutTableFragment[]
  auxiliaryContent?: LayoutAuxiliaryContent[]
}
export interface LayoutMeasurer { measure(text: string, font: string): number }
export const PAGE_GAP = 24
export const PAGE_INSET = 28
