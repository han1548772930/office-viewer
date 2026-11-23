import type { DocumentPageField } from '../../../domain/model'
import type { DocumentRun, TextTarget } from '../model/document-model'
import type { LatexLayout } from '../latex/latex-formula'

export interface LayoutRect { x: number; y: number; width: number; height: number }
export interface LayoutFlowLocation { sectionId: string; sectionFragment: number; column: number }
export interface LayoutGlyph extends LayoutRect { text: string; offset: number; end: number; run: DocumentRun; font: string; latex?: LatexLayout }
export interface LayoutLine extends LayoutRect, Partial<LayoutFlowLocation> {
  target: TextTarget; start: number; end: number; baseline: number; glyphLineHeight?: number; glyphs: LayoutGlyph[]; page: number
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
export interface LayoutCell extends LayoutRect, Partial<LayoutFlowLocation> { target: TextTarget; fill: string; page: number }
export interface LayoutImage extends LayoutRect, Partial<LayoutFlowLocation> { blockId: string; src: string; alt: string; page: number }
export interface LayoutDivider extends LayoutRect, Partial<LayoutFlowLocation> { page: number }
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
}
export interface DocumentLayout {
  width: number; height: number; pages: LayoutPage[]; pageContent: LayoutPageContent[]
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]; columnSeparators: LayoutDivider[]
  pageHeaderFooters: LayoutPageHeaderFooter[]; headings: Array<{ id: string; text: string; level: number; y: number; page: number }>
}
export interface LayoutMeasurer { measure(text: string, font: string): number }
export const PAGE_GAP = 24
export const PAGE_INSET = 28
