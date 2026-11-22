import type { DocumentBlock, DocumentModel, DocumentPageField, DocumentPageSetup } from '../../../domain/model'
import type { DocumentPoint, DocumentRun, TextTarget } from '../model/document-model'
import { hasEastAsianLetters } from '../model/document-font'
import { paragraphListMarkers, type ParagraphListMarker } from '../model/document-lists'
import { createRun, imageForBlock, runText, sameTarget, sliceRuns, tableForBlock, targetKey } from '../model/document-model'
import { DEFAULT_DOCUMENT_SECTION_ID, renderSectionStory, resolveSectionHeaderFooter } from '../services/section-header-footer'
import { documentSectionSetup } from '../services/document-settings'
import { codeBlockData } from '../code-blocks/code-blocks.ts'
import { getLatexFormula, layoutLatexFormula, type LatexLayout } from '../latex/latex-formula.ts'
import { documentStoryDisplay } from '../services/story-rendering'

export interface LayoutRect { x: number; y: number; width: number; height: number }
export interface LayoutGlyph extends LayoutRect { text: string; offset: number; end: number; run: DocumentRun; font: string; latex?: LatexLayout }
export interface LayoutLine extends LayoutRect { target: TextTarget; start: number; end: number; baseline: number; glyphLineHeight?: number; glyphs: LayoutGlyph[]; page: number; marker?: string; markerStyle?: { x: number; font: string; color: string }; checklist?: { blockId: string; checked: boolean }; codeBlock?: { language: string; lineNumber: number } }
export interface LayoutPage extends LayoutRect { index: number; content: LayoutRect }
export interface LayoutCell extends LayoutRect { target: TextTarget; fill: string; page: number }
export interface LayoutImage extends LayoutRect { blockId: string; src: string; alt: string; page: number }
export interface LayoutPageContent {
  lines: LayoutLine[]
  cells: LayoutCell[]
  images: LayoutImage[]
  dividers: Array<LayoutRect & { page: number }>
}
export const STORY_LINE_HEIGHT = 14
export type LayoutStoryLine = Pick<LayoutLine, 'x' | 'y' | 'width' | 'height' | 'start' | 'end' | 'baseline' | 'glyphs'>
export interface LayoutPageHeaderFooter extends LayoutRect { page: number; sectionId: string; headerStory: DocumentPageField['story']; footerStory: DocumentPageField['story']; header: string; footer: string; headerLines: string[]; footerLines: string[]; headerLayoutLines: LayoutStoryLine[]; footerLayoutLines: LayoutStoryLine[]; headerTop: number; footerBottom: number }
export interface DocumentLayout {
  width: number
  height: number
  pages: LayoutPage[]
  pageContent: LayoutPageContent[]
  lines: LayoutLine[]
  cells: LayoutCell[]
  images: LayoutImage[]
  dividers: Array<LayoutRect & { page: number }>
  pageHeaderFooters: LayoutPageHeaderFooter[]
  headings: Array<{ id: string; text: string; level: number; y: number; page: number }>
}
export interface LayoutMeasurer { measure(text: string, font: string): number }
export const PAGE_GAP = 24
export const PAGE_INSET = 28

function paragraphLineHeight(fontHeight: number, style: DocumentBlock['style']): number {
  if (style.spacingRule === 'exact') return Math.max(1, style.lineHeight)
  if (style.spacingRule === 'atLeast') return Math.max(fontHeight, style.lineHeight)
  return Math.max(1, fontHeight * Math.max(0.01, style.lineHeight))
}

export function runFont(run: DocumentRun, scale = 1, family = run.fontFamily): string {
  const verticalScale = run.verticalAlign === 'subscript' || run.verticalAlign === 'superscript' ? 0.65 : 1
  return `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${Math.max(1, run.fontSize) * 4 / 3 * verticalScale * scale}px ${family || 'Aptos'}, "Microsoft YaHei", sans-serif`
}

export function createTextMeasurer(): LayoutMeasurer {
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  const cache = new Map<string, number>()
  return { measure(text, font) {
    const key = `${font}\u0000${text}`
    const cached = cache.get(key)
    if (cached != null) return cached
    if (context) context.font = font
    const width = context?.measureText(text).width ?? text.length * 8
    if (cache.size > 40000) cache.clear()
    cache.set(key, width)
    return width
  } }
}

function pageSetup(source: DocumentPageSetup): DocumentPageSetup {
  const width = Math.max(200, source.width)
  const height = Math.max(240, source.height)
  return {
    width, height,
    marginLeft: Math.max(0, Math.min(source.marginLeft, width / 2)),
    marginRight: Math.max(0, Math.min(source.marginRight, width / 2)),
    marginTop: Math.max(0, Math.min(source.marginTop, height / 2)),
    marginBottom: Math.max(0, Math.min(source.marginBottom, height / 2)),
  }
}

type RawGlyph = Omit<LayoutGlyph, 'x' | 'y' | 'height'>
type RawLine = { glyphs: RawGlyph[]; start: number; end: number; width: number; fontHeight: number }

function wrapRuns(runs: readonly DocumentRun[], width: number | ((line: number) => number), measurer: LayoutMeasurer, wordWrap = false, minimumFontHeight = 14.67): RawLine[] {
  const tokens: RawGlyph[] = []
  const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' })
  let offset = 0
  for (const run of runs) {
    const latinFont = runFont(run)
    const eastAsianFont = run.eastAsiaFontFamily && runFont(run, 1, run.eastAsiaFontFamily)
    for (const part of segmenter.segment(run.text)) {
      const text = part.segment
      const font = eastAsianFont && hasEastAsianLetters(text) ? eastAsianFont : latinFont
      let latex: LatexLayout | undefined
      const formula = text === '\uFFFC' ? getLatexFormula(run) : undefined
      if (formula) {
        try { latex = layoutLatexFormula(formula.latex, { fontSize: run.fontSize * 4 / 3, measure: measurer.measure }) } catch { /* Keep malformed persisted formulas as a visible replacement glyph. */ }
      }
      tokens.push({ text, offset: offset + part.index, end: offset + part.index + text.length, width: text === '\n' || text === '\u2028' ? 0 : latex?.width ?? measurer.measure(text === '\t' ? '    ' : text, font), run, font, latex })
    }
    offset += run.text.length
  }
  const lines: RawLine[] = []
  let glyphs: RawGlyph[] = []
  let lineStart = 0
  let lineWidth = 0
  const commit = (end: number) => {
    lines.push({ glyphs, start: lineStart, end, width: lineWidth, fontHeight: Math.max(minimumFontHeight, ...(glyphs.length ? glyphs.map((glyph) => glyph.latex?.height ?? glyph.run.fontSize * 4 / 3) : [runs[0]?.fontSize ? runs[0].fontSize * 4 / 3 : minimumFontHeight])) })
    glyphs = []
    lineWidth = 0
    lineStart = end
  }
  const availableWidth = () => Math.max(1, typeof width === 'number' ? width : width(lines.length))
  const joinsWord = (token: RawGlyph) => /^[\p{L}\p{M}\p{N}_]+$/u.test(token.text) && !/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}]/u.test(token.text)
  for (let index = 0; index < tokens.length;) {
    const token = tokens[index]!
    if (token.text === '\n' || token.text === '\u2028') { commit(token.offset); lineStart = token.end; index++; continue }
    let end = index + 1
    if (!wordWrap && joinsWord(token)) while (end < tokens.length && joinsWord(tokens[end]!)) end++
    const word = tokens.slice(index, end)
    const wordWidth = word.reduce((sum, glyph) => sum + glyph.width, 0)
    if (glyphs.length && lineWidth + wordWidth > availableWidth()) commit(token.offset)
    for (const glyph of word) {
      // A word wider than the whole line still needs an emergency break.
      if (glyphs.length && lineWidth + glyph.width > availableWidth()) commit(glyph.offset)
      glyphs.push(glyph)
      lineWidth += glyph.width
    }
    index = end
  }
  commit(offset)
  return lines
}

function positionLine(raw: RawLine, target: TextTarget, x: number, y: number, width: number, lineHeight: number, page: number, align: DocumentBlock['style']['align'], justify: boolean, marker?: string): LayoutLine {
  const spare = Math.max(0, width - raw.width)
  const shift = align === 'center' ? spare / 2 : align === 'right' ? spare : 0
  const spaces = raw.glyphs.filter((glyph) => glyph.text === ' ').length
  const gap = align === 'justify' && justify && spaces ? spare / spaces : 0
  let cursor = x + shift
  const glyphs = raw.glyphs.map((glyph) => {
    const value: LayoutGlyph = { ...glyph, x: cursor, y, height: lineHeight, width: glyph.width + (glyph.text === ' ' ? gap : 0) }
    cursor += value.width
    return value
  })
  return { x: x + shift, y, width: Math.max(2, cursor - x - shift), height: lineHeight, glyphLineHeight: raw.fontHeight, target, start: raw.start, end: raw.end, baseline: y + (lineHeight - raw.fontHeight) / 2 + raw.fontHeight * 0.82, glyphs, page, marker }
}

interface ParagraphGeometry {
  indent: number
  firstIndent: number
  rightIndent: number
  markerIndent?: number
  markerFont?: string
  markerHeight: number
  markerWidth: number
  width: (index: number) => number
}

function paragraphGeometry(block: DocumentBlock, list: ParagraphListMarker | undefined, available: number, measurer: LayoutMeasurer): ParagraphGeometry {
  const legacyIndent = block.style.indent * 24
  const left = block.style.indentStart ?? legacyIndent
  let indent = left, firstIndent = left + (block.style.indentFirstLine ?? 0) - (block.style.hanging ?? 0)
  let rightIndent = block.style.indentEnd ?? 0
  let markerIndent: number | undefined, markerFont: string | undefined, markerHeight = 0, markerWidth = 0
  if (list) {
    const extra = block.style.indentStart === undefined ? legacyIndent : 0
    indent = list.inset + extra
    rightIndent = list.rightIndent
    markerIndent = list.markerIndent + extra
    markerFont = runFont(createRun('', { fontFamily: list.fontFamily, fontSize: list.fontSize, bold: list.bold, italic: list.italic }))
    markerHeight = list.picture ? list.picture.heightPx ?? list.fontSize * 4 / 3 : list.fontSize * 4 / 3
    markerWidth = list.picture ? list.picture.widthPx ?? markerHeight : measurer.measure(list.marker, markerFont)
    firstIndent = Math.max(indent, markerIndent + markerWidth + (markerWidth ? list.gap : 0))
  } else if (block.type === 'checklist') {
    indent += 28 * (block.level + 1)
    firstIndent = indent
  } else if (block.type === 'quote') {
    indent += 16
    firstIndent += 16
  }
  return { indent, firstIndent, rightIndent, markerIndent, markerFont, markerHeight, markerWidth,
    width: index => Math.max(1, available - (index === 0 ? firstIndent : indent) - rightIndent) }
}

function appendListMarker(line: LayoutLine, block: DocumentBlock, list: ParagraphListMarker | undefined, geometry: ParagraphGeometry, left: number, images: LayoutImage[]): void {
  if (!list || geometry.markerIndent === undefined) return
  if (list.picture?.src) {
    images.push({ blockId: block.id, src: list.picture.src, alt: '项目符号', x: left + geometry.markerIndent,
      y: line.y + (line.height - geometry.markerHeight) / 2, width: geometry.markerWidth, height: geometry.markerHeight, page: line.page })
  } else if (geometry.markerFont) line.markerStyle = { x: left + geometry.markerIndent, font: geometry.markerFont, color: list.color }
}

export function layoutDocument(model: DocumentModel, measurer: LayoutMeasurer): DocumentLayout {
  const setup = pageSetup(model.page)
  const result: DocumentLayout = { width: setup.width + PAGE_INSET * 2, height: 0, pages: [], pageContent: [], lines: [], cells: [], images: [], dividers: [], pageHeaderFooters: [], headings: [] }
  let currentPage: LayoutPage
  let cursor = 0
  let sectionId = DEFAULT_DOCUMENT_SECTION_ID
  const pageSections: string[] = []
  const addPage = () => {
    const index = result.pages.length
    const top = PAGE_INSET + index * (setup.height + PAGE_GAP)
    currentPage = { index, x: PAGE_INSET, y: top, width: setup.width, height: setup.height, content: { x: PAGE_INSET + setup.marginLeft, y: top + setup.marginTop, width: setup.width - setup.marginLeft - setup.marginRight, height: setup.height - setup.marginTop - setup.marginBottom } }
    result.pages.push(currentPage)
    pageSections.push(sectionId)
    cursor = currentPage.content.y
  }
  addPage()
  const bodyMarkers = paragraphListMarkers(model.blocks)
  const ensureRoom = (height: number) => { if (cursor + height > currentPage.content.y + currentPage.content.height && cursor > currentPage.content.y + 0.5) addPage() }
  for (const block of model.blocks) {
    if (block.type === 'pageBreak' || block.type === 'columnBreak') { addPage(); continue }
    if (block.type === 'sectionBreak') {
      const type = block.sectionBreak?.type === 'unspecified' ? 'nextPage' : block.sectionBreak?.type ?? 'nextPage'
      sectionId = block.id
      if (type !== 'continuous') {
        addPage()
        if (type === 'evenPage' && currentPage!.index % 2 === 0) addPage()
        if (type === 'oddPage' && currentPage!.index % 2 !== 0) addPage()
      } else pageSections[currentPage!.index] = sectionId
      continue
    }
    cursor += block.style.spaceBefore
    if (block.type === 'table') {
      const table = tableForBlock(block)
      if (!table?.rows.length) continue
      const colCount = Math.max(1, ...table.rows.map((row) => row.cells.reduce((sum, cell) => sum + Math.max(1, cell.colSpan), 0)))
      const available = currentPage!.content.width
      const widths = table.columnWidths?.length === colCount ? table.columnWidths.map((width) => width * available / table.columnWidths!.reduce((sum, value) => sum + value, 0)) : Array.from({ length: colCount }, () => available / colCount)
      table.rows.forEach((row, rowIndex) => {
        let column = 0
        const prepared = row.cells.map((cell, colIndex) => {
          const width = widths.slice(column, column + Math.max(1, cell.colSpan)).reduce((sum, value) => sum + value, 0)
          const xOffset = widths.slice(0, column).reduce((sum, value) => sum + value, 0)
          column += Math.max(1, cell.colSpan)
          const paragraphs = cell.blocks.filter((paragraph) => !['table', 'image', 'divider', 'pageBreak'].includes(paragraph.type)).map((paragraph) => {
            const list = bodyMarkers.get(paragraph.id)
            const geometry = paragraphGeometry(paragraph, list, width - 16, measurer)
            return { paragraph, list, geometry, lines: wrapRuns(paragraph.runs as DocumentRun[], geometry.width, measurer, paragraph.style.wordWrap === true) }
          })
          const height = Math.max(34, 16 + paragraphs.reduce((sum, item) => sum + item.paragraph.style.spaceBefore + item.lines.reduce((lineSum, line, index) => lineSum + paragraphLineHeight(Math.max(line.fontHeight, index === 0 ? item.geometry.markerHeight : 0), item.paragraph.style), 0) + item.paragraph.style.spaceAfter, 0))
          return { cell, colIndex, width, xOffset, paragraphs, height }
        })
        const rowHeight = Math.max(34, ...prepared.map((cell) => cell.height))
        ensureRoom(rowHeight)
        for (const preparedCell of prepared) {
          const { cell, colIndex, width, xOffset, paragraphs } = preparedCell
          const x = currentPage!.content.x + xOffset
          const target = { blockId: block.id, row: rowIndex, col: colIndex }
          result.cells.push({ x, y: cursor, width, height: rowHeight, page: currentPage!.index, target, fill: typeof cell.style.fill === 'string' ? cell.style.fill : typeof cell.style.background === 'string' ? cell.style.background : '#ffffff' })
          let cellY = cursor + 8
          for (const item of paragraphs) {
            cellY += item.paragraph.style.spaceBefore
            item.lines.forEach((line, index) => {
              const lineHeight = paragraphLineHeight(Math.max(line.fontHeight, index === 0 ? item.geometry.markerHeight : 0), item.paragraph.style)
              const indent = index === 0 ? item.geometry.firstIndent : item.geometry.indent
              const layoutLine = positionLine(line, { ...target, cellBlockId: item.paragraph.id }, x + 8 + indent, cellY, item.geometry.width(index), lineHeight, currentPage!.index, item.paragraph.style.align, index < item.lines.length - 1, index === 0 ? item.list?.marker : undefined)
              if (index === 0) appendListMarker(layoutLine, item.paragraph, item.list, item.geometry, x + 8, result.images)
              if (index === 0 && item.paragraph.type === 'checklist') layoutLine.checklist = { blockId: item.paragraph.id, checked: item.paragraph.checklist?.checked === true }
              result.lines.push(layoutLine)
              cellY += lineHeight
            })
            cellY += item.paragraph.style.spaceAfter
          }
        }
        cursor += rowHeight
      })
    } else if (block.type === 'image') {
      const image = imageForBlock(block)
      if (!image) continue
      const width = Math.min(currentPage!.content.width, image.width || 320)
      const height = Math.min(currentPage!.content.height, image.height ? image.height * width / (image.width || width) : width * 0.625)
      ensureRoom(height)
      const shift = block.style.align === 'center' ? (currentPage!.content.width - width) / 2 : block.style.align === 'right' ? currentPage!.content.width - width : 0
      result.images.push({ x: currentPage!.content.x + shift, y: cursor, width, height, blockId: block.id, src: image.src, alt: image.alt, page: currentPage!.index })
      cursor += height
    } else if (block.type === 'divider') {
      ensureRoom(18)
      result.dividers.push({ x: currentPage!.content.x, y: cursor + 9, width: currentPage!.content.width, height: 1, page: currentPage!.index })
      cursor += 18
    } else {
      const list = bodyMarkers.get(block.id)
      const geometry = paragraphGeometry(block, list, currentPage!.content.width, measurer)
      const codeData = codeBlockData(block)
      const lines = wrapRuns(block.runs as DocumentRun[], geometry.width, measurer, block.style.wordWrap === true)
      lines.forEach((line, index) => {
        const height = paragraphLineHeight(Math.max(line.fontHeight, index === 0 ? geometry.markerHeight : 0), block.style)
        ensureRoom(height)
        if (index === 0 && block.type === 'heading') result.headings.push({ id: block.id, text: runText(block.runs), level: block.level || 1, y: cursor, page: currentPage!.index })
        const marker = index === 0 ? list?.marker ?? (block.type === 'quote' ? '│' : undefined) : undefined
        const indent = index === 0 ? geometry.firstIndent : geometry.indent
        const layoutLine = positionLine(line, { blockId: block.id }, currentPage!.content.x + indent, cursor, geometry.width(index), height, currentPage!.index, block.style.align, index < lines.length - 1, marker)
        if (index === 0) appendListMarker(layoutLine, block, list, geometry, currentPage!.content.x, result.images)
        if (codeData) {
          layoutLine.codeBlock = { language: codeData.language, lineNumber: index + 1 }
          layoutLine.x += 8
          layoutLine.width = Math.max(2, layoutLine.width - 16)
          layoutLine.glyphs.forEach((glyph) => { glyph.x += 8 })
        }
        if (index === 0 && block.type === 'checklist') layoutLine.checklist = { blockId: block.id, checked: block.checklist?.checked === true }
        result.lines.push(layoutLine)
        cursor += height
      })
    }
    cursor += block.style.spaceAfter
  }
  result.height = result.pages.length * (setup.height + PAGE_GAP) + PAGE_INSET * 2 - PAGE_GAP
  result.pageContent = result.pages.map(() => ({ lines: [], cells: [], images: [], dividers: [] }))
  for (const line of result.lines) result.pageContent[line.page]?.lines.push(line)
  for (const cell of result.cells) result.pageContent[cell.page]?.cells.push(cell)
  for (const image of result.images) result.pageContent[image.page]?.images.push(image)
  for (const divider of result.dividers) result.pageContent[divider.page]?.dividers.push(divider)
  const sectionPageCounts = new Map<string, number>()
  result.pageHeaderFooters = result.pages.map((page) => {
    const pageSectionId = pageSections[page.index] ?? DEFAULT_DOCUMENT_SECTION_ID
    const section = resolveSectionHeaderFooter(model, pageSectionId)
    const sectionPageIndex = sectionPageCounts.get(pageSectionId) ?? 0
    sectionPageCounts.set(pageSectionId, sectionPageIndex + 1)
    const firstPage = section.differentFirstPage && sectionPageIndex === 0
    const evenPage = section.differentOddEven && page.index % 2 === 1
    const headerStory = firstPage ? 'firstPageHeader' : evenPage ? 'evenPageHeader' : 'header'
    const footerStory = firstPage ? 'firstPageFooter' : evenPage ? 'evenPageFooter' : 'footer'
    const header = renderSectionStory(section, headerStory, page.index + 1, result.pages.length)
    const footer = renderSectionStory(section, footerStory, page.index + 1, result.pages.length)
    const setup = documentSectionSetup(model, pageSectionId)
    const headerTop = page.y + (setup.marginHeader ?? 48), footerBottom = page.y + page.height - (setup.marginFooter ?? 48)
    const storyLines = (story: DocumentPageField['story'], top: number, bottom = false): LayoutStoryLine[] => {
      const display = documentStoryDisplay(section, story, page.index + 1, result.pages.length, model)
      const output: LayoutStoryLine[] = []
      let cursor = 0
      for (const paragraph of display.paragraphs) {
        const style = paragraph.style, runs = sliceRuns(display.runs, paragraph.start, paragraph.end)
        if (!runs.length) runs.push(paragraph.emptyRun)
        const left = style.indentStart ?? style.indent * 24, right = style.indentEnd ?? 0
        const first = (style.indentFirstLine ?? 0) - (style.hanging ?? 0)
        const available = (index: number) => Math.max(1, page.content.width - left - right - (index === 0 ? first : 0))
        cursor += style.spaceBefore
        const wrapped = wrapRuns(runs, available, measurer, style.wordWrap === true, 12)
        wrapped.forEach((raw, index) => {
          const height = paragraphLineHeight(raw.fontHeight, style)
          const line = positionLine(raw, { blockId: '' }, page.content.x + left + (index === 0 ? first : 0), cursor, available(index), height, page.index, style.align, index < wrapped.length - 1)
          line.glyphs = line.glyphs.map(glyph => { const range = display.inputRange(paragraph.start + glyph.offset, paragraph.start + glyph.end); return { ...glyph, offset: range.start, end: range.end } })
          const range = raw.glyphs.length ? display.inputRange(paragraph.start + raw.start, paragraph.start + raw.end)
            : raw.start > 0 ? display.inputRange(paragraph.start + raw.start, paragraph.start + raw.start)
              : { start: paragraph.bufferStart, end: paragraph.start === paragraph.end ? paragraph.bufferEnd : paragraph.bufferStart }
          output.push({ ...line, start: range.start, end: range.end })
          cursor += height
        })
        cursor += style.spaceAfter
      }
      const shift = bottom ? top - cursor : top
      return output.map(line => ({ ...line, y: line.y + shift, baseline: line.baseline + shift, glyphs: line.glyphs.map(glyph => ({ ...glyph, y: glyph.y + shift })) }))
    }
    const headerLayoutLines = storyLines(headerStory, headerTop), footerLayoutLines = storyLines(footerStory, footerBottom, true)
    return { page: page.index, sectionId: pageSectionId, headerStory, footerStory, x: page.content.x, y: page.y + 6, width: page.content.width, height: page.content.y - page.y - 12, header, footer,
      headerLines: headerLayoutLines.map(line => line.glyphs.map(glyph => glyph.text).join('')), footerLines: footerLayoutLines.map(line => line.glyphs.map(glyph => glyph.text).join('')),
      headerLayoutLines, footerLayoutLines, headerTop, footerBottom }
  })
  return result
}

export function lineForPoint(layout: DocumentLayout, point: DocumentPoint): LayoutLine | undefined {
  const lines = layout.lines.filter((line) => sameTarget(line.target, point))
  return lines.find((line, index) => point.offset >= line.start && (point.offset < line.end || index === lines.length - 1 || point.offset === line.end && line.end !== lines[index + 1]?.start)) ?? lines.at(-1)
}

export function documentStoryLayoutLines(layout: DocumentLayout, location: { page: number; area: 'header' | 'footer' }): LayoutStoryLine[] {
  const story = layout.pageHeaderFooters[location.page]
  return (location.area === 'header' ? story?.headerLayoutLines : story?.footerLayoutLines) ?? []
}

export function documentStoryCaretRect(lines: readonly LayoutStoryLine[], offset: number): LayoutRect | undefined {
  // A cached field may wrap across several lines, all carrying the same object
  // range. Its leading caret uses the first glyph and its trailing caret uses
  // the last glyph, rather than stopping at the first wrapped fragment.
  for (const line of lines) {
    const next = line.glyphs.find(glyph => glyph.offset === offset)
    if (next) return { x: next.x, y: line.y, width: 1, height: line.height }
    if (!line.glyphs.length && offset === line.start) return { x: line.x, y: line.y, width: 1, height: line.height }
  }
  for (const line of [...lines].reverse()) {
    const last = [...line.glyphs].reverse().find(glyph => glyph.end === offset)
    if (last) return { x: last.x + last.width, y: line.y, width: 1, height: line.height }
    if (!line.glyphs.length && offset === line.end) return { x: line.x, y: line.y, width: 1, height: line.height }
  }
  const line = lines.find((line, index) => offset >= line.start && (offset < line.end || index === lines.length - 1 || offset === line.end && line.end !== lines[index + 1]?.start)) ?? lines.at(-1)
  if (!line) return undefined
  const next = line.glyphs.find(glyph => glyph.offset >= offset)
  const last = line.glyphs.at(-1)
  return { x: next?.x ?? (last ? last.x + last.width : line.x), y: line.y, width: 1, height: line.height }
}

export function hitTestDocumentStory(lines: readonly LayoutStoryLine[], x: number, y: number): number {
  const line = lines.find(line => y >= line.y && y < line.y + line.height)
    ?? [...lines].sort((a, b) => Math.abs(y - a.y - a.height / 2) - Math.abs(y - b.y - b.height / 2))[0]
  if (!line) return 0
  for (let index = 0; index < line.glyphs.length;) {
    const first = line.glyphs[index]!
    let last = first
    while (index + 1 < line.glyphs.length && line.glyphs[index + 1]!.offset === first.offset && line.glyphs[index + 1]!.end === first.end) last = line.glyphs[++index]!
    if (x < first.x + (last.x + last.width - first.x) / 2) return first.offset
    if (x < last.x + last.width) return first.end
    index++
  }
  return line.end
}

export function caretRect(layout: DocumentLayout, point: DocumentPoint): LayoutRect | undefined {
  const line = lineForPoint(layout, point)
  if (!line) return undefined
  const glyph = line.glyphs.find((candidate) => point.offset <= candidate.offset)
  const x = glyph?.x ?? (line.glyphs.at(-1) ? line.glyphs.at(-1)!.x + line.glyphs.at(-1)!.width : line.x)
  return { x, y: line.y + 2, width: 1.3, height: Math.max(12, line.height - 4) }
}

export function hitTestDocument(layout: DocumentLayout, x: number, y: number): DocumentPoint | undefined {
  let nearest: LayoutLine | undefined
  let best = Number.POSITIVE_INFINITY
  for (const line of layout.lines) {
    const dy = y < line.y ? line.y - y : y > line.y + line.height ? y - line.y - line.height : 0
    const dx = x < line.x ? line.x - x : x > line.x + line.width ? x - line.x - line.width : 0
    const score = dy * 8 + dx
    if (score < best) { best = score; nearest = line }
  }
  if (!nearest) return undefined
  const glyph = nearest.glyphs.find((candidate) => x < candidate.x + candidate.width / 2)
  return { ...nearest.target, offset: glyph?.offset ?? nearest.end }
}

export function hitTestChecklist(layout: DocumentLayout, x: number, y: number): string | undefined {
  const line = layout.lines.find((candidate) => candidate.checklist && y >= candidate.y && y <= candidate.y + candidate.height && x >= candidate.x - 28 && x <= candidate.x - 6)
  return line?.checklist?.blockId
}

export function linesByTarget(layout: DocumentLayout): Map<string, LayoutLine[]> {
  const map = new Map<string, LayoutLine[]>()
  layout.lines.forEach((line) => map.set(targetKey(line.target), [...(map.get(targetKey(line.target)) ?? []), line]))
  return map
}

export const emptyLayout = (): DocumentLayout => ({ width: 850, height: 1179, pages: [], pageContent: [], lines: [], cells: [], images: [], dividers: [], pageHeaderFooters: [], headings: [] })
export const defaultRun = createRun
