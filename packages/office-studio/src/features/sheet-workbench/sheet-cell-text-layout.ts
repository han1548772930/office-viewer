import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellShrinkToFit, sheetCellTextRotation, sheetCellWrapStrategy } from './sheet-cell-layout.ts'
import { sheetCellNumberValue, sheetCellStoredValueType } from './sheet-number-presentation.ts'

export interface SheetTextLine {
  text: string
  width: number
  /** Offsets in the original display string, including whitespace and explicit line breaks. */
  start: number
  end: number
  paragraphEnd: boolean
}
export interface SheetTextLayout {
  lines: SheetTextLine[]
  width: number
  height: number
  lineHeight: number
}
export interface SheetTextLayoutOptions {
  width: number
  lineHeight: number
  wrap: boolean
  measure: (text: string) => number
}
export type SheetTextMeasure = (text: string, font: string) => number
export interface SheetCellTextFontOverrides { bold?: boolean; italic?: boolean }
export interface SheetCellTextFont { font: string; fontSize: number; lineHeight: number }

/** Office font sizes are points; render/layout measurements are CSS pixels at the editor zoom. */
export function sheetCellTextFont(cell: ExcelCellView | undefined, zoom = 1, overrides: SheetCellTextFontOverrides = {}): SheetCellTextFont {
  const points = cell?.fontSize && Number.isFinite(cell.fontSize) && cell.fontSize > 0 ? cell.fontSize : 11
  const fontSize = points * 4 / 3 * zoom
  const prefix = `${overrides.italic ?? cell?.italic ? 'italic ' : ''}${overrides.bold ?? cell?.bold ? '600 ' : ''}`
  return { font: `${prefix}${fontSize}px ${cell?.fontFamily ?? 'Aptos, "Microsoft YaHei", sans-serif'}`, fontSize, lineHeight: fontSize * 1.2 }
}

export function sheetCellTextIndent(cell: ExcelCellView | undefined, fontSize: number): number {
  return Math.max(0, Number.isFinite(cell?.indent) ? cell!.indent! : 0) * fontSize
}

export function rotatedSheetTextSize(width: number, height: number, degrees: number): { width: number; height: number } {
  const radians = degrees * Math.PI / 180, cosine = Math.abs(Math.cos(radians)), sine = Math.abs(Math.sin(radians))
  return { width: width * cosine + height * sine, height: width * sine + height * cosine }
}

let graphemeSegmenter: Intl.Segmenter | undefined
export function sheetTextGraphemes(text: string): Array<{ text: string; offset: number }> {
  if (typeof Intl.Segmenter === 'function') {
    graphemeSegmenter ??= new Intl.Segmenter(undefined, { granularity: 'grapheme' })
    return [...graphemeSegmenter.segment(text)].map(segment => ({ text: segment.segment, offset: segment.index }))
  }
  let offset = 0
  return Array.from(text, value => { const segment = { text: value, offset }; offset += value.length; return segment })
}

const closingPunctuation = /^[，。、；：！？）》」』】〕〉〗〙〛,.!?;:)\]}]/u
const openingPunctuation = /[（《「『【〔〈〖〘〚(\[{]$/u
const cjk = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u
function breakBetween(previous: string, next: string | undefined): boolean {
  if (!next || openingPunctuation.test(previous) || closingPunctuation.test(next)) return false
  return /[\s\-/]$/u.test(previous) || cjk.test(previous) || cjk.test(next)
}

function textParagraphs(text: string): Array<{ text: string; start: number }> {
  let offset = 0
  return text.split(/\r\n|\r|\n/u).map(value => {
    const paragraph = { text: value, start: offset }
    offset += value.length
    offset += text.slice(offset, offset + 2) === '\r\n' ? 2 : offset < text.length ? 1 : 0
    return paragraph
  })
}

/** Keep grapheme clusters and explicit paragraphs while breaking long words only when needed. */
export function layoutSheetCellText(text: string, options: SheetTextLayoutOptions): SheetTextLayout {
  const width = Math.max(0, options.width), lineHeight = Math.max(0, options.lineHeight)
  const lines: SheetTextLine[] = []
  const measure = (value: string) => {
    const measured = options.measure(value)
    return Number.isFinite(measured) ? Math.max(0, measured) : 0
  }
  for (const { text: paragraph, start: offset } of textParagraphs(text)) {
    const glyphs = sheetTextGraphemes(paragraph)
    if (!options.wrap || !glyphs.length) lines.push({ text: paragraph, width: measure(paragraph), start: offset, end: offset + paragraph.length, paragraphEnd: true })
    else {
      const breaks = glyphs.flatMap((glyph, index) => breakBetween(glyph.text, glyphs[index + 1]?.text) ? [index + 1] : [])
      const candidate = (start: number, end: number) => paragraph.slice(glyphs[start]!.offset, glyphs[end - 1]!.offset + glyphs[end - 1]!.text.length)
      const fitEnd = (start: number) => {
        let best = start + 1, upper = Math.min(glyphs.length, start + 2)
        if (measure(candidate(start, best)) > width || best === glyphs.length) return best
        while (measure(candidate(start, upper)) <= width) {
          best = upper
          if (upper === glyphs.length) return upper
          upper = Math.min(glyphs.length, start + (upper - start) * 2)
        }
        while (best + 1 < upper) {
          const middle = Math.floor((best + upper) / 2)
          if (measure(candidate(start, middle)) <= width) best = middle
          else upper = middle
        }
        return best
      }
      let start = 0
      while (start < glyphs.length) {
        const end = fitEnd(start)
        let low = 0, high = breaks.length
        while (low < high) {
          const middle = (low + high) >>> 1
          if (breaks[middle]! <= end) low = middle + 1
          else high = middle
        }
        const lastBreak = breaks[low - 1] ?? start
        let boundary = end < glyphs.length && lastBreak > start ? lastBreak : end
        // A cluster wider than the column must still make progress; keep paired punctuation with it.
        if (boundary < glyphs.length && boundary === start + 1
          && (openingPunctuation.test(glyphs[boundary - 1]!.text) || closingPunctuation.test(glyphs[boundary]!.text))) boundary++
        const startOffset = glyphs[start]!.offset
        const endOffset = glyphs[boundary - 1]!.offset + glyphs[boundary - 1]!.text.length
        const value = paragraph.slice(startOffset, endOffset)
        lines.push({ text: value, width: measure(value), start: offset + startOffset, end: offset + endOffset, paragraphEnd: boundary === glyphs.length })
        start = boundary
      }
    }
  }
  return { lines, width: lines.reduce((maximum, line) => Math.max(maximum, line.width), 0), height: lines.length * lineHeight, lineHeight }
}

export interface SheetStackedTextColumn {
  glyphs: Array<{ text: string; start: number; end: number }>
  width: number
  height: number
}
export interface SheetStackedTextLayout { columns: SheetStackedTextColumn[]; width: number; height: number; lineHeight: number }

/** Stacked text uses one grapheme per row and separate columns for explicit breaks or height wrapping. */
export function layoutSheetStackedText(text: string, options: { height: number; lineHeight: number; wrap: boolean; measure: (text: string) => number }): SheetStackedTextLayout {
  const columns: SheetStackedTextColumn[] = [], lineHeight = Math.max(0, options.lineHeight)
  const capacity = options.wrap && Number.isFinite(options.height) && lineHeight > 0 ? Math.max(1, Math.floor(Math.max(0, options.height) / lineHeight)) : Infinity
  for (const paragraph of textParagraphs(text)) {
    const glyphs = sheetTextGraphemes(paragraph.text).map(glyph => ({ text: glyph.text, start: paragraph.start + glyph.offset, end: paragraph.start + glyph.offset + glyph.text.length }))
    if (!glyphs.length) { columns.push({ glyphs: [], width: lineHeight, height: lineHeight }); continue }
    for (let offset = 0; offset < glyphs.length;) {
      const column = glyphs.slice(offset, offset + capacity)
      const width = column.reduce((maximum, glyph) => {
        const measured = options.measure(glyph.text)
        return Math.max(maximum, Number.isFinite(measured) ? Math.max(0, measured) : 0)
      }, 0)
      columns.push({ glyphs: column, width: Math.max(0, width), height: column.length * lineHeight })
      offset += column.length
    }
  }
  return { columns, width: columns.reduce((sum, column) => sum + column.width, 0), height: columns.reduce((maximum, column) => Math.max(maximum, column.height), 0), lineHeight }
}

/** Numeric and Boolean values, merged cells and full-width alignment never use horizontal overflow. */
export function canSheetCellTextOverflow(cell: ExcelCellView | undefined): boolean {
  if (!cell || sheetCellWrapStrategy(cell) !== 'overflow' || sheetCellShrinkToFit(cell) || cell.align === 'justify' || sheetCellTextRotation(cell) === 'v') return false
  const value = sheetCellNumberValue(cell), type = sheetCellStoredValueType(cell)
  return typeof value !== 'number' && typeof value !== 'boolean' && type !== 'number' && type !== 'boolean'
}

export function isSheetCellCoverable(cell: ExcelCellView | undefined): boolean {
  return !cell || !cell.formula && !cell.sparkline && (cell.rawValue === undefined || cell.rawValue === null || cell.rawValue === '') && cell.value === ''
}
