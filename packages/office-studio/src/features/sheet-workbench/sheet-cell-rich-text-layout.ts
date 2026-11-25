import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { sheetCellRichTextMatches } from './sheet-cell-rich-text-edit.ts'
import { sheetCellRichTextRunView } from './sheet-cell-rich-text-style.ts'
import { layoutSheetCellText, layoutSheetStackedText, sheetCellTextFont, sheetTextGraphemes,
  type SheetCellTextFont, type SheetCellTextFontOverrides, type SheetTextLine, type SheetTextMeasure } from './sheet-cell-text-layout.ts'

export interface SheetRichTextFragment extends SheetCellTextFont {
  text: string
  start: number
  end: number
  width: number
  color?: string | null
  underline?: boolean
  strike?: boolean
  hyperlink?: string | null
}
export interface SheetRichTextLine extends SheetTextLine {
  fragments: SheetRichTextFragment[]
  top: number
  lineHeight: number
  fontSize: number
}
export interface SheetRichTextColumn {
  glyphs: Array<SheetRichTextFragment & { top: number }>
  width: number
  height: number
}
export interface SheetCellRichTextLayout {
  lines: SheetRichTextLine[]
  columns?: SheetRichTextColumn[]
  width: number
  height: number
}
export interface SheetCellRichTextLayoutOptions {
  width: number
  height: number
  wrap: boolean
  stacked?: boolean
  zoom?: number
  fontOverrides?: SheetCellTextFontOverrides
  measure: SheetTextMeasure
}

/** Reuse the same source-aware grapheme/paragraph breaker for paint and axis measurements. */
export function layoutSheetCellRichText(cell: ExcelCellView | undefined, text: string,
  options: SheetCellRichTextLayoutOptions): SheetCellRichTextLayout | undefined {
  const runs = cell?.richText
  if (!cell || cell.formula || typeof cell.rawValue === 'number' || typeof cell.rawValue === 'boolean'
    || !runs?.length || !sheetCellRichTextMatches(runs, text)) return undefined
  const base = sheetCellTextFont(cell, options.zoom ?? 1, options.fontOverrides)
  let offset = 0
  const spans = runs.flatMap(run => {
    const start = offset; offset += (run.text as string).length
    if (offset === start) return []
    const view = sheetCellRichTextRunView(cell, run)
    return [{ start, end: offset, view, metrics: sheetCellTextFont(view, options.zoom ?? 1, options.fontOverrides) }]
  })
  // Imported boundaries inside a grapheme use the leading character's style instead of breaking its shape.
  let spanIndex = 0
  const segments: Array<{ start: number; end: number; view: ExcelCellView; metrics: SheetCellTextFont }> = []
  for (const glyph of sheetTextGraphemes(text)) {
    while (spanIndex < spans.length - 1 && spans[spanIndex]!.end <= glyph.offset) spanIndex++
    const span = spans[spanIndex]
    if (!span) continue
    const previous = segments.at(-1), end = glyph.offset + glyph.text.length
    if (previous?.view === span.view && previous.end === glyph.offset) previous.end = end
    else segments.push({ start: glyph.offset, end, view: span.view, metrics: span.metrics })
  }
  const finiteWidth = (value: string, font: string) => {
    const measured = options.measure(value, font)
    return Number.isFinite(measured) ? Math.max(0, measured) : 0
  }
  const firstSegment = (offset: number) => {
    let first = 0, last = segments.length
    while (first < last) {
      const middle = (first + last) >>> 1
      if (segments[middle]!.end <= offset) first = middle + 1
      else last = middle
    }
    return first
  }
  const fragments = (start: number, end: number): SheetRichTextFragment[] => {
    const result: SheetRichTextFragment[] = []
    for (let index = firstSegment(start); index < segments.length; index++) {
      const segment = segments[index]!
      if (segment.start >= end) break
      const first = Math.max(start, segment.start), last = Math.min(end, segment.end)
      const value = text.slice(first, last)
      result.push({ text: value, start: first, end: last, width: finiteWidth(value, segment.metrics.font), ...segment.metrics,
        color: segment.view.color, underline: segment.view.underline, strike: segment.view.strike, hyperlink: segment.view.hyperlink })
    }
    return result
  }
  const rangeWidth = (_value: string, start: number, end: number) => fragments(start, end).reduce((sum, fragment) => sum + fragment.width, 0)
  const at = (start: number) => segments[firstSegment(start)]?.metrics
    ?? (start === text.length ? segments.at(-1)?.metrics : undefined) ?? base
  const plain = layoutSheetCellText(text, { width: options.width, lineHeight: base.lineHeight, wrap: options.wrap,
    measure: value => finiteWidth(value, base.font), measureRange: rangeWidth })
  let top = 0
  const lines = plain.lines.map(line => {
    const pieces = fragments(line.start, line.end), fallback = at(line.start)
    const lineHeight = pieces.length ? pieces.reduce((maximum, piece) => Math.max(maximum, piece.lineHeight), 0) : fallback.lineHeight
    const fontSize = pieces.length ? pieces.reduce((maximum, piece) => Math.max(maximum, piece.fontSize), 0) : fallback.fontSize
    const result: SheetRichTextLine = { ...line, fragments: pieces, lineHeight, fontSize, top }
    top += lineHeight
    return result
  })
  if (!options.stacked) return { lines, width: plain.width, height: top }
  const stacked = layoutSheetStackedText(text, { height: options.height, lineHeight: base.fontSize, wrap: options.wrap,
    measure: value => finiteWidth(value, base.font), measureRange: rangeWidth,
    glyphHeight: (start, end) => fragments(start, end).reduce((maximum, fragment) => Math.max(maximum, fragment.fontSize), at(start).fontSize) })
  const columns: SheetRichTextColumn[] = stacked.columns.map(column => {
    let glyphTop = 0
    const glyphs = column.glyphs.flatMap(glyph => {
      const piece = fragments(glyph.start, glyph.end)[0]
      if (!piece) return []
      const result = { ...piece, top: glyphTop }; glyphTop += piece.fontSize
      return [result]
    })
    return { glyphs, width: column.width, height: column.height }
  })
  return { lines, columns, width: stacked.width, height: stacked.height }
}
