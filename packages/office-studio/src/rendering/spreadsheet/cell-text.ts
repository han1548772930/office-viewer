import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import type { DrawCommand, RenderRect } from '../contracts.ts'
import { sheetCellTextAlignment } from '../../features/sheet-workbench/sheet-cell-alignment.ts'
import { sheetCellShrinkToFit, sheetCellTextRotation, sheetCellWrapStrategy } from '../../features/sheet-workbench/sheet-cell-layout.ts'
import { layoutSheetCellText, layoutSheetStackedText, rotatedSheetTextSize, sheetCellTextFont, sheetCellTextIndent, sheetTextGraphemes,
  type SheetCellTextFontOverrides, type SheetTextMeasure } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'
import { positionSheetNumberFormatRuns, type SheetNumberPresentation } from '../../features/sheet-workbench/number-format-render.ts'
import { measureSpreadsheetText } from './text-metrics.ts'
import { layoutSheetCellRichText, type SheetRichTextFragment } from '../../features/sheet-workbench/sheet-cell-rich-text-layout.ts'

export interface SheetCellTextPaintOptions {
  cell?: ExcelCellView
  presentation: SheetNumberPresentation
  rect: RenderRect
  clip: RenderRect
  zoom: number
  iconPadding?: number
  color?: string
  underline?: boolean
  strike?: boolean
  fontOverrides?: SheetCellTextFontOverrides
  measure?: SheetTextMeasure
}

/** Layout creates backend-neutral commands, including complete text blocks and uniformly scaled fonts. */
export function createSheetCellTextCommands(options: SheetCellTextPaintOptions): DrawCommand[] {
  const { cell, presentation, rect, clip, zoom: z } = options
  if (clip.width <= 0 || clip.height <= 0 || rect.width <= 0 || rect.height <= 0) return []
  const text = presentation.text
  if (!text && !presentation.runs?.some(run => run.kind === 'fill' && run.character.trim())) return []
  const align = sheetCellTextAlignment(cell), direction = sheetCellTextRotation(cell), rotation = direction === 'v' ? 0 : direction
  const measure = options.measure ?? measureSpreadsheetText
  const base = sheetCellTextFont(cell, z, options.fontOverrides)
  const indent = align === 'center' ? 0 : sheetCellTextIndent(cell, base.fontSize)
  const left = rect.x + 6 * z + (options.iconPadding ?? 0) + (align === 'left' ? indent : 0)
  const right = rect.x + rect.width - 6 * z - (align === 'right' ? indent : 0)
  const availableWidth = Math.max(0, right - left), availableHeight = Math.max(0, rect.height - 6 * z)
  const wrap = sheetCellWrapStrategy(cell) === 'wrap'
  const hasRuns = direction !== 'v' && !/[\r\n]/u.test(text) && Boolean(presentation.runs)
  const layoutFor = (font: string, fontSize: number, lineHeight: number) => {
    const widthOf = (value: string) => measure(value, font)
    const rich = layoutSheetCellRichText(cell, text, { width: availableWidth, height: availableHeight, wrap, stacked: direction === 'v',
      zoom: z * fontSize / base.fontSize, fontOverrides: options.fontOverrides, measure })
    const horizontal = rich ? { lines: rich.lines, width: rich.width, height: rich.height, lineHeight }
      : layoutSheetCellText(text, { width: availableWidth, lineHeight, wrap, measure: widthOf })
    const stacked = !rich && direction === 'v' ? layoutSheetStackedText(text, { height: availableHeight, lineHeight: fontSize, wrap, measure: widthOf }) : undefined
    const runs = !rich && hasRuns ? positionSheetNumberFormatRuns(presentation.runs!, availableWidth, widthOf) : undefined
    const justified = cell?.align === 'justify' && horizontal.lines.some(line => !line.paragraphEnd)
    const width = stacked?.width ?? runs?.width ?? (justified && direction !== 'v' ? Math.max(availableWidth, horizontal.width) : horizontal.width)
    const height = rich?.height ?? stacked?.height ?? (runs ? lineHeight : horizontal.height)
    return { horizontal, stacked, rich, runs, width, height, size: rotatedSheetTextSize(width, height, rotation) }
  }
  let font = base.font, fontSize = base.fontSize, lineHeight = base.lineHeight
  let layout = layoutFor(font, fontSize, lineHeight)
  if (sheetCellShrinkToFit(cell) && availableWidth > 0) {
    const widthScale = layout.size.width > 0 ? availableWidth / layout.size.width : 1
    const heightScale = (rotation !== 0 || direction === 'v' || layout.horizontal.lines.length > 1) && layout.size.height > 0
      ? availableHeight / layout.size.height : 1
    const scale = Math.max(0, Math.min(1, widthScale, heightScale))
    if (scale === 0) return []
    if (scale < 1) {
      fontSize *= scale; lineHeight *= scale
      font = base.font.replace(/([\d.]+)px/u, `${fontSize}px`)
      layout = layoutFor(font, fontSize, lineHeight)
    }
  }
  const color = options.color ?? presentation.color ?? (cell?.hyperlink ? '#0563c1' : cell?.color) ?? '#243247'
  const underline = options.underline ?? (cell?.hyperlink ? true : cell?.underline), strike = options.strike ?? cell?.strike
  const radians = rotation * Math.PI / 180, cosine = Math.cos(radians), sine = Math.sin(radians)
  const rotate = (x: number, y: number): [number, number] => [x * cosine - y * sine, x * sine + y * cosine]
  const corners = [rotate(0, 0), rotate(layout.width, 0), rotate(0, layout.height), rotate(layout.width, layout.height)]
  const minimumX = Math.min(...corners.map(point => point[0])), minimumY = Math.min(...corners.map(point => point[1]))
  const blockLeft = align === 'center' ? (left + right - layout.size.width) / 2 : align === 'right' ? right - layout.size.width : left
  const blockTop = cell?.vertical === 'top' ? rect.y + 3 * z
    : cell?.vertical === 'bottom' ? rect.y + rect.height - 3 * z - layout.size.height : rect.y + (rect.height - layout.size.height) / 2
  const position = (x: number, y: number): [number, number] => {
    const point = rotate(x, y)
    return [point[0] + blockLeft - minimumX, point[1] + blockTop - minimumY]
  }
  const commands: DrawCommand[] = []
  const paint = (value: string, x: number, y: number, width: number, alignment: 'left' | 'center' | 'right' = 'left', fragment?: SheetRichTextFragment) => {
    if (!value) return
    const fragmentColor = options.color ?? presentation.color ?? (fragment?.hyperlink ? '#0563c1' : fragment?.color) ?? color
    const fragmentUnderline = options.underline ?? (fragment?.hyperlink ? true : fragment?.underline) ?? underline
    const fragmentStrike = options.strike ?? fragment?.strike ?? strike
    const [textX, textY] = position(x, y)
    commands.push({ kind: 'text', text: value, x: textX, y: textY, font: fragment?.font ?? font, color: fragmentColor, align: alignment, baseline: 'middle', rotation, clip })
    const localLeft = alignment === 'center' ? x - width / 2 : alignment === 'right' ? x - width : x
    const decoration = (offset: number) => {
      const [x1, y1] = position(localLeft, y + offset), [x2, y2] = position(localLeft + width, y + offset)
      commands.push({ kind: 'line', x1, y1, x2, y2, color: fragmentColor, clip })
    }
    if (fragmentUnderline) decoration((fragment?.fontSize ?? fontSize) * 0.55)
    if (fragmentStrike) decoration(0)
  }
  if (layout.rich?.columns) {
    let columnLeft = 0
    for (const column of layout.rich.columns) {
      const top = cell?.vertical === 'top' ? 0 : cell?.vertical === 'bottom' ? layout.height - column.height : (layout.height - column.height) / 2
      for (const glyph of column.glyphs) paint(glyph.text, columnLeft + column.width / 2, top + glyph.top + glyph.fontSize / 2, glyph.width, 'center', glyph)
      columnLeft += column.width
    }
  } else if (layout.rich) {
    for (const line of layout.rich.lines) {
      const words = line.text.match(/\s+|\S+\s*/gu) ?? []
      const cjk = words.length <= 1 && /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(line.text)
      const pieces = cjk ? sheetTextGraphemes(line.text).map(glyph => glyph.text) : words
      const justify = cell?.align === 'justify' && !line.paragraphEnd && line.width < availableWidth && pieces.length > 1
      const ranges = justify ? pieces : [line.text]
      const extra = justify ? (availableWidth - line.width) / (pieces.length - 1) : 0
      let offset = justify ? 0 : align === 'center' ? (layout.width - line.width) / 2 : align === 'right' ? layout.width - line.width : 0
      let start = line.start
      for (const piece of ranges) {
        const end = start + piece.length
        for (const fragment of line.fragments) {
          const first = Math.max(start, fragment.start), last = Math.min(end, fragment.end)
          if (first >= last) continue
          const value = text.slice(first, last), width = measure(value, fragment.font)
          const y = line.top + line.lineHeight / 2 + (line.fontSize - fragment.fontSize) * .3
          paint(value, offset, y, width, 'left', fragment); offset += width
        }
        offset += extra; start = end
      }
    }
  } else if (layout.stacked) {
    let columnLeft = 0
    for (const column of layout.stacked.columns) {
      const top = cell?.vertical === 'top' ? 0 : cell?.vertical === 'bottom' ? layout.height - column.height : (layout.height - column.height) / 2
      column.glyphs.forEach((glyph, index) => paint(glyph.text, columnLeft + column.width / 2, top + (index + .5) * fontSize, measure(glyph.text, font), 'center'))
      columnLeft += column.width
    }
  } else if (layout.runs) {
    for (const run of layout.runs.runs) paint(run.text, run.offset, lineHeight / 2, measure(run.text, font))
  } else {
    const anchor = align === 'center' ? layout.width / 2 : align === 'right' ? layout.width : 0
    for (const [index, line] of layout.horizontal.lines.entries()) {
      const y = (index + .5) * lineHeight
      if (cell?.align === 'justify' && !line.paragraphEnd && line.width < availableWidth) {
        const words = line.text.match(/\s+|\S+\s*/gu) ?? []
        const cjk = words.length <= 1 && /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(line.text)
        const pieces = cjk ? sheetTextGraphemes(line.text).map(glyph => glyph.text) : words
        if (pieces.length > 1) {
          const extra = (availableWidth - line.width) / (pieces.length - 1)
          let offset = 0
          for (const piece of pieces) {
            const width = measure(piece, font)
            paint(piece, offset, y, width); offset += width + extra
          }
          continue
        }
      }
      paint(line.text, anchor, y, line.width, align)
    }
  }
  return commands
}
