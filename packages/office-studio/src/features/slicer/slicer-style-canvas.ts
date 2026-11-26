import type { SlicerStyleBorderSide, SlicerStyleTextDecoration } from '../../domain/slicer-styles.ts'
import type { DrawCommand, LineCommand, RenderRect } from '../../rendering/contracts.ts'
import { createSheetBorderLines } from '../../rendering/spreadsheet/borders.ts'
import { measureSpreadsheetText } from '../../rendering/spreadsheet/text-metrics.ts'
import { sheetTextGraphemes } from '../sheet-workbench/sheet-cell-text-layout.ts'
import type { SlicerRenderRegion } from './slicer-style-rendering.ts'

/** Shared border segments keep Canvas2D, WebGPU and the Ant control's passive decoration identical. */
export function slicerRegionBorderCommands(rect: RenderRect, style: SlicerRenderRegion, diagonalsOnly = false): LineCommand[] {
  if (rect.width <= 0 || rect.height <= 0) return []
  const { x, y, width: w, height: h } = rect, sides: Record<SlicerStyleBorderSide, [number, number, number, number]> = {
    t: [x, y, x + w, y], r: [x + w, y, x + w, y + h], b: [x, y + h, x + w, y + h], l: [x, y, x, y + h],
    tl_br: [x, y, x + w, y + h], tl_bc: [x, y, x + w / 2, y + h], tl_mr: [x, y, x + w, y + h / 2],
    bl_tr: [x, y + h, x + w, y], ml_tr: [x, y + h / 2, x + w, y], bc_tr: [x + w / 2, y + h, x + w, y],
  }
  return (Object.keys(sides) as SlicerStyleBorderSide[]).flatMap(side => {
    const border = style.borders[side]
    if (!border || diagonalsOnly && ['t', 'r', 'b', 'l'].includes(side)) return []
    return createSheetBorderLines(...sides[side], border.style, border.color, 1, rect)
  })
}
function overlineCommands(x: number, y: number, width: number, decoration: SlicerStyleTextDecoration & { color: string }, clip: RenderRect): LineCommand[] {
  if (decoration.s !== 1 || decoration.t === 11 || width <= 0) return []
  const type = decoration.t ?? 12, heavy = [1, 2, 3, 5, 9, 13, 16].includes(type), thickness = heavy ? 2 : 1
  const line = (first: number, last: number, top: number, bottom = top): LineCommand => ({ kind: 'line', x1: first, y1: top, x2: last, y2: bottom,
    color: decoration.color, width: thickness, clip })
  if ([14, 15, 16].includes(type)) {
    const lines: LineCommand[] = []
    for (const offset of type === 15 ? [-1.5, 1.5] : [0]) for (let step = 0; step < width; step += 2) {
      const last = Math.min(width, step + 2)
      lines.push(line(x + step, x + last, y + offset + Math.sin(step * Math.PI / 4), y + offset + Math.sin(last * Math.PI / 4)))
    }
    return lines
  }
  if ([10, 19].includes(type)) return [line(x, x + width, y - 1), line(x, x + width, y + 1)]
  const patterns: Partial<Record<number, number[]>> = { 0: [4, 2], 1: [4, 2, 1, 2, 1, 2], 2: [4, 2, 1, 2], 3: [4, 2],
    4: [8, 3], 5: [8, 3], 6: [1, 2, 4, 2], 7: [1, 2, 1, 2, 4, 2], 8: [1, 2], 9: [1, 2] }
  const pattern = patterns[type]
  if (!pattern) return [line(x, x + width, y)]
  const lines: LineCommand[] = []
  for (let offset = 0, index = 0; offset < width; index++) {
    const length = pattern[index % pattern.length]!, last = Math.min(width, offset + length)
    if (index % 2 === 0) lines.push(line(x + offset, x + last, y))
    offset += length
  }
  return lines
}
/** Measured labels and decoration positions use the same CSS-pixel contract for both Canvas backends. */
export function slicerRegionTextCommands(rect: RenderRect, text: string, style: SlicerRenderRegion, fontSize = 14, padding = 8): DrawCommand[] {
  const available = Math.max(0, rect.width - padding * 2), font = `${style.bold ? 700 : 400} ${fontSize}px Arial, sans-serif`
  if (!text || available <= 0 || rect.height <= 0) return []
  const measured = measureSpreadsheetText(text, font), width = Math.min(measured, available), left = rect.x + padding
  const y = style.vertical === 1 ? rect.y + fontSize / 2 + 2 : style.vertical === 3 ? rect.y + rect.height - fontSize / 2 - 2 : rect.y + rect.height / 2
  const align = style.horizontal === 3 ? 'right' : style.horizontal === 2 ? 'center' : 'left', x = align === 'right' ? left + available : align === 'center' ? left + available / 2 : left
  const glyphs = sheetTextGraphemes(text), widths = glyphs.map(glyph => measureSpreadsheetText(glyph.text, font))
  const sum = widths.reduce((total, value) => total + value, 0), gaps = style.horizontal === 4
    ? glyphs.slice(0, -1).filter(glyph => /^\s+$/u.test(glyph.text)).length : glyphs.length - 1
  const distributed = style.horizontal >= 4 && sum < available && gaps > 0
  const commands: DrawCommand[] = [], overline = style.overline, top = y - fontSize / 2 - 1
  let cursor = distributed ? left : align === 'right' ? x - width : align === 'center' ? x - width / 2 : x
  if (distributed) for (let index = 0; index < glyphs.length; index++) {
    const glyph = glyphs[index]!, glyphWidth = widths[index]!
    commands.push({ kind: 'text', x: cursor, y, text: glyph.text, font, color: style.color, baseline: 'middle', maxWidth: glyphWidth, clip: rect })
    if (overline && (overline.t !== 17 || !/^\s+$/u.test(glyph.text))) commands.push(...overlineCommands(cursor, top, glyphWidth, overline, rect))
    cursor += glyphWidth + (index < glyphs.length - 1 && (style.horizontal !== 4 || /^\s+$/u.test(glyph.text)) ? (available - sum) / gaps : 0)
  } else {
    commands.push({ kind: 'text', x, y, text, font, color: style.color, align, baseline: 'middle', maxWidth: available, clip: rect })
    if (overline) {
      if (overline.t !== 17) commands.push(...overlineCommands(cursor, top, width, overline, rect))
      else for (let index = 0; index < glyphs.length; index++) {
        const glyph = glyphs[index]!, glyphWidth = widths[index]! * (measured > available ? available / measured : 1)
        if (!/^\s+$/u.test(glyph.text)) commands.push(...overlineCommands(cursor, top, glyphWidth, overline, rect))
        cursor += glyphWidth
      }
    }
  }
  return commands
}

/** Diagonal borders are passive graphics inside an Ant Button; interactive state stays in the library control. */
export function slicerRegionDecorationBackground(style: SlicerRenderRegion, width: number, height: number, text = '', fontSize = 14, padding = 8): string | undefined {
  const rect = { x: 0, y: 0, width, height }, commands = [...slicerRegionBorderCommands(rect, style, true),
    ...slicerRegionTextCommands(rect, text, style, fontSize, padding).filter((command): command is LineCommand => command.kind === 'line')]
  if (!commands.length) return undefined
  const escape = (value: string) => value.replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[character]!)
  const lines = commands.map(line => `<path d="M${line.x1} ${line.y1} L${line.x2} ${line.y2}" stroke="${escape(line.color)}" stroke-width="${line.width ?? 1}"/>`).join('')
  return `url("data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${lines}</svg>`)}")`
}
