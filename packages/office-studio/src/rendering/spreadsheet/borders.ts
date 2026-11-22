import type { LineCommand, RenderRect } from '../contracts.ts'

function visibleLineInterval(x: number, y: number, dx: number, dy: number, clip: RenderRect, padding: number): { start: number; end: number } | undefined {
  let start = 0, end = 1
  const axes: Array<[number, number, number, number]> = [
    [x, dx, clip.x - padding, clip.x + clip.width + padding],
    [y, dy, clip.y - padding, clip.y + clip.height + padding],
  ]
  for (const [origin, delta, minimum, maximum] of axes) {
    if (delta === 0) { if (origin < minimum || origin > maximum) return undefined; continue }
    const first = (minimum - origin) / delta, last = (maximum - origin) / delta
    start = Math.max(start, Math.min(first, last)); end = Math.min(end, Math.max(first, last))
    if (start > end) return undefined
  }
  return { start, end }
}

export function sheetBorderDashPattern(style: string): number[] {
  if (style.includes('dash-dot-dot')) return [4, 2, 1, 2, 1, 2]
  if (style.includes('dash-dot')) return [4, 2, 1, 2]
  if (style.includes('dotted')) return [1, 2]
  if (style.includes('dashed')) return [4, 2]
  return []
}

/** Both render backends consume the same dash geometry; phase is relative to the original side origin. */
export function createSheetBorderLines(
  x1: number, y1: number, x2: number, y2: number, style: string, color: string, zoom: number, clip: RenderRect,
): LineCommand[] {
  if (![x1, y1, x2, y2, zoom, clip.x, clip.y, clip.width, clip.height].every(Number.isFinite)
    || zoom <= 0 || clip.width <= 0 || clip.height <= 0 || style === 'none') return []
  const width = (style.startsWith('medium') ? 2 : style === 'thick' ? 3 : style === 'hair' ? 0.5 : 1) * zoom
  const pattern = sheetBorderDashPattern(style).map(value => value * zoom)
  const dx = x2 - x1, dy = y2 - y1, length = Math.hypot(dx, dy)
  if (!Number.isFinite(width) || !Number.isFinite(length) || length <= 0 || !pattern.every(Number.isFinite)) return []
  const perpendicularX = -dy / length, perpendicularY = dx / length
  const lines: LineCommand[] = []
  for (const offset of style === 'double' ? [-0.8 * zoom, 0.8 * zoom] : [0]) {
    const startX = x1 + perpendicularX * offset, startY = y1 + perpendicularY * offset
    const interval = visibleLineInterval(startX, startY, dx, dy, clip, width / 2)
    if (!interval) continue
    if (!pattern.length) {
      lines.push({ kind: 'line', x1: startX, y1: startY, x2: x2 + perpendicularX * offset, y2: y2 + perpendicularY * offset, color, width, clip })
      continue
    }
    const minimum = interval.start * length, maximum = interval.end * length
    const cycle = pattern.reduce((sum, value) => sum + value, 0)
    let position = Math.floor(minimum / cycle) * cycle, index = 0
    while (position < maximum) {
      const segment = pattern[index % pattern.length]!
      const first = Math.max(minimum, position), last = Math.min(maximum, position + segment)
      if (index % 2 === 0 && last > first) lines.push({ kind: 'line',
        x1: startX + dx / length * first, y1: startY + dy / length * first,
        x2: startX + dx / length * last, y2: startY + dy / length * last, color, width, clip })
      position += segment; index++
    }
  }
  return lines
}
