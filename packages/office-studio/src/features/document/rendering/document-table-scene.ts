import type { DrawCommand, RenderRect } from '../../../rendering/contracts'
import { intersectRect } from '../../../rendering/geometry'
import type { DocumentTableBorder, DocumentTableBorderSide } from '../model/document-model'
import type { LayoutCell } from '../layout/document-layout-types'

/** Both renderer backends receive the same resolved cell edges and clipping. */
export function appendDocumentTableCell(commands: DrawCommand[], cell: LayoutCell, rectangle: RenderRect, page: RenderRect, scale: number, parentClip?: RenderRect): void {
  const clip = parentClip ? intersectRect(page, parentClip) : page
  if (clip.width <= 0 || clip.height <= 0) return
  commands.push({ kind: 'rect', ...rectangle, fill: cell.fill, clip })
  const fallback: DocumentTableBorder = { width: 1, color: '#aab3bf', style: 'solid' }
  const edge = (side: DocumentTableBorderSide, x1: number, y1: number, x2: number, y2: number) => {
    if (side === 'top' && cell.continuedBefore || side === 'bottom' && cell.continuedAfter) return
    const border = cell.borders?.[side] ?? fallback
    if (border.width <= 0) return
    const width = border.width * scale
    if (border.style === 'solid') { commands.push({ kind: 'line', x1, y1, x2, y2, color: border.color, width, clip }); return }
    const horizontal = Math.abs(y2 - y1) < 1e-6, start = horizontal ? x1 : y1, end = horizontal ? x2 : y2
    const length = Math.max(0, end - start), segment = border.style === 'dot' ? Math.max(1, width) : Math.max(4, width * 3)
    const gap = Math.max(2, width * 2), period = segment + gap
    const visibleStart = Math.max(start, horizontal ? clip.x : clip.y), visibleEnd = Math.min(end, horizontal ? clip.x + clip.width : clip.y + clip.height)
    const first = Math.max(0, Math.floor((visibleStart - start) / period))
    for (let index = first; index * period < length && start + index * period < visibleEnd; index++) {
      const from = start + index * period, to = Math.min(end, from + segment)
      commands.push({ kind: 'line', x1: horizontal ? from : x1, y1: horizontal ? y1 : from,
        x2: horizontal ? to : x2, y2: horizontal ? y2 : to, color: border.color, width, clip })
    }
  }
  edge('top', rectangle.x, rectangle.y, rectangle.x + rectangle.width, rectangle.y)
  edge('right', rectangle.x + rectangle.width, rectangle.y, rectangle.x + rectangle.width, rectangle.y + rectangle.height)
  edge('bottom', rectangle.x, rectangle.y + rectangle.height, rectangle.x + rectangle.width, rectangle.y + rectangle.height)
  edge('left', rectangle.x, rectangle.y, rectangle.x, rectangle.y + rectangle.height)
}
