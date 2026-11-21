import type { ExcelCellView } from '../../services/excelViewAdapter'
import type { DrawCommand, RenderRect, RenderScene } from '../contracts'
import { intersectRect } from '../geometry.ts'
import { cellAddress, columnLabel, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from './layout.ts'
import type { GridRange, GridViewport } from './layout'
import { formatSheetValue, SHEET_NUMBER_FORMATS } from '../../features/sheet-workbench/sheet-number-format.ts'
import { sparklineCommandsForCell } from '../../features/sparkline/sparkline-render.ts'

let textMeasureContext: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null | undefined
function measureTextWidth(text: string, font: string): number | undefined {
  if (textMeasureContext === undefined) textMeasureContext = typeof OffscreenCanvas !== 'undefined'
    ? new OffscreenCanvas(1, 1).getContext('2d') : typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  if (!textMeasureContext) return undefined
  textMeasureContext.font = font
  return textMeasureContext.measureText(text).width
}
function decorationWidth(text: string, font: string, available: number): number {
  return Math.min(available, measureTextWidth(text, font) ?? available)
}
function rotatePoint(x: number, y: number, originX: number, originY: number, degrees: number): [number, number] {
  const radians = degrees * Math.PI / 180
  const cosine = Math.cos(radians), sine = Math.sin(radians)
  return [originX + (x - originX) * cosine - (y - originY) * sine,
    originY + (x - originX) * sine + (y - originY) * cosine]
}
function graphemes(text: string): string[] {
  if (typeof Intl.Segmenter === 'function') {
    return [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].map(item => item.segment)
  }
  return Array.from(text)
}

export interface SpreadsheetSceneOptions {
  activeCell: string
  selectedCells: readonly string[]
  selectedRanges?: readonly GridRange[]
  searchText?: string
  /** Drawing-object commands are expressed in unscaled worksheet CSS pixels. */
  overlayCommands?: readonly DrawCommand[]
}
export function displayCell(cell: ExcelCellView | undefined, date1904 = false): string {
  if (!cell) return ''
  const value = cell.formula ? cell.cachedValue ?? cell.formula : cell.value
  const scalar = cell.rawValue !== undefined ? cell.rawValue : value !== '' && /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(value) ? Number(value) : value
  const pattern = cell.rawNumberFormat ?? (cell.numberFormat ? SHEET_NUMBER_FORMATS[cell.numberFormat] : 'General')
  return formatSheetValue(scalar, pattern, date1904) ?? cell.displayValue ?? value
}

/** Layout and hit-testing share the same sparse axes and frozen-pane transforms. */
export function createSpreadsheetScene(geometry: SpreadsheetGeometry, viewport: GridViewport, options: SpreadsheetSceneOptions): RenderScene {
  const commands: DrawCommand[] = [], selected = new Set(options.selectedCells), search = options.searchText?.toLocaleLowerCase()
  const panes = geometry.panes(viewport), z = viewport.zoom
  const isSelected = (row: number, column: number) => selected.has(cellAddress(row, column)) || options.selectedRanges?.some(range => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn)
  const paintCell = (row: number, column: number, clip: RenderRect) => {
    const key = cellAddress(row, column), cell = geometry.sheet.cells[key], rect = geometry.cellRect(row, column, viewport)
    const visible = intersectRect(rect, clip)
    if (visible.width <= 0 || visible.height <= 0) return
    commands.push({ kind: 'rect', ...rect, fill: cell?.background ?? '#ffffff', clip })
    if (geometry.sheet.showGridlines !== false) commands.push({ kind: 'rect', ...rect, fill: 'transparent', stroke: '#e5e9f0', strokeWidth: 1, clip })
    if (isSelected(row, column)) commands.push({ kind: 'rect', ...rect, fill: 'rgba(36,107,253,0.09)', clip })
    if (cell?.border) for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const raw = cell.border[side]
      const edge = raw && typeof raw === 'object' ? raw as Record<string, unknown> : undefined
      const style = typeof raw === 'string' ? raw.split(':')[0] : edge?.style
      if (!raw || style === 'none' || style === '') continue
      const color = typeof edge?.color === 'string' ? edge.color : typeof raw === 'string' ? raw.split(':')[1] ?? '#1f2937' : '#1f2937'
      const width = style === 'thick' ? 3 : style === 'medium' ? 2 : 1
      const horizontal = side === 'top' || side === 'bottom'
      const borderX = side === 'right' ? rect.x + rect.width : rect.x
      const borderY = side === 'bottom' ? rect.y + rect.height : rect.y
      commands.push({ kind: 'line', x1: borderX, y1: borderY, x2: horizontal ? rect.x + rect.width : borderX,
        y2: horizontal ? borderY : rect.y + rect.height, color, width: width * z, clip })
    }
    const text = displayCell(cell, geometry.sheet.date1904)
    if (search && text.toLocaleLowerCase().includes(search)) commands.push({ kind: 'rect', ...rect, fill: 'rgba(250,204,21,0.28)', clip })
    if (text) {
      const align = cell?.align ?? (column === 0 ? 'left' : Number.isFinite(Number(text)) ? 'right' : 'left')
      const baseFontSize = (cell?.fontSize ?? 11) * 4 / 3 * z
      const fontFamily = cell?.fontFamily ?? 'Aptos, "Microsoft YaHei", sans-serif'
      const fontPrefix = `${cell?.italic ? 'italic ' : ''}${cell?.bold ? '600 ' : ''}`
      const font = `${fontPrefix}${baseFontSize}px ${fontFamily}`
      const availableTextWidth = Math.max(0, rect.width - 12 * z)
      const measuredWidth = cell?.shrinkToFit && cell.wrap !== true ? measureTextWidth(text, font) : undefined
      const maxWidth = measuredWidth !== undefined && measuredWidth > availableTextWidth && availableTextWidth > 0
        ? availableTextWidth
        : undefined
      const fontSize = baseFontSize
      const x = align === 'center' ? rect.x + rect.width / 2 : align === 'right' ? rect.x + rect.width - 6 * z : rect.x + 6 * z
      const rotation = cell?.rotation ?? 0
      const textClip = intersectRect(clip, { x: rect.x + 3 * z, y: rect.y + 1, width: Math.max(0, rect.width - 6 * z), height: Math.max(0, rect.height - 2) })
      const y = cell?.vertical === 'top' ? rect.y + 3 * z + fontSize / 2 : cell?.vertical === 'bottom' ? rect.y + rect.height - 3 * z - fontSize / 2 : rect.y + rect.height / 2
      const color = cell?.color ?? '#243247'
      if (cell?.rotationVertical) {
        const glyphs = graphemes(text)
        const lineHeight = fontSize
        const blockHeight = glyphs.length * lineHeight
        const top = cell.vertical === 'top' ? rect.y + 3 * z
          : cell.vertical === 'bottom' ? rect.y + rect.height - 3 * z - blockHeight
            : rect.y + (rect.height - blockHeight) / 2
        const centerX = rect.x + rect.width / 2
        for (const [index, glyph] of glyphs.entries()) {
          const glyphY = top + index * lineHeight + lineHeight / 2
          commands.push({ kind: 'text', x: centerX, y: glyphY, text: glyph, font, color, align: 'center', baseline: 'middle', clip: textClip })
          if (cell.underline || cell.strike) {
            const width = decorationWidth(glyph, font, availableTextWidth)
            const left = centerX - width / 2
            if (cell.underline) commands.push({ kind: 'line', x1: left, y1: glyphY + lineHeight * 0.55, x2: left + width, y2: glyphY + lineHeight * 0.55, color, clip: textClip })
            if (cell.strike) commands.push({ kind: 'line', x1: left, y1: glyphY, x2: left + width, y2: glyphY, color, clip: textClip })
          }
        }
      } else {
        if (cell?.numberFormat === 'accounting' && /^[$¥￥€£] /.test(text)) {
          commands.push({ kind: 'text', x: rect.x + 6 * z, y, text: text[0]!, font, color, align: 'left', baseline: 'middle', rotation, maxWidth, clip: textClip })
          commands.push({ kind: 'text', x: rect.x + rect.width - 6 * z, y, text: text.slice(2), font, color, align: 'right', baseline: 'middle', rotation, maxWidth, clip: textClip })
        } else commands.push({ kind: 'text', x, y, text, font, color, align, baseline: 'middle', rotation, maxWidth, clip: textClip })
      }
      if (!cell?.rotationVertical && (cell?.underline || cell?.strike)) {
        const width = decorationWidth(text, font, Math.max(0, rect.width - 12 * z))
        const left = align === 'center' ? x - width / 2 : align === 'right' ? x - width : x
        const decoration = (offsetY: number) => {
          const [x1, y1] = rotatePoint(left, y + offsetY, x, y, rotation)
          const [x2, y2] = rotatePoint(left + width, y + offsetY, x, y, rotation)
          commands.push({ kind: 'line', x1, y1, x2, y2, color, clip: textClip })
        }
        if (cell.underline) decoration(fontSize * 0.55)
        if (cell.strike) decoration(0)
      }
    }
    commands.push(...sparklineCommandsForCell(geometry.sheet, cell, rect))
  }
  for (const pane of panes) {
    if (pane.clip.width <= 0 || pane.clip.height <= 0) continue
    const anchors = new Set<string>()
    for (const row of pane.rows) for (const column of pane.columns) {
      const merge = geometry.mergeAt(row, column)
      if (merge) anchors.add(cellAddress(merge.startRow, merge.startColumn))
      else paintCell(row, column, pane.clip)
    }
    for (const address of anchors) { const cell = parseAddress(address)!; paintCell(cell.row, cell.column, pane.clip) }
    const active = parseAddress(options.activeCell)
    if (active) {
      const rect = geometry.cellRect(active.row, active.column, viewport)
      const intersection = intersectRect(rect, pane.clip)
      if (intersection.width > 0 && intersection.height > 0) {
        commands.push({ kind: 'rect', ...rect, fill: 'transparent', stroke: '#246bfd', strokeWidth: 2, clip: pane.clip })
        commands.push({ kind: 'rect', x: rect.x + rect.width - 3, y: rect.y + rect.height - 3, width: 6, height: 6, fill: '#246bfd', stroke: '#ffffff', clip: pane.clip })
      }
    }
  }
  // Overlay drawing objects share the same scroll/zoom transform as the sheet;
  // headers are painted afterwards and therefore stay above floating objects.
  for (const command of options.overlayCommands ?? []) commands.push(transformOverlayCommand(command, viewport))
  commands.push({ kind: 'rect', x: 0, y: 0, width: viewport.width, height: GRID_HEADER_HEIGHT, fill: '#f5f7fa' })
  commands.push({ kind: 'rect', x: 0, y: GRID_HEADER_HEIGHT, width: GRID_HEADER_WIDTH, height: Math.max(0, viewport.height - GRID_HEADER_HEIGHT), fill: '#f5f7fa' })
  const active = parseAddress(options.activeCell)
  for (const pane of panes) {
    for (const column of pane.columns) {
      const rect = geometry.cellRect(0, column, viewport, false), clip = { x: pane.clip.x, y: 0, width: pane.clip.width, height: GRID_HEADER_HEIGHT }
      commands.push({ kind: 'rect', x: rect.x, y: 0, width: rect.width, height: GRID_HEADER_HEIGHT, fill: active?.column === column ? '#e8f0ff' : '#f5f7fa', stroke: '#dfe5ee', clip })
      commands.push({ kind: 'text', x: rect.x + rect.width / 2, y: GRID_HEADER_HEIGHT / 2, text: columnLabel(column), font: '12px "Segoe UI", sans-serif', color: active?.column === column ? '#246bfd' : '#69778c', align: 'center', baseline: 'middle', clip })
    }
    for (const row of pane.rows) {
      const rect = geometry.cellRect(row, 0, viewport, false), clip = { x: 0, y: pane.clip.y, width: GRID_HEADER_WIDTH, height: pane.clip.height }
      commands.push({ kind: 'rect', x: 0, y: rect.y, width: GRID_HEADER_WIDTH, height: rect.height, fill: active?.row === row ? '#e8f0ff' : '#f5f7fa', stroke: '#dfe5ee', clip })
      commands.push({ kind: 'text', x: GRID_HEADER_WIDTH / 2, y: rect.y + rect.height / 2, text: String(row + 1), font: '12px "Segoe UI", sans-serif', color: active?.row === row ? '#246bfd' : '#69778c', align: 'center', baseline: 'middle', clip })
    }
  }
  commands.push({ kind: 'rect', x: 0, y: 0, width: GRID_HEADER_WIDTH, height: GRID_HEADER_HEIGHT, fill: '#eef2f7', stroke: '#dfe5ee' })
  commands.push({ kind: 'line', x1: GRID_HEADER_WIDTH - 13, y1: GRID_HEADER_HEIGHT - 5, x2: GRID_HEADER_WIDTH - 5, y2: GRID_HEADER_HEIGHT - 13, color: '#9aa8b9', width: 3 })
  if (geometry.frozenRows) { const y = GRID_HEADER_HEIGHT + geometry.rows.offset(geometry.frozenRows) * z; commands.push({ kind: 'line', x1: GRID_HEADER_WIDTH, y1: y, x2: viewport.width, y2: y, color: '#9eacbf', width: 2 }) }
  if (geometry.frozenColumns) { const x = GRID_HEADER_WIDTH + geometry.columns.offset(geometry.frozenColumns) * z; commands.push({ kind: 'line', x1: x, y1: GRID_HEADER_HEIGHT, x2: x, y2: viewport.height, color: '#9eacbf', width: 2 }) }
  return { width: viewport.width, height: viewport.height, background: '#ffffff', commands }
}

function transformOverlayCommand(command: DrawCommand, viewport: GridViewport): DrawCommand {
  const transformRect = (rect: RenderRect): RenderRect => ({
    x: rect.x * viewport.zoom - viewport.scrollLeft,
    y: rect.y * viewport.zoom - viewport.scrollTop,
    width: rect.width * viewport.zoom,
    height: rect.height * viewport.zoom,
  })
  const clip = command.clip ? transformRect(command.clip) : undefined
  if (command.kind === 'rect') return { ...command, ...transformRect(command), strokeWidth: command.strokeWidth === undefined ? undefined : command.strokeWidth * viewport.zoom, ...(clip ? { clip } : {}) }
  if (command.kind === 'line') return {
    ...command,
    x1: command.x1 * viewport.zoom - viewport.scrollLeft,
    y1: command.y1 * viewport.zoom - viewport.scrollTop,
    x2: command.x2 * viewport.zoom - viewport.scrollLeft,
    y2: command.y2 * viewport.zoom - viewport.scrollTop,
    width: command.width === undefined ? undefined : command.width * viewport.zoom,
    ...(clip ? { clip } : {}),
  }
  if (command.kind === 'text') return {
    ...command,
    x: command.x * viewport.zoom - viewport.scrollLeft,
    y: command.y * viewport.zoom - viewport.scrollTop,
    font: command.font.replace(/(\d+(?:\.\d+)?)px/g, (_match, size: string) => `${Number(size) * viewport.zoom}px`),
    maxWidth: command.maxWidth === undefined ? undefined : command.maxWidth * viewport.zoom,
    ...(clip ? { clip } : {}),
  }
  return { ...command, ...transformRect(command), ...(clip ? { clip } : {}) }
}
