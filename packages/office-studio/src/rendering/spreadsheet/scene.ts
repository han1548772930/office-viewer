import type { ExcelCellView } from '../../services/excelViewAdapter'
import type { DrawCommand, RenderRect, RenderScene } from '../contracts'
import { intersectRect } from '../geometry'
import { cellAddress, columnLabel, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from './layout'
import type { GridRange, GridViewport } from './layout'

export interface SpreadsheetSceneOptions {
  activeCell: string
  selectedCells: readonly string[]
  selectedRanges?: readonly GridRange[]
  searchText?: string
  /** Drawing-object commands are expressed in unscaled worksheet CSS pixels. */
  overlayCommands?: readonly DrawCommand[]
}
export function displayCell(cell: ExcelCellView | undefined): string {
  if (!cell) return ''
  const value = cell.formula ? cell.cachedValue ?? cell.formula : cell.value
  if (cell.numberFormat === 'percent' && value !== '' && Number.isFinite(Number(value))) return `${(Number(value) * 100).toLocaleString('zh-CN')}%`
  if (cell.numberFormat === 'number' && value !== '' && Number.isFinite(Number(value))) return Number(value).toLocaleString('zh-CN')
  return value
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
    const text = displayCell(cell)
    if (search && text.toLocaleLowerCase().includes(search)) commands.push({ kind: 'rect', ...rect, fill: 'rgba(250,204,21,0.28)', clip })
    if (text) {
      const align = cell?.align ?? (column === 0 ? 'left' : Number.isFinite(Number(text)) ? 'right' : 'left')
      const fontSize = (cell?.fontSize ?? 11) * 4 / 3 * z
      const x = align === 'center' ? rect.x + rect.width / 2 : align === 'right' ? rect.x + rect.width - 6 * z : rect.x + 6 * z
      const font = `${cell?.italic ? 'italic ' : ''}${cell?.bold ? '600 ' : ''}${fontSize}px ${cell?.fontFamily ?? 'Aptos, "Microsoft YaHei", sans-serif'}`
      const textClip = intersectRect(clip, { x: rect.x + 3 * z, y: rect.y + 1, width: Math.max(0, rect.width - 6 * z), height: Math.max(0, rect.height - 2) })
      commands.push({ kind: 'text', x, y: rect.y + rect.height / 2, text, font, color: cell?.color ?? '#243247', align, baseline: 'middle', clip: textClip })
      if (cell?.underline) commands.push({ kind: 'line', x1: rect.x + 6 * z, y1: rect.y + rect.height / 2 + fontSize * 0.55, x2: rect.x + rect.width - 6 * z, y2: rect.y + rect.height / 2 + fontSize * 0.55, color: cell.color ?? '#243247', clip: textClip })
    }
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
