import type { ExcelCellView } from '../../services/excelViewAdapter'
import { sheetViewCellAt } from '../../services/excelViewAdapter.ts'
import type { DrawCommand, RenderRect, RenderScene } from '../contracts'
import { intersectRect } from '../geometry.ts'
import { cellAddress, columnLabel, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from './layout.ts'
import type { GridRange, GridViewport } from './layout'
import { presentSheetCellNumber } from '../../features/sheet-workbench/sheet-number-presentation.ts'
import { sheetCellTextAlignment } from '../../features/sheet-workbench/sheet-cell-alignment.ts'
import { canSheetCellTextOverflow } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'
import { createSheetCellTextCommands } from './cell-text.ts'
import { createSheetBorderLines as createBorderLines } from './borders.ts'
import { sparklineCommandsForCell } from '../../features/sparkline/sparkline-render.ts'
import { conditionalFormatStyleAt } from '../../features/conditional-formatting/conditional-formatting.ts'
import { conditionalIconPreset } from '../../features/conditional-formatting/conditional-icon-presets.ts'
import { evaluateFormula } from '../../features/formula/formula-engine.ts'
import { conditionalFormatFormulaAt } from '../../features/conditional-formatting/conditional-format-formulas.ts'
import { createBlankSheet, defaultCellStyle, type SpreadsheetCell, type WorkbookModel } from '../../domain/model.ts'
import { sheetRangeDashedOutlineCommands, SHEET_SELECTION_HANDLE_SIZE, SHEET_SELECTION_STROKE_WIDTH } from './selection.ts'
import { crosshairHighlightRanges } from '../../features/sheet-crosshair/crosshair-highlight.ts'
import { sheetMergedBorderSegments } from '../../features/sheet-workbench/sheet-cell-border.ts'
import type { CellRange } from '../../domain/model.ts'
import type { SheetImageResource } from '../../domain/sheet-images.ts'
import type { SheetDrawingLeafRef } from '../../domain/sheet-drawings.ts'
import { orderSheetDrawingCommands } from './drawing-order.ts'
import { fitSheetCellImage } from '../../features/sheet-images/sheet-image-layout.ts'

function gradientDataBarColor(color: string, strength: number): string {
  const channels = color.match(/^#([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i)
  if (!channels) return color
  return `#${channels.slice(1).map(channel => Math.round(255 + (parseInt(channel, 16) - 255) * (0.2 + 0.8 * strength)).toString(16).padStart(2, '0')).join('')}`
}

export interface SheetNumberFormatPreview { pattern: string; ranges: readonly GridRange[] }

export function sheetNumberFormatPreviewPatternAt(preview: SheetNumberFormatPreview | undefined, row: number, column: number): string | undefined {
  return preview?.ranges.some(range => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn)
    ? preview.pattern : undefined
}

export interface SpreadsheetSceneOptions {
  region?: string
  activeCell: string
  editing?: boolean
  selectedCells: readonly string[]
  selectedRanges?: readonly GridRange[]
  /** Helper outline shown while the fill handle is being dragged. */
  fillPreviewRange?: GridRange
  /** Format-painter source is a transient view mark, rendered only by its owning worksheet. */
  formatPainterSource?: GridRange
  formatPainterColor?: string
  suppressFillHandle?: boolean
  searchText?: string
  /** Temporary presentation for selected numeric cells; it is excluded from document history. */
  numberFormatPreview?: SheetNumberFormatPreview
  /** Drawing-object commands are expressed in unscaled worksheet CSS pixels. */
  overlayCommands?: readonly DrawCommand[]
  imageSource?: (resource: SheetImageResource) => CanvasImageSource | undefined
  imageVisible?: (row: number, column: number) => boolean
  backgroundImageSource?: CanvasImageSource
  backgroundImageVersion?: number
  imageCommands?: readonly DrawCommand[]
  drawingOrder?: readonly SheetDrawingLeafRef[]
  crosshairHighlight?: { enabled: boolean; color: string }
}
export function displayCell(cell: ExcelCellView | undefined, date1904 = false, region?: string, previewPattern?: string): string {
  return presentSheetCellNumber(cell, date1904, region, previewPattern).text
}

/** Layout and hit-testing share the same sparse axes and frozen-pane transforms. */
export function createSpreadsheetScene(geometry: SpreadsheetGeometry, viewport: GridViewport, options: SpreadsheetSceneOptions): RenderScene {
  const commands: DrawCommand[] = options.backgroundImageSource ? [{ kind: 'image', x: 0, y: 0, width: viewport.width, height: viewport.height,
    source: options.backgroundImageSource, sourceVersion: options.backgroundImageVersion }] : []
  const selected = new Set(options.selectedCells), search = options.searchText?.toLocaleLowerCase()
  const panes = geometry.panes(viewport), z = viewport.zoom
  const conditionalScopeValues = Object.entries(geometry.sheet.cells).flatMap(([address, cell]) => {
    const position = parseAddress(address)
    return position ? [{ row: position.row, column: position.column, value: cell.rawValue }] : []
  })
  const formulaSheet = createBlankSheet(geometry.sheet.name)
  formulaSheet.id = geometry.sheet.id
  formulaSheet.cells = Object.fromEntries(Object.entries(geometry.sheet.cells).flatMap(([address, cell]) => {
    const position = parseAddress(address)
    if (!position) return []
    const style = defaultCellStyle()
    style.font.bold = cell.bold ?? false
    style.font.italic = cell.italic ?? false
    style.font.color = cell.hyperlink ? '#0563c1' : cell.color ?? '#1f2937'
    style.fill = cell.background ?? '#ffffff'
    const modelCell: SpreadsheetCell = {
      id: `${position.row}:${position.column}`, row: position.row, col: position.column,
      value: cell.rawValue ?? cell.value, formula: cell.formula ? (cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`) : null,
      display: cell.displayValue ?? cell.value, style, hyperlink: cell.hyperlink,
    }
    return [[`${position.row}:${position.column}`, modelCell]]
  }))
  const formulaWorkbook: WorkbookModel = {
    kind: 'workbook', name: geometry.sheet.name, sheets: [formulaSheet], activeSheetId: formulaSheet.id,
    date1904: geometry.sheet.date1904 ?? false, modified: false, pivotTables: [],
  }
  const active = parseAddress(options.activeCell)
  const selectionRanges = options.selectedRanges?.length
    ? options.selectedRanges
    : active ? [{ startRow: active.row, endRow: active.row, startColumn: active.column, endColumn: active.column }] : []
  const isSelected = (row: number, column: number) => selected.has(cellAddress(row, column)) || options.selectedRanges?.some(range => row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn)
  const paintCell = (row: number, column: number, clip: RenderRect, foreground: DrawCommand[], borders: DrawCommand[], visibleRange: CellRange, foregroundOnly = false) => {
    const key = cellAddress(row, column), cell = sheetViewCellAt(geometry.sheet, row, column), rect = geometry.cellRect(row, column, viewport)
    const merge = geometry.mergeAt(row, column)
    const visible = intersectRect(rect, clip)
    if (rect.width <= 0 || rect.height <= 0 || !foregroundOnly && (visible.width <= 0 || visible.height <= 0)) return
    if (!foregroundOnly) commands.push({ kind: 'rect', ...rect, fill: cell?.background ?? (geometry.sheet.backgroundImage ? 'transparent' : '#ffffff'), clip })
    const conditionalFormulaValue = (formula: string, rule: NonNullable<typeof geometry.sheet.conditionalFormats>[number]) => {
        try {
          const translated = conditionalFormatFormulaAt(formula, rule, row, column)
          if (translated === undefined) return undefined
          return evaluateFormula(translated.startsWith('=') ? translated : `=${translated}`, formulaWorkbook, formulaSheet.id, { formulaCell: { row, column } })
        } catch { return undefined }
      }
    const conditionalStyle = conditionalFormatStyleAt(geometry.sheet.conditionalFormats ?? [], row, column, cell?.rawValue, [], conditionalScopeValues,
      (formula, rule) => { const result = conditionalFormulaValue(formula, rule); return result === true || typeof result === 'number' && result !== 0 }, conditionalFormulaValue)
    if (conditionalStyle.fill && !foregroundOnly) commands.push({ kind: 'rect', ...rect, fill: conditionalStyle.fill, clip })
    if (conditionalStyle.dataBar && !foregroundOnly) {
      const bar = conditionalStyle.dataBar
      const width = rect.width * bar.fraction
      const x = rect.x + rect.width * bar.start - (bar.negative ? width : 0)
      if (bar.gradient) {
        const segments = Math.max(1, Math.min(32, Math.ceil(width / 4)))
        for (let index = 0; index < segments; index++) {
          const fraction = (index + 0.5) / segments
          const color = gradientDataBarColor(bar.color, bar.negative ? fraction : 1 - fraction)
          commands.push({ kind: 'rect', x: x + width * index / segments, y: rect.y + rect.height * 0.18, width: width / segments, height: rect.height * 0.64, fill: color, clip })
        }
      } else commands.push({ kind: 'rect', x, y: rect.y + rect.height * 0.18, width, height: rect.height * 0.64, fill: bar.color, clip })
    }
    const conditionalIcon = conditionalStyle.icon ? conditionalIconPreset(conditionalStyle.icon.iconType, conditionalStyle.icon.iconId) : undefined
    if (conditionalIcon && !foregroundOnly) foreground.push({ kind: 'text', x: rect.x + 10 * z, y: rect.y + rect.height / 2, text: conditionalIcon.symbol, font: `${16 * z}px "Segoe UI Symbol", sans-serif`, color: conditionalIcon.color, align: 'center', baseline: 'middle', clip: intersectRect(rect, clip) })
    if (geometry.sheet.showGridlines !== false && !foregroundOnly) commands.push({ kind: 'rect', ...rect, fill: 'transparent', stroke: '#e5e9f0', strokeWidth: 1, clip })
    const decodedImage = cell?.image && (!options.imageVisible || options.imageVisible(row, column)) && options.imageSource?.(cell.image.resource)
    if (decodedImage && cell?.image && !foregroundOnly) foreground.push({ kind: 'image', ...fitSheetCellImage(rect,
      cell.image.resource.width, cell.image.resource.height, 2 * z), source: decodedImage, clip: intersectRect(rect, clip) })
    if (isSelected(row, column) && !foregroundOnly) commands.push({ kind: 'rect', ...rect, fill: 'rgba(36,107,253,0.09)', clip })
    if ((cell?.border || merge) && !foregroundOnly) {
      const segments = merge ? sheetMergedBorderSegments({ startRow: merge.startRow, endRow: merge.endRow, startCol: merge.startColumn, endCol: merge.endColumn },
        (segmentRow, segmentCol) => sheetViewCellAt(geometry.sheet, segmentRow, segmentCol).border, visibleRange)
        : (['top', 'bottom', 'left', 'right'] as const).map(side => ({ side, row, col: column, edge: cell?.border?.[side] }))
      for (const segment of segments) {
        const { side } = segment, edge = borderEdge(segment.edge)
        if (!edge) continue
        const edgeRect = merge ? geometry.cellRect(segment.row, segment.col, viewport, false) : rect
        const horizontal = side === 'top' || side === 'bottom'
        const borderX = side === 'right' ? edgeRect.x + edgeRect.width : edgeRect.x
        const borderY = side === 'bottom' ? edgeRect.y + edgeRect.height : edgeRect.y
        borders.push(...createBorderLines(borderX, borderY, horizontal ? edgeRect.x + edgeRect.width : borderX,
          horizontal ? borderY : edgeRect.y + edgeRect.height, edge.style, edge.color, z, clip))
      }
      const diagonalDown = borderEdge(cell?.border?.diagonalDown ?? cell?.border?.diagonal)
      if (diagonalDown) borders.push(...createBorderLines(rect.x, rect.y, rect.x + rect.width, rect.y + rect.height, diagonalDown.style, diagonalDown.color, z, clip))
      const diagonalUp = borderEdge(cell?.border?.diagonalUp)
      if (diagonalUp) borders.push(...createBorderLines(rect.x, rect.y + rect.height, rect.x + rect.width, rect.y, diagonalUp.style, diagonalUp.color, z, clip))
      for (const [key, x1, y1, x2, y2] of compoundDiagonalSegments) {
        const edge = borderEdge(cell?.border?.[key])
        if (edge) borders.push(...createBorderLines(rect.x + x1 * rect.width, rect.y + y1 * rect.height,
          rect.x + x2 * rect.width, rect.y + y2 * rect.height, edge.style, edge.color, z, clip))
      }
    }
    const presentation = presentSheetCellNumber(cell, geometry.sheet.date1904, options.region,
      sheetNumberFormatPreviewPatternAt(options.numberFormatPreview, row, column))
    const hideValue = Boolean(conditionalStyle.dataBar && !conditionalStyle.dataBar.showValue || conditionalStyle.icon && !conditionalStyle.icon.showValue)
    const text = hideValue ? '' : presentation.text
    if (search && text.toLocaleLowerCase().includes(search) && !foregroundOnly) commands.push({ kind: 'rect', ...rect, fill: 'rgba(250,204,21,0.28)', clip })
    if (text || !hideValue && presentation.runs?.some(run => run.kind === 'fill' && run.character.trim())) {
      const align = sheetCellTextAlignment(cell)
      const iconPadding = conditionalIcon ? 18 * z : 0
      let start = rect.x + 3 * z + iconPadding, end = rect.x + rect.width - 3 * z
      if (canSheetCellTextOverflow(cell) && !merge) {
        const bounds = geometry.overflowColumnBounds(row, column)
        if (align !== 'left') start = geometry.cellRect(row, bounds.startColumn, viewport, false).x + 3 * z + iconPadding
        if (align !== 'right') {
          const last = geometry.cellRect(row, bounds.endColumn, viewport, false)
          end = last.x + last.width - 3 * z
        }
      }
      const textClip = intersectRect(clip, { x: start, y: rect.y + z, width: Math.max(0, end - start), height: Math.max(0, rect.height - 2 * z) })
      foreground.push(...createSheetCellTextCommands({ cell, presentation: { ...presentation, text }, rect, clip: textClip, zoom: z, iconPadding,
        color: conditionalStyle.fontColor, underline: conditionalStyle.underline, strike: conditionalStyle.strike,
        fontOverrides: { bold: conditionalStyle.bold, italic: conditionalStyle.italic } }))
    }
    if (!foregroundOnly) foreground.push(...sparklineCommandsForCell(geometry.sheet, cell, rect).map(command => ({ ...command, clip: intersectRect(command.clip ?? rect, clip) })))
  }
  for (const pane of panes) {
    if (pane.clip.width <= 0 || pane.clip.height <= 0) continue
    const anchors = new Set<string>()
    const foreground: DrawCommand[] = [], borders: DrawCommand[] = []
    const visibleRange = { startRow: pane.rows[0] ?? 0, endRow: pane.rows.at(-1) ?? -1,
      startCol: pane.columns[0] ?? 0, endCol: pane.columns.at(-1) ?? -1 }
    for (const row of pane.rows) for (const column of pane.columns) {
      const merge = geometry.mergeAt(row, column)
      if (merge) anchors.add(cellAddress(merge.startRow, merge.startColumn))
      else paintCell(row, column, pane.clip, foreground, borders, visibleRange)
    }
    for (const address of anchors) { const cell = parseAddress(address)!; paintCell(cell.row, cell.column, pane.clip, foreground, borders, visibleRange) }
    const firstColumn = pane.columns[0], lastColumn = pane.columns.at(-1)
    if (firstColumn !== undefined && lastColumn !== undefined) {
      for (const row of pane.rows) {
        for (const column of [geometry.contentColumnBefore(row, firstColumn), geometry.contentColumnAfter(row, lastColumn)]) {
          if (column === undefined || (column < geometry.frozenColumns) !== (firstColumn < geometry.frozenColumns)
            || geometry.mergeAt(row, column) || !canSheetCellTextOverflow(geometry.sheet.cells[cellAddress(row, column)])) continue
          paintCell(row, column, pane.clip, foreground, borders, visibleRange, true)
        }
      }
    }
    if (options.crosshairHighlight?.enabled) {
      const ranges = crosshairHighlightRanges(selectionRanges, geometry.sheet.rowCount, geometry.sheet.columnCount)
      for (const range of ranges) {
        const rows = pane.rows.filter(row => row >= range.startRow && row <= range.endRow)
        const columns = pane.columns.filter(column => column >= range.startColumn && column <= range.endColumn)
        if (!rows.length || !columns.length) continue
        const start = geometry.cellRect(rows[0]!, columns[0]!, viewport, false)
        const end = geometry.cellRect(rows[rows.length - 1]!, columns[columns.length - 1]!, viewport, false)
        commands.push({ kind: 'rect', x: start.x, y: start.y, width: end.x + end.width - start.x, height: end.y + end.height - start.y, fill: options.crosshairHighlight.color, clip: pane.clip })
      }
    }
    commands.push(...borders, ...foreground)
    for (let rangeIndex = 0; rangeIndex < selectionRanges.length; rangeIndex += 1) {
      const range = selectionRanges[rangeIndex]!
      const rows = pane.rows.filter(row => row >= range.startRow && row <= range.endRow)
      const columns = pane.columns.filter(column => column >= range.startColumn && column <= range.endColumn)
      if (!rows.length || !columns.length) continue
      const rangeStart = geometry.cellRect(range.startRow, range.startColumn, viewport, false)
      const rangeEnd = geometry.cellRect(range.endRow, range.endColumn, viewport, false)
      const left = rangeStart.x, top = rangeStart.y
      const right = rangeEnd.x + rangeEnd.width, bottom = rangeEnd.y + rangeEnd.height
      const strokeOffset = SHEET_SELECTION_STROKE_WIDTH / 2
      const editingCellRange = options.editing && active && range.startRow === range.endRow && range.startColumn === range.endColumn && active.row === range.startRow && active.column === range.startColumn
      if (!editingCellRange) {
        commands.push({ kind: 'line', x1: left + strokeOffset, y1: top + strokeOffset, x2: right + strokeOffset, y2: top + strokeOffset, color: '#246bfd', width: SHEET_SELECTION_STROKE_WIDTH, clip: pane.clip })
        commands.push({ kind: 'line', x1: left + strokeOffset, y1: bottom + strokeOffset, x2: right + strokeOffset, y2: bottom + strokeOffset, color: '#246bfd', width: SHEET_SELECTION_STROKE_WIDTH, clip: pane.clip })
        commands.push({ kind: 'line', x1: left + strokeOffset, y1: top + strokeOffset, x2: left + strokeOffset, y2: bottom + strokeOffset, color: '#246bfd', width: SHEET_SELECTION_STROKE_WIDTH, clip: pane.clip })
        commands.push({ kind: 'line', x1: right + strokeOffset, y1: top + strokeOffset, x2: right + strokeOffset, y2: bottom + strokeOffset, color: '#246bfd', width: SHEET_SELECTION_STROKE_WIDTH, clip: pane.clip })
      }
      const endpointIsFrozenRow = range.endRow < geometry.frozenRows
      const endpointIsFrozenColumn = range.endColumn < geometry.frozenColumns
      const paneHasEndpoint = pane.rows.some(row => (row < geometry.frozenRows) === endpointIsFrozenRow)
        && pane.columns.some(column => (column < geometry.frozenColumns) === endpointIsFrozenColumn)
      if (rangeIndex === selectionRanges.length - 1 && !editingCellRange && !options.suppressFillHandle && paneHasEndpoint) {
        const handleSize = SHEET_SELECTION_HANDLE_SIZE
        commands.push({ kind: 'rect', x: right - handleSize / 2, y: bottom - handleSize / 2, width: handleSize - SHEET_SELECTION_STROKE_WIDTH, height: handleSize - SHEET_SELECTION_STROKE_WIDTH, fill: '#246bfd', stroke: '#ffffff', strokeWidth: SHEET_SELECTION_STROKE_WIDTH, clip: pane.clip })
      }
    }
    const preview = options.fillPreviewRange
    if (preview) {
      const rows = pane.rows.filter(row => row >= preview.startRow && row <= preview.endRow)
      const columns = pane.columns.filter(column => column >= preview.startColumn && column <= preview.endColumn)
      if (rows.length && columns.length) {
        const start = geometry.cellRect(preview.startRow, preview.startColumn, viewport, false)
        const end = geometry.cellRect(preview.endRow, preview.endColumn, viewport, false)
        const left = start.x, top = start.y
        const right = end.x + end.width, bottom = end.y + end.height
        const width = SHEET_SELECTION_STROKE_WIDTH
        const strokeOffset = width / 2
        const color = '#246bfd'
        commands.push({ kind: 'line', x1: left + strokeOffset, y1: top + strokeOffset, x2: right + strokeOffset, y2: top + strokeOffset, color, width, clip: pane.clip })
        commands.push({ kind: 'line', x1: left + strokeOffset, y1: bottom + strokeOffset, x2: right + strokeOffset, y2: bottom + strokeOffset, color, width, clip: pane.clip })
        commands.push({ kind: 'line', x1: left + strokeOffset, y1: top + strokeOffset, x2: left + strokeOffset, y2: bottom + strokeOffset, color, width, clip: pane.clip })
        commands.push({ kind: 'line', x1: right + strokeOffset, y1: top + strokeOffset, x2: right + strokeOffset, y2: bottom + strokeOffset, color, width, clip: pane.clip })
      }
    }
  }
  if (options.formatPainterSource) commands.push(...sheetRangeDashedOutlineCommands(geometry, options.formatPainterSource, viewport, options.formatPainterColor))
  // Overlay drawing objects share the same scroll/zoom transform as the sheet;
  // headers are painted afterwards and therefore stay above floating objects.
  const drawingCommands: DrawCommand[] = []
  for (const command of options.overlayCommands ?? []) {
    if (!command.drawingRef) { drawingCommands.push(transformOverlayCommand(command, viewport)); continue }
    for (const pane of geometry.panes(viewport)) {
      if (!pane.clip.width || !pane.clip.height) continue
      const frozenX = geometry.frozenColumns > 0 && pane.clip.x === GRID_HEADER_WIDTH,
        frozenY = geometry.frozenRows > 0 && pane.clip.y === GRID_HEADER_HEIGHT,
        projected = transformOverlayCommand(command, { ...viewport, scrollLeft: frozenX ? 0 : viewport.scrollLeft, scrollTop: frozenY ? 0 : viewport.scrollTop })
      projected.clip = projected.clip ? intersectRect(projected.clip, pane.clip) : pane.clip
      if (projected.clip.width > 0 && projected.clip.height > 0) drawingCommands.push(projected)
    }
  }
  drawingCommands.push(...(options.imageCommands ?? []))
  commands.push(...orderSheetDrawingCommands(drawingCommands, options.drawingOrder))
  commands.push({ kind: 'rect', x: 0, y: 0, width: viewport.width, height: GRID_HEADER_HEIGHT, fill: '#f5f7fa' })
  commands.push({ kind: 'rect', x: 0, y: GRID_HEADER_HEIGHT, width: GRID_HEADER_WIDTH, height: Math.max(0, viewport.height - GRID_HEADER_HEIGHT), fill: '#f5f7fa' })
  for (const pane of panes) {
    for (const column of pane.columns) {
      const rect = geometry.cellRect(0, column, viewport, false), clip = { x: pane.clip.x, y: 0, width: pane.clip.width, height: GRID_HEADER_HEIGHT }
      commands.push({ kind: 'rect', x: rect.x, y: 0, width: rect.width, height: GRID_HEADER_HEIGHT, fill: active?.column === column ? '#e8f0ff' : '#f5f7fa', stroke: '#dfe5ee', clip })
      commands.push({ kind: 'text', x: rect.x + rect.width / 2, y: GRID_HEADER_HEIGHT / 2, text: columnLabel(column), font: '12px "Segoe UI", sans-serif', color: active?.column === column ? '#246bfd' : '#69778c', align: 'center', baseline: 'middle', clip })
    }
    for (const row of pane.rows) {
      const rect = geometry.cellRect(row, 0, viewport, false), clip = { x: 0, y: pane.clip.y, width: GRID_HEADER_WIDTH, height: pane.clip.height }
      commands.push({ kind: 'rect', x: 0, y: rect.y, width: GRID_HEADER_WIDTH, height: rect.height, fill: active?.row === row ? '#e8f0ff' : '#f5f7fa', stroke: '#dfe5ee', clip })
      commands.push({ kind: 'text', x: GRID_HEADER_WIDTH - 7, y: rect.y + rect.height / 2, text: String(row + 1), font: '12px "Segoe UI", sans-serif', color: active?.row === row ? '#246bfd' : '#69778c', align: 'right', baseline: 'middle', clip })
    }
  }
  commands.push({ kind: 'rect', x: 0, y: 0, width: GRID_HEADER_WIDTH, height: GRID_HEADER_HEIGHT, fill: '#eef2f7', stroke: '#dfe5ee' })
  commands.push({ kind: 'line', x1: GRID_HEADER_WIDTH - 13, y1: GRID_HEADER_HEIGHT - 5, x2: GRID_HEADER_WIDTH - 5, y2: GRID_HEADER_HEIGHT - 13, color: '#9aa8b9', width: 3 })
  for (const line of geometry.outlineBracketLines(viewport)) commands.push({ kind: 'line', x1: line.x1, y1: line.y1, x2: line.x2, y2: line.y2, color: '#8593a5', width: 1 })
  if (geometry.frozenRows) { const y = GRID_HEADER_HEIGHT + geometry.rows.offset(geometry.frozenRows) * z; commands.push({ kind: 'line', x1: GRID_HEADER_WIDTH, y1: y, x2: viewport.width, y2: y, color: '#9eacbf', width: 2 }) }
  if (geometry.frozenColumns) { const x = GRID_HEADER_WIDTH + geometry.columns.offset(geometry.frozenColumns) * z; commands.push({ kind: 'line', x1: x, y1: GRID_HEADER_HEIGHT, x2: x, y2: viewport.height, color: '#9eacbf', width: 2 }) }
  return { width: viewport.width, height: viewport.height, background: '#ffffff', commands }
}

const compoundDiagonalSegments = [
  ['topLeftBottomCenter', 0, 0, .5, 1], ['topLeftMiddleRight', 0, 0, 1, .5],
  ['middleLeftTopRight', 0, .5, 1, 0], ['bottomCenterTopRight', .5, 1, 1, 0],
] as const

function borderEdge(raw: unknown): { style: string; color: string } | undefined {
  if (!raw) return undefined
  const edge = typeof raw === 'object' ? raw as Record<string, unknown> : undefined
  const style = typeof raw === 'string' ? raw.split(':')[0] : edge?.style
  if (typeof style !== 'string' || style === 'none' || style === '') return undefined
  const color = typeof edge?.color === 'string' ? edge.color : typeof raw === 'string' ? raw.split(':')[1] ?? '#1f2937' : '#1f2937'
  return { style, color }
}

function transformOverlayCommand(command: DrawCommand, viewport: GridViewport): DrawCommand {
  const transformRect = (rect: RenderRect): RenderRect => ({
    x: GRID_HEADER_WIDTH + (rect.x - GRID_HEADER_WIDTH) * viewport.zoom - viewport.scrollLeft,
    y: GRID_HEADER_HEIGHT + (rect.y - GRID_HEADER_HEIGHT) * viewport.zoom - viewport.scrollTop,
    width: rect.width * viewport.zoom,
    height: rect.height * viewport.zoom,
  })
  const clip = command.clip ? transformRect(command.clip) : undefined
  if (command.kind === 'image' && command.matrix) {
    const m = command.matrix
    return { ...command, matrix: { a: m.a * viewport.zoom, b: m.b * viewport.zoom, c: m.c * viewport.zoom, d: m.d * viewport.zoom,
      e: GRID_HEADER_WIDTH + (m.e - GRID_HEADER_WIDTH) * viewport.zoom - viewport.scrollLeft,
      f: GRID_HEADER_HEIGHT + (m.f - GRID_HEADER_HEIGHT) * viewport.zoom - viewport.scrollTop }, ...(clip ? { clip } : {}) }
  }
  if (command.kind === 'rect') return { ...command, ...transformRect(command), strokeWidth: command.strokeWidth === undefined ? undefined : command.strokeWidth * viewport.zoom, ...(clip ? { clip } : {}) }
  if (command.kind === 'line') return {
    ...command,
    x1: GRID_HEADER_WIDTH + (command.x1 - GRID_HEADER_WIDTH) * viewport.zoom - viewport.scrollLeft,
    y1: GRID_HEADER_HEIGHT + (command.y1 - GRID_HEADER_HEIGHT) * viewport.zoom - viewport.scrollTop,
    x2: GRID_HEADER_WIDTH + (command.x2 - GRID_HEADER_WIDTH) * viewport.zoom - viewport.scrollLeft,
    y2: GRID_HEADER_HEIGHT + (command.y2 - GRID_HEADER_HEIGHT) * viewport.zoom - viewport.scrollTop,
    width: command.width === undefined ? undefined : command.width * viewport.zoom,
    ...(clip ? { clip } : {}),
  }
  if (command.kind === 'polygon') return {
    ...command,
    points: command.points.map((point) => ({ x: GRID_HEADER_WIDTH + (point.x - GRID_HEADER_WIDTH) * viewport.zoom - viewport.scrollLeft,
      y: GRID_HEADER_HEIGHT + (point.y - GRID_HEADER_HEIGHT) * viewport.zoom - viewport.scrollTop })),
    ...(clip ? { clip } : {}),
  }
  if (command.kind === 'text') return {
    ...command,
    x: GRID_HEADER_WIDTH + (command.x - GRID_HEADER_WIDTH) * viewport.zoom - viewport.scrollLeft,
    y: GRID_HEADER_HEIGHT + (command.y - GRID_HEADER_HEIGHT) * viewport.zoom - viewport.scrollTop,
    font: command.font.replace(/(\d+(?:\.\d+)?)px/g, (_match, size: string) => `${Number(size) * viewport.zoom}px`),
    maxWidth: command.maxWidth === undefined ? undefined : command.maxWidth * viewport.zoom,
    ...(clip ? { clip } : {}),
  }
  return { ...command, ...transformRect(command), ...(clip ? { clip } : {}) }
}
