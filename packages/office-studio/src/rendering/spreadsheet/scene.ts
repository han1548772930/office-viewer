import type { ExcelCellView } from '../../services/excelViewAdapter'
import type { DrawCommand, RenderRect, RenderScene } from '../contracts'
import { intersectRect } from '../geometry.ts'
import { cellAddress, columnLabel, GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, parseAddress, SpreadsheetGeometry } from './layout.ts'
import type { GridRange, GridViewport } from './layout'
import { formatSheetValue, SHEET_NUMBER_FORMATS } from '../../features/sheet-workbench/sheet-number-format.ts'
import { sparklineCommandsForCell } from '../../features/sparkline/sparkline-render.ts'
import { conditionalFormatStyleAt } from '../../features/conditional-formatting/conditional-formatting.ts'
import { conditionalIconPreset } from '../../features/conditional-formatting/conditional-icon-presets.ts'
import { evaluateFormula } from '../../features/formula/formula-engine.ts'
import { translateFormulaReferences } from '../../features/formula/formula-references.ts'
import { createBlankSheet, defaultCellStyle, type SpreadsheetCell, type WorkbookModel } from '../../domain/model.ts'
import { SHEET_SELECTION_HANDLE_SIZE, SHEET_SELECTION_STROKE_WIDTH } from './selection.ts'
import { crosshairHighlightRanges } from '../../features/sheet-crosshair/crosshair-highlight.ts'

let textMeasureContext: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null | undefined
function gradientDataBarColor(color: string, strength: number): string {
  const channels = color.match(/^#([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i)
  if (!channels) return color
  return `#${channels.slice(1).map(channel => Math.round(255 + (parseInt(channel, 16) - 255) * (0.2 + 0.8 * strength)).toString(16).padStart(2, '0')).join('')}`
}
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
  editing?: boolean
  selectedCells: readonly string[]
  selectedRanges?: readonly GridRange[]
  /** Helper outline shown while the fill handle is being dragged. */
  fillPreviewRange?: GridRange
  searchText?: string
  /** Drawing-object commands are expressed in unscaled worksheet CSS pixels. */
  overlayCommands?: readonly DrawCommand[]
  crosshairHighlight?: { enabled: boolean; color: string }
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
  const conditionalScopeValues = Object.values(geometry.sheet.cells).map(cell => ({ row: cell.row, column: cell.col, value: cell.rawValue }))
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
  const paintCell = (row: number, column: number, clip: RenderRect) => {
    const key = cellAddress(row, column), cell = geometry.sheet.cells[key], rect = geometry.cellRect(row, column, viewport)
    const visible = intersectRect(rect, clip)
    if (visible.width <= 0 || visible.height <= 0) return
    commands.push({ kind: 'rect', ...rect, fill: cell?.background ?? '#ffffff', clip })
    const conditionalFormulaValue = (formula: string, rule: NonNullable<typeof geometry.sheet.conditionalFormats>[number]) => {
        const firstRange = [...rule.ranges].sort((left, right) => left.startRow - right.startRow || left.startCol - right.startCol)[0]
        if (!firstRange) return undefined
        try {
          const translated = translateFormulaReferences(formula, row - firstRange.startRow, column - firstRange.startCol)
          return evaluateFormula(translated.startsWith('=') ? translated : `=${translated}`, formulaWorkbook, formulaSheet.id)
        } catch { return undefined }
      }
    const conditionalStyle = conditionalFormatStyleAt(geometry.sheet.conditionalFormats ?? [], row, column, cell?.rawValue, [], conditionalScopeValues,
      (formula, rule) => { const result = conditionalFormulaValue(formula, rule); return result === true || typeof result === 'number' && result !== 0 }, conditionalFormulaValue)
    if (conditionalStyle.fill) commands.push({ kind: 'rect', ...rect, fill: conditionalStyle.fill, clip })
    if (conditionalStyle.dataBar) {
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
    if (conditionalIcon) commands.push({ kind: 'text', x: rect.x + 10 * z, y: rect.y + rect.height / 2, text: conditionalIcon.symbol, font: `${16 * z}px "Segoe UI Symbol", sans-serif`, color: conditionalIcon.color, align: 'center', baseline: 'middle', clip })
    if (geometry.sheet.showGridlines !== false) commands.push({ kind: 'rect', ...rect, fill: 'transparent', stroke: '#e5e9f0', strokeWidth: 1, clip })
    if (isSelected(row, column)) commands.push({ kind: 'rect', ...rect, fill: 'rgba(36,107,253,0.09)', clip })
    if (cell?.border) {
      for (const side of ['top', 'bottom', 'left', 'right'] as const) {
        const raw = cell.border[side]
        const edge = borderEdge(raw)
        if (!edge) continue
        const horizontal = side === 'top' || side === 'bottom'
        const borderX = side === 'right' ? rect.x + rect.width : rect.x
        const borderY = side === 'bottom' ? rect.y + rect.height : rect.y
        commands.push(...createBorderLines(borderX, borderY, horizontal ? rect.x + rect.width : borderX,
          horizontal ? borderY : rect.y + rect.height, edge.style, edge.color, z, clip))
      }
      const diagonalDown = borderEdge(cell.border.diagonalDown)
      if (diagonalDown) commands.push(...createBorderLines(rect.x, rect.y, rect.x + rect.width, rect.y + rect.height, diagonalDown.style, diagonalDown.color, z, clip))
      const diagonalUp = borderEdge(cell.border.diagonalUp)
      if (diagonalUp) commands.push(...createBorderLines(rect.x, rect.y + rect.height, rect.x + rect.width, rect.y, diagonalUp.style, diagonalUp.color, z, clip))
    }
    const text = conditionalStyle.dataBar && !conditionalStyle.dataBar.showValue || conditionalStyle.icon && !conditionalStyle.icon.showValue ? '' : displayCell(cell, geometry.sheet.date1904)
    if (search && text.toLocaleLowerCase().includes(search)) commands.push({ kind: 'rect', ...rect, fill: 'rgba(250,204,21,0.28)', clip })
    if (text) {
      const align = cell?.align ?? (column === 0 ? 'left' : Number.isFinite(Number(text)) ? 'right' : 'left')
      const baseFontSize = (cell?.fontSize ?? 11) * 4 / 3 * z
      const fontFamily = cell?.fontFamily ?? 'Aptos, "Microsoft YaHei", sans-serif'
      const fontPrefix = `${conditionalStyle.italic ?? cell?.italic ? 'italic ' : ''}${conditionalStyle.bold ?? cell?.bold ? '600 ' : ''}`
      const font = `${fontPrefix}${baseFontSize}px ${fontFamily}`
      const iconPadding = conditionalIcon ? 18 * z : 0
      const availableTextWidth = Math.max(0, rect.width - 12 * z - iconPadding)
      const measuredWidth = cell?.shrinkToFit && cell.wrap !== true ? measureTextWidth(text, font) : undefined
      const maxWidth = measuredWidth !== undefined && measuredWidth > availableTextWidth && availableTextWidth > 0
        ? availableTextWidth
        : undefined
      const fontSize = baseFontSize
      const x = align === 'center' ? rect.x + (rect.width + iconPadding) / 2 : align === 'right' ? rect.x + rect.width - 6 * z : rect.x + 6 * z + iconPadding
      const rotation = cell?.rotation ?? 0
      const textClip = intersectRect(clip, { x: rect.x + 3 * z + iconPadding, y: rect.y + 1, width: Math.max(0, rect.width - 6 * z - iconPadding), height: Math.max(0, rect.height - 2) })
      const y = cell?.vertical === 'top' ? rect.y + 3 * z + fontSize / 2 : cell?.vertical === 'bottom' ? rect.y + rect.height - 3 * z - fontSize / 2 : rect.y + rect.height / 2
      const color = conditionalStyle.fontColor ?? (cell?.hyperlink ? '#0563c1' : cell?.color) ?? '#243247'
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
          if ((conditionalStyle.underline ?? (cell.hyperlink ? true : cell.underline)) || (conditionalStyle.strike ?? cell.strike)) {
            const width = decorationWidth(glyph, font, availableTextWidth)
            const left = centerX - width / 2
            if (conditionalStyle.underline ?? (cell.hyperlink ? true : cell.underline)) commands.push({ kind: 'line', x1: left, y1: glyphY + lineHeight * 0.55, x2: left + width, y2: glyphY + lineHeight * 0.55, color, clip: textClip })
            if (conditionalStyle.strike ?? cell.strike) commands.push({ kind: 'line', x1: left, y1: glyphY, x2: left + width, y2: glyphY, color, clip: textClip })
          }
        }
      } else {
        if (cell?.numberFormat === 'accounting' && /^[$¥￥€£] /.test(text)) {
          commands.push({ kind: 'text', x: rect.x + 6 * z, y, text: text[0]!, font, color, align: 'left', baseline: 'middle', rotation, maxWidth, clip: textClip })
          commands.push({ kind: 'text', x: rect.x + rect.width - 6 * z, y, text: text.slice(2), font, color, align: 'right', baseline: 'middle', rotation, maxWidth, clip: textClip })
        } else commands.push({ kind: 'text', x, y, text, font, color, align, baseline: 'middle', rotation, maxWidth, clip: textClip })
      }
      if (!cell?.rotationVertical && ((conditionalStyle.underline ?? (cell?.hyperlink ? true : cell?.underline)) || (conditionalStyle.strike ?? cell?.strike))) {
        const width = decorationWidth(text, font, Math.max(0, rect.width - 12 * z))
        const left = align === 'center' ? x - width / 2 : align === 'right' ? x - width : x
        const decoration = (offsetY: number) => {
          const [x1, y1] = rotatePoint(left, y + offsetY, x, y, rotation)
          const [x2, y2] = rotatePoint(left + width, y + offsetY, x, y, rotation)
          commands.push({ kind: 'line', x1, y1, x2, y2, color, clip: textClip })
        }
        if (conditionalStyle.underline ?? (cell?.hyperlink ? true : cell?.underline)) decoration(fontSize * 0.55)
        if (conditionalStyle.strike ?? cell?.strike) decoration(0)
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
      if (rangeIndex === selectionRanges.length - 1 && !editingCellRange && paneHasEndpoint) {
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
  // Overlay drawing objects share the same scroll/zoom transform as the sheet;
  // headers are painted afterwards and therefore stay above floating objects.
  for (const command of options.overlayCommands ?? []) commands.push(transformOverlayCommand(command, viewport))
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
  for (const control of geometry.outlineToggleRects(viewport)) {
    commands.push({ kind: 'rect', x: control.x, y: control.y, width: control.width, height: control.height, fill: '#ffffff', stroke: '#7c899a', strokeWidth: 1 })
    const centerX = control.x + control.width / 2, centerY = control.y + control.height / 2
    commands.push({ kind: 'line', x1: centerX - 3, y1: centerY, x2: centerX + 3, y2: centerY, color: '#465568', width: 1 })
    if (control.collapsed) commands.push({ kind: 'line', x1: centerX, y1: centerY - 3, x2: centerX, y2: centerY + 3, color: '#465568', width: 1 })
  }
  const outline = geometry.sheet.outline
  if (outline) {
    for (const axis of ['row', 'column'] as const) {
      const groups = outline.groups.filter(group => group.axis === axis)
      if (!groups.length) continue
      const maxLevel = Math.max(...groups.map(group => groups.filter(parent => parent.start <= group.start && parent.end >= group.end && (parent.start < group.start || parent.end > group.end)).length)) + 2
      for (let level = 1; level <= maxLevel; level += 1) {
        const x = 1 + (level - 1) * 12
        const y = axis === 'row' ? 1 : 13
        commands.push({ kind: 'rect', x, y, width: 11, height: 11, fill: '#ffffff', stroke: '#9aa6b4', strokeWidth: 1 })
        commands.push({ kind: 'text', x: x + 5.5, y: y + 5.5, text: String(level), font: '8px "Segoe UI", sans-serif', color: '#536174', align: 'center', baseline: 'middle' })
      }
    }
  }
  if (geometry.frozenRows) { const y = GRID_HEADER_HEIGHT + geometry.rows.offset(geometry.frozenRows) * z; commands.push({ kind: 'line', x1: GRID_HEADER_WIDTH, y1: y, x2: viewport.width, y2: y, color: '#9eacbf', width: 2 }) }
  if (geometry.frozenColumns) { const x = GRID_HEADER_WIDTH + geometry.columns.offset(geometry.frozenColumns) * z; commands.push({ kind: 'line', x1: x, y1: GRID_HEADER_HEIGHT, x2: x, y2: viewport.height, color: '#9eacbf', width: 2 }) }
  return { width: viewport.width, height: viewport.height, background: '#ffffff', commands }
}

function borderEdge(raw: unknown): { style: string; color: string } | undefined {
  if (!raw) return undefined
  const edge = typeof raw === 'object' ? raw as Record<string, unknown> : undefined
  const style = typeof raw === 'string' ? raw.split(':')[0] : edge?.style
  if (typeof style !== 'string' || style === 'none' || style === '') return undefined
  const color = typeof edge?.color === 'string' ? edge.color : typeof raw === 'string' ? raw.split(':')[1] ?? '#1f2937' : '#1f2937'
  return { style, color }
}

function createBorderLines(
  x1: number, y1: number, x2: number, y2: number, style: string, color: string, zoom: number, clip: RenderRect,
): DrawCommand[] {
  const width = style.startsWith('medium') ? 2 : style === 'thick' ? 3 : style === 'hair' ? 0.5 : 1
  const pattern = borderDashPattern(style).map(value => value * zoom)
  const double = style === 'double'
  const dx = x2 - x1, dy = y2 - y1
  const length = Math.hypot(dx, dy)
  if (!length) return []
  const perpendicularX = -dy / length, perpendicularY = dx / length
  const offsets = double ? [-0.8 * zoom, 0.8 * zoom] : [0]
  const lines: DrawCommand[] = []
  for (const offset of offsets) {
    const startX = x1 + perpendicularX * offset, startY = y1 + perpendicularY * offset
    if (!pattern.length) {
      lines.push({ kind: 'line', x1: startX, y1: startY, x2: x2 + perpendicularX * offset, y2: y2 + perpendicularY * offset, color, width: width * zoom, clip })
      continue
    }
    const cycle = pattern.reduce((sum, value) => sum + value, 0)
    for (let position = 0, index = 0; position < length; index += 1) {
      const segment = Math.min(pattern[index % pattern.length]!, length - position)
      if (index % 2 === 0 && segment > 0) {
        lines.push({ kind: 'line', x1: startX + dx / length * position, y1: startY + dy / length * position,
          x2: startX + dx / length * (position + segment), y2: startY + dy / length * (position + segment), color, width: width * zoom, clip })
      }
      position += Math.max(segment, Number.EPSILON)
      if (position >= length) break
      if (index > length / cycle * pattern.length + pattern.length) break
    }
  }
  return lines
}

function borderDashPattern(style: string): number[] {
  if (style.includes('dash-dot-dot')) return [4, 2, 1, 2, 1, 2]
  if (style.includes('dash-dot')) return [4, 2, 1, 2]
  if (style.includes('dotted')) return [1, 2]
  if (style.includes('dashed')) return [4, 2]
  return []
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
  if (command.kind === 'polygon') return {
    ...command,
    points: command.points.map((point) => ({ x: point.x * viewport.zoom - viewport.scrollLeft, y: point.y * viewport.zoom - viewport.scrollTop })),
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
