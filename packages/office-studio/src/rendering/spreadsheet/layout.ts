import type { ExcelSheetView } from '../../services/excelViewAdapter'
import type { RenderRect } from '../contracts'
import { hiddenSheetRows } from '../../features/sheet-workbench/sheet-row-visibility.ts'
import { isSheetCellCoverable } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'
import { effectiveSheetRowSizes, sheetAxisSpanAt, SparseAxisIndex as AxisIndex } from '../../domain/sheet-axis.ts'
export { SparseAxisIndex as AxisIndex } from '../../domain/sheet-axis.ts'

export type SpreadsheetRenderSheet = ExcelSheetView & {
  columnWidths?: Record<number, number>
  rowHeights?: Record<number, number>
  hiddenRows?: number[]
  hiddenColumns?: number[]
  outlineHiddenColumns?: number[]
  showGridlines?: boolean
}
export interface GridViewport { width: number; height: number; scrollLeft: number; scrollTop: number; zoom: number }
export interface GridRange { startRow: number; startColumn: number; endRow: number; endColumn: number; rangeType?: 0 | 1 | 2 | 3 }
export type GridHit = { kind: 'cell'; row: number; column: number; key: string } | { kind: 'fill-handle'; range: GridRange } | { kind: 'row'; row: number } | { kind: 'column'; column: number } | { kind: 'all' } | { kind: 'resize-row'; row: number; size: number } | { kind: 'resize-column'; column: number; size: number } | { kind: 'outline-toggle'; axis: 'row' | 'column'; groupId: string; collapsed: boolean } | { kind: 'outline-level'; axis: 'row' | 'column'; level: number }
export interface OutlineControlRect extends RenderRect { axis: 'row' | 'column'; groupId: string; collapsed: boolean }
export const GRID_HEADER_WIDTH = 44
export const GRID_HEADER_HEIGHT = 26

interface MergeEntry { range: GridRange; order: number }
interface MergeBand {
  startRow: number
  endRow: number
  columns: MergeEntry[]
  hasColumnOverlap: boolean
}
interface MergeRowNode {
  center: number
  byStart: MergeBand[]
  byEnd: MergeBand[]
  left?: MergeRowNode
  right?: MergeRowNode
}

function lowerBound(values: readonly number[], value: number): number {
  let low = 0, high = values.length
  while (low < high) {
    const middle = (low + high) >>> 1
    if (values[middle]! < value) low = middle + 1
    else high = middle
  }
  return low
}

function buildMergeRowIndex(bands: MergeBand[]): MergeRowNode | undefined {
  if (!bands.length) return undefined
  const ordered = [...bands].sort((a, b) => a.startRow - b.startRow || a.endRow - b.endRow)
  const center = ordered[Math.floor(ordered.length / 2)]!.startRow
  const crossing: MergeBand[] = [], left: MergeBand[] = [], right: MergeBand[] = []
  for (const band of ordered) {
    if (band.endRow < center) left.push(band)
    else if (band.startRow > center) right.push(band)
    else crossing.push(band)
  }
  return {
    center,
    byStart: crossing.sort((a, b) => a.startRow - b.startRow || a.endRow - b.endRow),
    byEnd: [...crossing].sort((a, b) => b.endRow - a.endRow || a.startRow - b.startRow),
    left: buildMergeRowIndex(left),
    right: buildMergeRowIndex(right),
  }
}

function findBandMerge(band: MergeBand, row: number, column: number): MergeEntry | undefined {
  if (band.hasColumnOverlap) {
    return band.columns.find(({ range }) => column >= range.startColumn && column <= range.endColumn)
  }
  let low = 0, high = band.columns.length
  while (low < high) {
    const middle = (low + high) >>> 1
    if (band.columns[middle]!.range.startColumn <= column) low = middle + 1
    else high = middle
  }
  const candidate = band.columns[low - 1]
  return candidate && column <= candidate.range.endColumn && row >= candidate.range.startRow && row <= candidate.range.endRow
    ? candidate
    : undefined
}

function visitMergeBands(node: MergeRowNode | undefined, row: number, visit: (band: MergeBand) => void): void {
  if (!node) return
  if (row < node.center) {
    for (const band of node.byStart) {
      if (band.startRow > row) break
      visit(band)
    }
    visitMergeBands(node.left, row, visit)
  } else if (row > node.center) {
    for (const band of node.byEnd) {
      if (band.endRow < row) break
      visit(band)
    }
    visitMergeBands(node.right, row, visit)
  } else {
    for (const band of node.byStart) visit(band)
  }
}

export function columnLabel(index: number): string {
  let label = '', cursor = index
  while (cursor >= 0) { label = String.fromCharCode(cursor % 26 + 65) + label; cursor = Math.floor(cursor / 26) - 1 }
  return label
}
export function cellAddress(row: number, column: number): string { return `${columnLabel(column)}${row + 1}` }
export function parseAddress(address: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address)
  if (!match) return undefined
  let column = 0
  for (const char of match[1].toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

export class SpreadsheetGeometry {
  readonly sheet: SpreadsheetRenderSheet
  readonly rows: AxisIndex
  readonly columns: AxisIndex
  readonly merges: GridRange[]
  private readonly mergeRowIndex: MergeRowNode | undefined
  private readonly contentColumns = new Map<number, number[]>()
  readonly frozenRows: number
  readonly frozenColumns: number
  constructor(sheet: SpreadsheetRenderSheet) {
    this.sheet = sheet
    // Parsed Office sheets carry their default axis sizes. The workbench keeps
    // its wider first column only for sheets without an explicit default.
    const defaultRowHeight = sheet.defaultRowHeight ?? 26
    const defaultColumnWidth = sheet.defaultColumnWidth ?? 112
    const columnWidths = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0)
      ? { 0: 152, ...sheet.columnWidths }
      : sheet.columnWidths
    const rowSizes = effectiveSheetRowSizes(sheet)
    this.rows = new AxisIndex(sheet.rowCount, defaultRowHeight, rowSizes.sizes, [...hiddenSheetRows(sheet)], rowSizes.spans, sheet.rowHiddenSpans)
    this.columns = new AxisIndex(sheet.columnCount, defaultColumnWidth, columnWidths, [...(sheet.hiddenColumns ?? []), ...(sheet.outlineHiddenColumns ?? [])], sheet.columnWidthSpans, sheet.columnHiddenSpans)
    this.frozenRows = Math.max(0, Math.min(sheet.rowCount, sheet.freeze?.rows ?? 0))
    this.frozenColumns = Math.max(0, Math.min(sheet.columnCount, sheet.freeze?.columns ?? 0))
    this.merges = (sheet.merges ?? []).flatMap(value => { const [a, b] = value.split(':'); const start = parseAddress(a), end = parseAddress(b ?? a); return start && end ? [{ startRow: Math.min(start.row, end.row), startColumn: Math.min(start.column, end.column), endRow: Math.max(start.row, end.row), endColumn: Math.max(start.column, end.column) }] : [] })
    const bands = new Map<string, MergeBand>()
    this.merges.forEach((range, order) => {
      const key = `${range.startRow}:${range.endRow}`
      let band = bands.get(key)
      if (!band) {
        band = { startRow: range.startRow, endRow: range.endRow, columns: [], hasColumnOverlap: false }
        bands.set(key, band)
      }
      band.columns.push({ range, order })
    })
    for (const band of bands.values()) {
      band.columns.sort((a, b) => a.range.startColumn - b.range.startColumn || a.order - b.order)
      for (let index = 1; index < band.columns.length; index += 1) {
        if (band.columns[index - 1]!.range.endColumn >= band.columns[index]!.range.startColumn) {
          band.hasColumnOverlap = true
          band.columns.sort((a, b) => a.order - b.order)
          break
        }
      }
    }
    this.mergeRowIndex = buildMergeRowIndex([...bands.values()])
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const position = parseAddress(address)
      if (!position || position.row >= this.rows.count || position.column >= this.columns.count
        || this.rows.size(position.row) <= 0 || this.columns.size(position.column) <= 0 || isSheetCellCoverable(cell)) continue
      const columns = this.contentColumns.get(position.row) ?? []
      columns.push(position.column)
      this.contentColumns.set(position.row, columns)
    }
    for (const columns of this.contentColumns.values()) columns.sort((a, b) => a - b)
  }
  mergeAt(row: number, column: number): GridRange | undefined {
    let match: MergeEntry | undefined
    visitMergeBands(this.mergeRowIndex, row, band => {
      const candidate = findBandMerge(band, row, column)
      if (candidate && (!match || candidate.order < match.order)) match = candidate
    })
    return match?.range
  }
  /** Nearest visible content anchors can be outside the viewport; empty styled cells remain coverable. */
  contentColumnBefore(row: number, column: number): number | undefined {
    const columns = this.contentColumns.get(row) ?? []
    return columns[lowerBound(columns, column) - 1]
  }
  contentColumnAfter(row: number, column: number): number | undefined {
    const columns = this.contentColumns.get(row) ?? []
    return columns[lowerBound(columns, column + 1)]
  }
  /** Content, merged ranges and frozen-column boundaries delimit a text overflow lane. */
  overflowColumnBounds(row: number, column: number): { startColumn: number; endColumn: number } {
    const minimum = column < this.frozenColumns ? 0 : this.frozenColumns
    const maximum = column < this.frozenColumns ? this.frozenColumns - 1 : this.columns.count - 1
    const left = this.contentColumnBefore(row, column), right = this.contentColumnAfter(row, column)
    let startColumn = Math.max(minimum, left === undefined ? minimum : left + 1)
    let endColumn = Math.min(maximum, right === undefined ? maximum : right - 1)
    const inspect = ({ range }: MergeEntry) => {
      if (this.columns.offset(range.endColumn + 1) <= this.columns.offset(range.startColumn)) return
      if (range.endColumn < column) startColumn = Math.max(startColumn, range.endColumn + 1)
      else if (range.startColumn > column) endColumn = Math.min(endColumn, range.startColumn - 1)
      else { startColumn = column; endColumn = column }
    }
    visitMergeBands(this.mergeRowIndex, row, band => {
      if (band.hasColumnOverlap) { for (const entry of band.columns) inspect(entry); return }
      let low = 0, high = band.columns.length
      while (low < high) {
        const middle = (low + high) >>> 1
        if (band.columns[middle]!.range.startColumn <= column) low = middle + 1
        else high = middle
      }
      for (let index = low - 1; index >= 0; index--) {
        const entry = band.columns[index]!
        if (this.columns.offset(entry.range.endColumn + 1) <= this.columns.offset(entry.range.startColumn)) continue
        inspect(entry); break
      }
      for (let index = low; index < band.columns.length; index++) {
        const entry = band.columns[index]!
        if (this.columns.offset(entry.range.endColumn + 1) <= this.columns.offset(entry.range.startColumn)) continue
        inspect(entry); break
      }
    })
    return { startColumn, endColumn }
  }
  cellRect(row: number, column: number, viewport: GridViewport, merged = true): RenderRect {
    const range = merged ? this.mergeAt(row, column) : undefined
    const startRow = range?.startRow ?? row, startColumn = range?.startColumn ?? column
    const endRow = range?.endRow ?? row, endColumn = range?.endColumn ?? column
    const scale = viewport.zoom
    return {
      x: GRID_HEADER_WIDTH + this.columns.offset(startColumn) * scale - (startColumn < this.frozenColumns ? 0 : viewport.scrollLeft),
      y: GRID_HEADER_HEIGHT + this.rows.offset(startRow) * scale - (startRow < this.frozenRows ? 0 : viewport.scrollTop),
      width: (this.columns.offset(endColumn + 1) - this.columns.offset(startColumn)) * scale,
      height: (this.rows.offset(endRow + 1) - this.rows.offset(startRow)) * scale,
    }
  }
  /** Scroll offsets in CSS pixels, shared by keyboard navigation and the input bridge. */
  scrollForCell(row: number, column: number, viewport: GridViewport): { scrollLeft: number; scrollTop: number } {
    const current = { scrollLeft: viewport.scrollLeft, scrollTop: viewport.scrollTop }
    if (row < 0 || column < 0 || row >= this.rows.count || column >= this.columns.count || viewport.zoom <= 0 || !Number.isFinite(viewport.zoom)) return current
    const merge = this.mergeAt(row, column)
    const startRow = merge?.startRow ?? row, startColumn = merge?.startColumn ?? column
    const rect = this.cellRect(row, column, viewport)
    const reveal = (offset: number, position: number, size: number, leading: number, trailing: number, maximum: number) => {
      if (trailing <= leading || size <= 0) return offset
      // A wide column/tall row (including a merged range) cannot fit entirely.
      // Reveal its beginning instead of scrolling to the far edge and hiding
      // the active cell's input position behind the frozen panes.
      const next = size >= trailing - leading || position < leading
        ? offset + position - leading
        : position + size > trailing ? offset + position + size - trailing : offset
      return Math.max(0, Math.min(maximum, next))
    }
    return {
      scrollLeft: startColumn < this.frozenColumns ? viewport.scrollLeft : reveal(
        viewport.scrollLeft, rect.x, rect.width,
        GRID_HEADER_WIDTH + this.columns.offset(this.frozenColumns) * viewport.zoom, viewport.width,
        Math.max(0, GRID_HEADER_WIDTH + this.columns.offset(this.columns.count) * viewport.zoom - viewport.width),
      ),
      scrollTop: startRow < this.frozenRows ? viewport.scrollTop : reveal(
        viewport.scrollTop, rect.y, rect.height,
        GRID_HEADER_HEIGHT + this.rows.offset(this.frozenRows) * viewport.zoom, viewport.height,
        Math.max(0, GRID_HEADER_HEIGHT + this.rows.offset(this.rows.count) * viewport.zoom - viewport.height),
      ),
    }
  }
  hitTest(x: number, y: number, viewport: GridViewport): GridHit | undefined {
    if (x < 0 || y < 0 || x >= viewport.width || y >= viewport.height) return undefined
    if (!this.rows.count || !this.columns.count) return undefined
    if (x < GRID_HEADER_WIDTH && y < GRID_HEADER_HEIGHT) {
      const rowLevel = this.outlineLevelAt('row', x, y)
      if (rowLevel) return rowLevel
      const columnLevel = this.outlineLevelAt('column', x, y)
      if (columnLevel) return columnLevel
      for (const control of this.outlineToggleRects(viewport)) if (containsPoint(control, x, y)) return { kind: 'outline-toggle', axis: control.axis, groupId: control.groupId, collapsed: control.collapsed }
      return { kind: 'all' }
    }
    const scale = viewport.zoom
    const frozenWidth = this.columns.offset(this.frozenColumns) * scale, frozenHeight = this.rows.offset(this.frozenRows) * scale
    const columnOffset = (x - GRID_HEADER_WIDTH + (x >= GRID_HEADER_WIDTH + frozenWidth ? viewport.scrollLeft : 0)) / scale
    const rowOffset = (y - GRID_HEADER_HEIGHT + (y >= GRID_HEADER_HEIGHT + frozenHeight ? viewport.scrollTop : 0)) / scale
    // AxisIndex.locate deliberately clamps for visible-range queries. A pointer
    // in the unused canvas area must not select that clamped (possibly hidden)
    // final row or column.
    if (x >= GRID_HEADER_WIDTH && (!Number.isFinite(columnOffset) || columnOffset < 0 || columnOffset >= this.columns.offset(this.columns.count))) return undefined
    if (y >= GRID_HEADER_HEIGHT && (!Number.isFinite(rowOffset) || rowOffset < 0 || rowOffset >= this.rows.offset(this.rows.count))) return undefined
    const column = this.columns.locate(columnOffset)
    const row = this.rows.locate(rowOffset)
    if (y < GRID_HEADER_HEIGHT) {
      for (const control of this.outlineToggleRects(viewport)) if (control.axis === 'column' && containsPoint(control, x, y)) return { kind: 'outline-toggle', axis: control.axis, groupId: control.groupId, collapsed: control.collapsed }
      const rect = this.cellRect(0, column, viewport, false)
      if (Math.abs(x - rect.x - rect.width) < 4) return { kind: 'resize-column', column, size: this.columns.size(column) }
      return { kind: 'column', column }
    }
    if (x < GRID_HEADER_WIDTH) {
      for (const control of this.outlineToggleRects(viewport)) if (control.axis === 'row' && containsPoint(control, x, y)) return { kind: 'outline-toggle', axis: control.axis, groupId: control.groupId, collapsed: control.collapsed }
      const rect = this.cellRect(row, 0, viewport, false)
      if (Math.abs(y - rect.y - rect.height) < 4) return { kind: 'resize-row', row, size: this.rows.size(row) }
      return { kind: 'row', row }
    }
    const merge = this.mergeAt(row, column)
    const targetRow = merge?.startRow ?? row, targetColumn = merge?.startColumn ?? column
    return { kind: 'cell', row: targetRow, column: targetColumn, key: cellAddress(targetRow, targetColumn) }
  }
  outlineToggleRects(viewport: GridViewport): OutlineControlRect[] {
    const outline = this.sheet.outline
    if (!outline) return []
    const rects: OutlineControlRect[] = []
    const rowGroups = outline.groups.filter(group => group.axis === 'row')
    for (const group of rowGroups) {
      const summaryRow = outline.summaryAbove ? group.start - 1 : group.end + 1
      if (summaryRow < 0 || summaryRow >= this.rows.count || this.rows.size(summaryRow) <= 0) continue
      const depth = groupDepth(group, rowGroups)
      const rect = this.cellRect(summaryRow, 0, viewport, false)
      rects.push({ axis: 'row', groupId: group.id, collapsed: group.collapsed, x: Math.max(1, 2 + depth * 8), y: rect.y + (rect.height - 12) / 2, width: 12, height: 12 })
    }
    const columnGroups = outline.groups.filter(group => group.axis === 'column')
    for (const group of columnGroups) {
      const summaryColumn = outline.summaryLeft ? group.start - 1 : group.end + 1
      if (summaryColumn < 0 || summaryColumn >= this.columns.count || this.columns.size(summaryColumn) <= 0) continue
      const depth = groupDepth(group, columnGroups)
      const rect = this.cellRect(0, summaryColumn, viewport, false)
      rects.push({ axis: 'column', groupId: group.id, collapsed: group.collapsed, x: rect.x + rect.width / 2 - 6, y: Math.max(1, GRID_HEADER_HEIGHT - 13 - depth), width: 12, height: 12 })
    }
    return rects
  }
  outlineBracketLines(viewport: GridViewport): Array<{ axis: 'row' | 'column'; x1: number; y1: number; x2: number; y2: number; groupId: string }> {
    const outline = this.sheet.outline
    if (!outline) return []
    const lines: Array<{ axis: 'row' | 'column'; x1: number; y1: number; x2: number; y2: number; groupId: string }> = []
    const rowGroups = outline.groups.filter(group => group.axis === 'row')
    for (const group of rowGroups) {
      const depth = groupDepth(group, rowGroups), x = 17 + depth * 8
      const first = this.cellRect(group.start, 0, viewport, false), last = this.cellRect(group.end, 0, viewport, false)
      lines.push({ axis: 'row', groupId: group.id, x1: x, y1: first.y + first.height / 2, x2: x, y2: last.y + last.height / 2 })
      lines.push({ axis: 'row', groupId: group.id, x1: x, y1: first.y + first.height / 2, x2: x + 4, y2: first.y + first.height / 2 })
      lines.push({ axis: 'row', groupId: group.id, x1: x, y1: last.y + last.height / 2, x2: x + 4, y2: last.y + last.height / 2 })
    }
    const columnGroups = outline.groups.filter(group => group.axis === 'column')
    for (const group of columnGroups) {
      const depth = groupDepth(group, columnGroups), y = GRID_HEADER_HEIGHT - 2 - depth * 4
      const first = this.cellRect(0, group.start, viewport, false), last = this.cellRect(0, group.end, viewport, false)
      lines.push({ axis: 'column', groupId: group.id, x1: first.x + first.width / 2, y1: y, x2: last.x + last.width / 2, y2: y })
      lines.push({ axis: 'column', groupId: group.id, x1: first.x + first.width / 2, y1: y, x2: first.x + first.width / 2, y2: y - 4 })
      lines.push({ axis: 'column', groupId: group.id, x1: last.x + last.width / 2, y1: y, x2: last.x + last.width / 2, y2: y - 4 })
    }
    return lines
  }
  private outlineLevelAt(axis: 'row' | 'column', x: number, y: number): GridHit | undefined {
    const groups = this.sheet.outline?.groups.filter(group => group.axis === axis) ?? []
    if (!groups.length) return undefined
    const maxLevel = Math.max(...groups.map(group => groupDepth(group, groups))) + 2
    const level = Math.floor((x - 1) / 12) + 1
    const inBand = axis === 'row' ? y >= 1 && y < 13 : y >= 13 && y < 25
    if (!inBand || level > maxLevel) return undefined
    return { kind: 'outline-level', axis, level }
  }
  panes(viewport: GridViewport): Array<{ rows: number[]; columns: number[]; clip: RenderRect }> {
    const z = viewport.zoom, fw = Math.max(0, Math.min(viewport.width - GRID_HEADER_WIDTH, this.columns.offset(this.frozenColumns) * z)), fh = Math.max(0, Math.min(viewport.height - GRID_HEADER_HEIGHT, this.rows.offset(this.frozenRows) * z))
    const frozenRows = this.rows.visible(0, (viewport.height - GRID_HEADER_HEIGHT) / z, 0, this.frozenRows)
    const frozenColumns = this.columns.visible(0, (viewport.width - GRID_HEADER_WIDTH) / z, 0, this.frozenColumns)
    const rows = this.rows.visible((viewport.scrollTop + fh) / z, (viewport.scrollTop + viewport.height - GRID_HEADER_HEIGHT) / z, this.frozenRows)
    const columns = this.columns.visible((viewport.scrollLeft + fw) / z, (viewport.scrollLeft + viewport.width - GRID_HEADER_WIDTH) / z, this.frozenColumns)
    return [
      { rows, columns, clip: { x: GRID_HEADER_WIDTH + fw, y: GRID_HEADER_HEIGHT + fh, width: viewport.width - GRID_HEADER_WIDTH - fw, height: viewport.height - GRID_HEADER_HEIGHT - fh } },
      { rows: frozenRows, columns, clip: { x: GRID_HEADER_WIDTH + fw, y: GRID_HEADER_HEIGHT, width: viewport.width - GRID_HEADER_WIDTH - fw, height: fh } },
      { rows, columns: frozenColumns, clip: { x: GRID_HEADER_WIDTH, y: GRID_HEADER_HEIGHT + fh, width: fw, height: viewport.height - GRID_HEADER_HEIGHT - fh } },
      { rows: frozenRows, columns: frozenColumns, clip: { x: GRID_HEADER_WIDTH, y: GRID_HEADER_HEIGHT, width: fw, height: fh } },
    ].map(pane => ({ ...pane, clip: { ...pane.clip, width: Math.max(0, pane.clip.width), height: Math.max(0, pane.clip.height) } }))
  }
}

function groupDepth(group: { start: number; end: number }, groups: readonly { start: number; end: number }[]): number {
  return groups.filter(parent => parent.start <= group.start && parent.end >= group.end && (parent.start < group.start || parent.end > group.end)).length
}

function containsPoint(rect: RenderRect, x: number, y: number): boolean {
  return x >= rect.x && x < rect.x + rect.width && y >= rect.y && y < rect.y + rect.height
}
