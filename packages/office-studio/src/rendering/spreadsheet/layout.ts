import type { ExcelSheetView } from '../../services/excelViewAdapter'
import type { RenderRect } from '../contracts'
import { hiddenSheetRows, hiddenSheetColumns, sheetAxisViewHiddenSpans } from '../../features/sheet-workbench/sheet-row-visibility.ts'
import { isSheetCellCoverable } from '../../features/sheet-workbench/sheet-cell-text-layout.ts'
import { getSheetOutlineLevelControls, getSheetOutlineViewModels, type SheetOutlineViewModel } from '../../features/sheet-outline/sheet-outline-view.ts'
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
export interface GridViewport { width: number; height: number; scrollLeft: number; scrollTop: number; zoom: number; headerWidth?: number; headerHeight?: number }
export interface GridRange { startRow: number; startColumn: number; endRow: number; endColumn: number; rangeType?: 0 | 1 | 2 | 3 }
export type GridHit = { kind: 'cell'; row: number; column: number; key: string } | { kind: 'fill-handle'; range: GridRange } | { kind: 'row'; row: number } | { kind: 'column'; column: number } | { kind: 'all' } | { kind: 'resize-row'; row: number; size: number } | { kind: 'resize-column'; column: number; size: number } | { kind: 'outline-toggle'; axis: 'row' | 'column'; groupId: string; collapsed: boolean } | { kind: 'outline-level'; axis: 'row' | 'column'; level: number }
export interface OutlineControlRect extends RenderRect { axis: 'row' | 'column'; groupId: string; collapsed: boolean }
export interface OutlineLevelControlRect extends RenderRect { axis: 'row' | 'column'; level: number; active: boolean }
export const SHEET_OUTLINE_CONTROL_SIZE = 14
const OUTLINE_LANE_SIZE = 18

/** Persistent drawing coordinates retain the base header; these fields describe the viewport host. */
export function gridViewportHeaders(viewport: GridViewport): { width: number; height: number } {
  return { width: viewport.headerWidth ?? GRID_HEADER_WIDTH, height: viewport.headerHeight ?? GRID_HEADER_HEIGHT }
}
export const GRID_HEADER_WIDTH = 44
export const GRID_HEADER_HEIGHT = 26

export function sheetOutlineHeaderExtents(outline: SpreadsheetRenderSheet['outline']): { width: number; height: number } {
  const rows = outline ? getSheetOutlineLevelControls(outline, 'row').length : 0
  const columns = outline ? getSheetOutlineLevelControls(outline, 'column').length : 0
  return { width: GRID_HEADER_WIDTH + (rows ? rows * OUTLINE_LANE_SIZE + 4 : 0),
    height: GRID_HEADER_HEIGHT + (columns ? columns * OUTLINE_LANE_SIZE + 4 : 0) }
}

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
  readonly headerWidth: number
  readonly headerHeight: number
  readonly rowOutlineGutterWidth: number
  readonly columnOutlineGutterHeight: number
  private readonly outlineViews: Record<'row' | 'column', SheetOutlineViewModel[]>
  private readonly outlineLevels: Record<'row' | 'column', ReturnType<typeof getSheetOutlineLevelControls>>
  constructor(sheet: SpreadsheetRenderSheet) {
    this.sheet = sheet
    this.outlineViews = { row: sheet.outline ? getSheetOutlineViewModels(sheet.outline, 'row', sheet.rowCount) : [],
      column: sheet.outline ? getSheetOutlineViewModels(sheet.outline, 'column', sheet.columnCount) : [] }
    this.outlineLevels = { row: sheet.outline ? getSheetOutlineLevelControls(sheet.outline, 'row') : [],
      column: sheet.outline ? getSheetOutlineLevelControls(sheet.outline, 'column') : [] }
    this.rowOutlineGutterWidth = this.outlineLevels.row.length ? this.outlineLevels.row.length * OUTLINE_LANE_SIZE + 4 : 0
    this.columnOutlineGutterHeight = this.outlineLevels.column.length ? this.outlineLevels.column.length * OUTLINE_LANE_SIZE + 4 : 0
    this.headerWidth = GRID_HEADER_WIDTH + this.rowOutlineGutterWidth
    this.headerHeight = GRID_HEADER_HEIGHT + this.columnOutlineGutterHeight
    // Parsed Office sheets carry their default axis sizes. The workbench keeps
    // its wider first column only for sheets without an explicit default.
    const defaultRowHeight = sheet.defaultRowHeight ?? 26
    const defaultColumnWidth = sheet.defaultColumnWidth ?? 112
    const columnWidths = sheet.defaultColumnWidth === undefined && !sheetAxisSpanAt(sheet.columnWidthSpans, 0)
      ? { 0: 152, ...sheet.columnWidths }
      : sheet.columnWidths
    const rowSizes = effectiveSheetRowSizes(sheet)
    this.rows = new AxisIndex(sheet.rowCount, defaultRowHeight, rowSizes.sizes, [...hiddenSheetRows(sheet)], rowSizes.spans, sheetAxisViewHiddenSpans(sheet, 'row'))
    this.columns = new AxisIndex(sheet.columnCount, defaultColumnWidth, columnWidths, [...hiddenSheetColumns(sheet)], sheet.columnWidthSpans, sheetAxisViewHiddenSpans(sheet, 'column'))
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
      x: this.headerWidth + this.columns.offset(startColumn) * scale - (startColumn < this.frozenColumns ? 0 : viewport.scrollLeft),
      y: this.headerHeight + this.rows.offset(startRow) * scale - (startRow < this.frozenRows ? 0 : viewport.scrollTop),
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
        this.headerWidth + this.columns.offset(this.frozenColumns) * viewport.zoom, viewport.width,
        Math.max(0, this.headerWidth + this.columns.offset(this.columns.count) * viewport.zoom - viewport.width),
      ),
      scrollTop: startRow < this.frozenRows ? viewport.scrollTop : reveal(
        viewport.scrollTop, rect.y, rect.height,
        this.headerHeight + this.rows.offset(this.frozenRows) * viewport.zoom, viewport.height,
        Math.max(0, this.headerHeight + this.rows.offset(this.rows.count) * viewport.zoom - viewport.height),
      ),
    }
  }
  hitTest(x: number, y: number, viewport: GridViewport): GridHit | undefined {
    if (x < 0 || y < 0 || x >= viewport.width || y >= viewport.height) return undefined
    if (!this.rows.count || !this.columns.count) return undefined
    for (const control of this.outlineLevelRects()) if (containsPoint(control, x, y)) return { kind: 'outline-level', axis: control.axis, level: control.level }
    for (const control of this.outlineToggleRects(viewport)) if (containsPoint(control, x, y)) return { kind: 'outline-toggle', axis: control.axis, groupId: control.groupId, collapsed: control.collapsed }
    if (x < this.headerWidth && y < this.headerHeight) {
      return x >= this.rowOutlineGutterWidth && y >= this.columnOutlineGutterHeight ? { kind: 'all' } : undefined
    }
    if (x < this.rowOutlineGutterWidth || y < this.columnOutlineGutterHeight) return undefined
    const scale = viewport.zoom
    const frozenWidth = this.columns.offset(this.frozenColumns) * scale, frozenHeight = this.rows.offset(this.frozenRows) * scale
    const columnOffset = (x - this.headerWidth + (x >= this.headerWidth + frozenWidth ? viewport.scrollLeft : 0)) / scale
    const rowOffset = (y - this.headerHeight + (y >= this.headerHeight + frozenHeight ? viewport.scrollTop : 0)) / scale
    // AxisIndex.locate deliberately clamps for visible-range queries. A pointer
    // in the unused canvas area must not select that clamped (possibly hidden)
    // final row or column.
    if (x >= this.headerWidth && (!Number.isFinite(columnOffset) || columnOffset < 0 || columnOffset >= this.columns.offset(this.columns.count))) return undefined
    if (y >= this.headerHeight && (!Number.isFinite(rowOffset) || rowOffset < 0 || rowOffset >= this.rows.offset(this.rows.count))) return undefined
    const column = this.columns.locate(columnOffset)
    const row = this.rows.locate(rowOffset)
    if (y < this.headerHeight) {
      const rect = this.cellRect(0, column, viewport, false)
      if (Math.abs(x - rect.x - rect.width) < 4) return { kind: 'resize-column', column, size: this.columns.size(column) }
      return { kind: 'column', column }
    }
    if (x < this.headerWidth) {
      const rect = this.cellRect(row, 0, viewport, false)
      if (Math.abs(y - rect.y - rect.height) < 4) return { kind: 'resize-row', row, size: this.rows.size(row) }
      return { kind: 'row', row }
    }
    const merge = this.mergeAt(row, column)
    const targetRow = merge?.startRow ?? row, targetColumn = merge?.startColumn ?? column
    return { kind: 'cell', row: targetRow, column: targetColumn, key: cellAddress(targetRow, targetColumn) }
  }
  outlineLevelRects(): OutlineLevelControlRect[] {
    return (['row', 'column'] as const).flatMap(axis => this.outlineLevels[axis].map(control => ({
      ...control, x: axis === 'row' ? 2 + (control.level - 1) * OUTLINE_LANE_SIZE : this.rowOutlineGutterWidth + 14,
      y: axis === 'row' ? this.columnOutlineGutterHeight + 6 : 2 + (control.level - 1) * OUTLINE_LANE_SIZE,
      width: SHEET_OUTLINE_CONTROL_SIZE, height: SHEET_OUTLINE_CONTROL_SIZE,
    })))
  }
  outlineToggleRects(viewport: GridViewport): OutlineControlRect[] {
    const rects: OutlineControlRect[] = []
    for (const axis of ['row', 'column'] as const) {
      const index = axis === 'row' ? this.rows : this.columns
      const header = axis === 'row' ? this.headerHeight : this.headerWidth
      const extent = axis === 'row' ? viewport.height : viewport.width
      const frozen = axis === 'row' ? this.frozenRows : this.frozenColumns
      const scroll = axis === 'row' ? viewport.scrollTop : viewport.scrollLeft
      for (const view of this.outlineViews[axis]) {
        let center: number
        if (view.anchor !== null && index.size(view.anchor) > 0) {
          center = header + (index.offset(view.anchor) + index.size(view.anchor) / 2) * viewport.zoom
            - (view.anchor < frozen ? 0 : scroll)
        } else {
          const boundary = header + index.offset(view.anchorBoundary) * viewport.zoom
            - (view.anchorBoundary <= frozen ? 0 : scroll)
          center = view.summaryAtStart ? boundary + 2 + SHEET_OUTLINE_CONTROL_SIZE / 2
            : boundary - 2 - SHEET_OUTLINE_CONTROL_SIZE / 2
          // An outline spanning the whole axis has no adjacent summary cell.
          // Keep its collapsed boundary reachable even when every detail is hidden.
          if (view.collapsed) center = Math.max(header + 2 + SHEET_OUTLINE_CONTROL_SIZE / 2, center)
        }
        const position = center - SHEET_OUTLINE_CONTROL_SIZE / 2
        if (position < header || position + SHEET_OUTLINE_CONTROL_SIZE > extent) continue
        const lane = 2 + (view.depth - 1) * OUTLINE_LANE_SIZE
        rects.push({ axis, groupId: view.id, collapsed: view.collapsed,
          x: axis === 'row' ? lane : position, y: axis === 'row' ? position : lane,
          width: SHEET_OUTLINE_CONTROL_SIZE, height: SHEET_OUTLINE_CONTROL_SIZE })
      }
    }
    return rects
  }
  outlineBracketLines(viewport: GridViewport): Array<{ axis: 'row' | 'column'; x1: number; y1: number; x2: number; y2: number; groupId: string }> {
    const lines: Array<{ axis: 'row' | 'column'; x1: number; y1: number; x2: number; y2: number; groupId: string }> = []
    for (const axis of ['row', 'column'] as const) {
      const index = axis === 'row' ? this.rows : this.columns
      const header = axis === 'row' ? this.headerHeight : this.headerWidth
      const extent = axis === 'row' ? viewport.height : viewport.width
      const frozen = axis === 'row' ? this.frozenRows : this.frozenColumns
      const scroll = axis === 'row' ? viewport.scrollTop : viewport.scrollLeft
      const frozenEnd = Math.min(extent, header + index.offset(frozen) * viewport.zoom)
      for (const view of this.outlineViews[axis]) {
        if (view.collapsed) continue
        const lane = 9 + (view.depth - 1) * OUTLINE_LANE_SIZE
        const add = (start: number, end: number, offset: number, low: number, high: number) => {
          if (end < start || high <= low) return
          const rawStart = header + index.offset(start) * viewport.zoom - offset
          const rawEnd = header + index.offset(end + 1) * viewport.zoom - offset
          const first = Math.max(low, rawStart), last = Math.min(high, rawEnd)
          if (last <= first) return
          const line = (a: number, b: number, tick = false) => lines.push({ axis, groupId: view.id,
            x1: axis === 'row' ? lane : a, y1: axis === 'row' ? a : lane,
            x2: axis === 'row' ? lane + (tick ? 4 : 0) : b, y2: axis === 'row' ? b : lane + (tick ? 4 : 0) })
          line(first, last)
          if (rawStart >= low && start === view.start) line(first, first, true)
          if (rawEnd <= high && end === view.end) line(last, last, true)
        }
        add(view.start, Math.min(view.end, frozen - 1), 0, header, frozenEnd)
        add(Math.max(view.start, frozen), view.end, scroll, frozenEnd, extent)
      }
    }
    return lines
  }
  panes(viewport: GridViewport): Array<{ rows: number[]; columns: number[]; clip: RenderRect }> {
    const z = viewport.zoom, fw = Math.max(0, Math.min(viewport.width - this.headerWidth, this.columns.offset(this.frozenColumns) * z)), fh = Math.max(0, Math.min(viewport.height - this.headerHeight, this.rows.offset(this.frozenRows) * z))
    const frozenRows = this.rows.visible(0, (viewport.height - this.headerHeight) / z, 0, this.frozenRows)
    const frozenColumns = this.columns.visible(0, (viewport.width - this.headerWidth) / z, 0, this.frozenColumns)
    const rows = this.rows.visible((viewport.scrollTop + fh) / z, (viewport.scrollTop + viewport.height - this.headerHeight) / z, this.frozenRows)
    const columns = this.columns.visible((viewport.scrollLeft + fw) / z, (viewport.scrollLeft + viewport.width - this.headerWidth) / z, this.frozenColumns)
    return [
      { rows, columns, clip: { x: this.headerWidth + fw, y: this.headerHeight + fh, width: viewport.width - this.headerWidth - fw, height: viewport.height - this.headerHeight - fh } },
      { rows: frozenRows, columns, clip: { x: this.headerWidth + fw, y: this.headerHeight, width: viewport.width - this.headerWidth - fw, height: fh } },
      { rows, columns: frozenColumns, clip: { x: this.headerWidth, y: this.headerHeight + fh, width: fw, height: viewport.height - this.headerHeight - fh } },
      { rows: frozenRows, columns: frozenColumns, clip: { x: this.headerWidth, y: this.headerHeight, width: fw, height: fh } },
    ].map(pane => ({ ...pane, clip: { ...pane.clip, width: Math.max(0, pane.clip.width), height: Math.max(0, pane.clip.height) } }))
  }
}

function containsPoint(rect: RenderRect, x: number, y: number): boolean {
  return x >= rect.x && x < rect.x + rect.width && y >= rect.y && y < rect.y + rect.height
}
