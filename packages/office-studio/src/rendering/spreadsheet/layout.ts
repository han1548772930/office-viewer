import type { ExcelSheetView } from '../../services/excelViewAdapter'
import type { RenderRect } from '../contracts'
import { hiddenSheetRows } from '../../features/sheet-workbench/sheet-row-visibility.ts'

export type SpreadsheetRenderSheet = ExcelSheetView & {
  columnWidths?: Record<number, number>
  rowHeights?: Record<number, number>
  hiddenRows?: number[]
  hiddenColumns?: number[]
  showGridlines?: boolean
}
export interface GridViewport { width: number; height: number; scrollLeft: number; scrollTop: number; zoom: number }
export interface GridRange { startRow: number; startColumn: number; endRow: number; endColumn: number }
export type GridHit = { kind: 'cell'; row: number; column: number; key: string } | { kind: 'row'; row: number } | { kind: 'column'; column: number } | { kind: 'all' } | { kind: 'resize-row'; row: number; size: number } | { kind: 'resize-column'; column: number; size: number }
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

/** Sparse prefix-size index: storage is proportional to custom/hidden axes, not sheet size. */
export class AxisIndex {
  readonly count: number
  readonly defaultSize: number
  private readonly indices: number[]
  private readonly adjustments: number[]
  private readonly values: Map<number, number>
  constructor(count: number, defaultSize: number, sizes: Record<number, number> = {}, hidden: readonly number[] = []) {
    this.count = Number.isFinite(count) ? Math.max(0, Math.floor(count)) : 0
    this.defaultSize = Number.isFinite(defaultSize) ? Math.max(0, defaultSize) : 0
    this.values = new Map(Object.entries(sizes).filter(([key, value]) => Number.isInteger(Number(key)) && Number(key) >= 0 && Number(key) < this.count && Number.isFinite(value)).map(([key, value]) => [Number(key), Math.max(0, value)]))
    for (const index of hidden) if (Number.isInteger(index) && index >= 0 && index < this.count) this.values.set(index, 0)
    this.indices = [...this.values.keys()].filter(index => index >= 0 && index < this.count).sort((a, b) => a - b)
    this.adjustments = [0]
    for (const index of this.indices) this.adjustments.push(this.adjustments[this.adjustments.length - 1] + this.values.get(index)! - this.defaultSize)
  }
  size(index: number): number { return this.values.get(index) ?? this.defaultSize }
  offset(index: number): number {
    let low = 0, high = this.indices.length
    while (low < high) { const middle = (low + high) >>> 1; if (this.indices[middle] < index) low = middle + 1; else high = middle }
    return index * this.defaultSize + this.adjustments[low]
  }
  locate(coordinate: number): number {
    let low = 0, high = this.count
    while (low < high) { const middle = (low + high) >>> 1; if (this.offset(middle + 1) <= coordinate) low = middle + 1; else high = middle }
    return Math.min(this.count - 1, Math.max(0, low))
  }
  visible(start: number, end: number, minimum = 0, maximum = this.count): number[] {
    const result: number[] = []
    let index = Math.max(minimum, this.locate(Math.max(0, start)))
    while (index < Math.min(maximum, this.count) && this.offset(index) < end) { if (this.size(index) > 0) result.push(index); index += 1 }
    return result
  }
}

export class SpreadsheetGeometry {
  readonly sheet: SpreadsheetRenderSheet
  readonly rows: AxisIndex
  readonly columns: AxisIndex
  readonly merges: GridRange[]
  private readonly mergeRowIndex: MergeRowNode | undefined
  readonly frozenRows: number
  readonly frozenColumns: number
  constructor(sheet: SpreadsheetRenderSheet) {
    this.sheet = sheet
    // Parsed Office sheets carry their default axis sizes. The workbench keeps
    // its wider first column only for sheets without an explicit default.
    const defaultRowHeight = sheet.defaultRowHeight ?? 26
    const defaultColumnWidth = sheet.defaultColumnWidth ?? 112
    const columnWidths = sheet.defaultColumnWidth === undefined
      ? { 0: 152, ...sheet.columnWidths }
      : sheet.columnWidths
    this.rows = new AxisIndex(sheet.rowCount, defaultRowHeight, sheet.rowHeights, [...hiddenSheetRows(sheet)])
    this.columns = new AxisIndex(sheet.columnCount, defaultColumnWidth, columnWidths, sheet.hiddenColumns)
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
  }
  mergeAt(row: number, column: number): GridRange | undefined {
    let match: MergeEntry | undefined
    visitMergeBands(this.mergeRowIndex, row, band => {
      const candidate = findBandMerge(band, row, column)
      if (candidate && (!match || candidate.order < match.order)) match = candidate
    })
    return match?.range
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
    if (x < GRID_HEADER_WIDTH && y < GRID_HEADER_HEIGHT) return { kind: 'all' }
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
      const rect = this.cellRect(0, column, viewport, false)
      if (Math.abs(x - rect.x - rect.width) < 4) return { kind: 'resize-column', column, size: this.columns.size(column) }
      return { kind: 'column', column }
    }
    if (x < GRID_HEADER_WIDTH) {
      const rect = this.cellRect(row, 0, viewport, false)
      if (Math.abs(y - rect.y - rect.height) < 4) return { kind: 'resize-row', row, size: this.rows.size(row) }
      return { kind: 'row', row }
    }
    const merge = this.mergeAt(row, column)
    const targetRow = merge?.startRow ?? row, targetColumn = merge?.startColumn ?? column
    return { kind: 'cell', row: targetRow, column: targetColumn, key: cellAddress(targetRow, targetColumn) }
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
