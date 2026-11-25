import type { DocumentBlock } from '../../../domain/model'
import { tableForBlock } from '../model/document-model'
import { documentTableCellTarget, documentTableGrid, type DocumentTableGridCell } from '../model/document-table-grid'
import { documentTableCellBorders, documentTableCellFill, documentTableCellMargin,
  documentTableRowProperties, DOCUMENT_TABLE_MIN_ROW_HEIGHT } from '../model/document-table-style'
import type { ParagraphListMarker } from '../model/document-lists'
import type { DocumentTextLayoutDisplay, LayoutCell, LayoutDivider, LayoutImage, LayoutLine, LayoutMeasurer, LayoutRect, LayoutTableRowMetric } from './document-layout-types'
import { measureDocumentTableColumnWidths } from './document-table-intrinsic-widths'
import { documentTableCellContentContainers } from '../model/document-table-continuations'
import { measureDocumentBlockCollections, intersectDocumentRect as intersectTableRect, translateDocumentLine as translateTableLine,
  translateDocumentRect as translateTableRect, type DocumentBlockContentMeasurement } from './document-block-measurement'
import type { LayoutContainerDecoration } from './document-layout-types'

export { intersectDocumentRect as intersectTableRect, translateDocumentLine as translateTableLine, translateDocumentRect as translateTableRect } from './document-block-measurement'

export interface TableVerticalInterval { start: number; end: number; kind: 'line' | 'image' | 'row' }
export interface MeasuredDocumentTable extends LayoutRect {
  tableId: string; depth: number; columnWidths: number[]; rows: LayoutTableRowMetric[]
  clip?: LayoutRect
}
export interface DocumentTableMeasurement {
  width: number; height: number; rowHeights: number[]; rowStarts: number[]; headerHeight: number; headerRowCount: number
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]; tables: MeasuredDocumentTable[]; containers: LayoutContainerDecoration[]
  /** Page cuts cannot pass through visible text lines, images, or nonsplitting rows. */
  indivisible: TableVerticalInterval[]
}
interface CellContent extends DocumentBlockContentMeasurement { owner: DocumentTableGridCell; width: number; x: number }
const EPSILON = 1e-6

function translatedClipped<T extends LayoutRect & { clip?: LayoutRect }>(item: T, dx: number, dy: number, parent: LayoutRect): T {
  const shifted = translateTableRect(item, dx, dy)
  return { ...shifted, clip: item.clip ? intersectTableRect(translateTableRect(item.clip, dx, dy), parent) : { ...parent } }
}
function prefix(values: readonly number[]): number[] {
  const positions = [0]
  for (const value of values) positions.push(positions.at(-1)! + value)
  return positions
}
/** Relative measurement is shared by pagination and table resize controls, without mutating the model. */
export function measureDocumentTable(block: DocumentBlock, available: number, markers: Map<string, ParagraphListMarker>, measurer: LayoutMeasurer, depth = 0, maximumObjectHeight = Number.POSITIVE_INFINITY,
  display?: DocumentTextLayoutDisplay): DocumentTableMeasurement {
  const table = tableForBlock(block)
  if (!table) throw new RangeError('表格模型不存在')
  const grid = documentTableGrid(table), widths = measureDocumentTableColumnWidths(block, available, markers, measurer, display), columns = prefix(widths)
  const contents = grid.anchors.map((owner): CellContent => {
    const margin = documentTableCellMargin(table, owner.cell), width = columns[owner.column + owner.colSpan]! - columns[owner.column]!
    const contentWidth = Math.max(1, width - margin.left - margin.right)
    const content = measureDocumentBlockCollections(documentTableCellContentContainers(owner.cell), { width: contentWidth, markers, measurer,
      measureTable: measureDocumentTable, target: paragraphId => documentTableCellTarget(block.id, owner, paragraphId), tableCellId: owner.cell.id,
      tableDepth: depth + 1, nextTableDepth: depth + 1, maximumObjectHeight,
      maximumImageHeight: maximumObjectHeight - margin.top - margin.bottom, display })
    return { ...content, owner, width, x: columns[owner.column]!, height: content.height + margin.top + margin.bottom }
  })
  const properties = table.rows.map((_row, index) => documentTableRowProperties(table, index))
  const heights = properties.map(row => row.heightRule === 'exact' ? row.height ?? DOCUMENT_TABLE_MIN_ROW_HEIGHT
    : Math.max(DOCUMENT_TABLE_MIN_ROW_HEIGHT, row.heightRule === 'auto' ? 0 : row.height ?? 0))
  for (const content of contents) if (content.owner.rowSpan === 1 && properties[content.owner.row]!.heightRule !== 'exact') {
    heights[content.owner.row] = Math.max(heights[content.owner.row]!, content.height)
  }
  // A rowSpan imposes a constraint on the sum of covered rows. Exact rows keep
  // their physical size; distribute any deficit only over adjustable rows.
  for (const content of contents.filter(content => content.owner.rowSpan > 1).sort((a, b) => a.owner.rowSpan - b.owner.rowSpan)) {
    const indices = Array.from({ length: content.owner.rowSpan }, (_, i) => content.owner.row + i)
    const deficit = content.height - indices.reduce((sum, row) => sum + heights[row]!, 0)
    const adjustable = indices.filter(row => properties[row]!.heightRule !== 'exact')
    if (deficit > EPSILON && adjustable.length) for (const row of adjustable) heights[row] = heights[row]! + deficit / adjustable.length
  }
  let headerRowCount = table.headerRowCount ?? 0
  // Imported header flags can end inside a rowSpan. Repeat its complete
  // merged owner in the view, without changing the persisted declaration.
  while (grid.anchors.some(owner => owner.row < headerRowCount && owner.row + owner.rowSpan > headerRowCount)) {
    headerRowCount = Math.max(headerRowCount, ...grid.anchors.filter(owner => owner.row < headerRowCount).map(owner => owner.row + owner.rowSpan))
  }
  const starts = prefix(heights), result: DocumentTableMeasurement = { width: columns.at(-1)!, height: starts.at(-1)!,
    rowHeights: heights, rowStarts: starts, headerRowCount,
    headerHeight: starts[headerRowCount]!, lines: [], cells: [], images: [], dividers: [], tables: [], containers: [], indivisible: [] }
  result.tables.push({ tableId: block.id, depth, x: 0, y: 0, width: result.width, height: result.height,
    columnWidths: widths, rows: table.rows.map((row, index) => ({ row: index, id: row.id, y: starts[index]!, height: heights[index]!, contentHeight: heights[index]! })) })
  for (const content of contents) {
    const owner = content.owner, margin = documentTableCellMargin(table, owner.cell), y = starts[owner.row]!, height = starts[owner.row + owner.rowSpan]! - y
    const extra = Math.max(0, height - content.height), align = owner.cell.style.verticalAlign ?? 'top'
    const dy = y + margin.top + (align === 'center' ? extra / 2 : align === 'bottom' ? extra : 0), dx = content.x + margin.left
    const clip = { x: dx, y: y + margin.top, width: Math.max(0, content.width - margin.left - margin.right), height: Math.max(0, height - margin.top - margin.bottom) }
    result.cells.push({ x: content.x, y, width: content.width, height, target: documentTableCellTarget(block.id, owner), page: 0,
      fill: documentTableCellFill(table, owner.cell), borders: documentTableCellBorders(grid, owner), tableId: block.id, cellId: owner.cell.id,
      row: owner.row, columnIndex: owner.column, rowSpan: owner.rowSpan, colSpan: owner.colSpan, depth })
    for (const source of content.lines) {
      const line = translateTableLine(source, dx, dy)
      line.clip = line.clip ? intersectTableRect(line.clip, clip) : { ...clip }
      result.lines.push(line)
    }
    result.cells.push(...content.cells.map(item => translatedClipped(item, dx, dy, clip)))
    result.images.push(...content.images.map(item => translatedClipped(item, dx, dy, clip)))
    result.dividers.push(...content.dividers.map(item => translatedClipped(item, dx, dy, clip)))
    result.containers.push(...content.containers.map(item => translatedClipped(item, dx, dy, clip)))
    result.tables.push(...content.tables.map(item => ({ ...translatedClipped(item, dx, dy, clip), rows: item.rows.map(row => ({ ...row, y: row.y + dy })) })))
    for (const interval of content.indivisible) {
      const start = Math.max(clip.y, interval.start + dy), end = Math.min(clip.y + clip.height, interval.end + dy)
      if (end > start + EPSILON) result.indivisible.push({ ...interval, start, end })
    }
  }
  for (const item of [...result.lines, ...result.images]) {
    const visible = item.clip ? intersectTableRect(item, item.clip) : item
    if (visible.width > 0 && visible.height > EPSILON) result.indivisible.push({ start: visible.y, end: visible.y + visible.height, kind: 'glyphs' in item ? 'line' : 'image' })
  }
  properties.forEach((row, index) => { if (!row.allowSplit) result.indivisible.push({ start: starts[index]!, end: starts[index + 1]!, kind: 'row' }) })
  result.indivisible.sort((a, b) => a.start - b.start || a.end - b.end)
  return result
}

/**
 * Staggered cell lines need not share a natural page boundary. Insert derived
 * space before all crossing objects, so every line moves whole to the next
 * fragment. Stretch intersected row/cell rectangles and shift later rows;
 * this affects the layout only, never row settings or text in the document.
 */
export function insertDocumentTablePageGap(measurement: DocumentTableMeasurement, at: number, height: number): void {
  if (height <= EPSILON) return
  const shiftedRect = <T extends LayoutRect>(rect: T): T => rect.y >= at - EPSILON ? translateTableRect(rect, 0, height)
    : rect.y + rect.height > at + EPSILON ? { ...rect, height: rect.height + height } : { ...rect }
  const shiftedClip = <T extends LayoutRect & { clip?: LayoutRect }>(item: T): T => ({ ...shiftedRect(item), ...(item.clip ? { clip: shiftedRect(item.clip) } : {}) })
  measurement.lines = measurement.lines.map(line => {
    const shifted = line.y >= at - EPSILON ? translateTableLine(line, 0, height) : { ...line }
    if (line.clip) shifted.clip = shiftedRect(line.clip)
    return shifted
  })
  measurement.containers = measurement.containers.map(shiftedClip)
  measurement.images = measurement.images.map(image => ({ ...image, ...(image.y >= at - EPSILON ? { y: image.y + height } : {}),
    ...(image.clip ? { clip: shiftedRect(image.clip) } : {}) }))
  measurement.cells = measurement.cells.map(shiftedClip)
  measurement.dividers = measurement.dividers.map(shiftedClip)
  measurement.tables = measurement.tables.map(table => ({ ...shiftedClip(table), rows: table.rows.map(row => shiftedRect({ ...row, x: 0, width: 0 })) }))
  measurement.indivisible = measurement.indivisible.map(interval => interval.start >= at - EPSILON
    ? { ...interval, start: interval.start + height, end: interval.end + height }
    : interval.kind === 'row' && interval.end > at + EPSILON ? { ...interval, end: interval.end + height } : { ...interval })
  const root = measurement.tables[0]!
  measurement.height += height
  measurement.rowHeights = root.rows.map(row => row.height)
  measurement.rowStarts = [...root.rows.map(row => row.y), root.y + root.height]
  measurement.headerHeight = measurement.headerRowCount ? measurement.rowStarts[measurement.headerRowCount]! : 0
  measurement.indivisible.sort((a, b) => a.start - b.start || a.end - b.end)
}
