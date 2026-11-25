import type { DocumentBlock } from '../../../domain/model'
import { tableForBlock } from '../model/document-model'
import { documentTableHorizontalOffset } from '../model/document-table-style'
import type { ParagraphListMarker } from '../model/document-lists'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutMeasurer, LayoutRect } from './document-layout-types'
import type { DocumentSectionFlow } from './document-section-flow'
import { insertDocumentTablePageGap, intersectTableRect, measureDocumentTable, translateTableLine, translateTableRect,
  type DocumentTableMeasurement } from './document-table-measurement'

const EPSILON = 1e-6

function visiblePart(item: LayoutRect & { clip?: LayoutRect }): LayoutRect {
  return item.clip ? intersectTableRect(item, item.clip) : item
}

/** Map one measured source strip to a physical page/column, retaining cell owners. */
function appendSlice(measurement: DocumentTableMeasurement, start: number, end: number, flow: DocumentSectionFlow,
  result: DocumentLayout, left: number, repeatedHeader: boolean): void {
  const top = flow.cursor, dx = left, dy = top - start
  const band = { x: left, y: top, width: measurement.width, height: end - start }
  const location = { ...flow.location, page: flow.page.index }
  for (const cell of measurement.cells) {
    const from = Math.max(start, cell.y), to = Math.min(end, cell.y + cell.height)
    if (to <= from + EPSILON) continue
    const rectangle = { ...cell, x: left + cell.x, y: top + from - start, height: to - from, ...location,
      continuedBefore: from > cell.y + EPSILON, continuedAfter: to < cell.y + cell.height - EPSILON, repeatedHeader }
    if (cell.clip) rectangle.clip = intersectTableRect(translateTableRect(cell.clip, dx, dy), band)
    result.cells.push(rectangle)
  }
  for (const source of measurement.lines) {
    const visible = visiblePart(source)
    if (visible.width <= 0 || visible.height <= 0 || visible.y < start - EPSILON || visible.y + visible.height > end + EPSILON) continue
    const line = translateTableLine(source, dx, dy)
    if (repeatedHeader) line.glyphs = line.glyphs.filter(glyph => glyph.note?.role !== 'bodyLabel')
    Object.assign(line, location)
    line.repeatedHeader = repeatedHeader
    line.clip = line.clip ? intersectTableRect(line.clip, band) : { ...band }
    result.lines.push(line)
  }
  for (const source of measurement.images) {
    const visible = visiblePart(source)
    if (visible.width <= 0 || visible.height <= 0 || visible.y < start - EPSILON || visible.y + visible.height > end + EPSILON) continue
    result.images.push({ ...translateTableRect(source, dx, dy), ...location, repeatedHeader,
      clip: source.clip ? intersectTableRect(translateTableRect(source.clip, dx, dy), band) : { ...band } })
  }
  for (const source of measurement.dividers) {
    if (source.y < start - EPSILON || source.y + source.height > end + EPSILON) continue
    result.dividers.push({ ...translateTableRect(source, dx, dy), ...location,
      clip: source.clip ? intersectTableRect(translateTableRect(source.clip, dx, dy), band) : { ...band } })
  }
  result.containers ??= []
  for (const source of measurement.containers) {
    const from = Math.max(start, source.y), to = Math.min(end, source.y + source.height)
    if (to <= from + EPSILON) continue
    result.containers.push({ ...source, x: left + source.x, y: top + from - start, height: to - from, ...location,
      first: source.first && from <= source.y + EPSILON, last: source.last && to >= source.y + source.height - EPSILON,
      clip: source.clip ? intersectTableRect(translateTableRect(source.clip, dx, dy), band) : { ...band } })
  }
  result.tableFragments ??= []
  for (const source of measurement.tables) {
    const from = Math.max(start, source.y), to = Math.min(end, source.y + source.height)
    if (to <= from + EPSILON) continue
    result.tableFragments.push({ tableId: source.tableId, depth: source.depth, x: left + source.x, y: top + from - start,
      width: source.width, height: to - from, ...location, sourceStart: from - source.y, sourceEnd: to - source.y,
      columnWidths: [...source.columnWidths], rows: source.rows.map(row => ({ ...row, y: top + row.y - start })), repeatedHeader,
      ...(source.clip ? { clip: intersectTableRect(translateTableRect(source.clip, dx, dy), band) } : {}),
      continuedBefore: from > source.y + EPSILON, continuedAfter: to < source.y + source.height - EPSILON })
  }
  flow.consume(end - start)
}

function atomicRoom(measurement: DocumentTableMeasurement, start: number): number {
  const next = measurement.indivisible.find(interval => interval.kind !== 'row' && interval.end > start + EPSILON)
  return next ? Math.max(1, next.end - Math.max(start, next.start)) : 1
}

/**
 * Move whole nonsplitting rows when they fit a fresh column. A row taller than
 * a page must still make progress; its text is fragmented at whole lines.
 * This exceptional Office behavior remains a sample-validation requirement.
 */
function tableCut(measurement: DocumentTableMeasurement, start: number, room: number, fullRoom: number): number {
  let end = Math.min(measurement.height, start + room)
  if (measurement.indivisible.some(interval => interval.kind === 'row' && interval.start <= start + EPSILON
    && interval.end > end + EPSILON && interval.end - start <= fullRoom + EPSILON)) return start
  while (true) {
    const blocked = measurement.indivisible.filter(interval => interval.kind === 'row' && interval.start < end - EPSILON && interval.end > end + EPSILON
      && interval.start > start + EPSILON && interval.end - interval.start <= fullRoom + EPSILON)
    if (!blocked.length) break
    end = Math.min(...blocked.map(interval => interval.start))
  }
  const crossing = measurement.indivisible.filter(interval => interval.kind !== 'row' && interval.start < end - EPSILON && interval.end > end + EPSILON)
  if (crossing.length) {
    const first = Math.min(...crossing.map(interval => interval.start))
    if (first > start + EPSILON) insertDocumentTablePageGap(measurement, first, end - first)
    else return start
  }
  return end
}

export function layoutDocumentTable(block: DocumentBlock, markers: Map<string, ParagraphListMarker>, flow: DocumentSectionFlow,
  result: DocumentLayout, measurer: LayoutMeasurer, display?: DocumentTextLayoutDisplay): void {
  const table = tableForBlock(block)
  if (!table?.rows.length) return
  const offsetY = table.style?.offsetY ?? 0
  if (!Number.isFinite(offsetY)) throw new RangeError('表格纵向偏移必须是有限 CSS px')
  // Offset is retained in the model; a negative derived position is bounded by
  // this section column's top so it cannot paint into another physical page.
  flow.addSpace(Math.max(offsetY, flow.column.y - flow.cursor))
  let available = flow.column.width
  let maximum = flow.fullPageContentHeight
  let measurement = measureDocumentTable(block, available, markers, measurer, 0, maximum, display)
  // Before publishing the first fragment, remeasure if a continuous section
  // sends the beginning of the table to a column with different margins.
  while (true) {
    flow.ensureRoom(Math.min(measurement.rowHeights[0]!, Math.max(1, atomicRoom(measurement, 0))))
    if (Math.abs(available - flow.column.width) <= EPSILON
      && (maximum === flow.fullPageContentHeight || Math.abs(maximum - flow.fullPageContentHeight) <= EPSILON)) break
    available = flow.column.width
    maximum = flow.fullPageContentHeight
    measurement = measureDocumentTable(block, available, markers, measurer, 0, maximum, display)
  }
  let start = 0, continued = false
  while (start < measurement.height - EPSILON) {
    let room = flow.column.y + flow.column.height - flow.cursor
    if (room <= EPSILON) { flow.nextColumn(); continued = true; continue }
    const futureReferences = measurement.lines.flatMap(line => {
      if (line.y + line.height <= start + EPSILON || line.y >= start + room - EPSILON) return []
      const floor = flow.referencePageFloor(line.glyphs.filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id))
      if (floor <= flow.page.index) return []
      const row = measurement.indivisible.find(interval => interval.kind === 'row' && interval.start <= line.y + EPSILON && interval.end >= line.y + line.height - EPSILON)
      return [{ at: row?.start ?? line.y, page: floor }]
    }).sort((left, right) => left.at - right.at)
    const future = futureReferences[0]
    if (future && future.at <= start + EPSILON) { flow.advanceToPage(future.page); continued = start > EPSILON; continue }
    if (future) room = Math.min(room, future.at - start)
    const header = measurement.headerHeight
    const fullRoom = flow.column.height
    // Oversized headers or large objects keep their full source content. Skip
    // the derived repeat when it would leave no space for the next object.
    const unsplit = measurement.indivisible.filter(interval => interval.kind === 'row' && interval.start <= start + EPSILON
      && interval.end > start + EPSILON && interval.end - start <= fullRoom + EPSILON)
    const required = Math.max(atomicRoom(measurement, start), ...unsplit.map(interval => interval.end - start))
    const repeat = continued && start >= header - EPSILON && header > EPSILON
      && header + required <= fullRoom + EPSILON && header < room - EPSILON
    let left = flow.column.x + documentTableHorizontalOffset(table, measurement.width, flow.column.width)
    if (repeat) { appendSlice(measurement, 0, header, flow, result, left, true); room -= header }
    let end = tableCut(measurement, start, room, fullRoom - (repeat ? header : 0))
    if (end <= start + EPSILON) {
      const fresh = flow.cursor <= flow.column.y + (repeat ? header : 0) + EPSILON
      if (!fresh) { flow.nextColumn(); continued = true; continue }
      // A single glyph line larger than the physical page cannot be broken
      // without changing its text/size. Use the same one-time overflow rule
      // as ordinary paragraph layout, retaining the complete layout object.
      const object = measurement.indivisible.find(interval => interval.kind !== 'row' && interval.start <= start + EPSILON && interval.end > start + EPSILON)
      if (!object) throw new RangeError('表格分页无法推进')
      end = object.end
      flow.ensureRoom(end - start)
      left = flow.column.x + documentTableHorizontalOffset(table, measurement.width, flow.column.width)
    }
    appendSlice(measurement, start, end, flow, result, left, false)
    start = end
    if (start < measurement.height - EPSILON) { flow.nextColumn(); continued = true }
  }
}
