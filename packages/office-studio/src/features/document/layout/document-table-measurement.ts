import type { DocumentBlock } from '../../../domain/model'
import { imageForBlock, tableForBlock, type DocumentRun } from '../model/document-model'
import { documentTableCellTarget, documentTableGrid, type DocumentTableGridCell } from '../model/document-table-grid'
import { documentTableCellBorders, documentTableCellFill, documentTableCellMargin,
  documentTableHorizontalOffset, documentTableRowProperties, DOCUMENT_TABLE_MIN_ROW_HEIGHT } from '../model/document-table-style'
import type { ParagraphListMarker } from '../model/document-lists'
import { codeBlockData } from '../code-blocks/code-blocks'
import { documentTocPageTab } from '../toc/document-toc-layout'
import type { DocumentTextLayoutDisplay, LayoutCell, LayoutDivider, LayoutImage, LayoutLine, LayoutMeasurer, LayoutRect, LayoutTableRowMetric } from './document-layout-types'
import { appendListMarker, paragraphGeometry } from './document-paragraph-layout'
import { paragraphLineHeight, positionLine, wrapRuns } from './document-text-layout'
import { measureDocumentTableColumnWidths } from './document-table-intrinsic-widths'
import { documentTableCellBlocks, documentTableCellContentContainers } from '../model/document-table-continuations'
import { documentPictureInputAnchors, layoutDocumentPictureInputAnchor } from './document-picture-input-anchor'
import { isDocumentStructuredBlock } from '../model/document-structural-blocks'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { documentColumnWidths } from './document-structural-layout'
import type { LayoutContainerDecoration } from './document-layout-types'

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
interface CellContent {
  owner: DocumentTableGridCell; width: number; x: number; height: number
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]; tables: MeasuredDocumentTable[]; containers: LayoutContainerDecoration[]
  indivisible: TableVerticalInterval[]
}
const EPSILON = 1e-6

export function intersectTableRect(a: LayoutRect, b: LayoutRect): LayoutRect {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y)
  return { x, y, width: Math.max(0, Math.min(a.x + a.width, b.x + b.width) - x),
    height: Math.max(0, Math.min(a.y + a.height, b.y + b.height) - y) }
}

export function translateTableLine(line: LayoutLine, dx: number, dy: number): LayoutLine {
  return { ...line, x: line.x + dx, y: line.y + dy, baseline: line.baseline + dy,
    ...(line.clip ? { clip: translateTableRect(line.clip, dx, dy) } : {}),
    ...(line.markerStyle ? { markerStyle: { ...line.markerStyle, x: line.markerStyle.x + dx } } : {}),
    glyphs: line.glyphs.map(glyph => ({ ...glyph, x: glyph.x + dx, y: glyph.y + dy })) }
}
export function translateTableRect<T extends LayoutRect>(rect: T, dx: number, dy: number): T {
  return { ...rect, x: rect.x + dx, y: rect.y + dy }
}
function translatedClipped<T extends LayoutRect & { clip?: LayoutRect }>(item: T, dx: number, dy: number, parent: LayoutRect): T {
  const shifted = translateTableRect(item, dx, dy)
  return { ...shifted, clip: item.clip ? intersectTableRect(translateTableRect(item.clip, dx, dy), parent) : { ...parent } }
}
function prefix(values: readonly number[]): number[] {
  const positions = [0]
  for (const value of values) positions.push(positions.at(-1)! + value)
  return positions
}
function assertSpace(value: number): number {
  if (!Number.isFinite(value) || value < 0) throw new RangeError('表格内容间距必须是非负 CSS px')
  return value
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
    const content: CellContent = { owner, width, x: columns[owner.column]!, height: 0, lines: [], cells: [], images: [], dividers: [], tables: [], containers: [], indivisible: [] }
    let cursor = 0
    const appendBlocks = (blocks: readonly DocumentBlock[], left: number, contentWidth: number, nesting = 0) => {
    if (nesting > 64) throw new RangeError('表格内高级块嵌套超过预算。')
    const pictureAnchors = documentPictureInputAnchors(blocks)
    const anchorIds = new Set([...pictureAnchors.values()].map(block => block.id))
    for (const paragraph of blocks) {
      if (anchorIds.has(paragraph.id)) continue
      const inputAnchor = pictureAnchors.get(paragraph.id), style = inputAnchor?.style ?? paragraph.style
      cursor += assertSpace(style.spaceBefore)
      if (isDocumentStructuredBlock(paragraph)) {
        const top = cursor, firstLine = content.lines.length
        if (paragraph.type === 'columns') {
          const dimensions = documentColumnWidths(paragraph.columns!, contentWidth), endings: number[] = []
          let x = left
          for (const [index, column] of paragraph.columns!.columns.entries()) {
            if (!dimensions.stacked) cursor = top
            appendBlocks(column.blocks, dimensions.stacked ? left : x, dimensions.widths[index]!, nesting + 1)
            endings.push(cursor)
            if (dimensions.stacked) cursor += dimensions.gap
            else x += dimensions.widths[index]! + dimensions.gap
          }
          if (!dimensions.stacked) cursor = Math.max(top, ...endings)
        } else {
          const config = paragraph.callout, paddingLeft = config ? config.paddingLeft + (config.showIcon ? config.iconSize + config.iconGap : 0)
            : paragraph.type === 'code' ? paragraph.code?.showLineNumbers ? 44 : 12 : 18
          const paddingRight = config?.paddingRight ?? 12, paddingTop = config?.paddingTop ?? 8, paddingBottom = config?.paddingBottom ?? 8
          cursor += paddingTop
          appendBlocks(paragraph.children!, left + Math.min(paddingLeft, Math.max(0, contentWidth - 1)), Math.max(1, contentWidth - paddingLeft - paddingRight), nesting + 1)
          cursor += paddingBottom
          content.containers.push({ blockId: paragraph.id, kind: paragraph.type as 'quote' | 'callout' | 'code', x: left, y: top, width: contentWidth,
            height: Math.max(1, cursor - top), page: 0, first: true, last: true, tableDepth: depth + 1, ...(config ? { callout: { ...config } } : {}) })
          if (paragraph.type === 'code') {
            const paragraphs = documentTextFieldParagraphs({ blocks: paragraph.children! }), source = paragraphs.map(value => value.text).join('\n')
            const byId = new Map(paragraphs.map(value => [value.block.id, value]))
            for (const line of content.lines.slice(firstLine)) {
              const value = byId.get(line.target.cellBlockId ?? line.target.blockId)
              if (!value) continue
              const offset = value.absoluteStart + line.start
              line.codeBlock = { language: paragraph.code!.language, lineNumber: 1 + (source.slice(0, offset).match(/\n/g)?.length ?? 0),
                containerId: paragraph.id, sourceOffset: value.absoluteStart, showLineNumbers: paragraph.code!.showLineNumbers,
                numberVisible: offset === 0 || source[offset - 1] === '\n' }
            }
          }
        }
      } else if (paragraph.type === 'table') {
        const nested = measureDocumentTable(paragraph, contentWidth, markers, measurer, depth + 1, maximumObjectHeight, display)
        const nestedTable = tableForBlock(paragraph)!, offset = left + documentTableHorizontalOffset(nestedTable, nested.width, contentWidth)
        const offsetY = nestedTable.style?.offsetY ?? 0
        if (!Number.isFinite(offsetY)) throw new RangeError('表格纵向偏移必须是有限 CSS px')
        cursor += Math.max(0, offsetY)
        content.lines.push(...nested.lines.map(line => translateTableLine(line, offset, cursor)))
        for (const cell of nested.cells) content.cells.push({ ...translateTableRect(cell, offset, cursor), ...(cell.clip ? { clip: translateTableRect(cell.clip, offset, cursor) } : {}) })
        for (const image of nested.images) content.images.push({ ...translateTableRect(image, offset, cursor), ...(image.clip ? { clip: translateTableRect(image.clip, offset, cursor) } : {}) })
        for (const divider of nested.dividers) content.dividers.push({ ...translateTableRect(divider, offset, cursor), ...(divider.clip ? { clip: translateTableRect(divider.clip, offset, cursor) } : {}) })
        content.tables.push(...nested.tables.map(item => ({ ...translateTableRect(item, offset, cursor), ...(item.clip ? { clip: translateTableRect(item.clip, offset, cursor) } : {}), rows: item.rows.map(row => ({ ...row, y: row.y + cursor })) })))
        content.containers.push(...nested.containers.map(item => ({ ...translateTableRect(item, offset, cursor), ...(item.clip ? { clip: translateTableRect(item.clip, offset, cursor) } : {}) })))
        content.indivisible.push(...nested.indivisible.map(interval => ({ ...interval, start: interval.start + cursor, end: interval.end + cursor })))
        cursor += nested.height
      } else if (paragraph.type === 'image') {
        const image = imageForBlock(paragraph)
        if (image) {
          const nativeWidth = Math.max(1, image.width || 320), nativeHeight = Math.max(1, image.height || nativeWidth * 0.625)
          const geometry = inputAnchor && paragraphGeometry(inputAnchor, markers.get(inputAnchor.id), contentWidth, measurer)
          const available = geometry ? geometry.width(0) : contentWidth
          const scale = Math.min(1, available / nativeWidth, Math.max(1, maximumObjectHeight - margin.top - margin.bottom) / nativeHeight)
          const width = nativeWidth * scale, height = nativeHeight * scale
          const shift = (geometry?.firstIndent ?? 0) + (style.align === 'center' ? (available - width) / 2 : style.align === 'right' ? available - width : 0)
          const positioned = { blockId: paragraph.id, src: image.src, alt: image.alt, x: left + shift, y: cursor, width, height, page: 0, tableCellId: owner.cell.id }
          content.images.push(positioned)
          if (inputAnchor) content.lines.push(layoutDocumentPictureInputAnchor(inputAnchor, positioned, documentTableCellTarget(block.id, owner, inputAnchor.id),
            contentWidth, left, measurer, content.images, markers.get(inputAnchor.id), display))
          cursor += height
        }
      } else if (paragraph.type === 'divider') {
        content.dividers.push({ x: left, y: cursor + 9, width: contentWidth, height: 1, page: 0 })
        cursor += 18
      } else if (!['pageBreak', 'columnBreak', 'sectionBreak'].includes(paragraph.type)) {
        const list = markers.get(paragraph.id), geometry = paragraphGeometry(paragraph, list, contentWidth, measurer), pageTab = documentTocPageTab(paragraph)
        const rightTabs = (index: number) => pageTab ? pageTab.stops.map(stop => ({ ...stop, offset: pageTab.offset,
          ...(stop.right === undefined ? {} : { right: Math.max(0, stop.right - (index === 0 ? geometry.firstIndent : geometry.indent)) }) })) : []
        const rawLines = wrapRuns(paragraph.runs as DocumentRun[], geometry.width, measurer, paragraph.style.wordWrap === true, undefined, pageTab ? rightTabs : undefined,
          { references: display?.references, leading: display?.leadingLabels?.get(paragraph.id) })
        const code = codeBlockData(paragraph)
        for (const [index, raw] of rawLines.entries()) {
          const height = paragraphLineHeight(Math.max(raw.fontHeight, index === 0 ? geometry.markerHeight : 0), paragraph.style)
          const indent = index === 0 ? geometry.firstIndent : geometry.indent
          const line = positionLine(raw, documentTableCellTarget(block.id, owner, paragraph.id), left + indent, cursor,
            geometry.width(index), height, 0, paragraph.style.align, index < rawLines.length - 1,
            index === 0 ? list?.marker ?? (paragraph.type === 'quote' ? '│' : undefined) : undefined)
          line.tableCellId = owner.cell.id
          if (index === 0) appendListMarker(line, paragraph, list, geometry, left, content.images)
          if (code) {
            line.codeBlock = { language: code.language, lineNumber: index + 1 }
            line.x += 8; line.width = Math.max(1, line.width - 16)
            for (const glyph of line.glyphs) glyph.x += 8
          }
          if (index === 0 && paragraph.type === 'checklist') line.checklist = { blockId: paragraph.id, checked: paragraph.checklist?.checked === true }
          content.lines.push(line)
          cursor += height
        }
      }
      cursor += assertSpace(style.spaceAfter)
    }
    }
    for (const blocks of documentTableCellContentContainers(owner.cell)) appendBlocks(blocks, 0, contentWidth)
    content.height = cursor + margin.top + margin.bottom
    return content
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
