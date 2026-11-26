import type { DocumentBlock, DocumentCodeConfig } from '../../../domain/model'
import { imageForBlock, tableForBlock, type TextTarget } from '../model/document-model'
import { documentHorizontalLineBorder } from '../model/document-horizontal-line'
import { documentTableHorizontalOffset } from '../model/document-table-style'
import type { ParagraphListMarker } from '../model/document-lists'
import { codeBlockData } from '../code-blocks/code-blocks'
import { isDocumentStructuredBlock } from '../model/document-structural-blocks'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { documentQuoteLineColor } from '../model/document-structured-appearance'
import { appendListMarker, measureDocumentImageTextRow, measureDocumentParagraphLines, paragraphGeometry } from './document-paragraph-layout'
import { paragraphLineHeight, positionLine } from './document-text-layout'
import { documentPictureInputAnchors, layoutDocumentPictureInputAnchor } from './document-picture-input-anchor'
import { appendDocumentInlineImages, documentImageBox, documentImageObstacle, documentImageWrapBands, positionDocumentImage, type DocumentImageObstacle } from './document-image-layout'
import { documentImageAtom } from '../images/document-image-atom'
import { documentColumnMetrics, documentColumnWidths, documentDecoratedBlockInsets } from './document-structural-layout'
import type { DocumentTextLayoutDisplay, LayoutCell, LayoutContainerDecoration, LayoutDivider, LayoutImage, LayoutLine, LayoutMeasurer, LayoutRect } from './document-layout-types'
import type { measureDocumentTable, MeasuredDocumentTable, TableVerticalInterval } from './document-table-measurement'

export interface DocumentBlockContentMeasurement {
  height: number
  lines: LayoutLine[]; cells: LayoutCell[]; images: LayoutImage[]; dividers: LayoutDivider[]
  tables: MeasuredDocumentTable[]; containers: LayoutContainerDecoration[]; indivisible: TableVerticalInterval[]
}
export interface DocumentBlockMeasurementOptions {
  width: number
  markers: Map<string, ParagraphListMarker>
  measurer: LayoutMeasurer
  measureTable: typeof measureDocumentTable
  target?: (blockId: string) => TextTarget
  tableCellId?: string
  tableDepth?: number
  nextTableDepth?: number
  maximumObjectHeight?: number
  maximumImageHeight?: number
  display?: DocumentTextLayoutDisplay
  code?: DocumentCodeConfig
}

export function intersectDocumentRect(a: LayoutRect, b: LayoutRect): LayoutRect {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y)
  return { x, y, width: Math.max(0, Math.min(a.x + a.width, b.x + b.width) - x),
    height: Math.max(0, Math.min(a.y + a.height, b.y + b.height) - y) }
}
export function translateDocumentLine(line: LayoutLine, dx: number, dy: number): LayoutLine {
  return { ...line, x: line.x + dx, y: line.y + dy, baseline: line.baseline + dy,
    ...(line.clip ? { clip: translateDocumentRect(line.clip, dx, dy) } : {}),
    ...(line.markerStyle ? { markerStyle: { ...line.markerStyle, x: line.markerStyle.x + dx } } : {}),
    ...(line.codeBlock ? { codeBlock: { ...line.codeBlock,
      ...(line.codeBlock.gutterX === undefined ? {} : { gutterX: line.codeBlock.gutterX + dx }),
      ...(line.codeBlock.gutterClip ? { gutterClip: translateDocumentRect(line.codeBlock.gutterClip, dx, dy) } : {}) } } : {}),
    glyphs: line.glyphs.map(glyph => ({ ...glyph, x: glyph.x + dx, y: glyph.y + dy })) }
}
export function translateDocumentRect<T extends LayoutRect>(rect: T, dx: number, dy: number): T {
  return { ...rect, x: rect.x + dx, y: rect.y + dy }
}
const assertSpace = (value: number): number => {
  if (!Number.isFinite(value) || value < 0) throw new RangeError('文档内容间距必须是非负 CSS px')
  return value
}

/** Tables and auxiliary hosts share relative block geometry; host addressing is supplied by the caller. */
export function measureDocumentBlockCollections(collections: readonly (readonly DocumentBlock[])[], options: DocumentBlockMeasurementOptions): DocumentBlockContentMeasurement {
  const { markers, measurer, display } = options, target = options.target ?? ((blockId: string): TextTarget => ({ blockId }))
  const maximumObjectHeight = options.maximumObjectHeight ?? Number.POSITIVE_INFINITY
  const maximumImageHeight = options.maximumImageHeight ?? maximumObjectHeight
  const output: DocumentBlockContentMeasurement = { height: 0, lines: [], cells: [], images: [], dividers: [], tables: [], containers: [], indivisible: [] }
  const containerDepth = options.tableDepth === undefined ? {} : { tableDepth: options.tableDepth }
  let cursor = 0
  const rootObstacles: DocumentImageObstacle[] = []
  const appendBlocks = (blocks: readonly DocumentBlock[], left: number, contentWidth: number, nesting = 0, parentCode?: DocumentCodeConfig,
    obstacles: DocumentImageObstacle[] = []): void => {
    if (nesting > 64) throw new RangeError('文档高级块嵌套超过预算。')
    const pictureAnchors = documentPictureInputAnchors(blocks), anchorIds = new Set([...pictureAnchors.values()].map(block => block.id))
    for (const paragraph of blocks) {
      if (anchorIds.has(paragraph.id)) continue
      const inputAnchor = pictureAnchors.get(paragraph.id), style = inputAnchor?.style ?? paragraph.style
      cursor += assertSpace(style.spaceBefore)
      if (isDocumentStructuredBlock(paragraph)) {
        const top = cursor, firstLine = output.lines.length
        if (paragraph.type === 'columns') {
          const dimensions = documentColumnWidths(paragraph.columns!, contentWidth), metrics = documentColumnMetrics(paragraph.columns!), endings: number[] = []
          let x = left
          for (const [index, column] of paragraph.columns!.columns.entries()) {
            if (!dimensions.stacked) cursor = top
            else if (index) cursor += dimensions.gap
            const begin = cursor, width = dimensions.widths[index]!, inset = Math.min(metrics.horizontalPadding, Math.max(0, width - 1) / 2)
            appendBlocks(column.blocks, (dimensions.stacked ? left : x) + inset, Math.max(1, width - inset * 2), nesting + 1, parentCode)
            cursor = Math.max(cursor, begin + metrics.minHeight); endings.push(cursor)
            if (!dimensions.stacked) x += width + dimensions.gap
          }
          if (!dimensions.stacked) cursor = Math.max(top, ...endings)
          output.containers.push({ blockId: paragraph.id, kind: 'columns', x: left, y: top, width: contentWidth,
            height: Math.max(1, cursor - top), page: 0, first: true, last: true, ...containerDepth,
            columnGroup: { columnIds: paragraph.columns!.columns.map(column => column.id), ...dimensions, ...metrics } })
        } else {
          const config = paragraph.callout, { left: paddingLeft, right: paddingRight, top: paddingTop, bottom: paddingBottom } = documentDecoratedBlockInsets(paragraph)
          const sum = paddingLeft + paddingRight, scale = sum > contentWidth - 1 ? Math.max(0, contentWidth - 1) / Math.max(1, sum) : 1
          cursor += paddingTop
          appendBlocks(paragraph.children!, left + paddingLeft * scale, Math.max(1, contentWidth - sum * scale), nesting + 1, paragraph.code ?? parentCode)
          cursor += paddingBottom
          output.containers.push({ blockId: paragraph.id, kind: paragraph.type as 'quote' | 'callout' | 'code', x: left, y: top, width: contentWidth,
            height: Math.max(1, cursor - top), page: 0, first: true, last: true, ...containerDepth, ...(config ? { callout: { ...config } } : {}),
            ...(paragraph.code ? { codeWrap: paragraph.code.wrap } : {}),
            ...(paragraph.type === 'quote' ? { quoteLineColor: documentQuoteLineColor(paragraph) } : {}) })
          if (paragraph.type === 'code') {
            const paragraphs = documentTextFieldParagraphs({ blocks: paragraph.children! }), source = paragraphs.map(value => value.text).join('\n')
            const byId = new Map(paragraphs.map(value => [value.block.id, value]))
            for (const line of output.lines.slice(firstLine)) {
              if (line.codeBlock?.containerId && line.codeBlock.containerId !== paragraph.id) continue
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
        const nested = options.measureTable(paragraph, contentWidth, markers, measurer, options.nextTableDepth ?? 0, maximumObjectHeight, display)
        const table = tableForBlock(paragraph)!, dx = left + documentTableHorizontalOffset(table, nested.width, contentWidth), offsetY = table.style?.offsetY ?? 0
        if (!Number.isFinite(offsetY)) throw new RangeError('表格纵向偏移必须是有限 CSS px')
        cursor += Math.max(0, offsetY)
        output.lines.push(...nested.lines.map(line => translateDocumentLine(line, dx, cursor)))
        const shifted = <T extends LayoutRect & { clip?: LayoutRect }>(item: T): T => ({ ...translateDocumentRect(item, dx, cursor),
          ...(item.clip ? { clip: translateDocumentRect(item.clip, dx, cursor) } : {}) })
        output.cells.push(...nested.cells.map(shifted)); output.images.push(...nested.images.map(shifted)); output.dividers.push(...nested.dividers.map(shifted))
        output.tables.push(...nested.tables.map(item => ({ ...shifted(item), rows: item.rows.map(row => ({ ...row, y: row.y + cursor })) })))
        output.containers.push(...nested.containers.map(shifted))
        output.indivisible.push(...nested.indivisible.map(interval => ({ ...interval, start: interval.start + cursor, end: interval.end + cursor })))
        cursor += nested.height
      } else if (paragraph.type === 'image') {
        const image = imageForBlock(paragraph)
        if (image) {
          const geometry = inputAnchor && paragraphGeometry(inputAnchor, markers.get(inputAnchor.id), contentWidth, measurer), available = geometry ? geometry.width(0) : contentWidth
          const box = documentImageBox(image, available, Math.max(1, maximumImageHeight)), { width, height } = box
          const shift = (geometry?.firstIndent ?? 0) + (style.align === 'center' ? (available - width) / 2 : style.align === 'right' ? available - width : 0)
          const floated = (image.wrappingStyle ?? 'inline') !== 'inline'
          // Relative hosts have no final page height yet. Keep their reference frame
          // finite; the outer pagination pass supplies the actual page/margin frame.
          const area = { x: left, y: 0, width: contentWidth,
            height: Math.max(height, Number.isFinite(maximumObjectHeight) ? maximumObjectHeight : cursor + height) }
          const position = floated ? positionDocumentImage(image, { ...area, content: area, pageNumber: 1 }, area, cursor, box) : { x: left + shift, y: cursor }
          const positioned = { blockId: paragraph.id, src: image.src, alt: image.alt, ...position, ...box, page: 0,
            objectImage: true, angle: image.angle, flipX: image.flipX, flipY: image.flipY, crop: image.crop, wrappingStyle: image.wrappingStyle ?? 'inline', layer: image.layer ?? 0,
            ...(options.tableCellId ? { tableCellId: options.tableCellId } : {}) }
          output.images.push(positioned)
          const obstacle = documentImageObstacle(positioned, image)
          if (obstacle) obstacles.push(obstacle)
          if (inputAnchor) output.lines.push(layoutDocumentPictureInputAnchor(inputAnchor, positioned, target(inputAnchor.id), contentWidth, left, measurer, output.images, markers.get(inputAnchor.id), display))
          if (!floated) cursor += height
        }
      } else if (paragraph.type === 'divider') {
        const border = documentHorizontalLineBorder(paragraph)
        output.dividers.push({ x: left, y: cursor + border.padding + border.width / 2, width: contentWidth, height: border.width, page: 0,
          color: border.color, style: border.style })
        cursor += Math.max(18, border.padding * 2 + border.width)
      } else if (!['pageBreak', 'columnBreak', 'sectionBreak'].includes(paragraph.type)) {
        const list = markers.get(paragraph.id), code = codeBlockData(paragraph)
        let sourceOffset = 0
        for (const run of paragraph.runs) {
          const atom = documentImageAtom(run)
          if (atom && (atom.image.wrappingStyle ?? 'inline') !== 'inline') {
            const image = atom.image, box = documentImageBox(image, contentWidth, maximumImageHeight)
            const area = { x: left, y: 0, width: contentWidth, height: Math.max(box.height, Number.isFinite(maximumObjectHeight) ? maximumObjectHeight : cursor + box.height) }
            const positioned: LayoutImage = { ...positionDocumentImage(image, { ...area, content: area, pageNumber: 1 }, area, cursor, box), ...box,
              blockId: atom.imageId, src: image.src, alt: image.alt, page: 0, objectImage: true,
              angle: image.angle, flipX: image.flipX, flipY: image.flipY, crop: image.crop, wrappingStyle: image.wrappingStyle, layer: image.layer ?? 0,
              ...(options.tableCellId ? { tableCellId: options.tableCellId } : {}),
              anchor: { target: target(paragraph.id), startOffset: sourceOffset, endOffset: sourceOffset + run.text.length } }
            output.images.push(positioned)
            const obstacle = documentImageObstacle(positioned, image)
            if (obstacle) obstacles.push(obstacle)
          }
          sourceOffset += run.text.length
        }
        let available = contentWidth, measured = measureDocumentParagraphLines(paragraph, list, available, measurer, display, parentCode, 0, 0, maximumImageHeight)
        let geometry = measured.geometry, rawLines = measured.lines, rawIndex = 0, lineIndex = 0, baseOffset = 0, heightFloor = 0
        while (rawIndex < rawLines.length) {
          const raw = rawLines[rawIndex]!
          const wholeGeometry = paragraphGeometry(paragraph, list, contentWidth, measurer), wholeIndent = lineIndex === 0 ? wholeGeometry.firstIndent : wholeGeometry.indent
          const minimumHeight = paragraphLineHeight(Math.max(paragraph.runs[0]?.fontSize ? paragraph.runs[0].fontSize * 4 / 3 : 14.67,
            lineIndex === 0 ? wholeGeometry.markerHeight : 0), paragraph.style)
          const bands = documentImageWrapBands({ x: left + wholeIndent, y: 0,
            width: Math.max(1, contentWidth - wholeIndent - wholeGeometry.rightIndent), height: 0 }, cursor, Math.max(minimumHeight, heightFloor), obstacles)
          cursor = bands[0]!.y
          if (bands.length > 1) {
            const rowStart = baseOffset + raw.start, row = measureDocumentImageTextRow(paragraph, list, contentWidth, bands,
              measurer, display, parentCode, lineIndex, rowStart, maximumImageHeight)
            if (row.height > Math.max(minimumHeight, heightFloor) + 1e-6) { heightFloor = row.height; continue }
            const rowKey = `${paragraph.id}:relative:${cursor}:${rowStart}`
            for (const [fragmentIndex, fragment] of row.fragments.entries()) {
              const first = lineIndex === 0 && fragmentIndex === 0
              const line = positionLine(fragment.raw, target(paragraph.id), fragment.band.x, cursor, fragment.band.width, row.height,
                0, paragraph.style.align, !row.complete, first ? list?.marker ?? (paragraph.type === 'quote' ? '│' : undefined) : undefined)
              Object.assign(line, { rowKey, fragmentIndex, fragmentCount: row.fragments.length })
              if (options.tableCellId) line.tableCellId = options.tableCellId
              if (first) appendListMarker(line, paragraph, list, wholeGeometry, fragment.band.x - wholeIndent, output.images)
              if (code) {
                line.codeBlock = { language: code.language, lineNumber: lineIndex + 1, numberVisible: fragmentIndex === 0 }
                line.x += 8; line.width = Math.max(1, line.width - 16); for (const glyph of line.glyphs) glyph.x += 8
              }
              if (first && paragraph.type === 'checklist') line.checklist = { blockId: paragraph.id, checked: paragraph.checklist?.checked === true }
              appendDocumentInlineImages(line, output.images); output.lines.push(line)
            }
            cursor += row.height; baseOffset = row.nextOffset; lineIndex++; rawIndex = 0; heightFloor = 0
            available = contentWidth
            measured = measureDocumentParagraphLines(paragraph, list, available, measurer, display, parentCode, lineIndex, baseOffset, maximumImageHeight)
            geometry = measured.geometry; rawLines = row.complete ? [] : measured.lines
            continue
          }
          const height = paragraphLineHeight(Math.max(raw.fontHeight, lineIndex === 0 ? geometry.markerHeight : 0), paragraph.style)
          heightFloor = Math.max(heightFloor, height)
          const actualBands = documentImageWrapBands({ x: left + wholeIndent, y: 0,
            width: Math.max(1, contentWidth - wholeIndent - wholeGeometry.rightIndent), height: 0 }, cursor, heightFloor, obstacles), band = actualBands[0]!
          cursor = band.y
          if (actualBands.length > 1) continue
          const measurementWidth = band.width + wholeIndent + wholeGeometry.rightIndent
          if (Math.abs(available - measurementWidth) > 1e-6) {
            // Rewrap only the unplaced suffix, retaining original UTF-16 offsets
            // for selection, note references, fields and nested-cell commands.
            baseOffset += raw.start; available = measurementWidth
            measured = measureDocumentParagraphLines(paragraph, list, available, measurer, display, parentCode, lineIndex, baseOffset, maximumImageHeight)
            geometry = measured.geometry; rawLines = measured.lines; rawIndex = 0
            continue
          }
          const line = positionLine({ ...raw, start: baseOffset + raw.start, end: baseOffset + raw.end,
            glyphs: raw.glyphs.map(glyph => ({ ...glyph, offset: baseOffset + glyph.offset, end: baseOffset + glyph.end })) },
          target(paragraph.id), band.x, cursor, geometry.width(lineIndex), height, 0, paragraph.style.align,
          rawIndex < rawLines.length - 1, lineIndex === 0 ? list?.marker ?? (paragraph.type === 'quote' ? '│' : undefined) : undefined)
          if (options.tableCellId) line.tableCellId = options.tableCellId
          if (lineIndex === 0) appendListMarker(line, paragraph, list, geometry, band.x - wholeIndent, output.images)
          if (code) { line.codeBlock = { language: code.language, lineNumber: lineIndex + 1 }; line.x += 8; line.width = Math.max(1, line.width - 16); for (const glyph of line.glyphs) glyph.x += 8 }
          if (lineIndex === 0 && paragraph.type === 'checklist') line.checklist = { blockId: paragraph.id, checked: paragraph.checklist?.checked === true }
          appendDocumentInlineImages(line, output.images)
          output.lines.push(line); cursor += height
          rawIndex++; lineIndex++; heightFloor = 0
        }
      }
      cursor += assertSpace(style.spaceAfter)
    }
  }
  for (const blocks of collections) appendBlocks(blocks, 0, Math.max(1, options.width), 0, options.code, rootObstacles)
  output.height = output.images.reduce((height, image) => image.objectImage ? Math.max(height, image.y + image.height) : height, cursor)
  return output
}
