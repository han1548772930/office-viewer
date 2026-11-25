import type { DocumentBlock, DocumentCodeConfig } from '../../../domain/model'
import { createRun, runText, sliceRuns, type DocumentRun } from '../model/document-model'
import type { ParagraphListMarker } from '../model/document-lists'
import { codeBlockData } from '../code-blocks/code-blocks'
import type { DocumentLayout, DocumentTextLayoutDisplay, LayoutImage, LayoutLine, LayoutMeasurer, LayoutRect } from './document-layout-types'
import type { DocumentSectionFlow } from './document-section-flow'
import { paragraphLineHeight, positionLine, runFont, wrapRuns, type RawLine } from './document-text-layout'
import { documentTocPageTab } from '../toc/document-toc-layout'
import { planDocumentParagraphFragment } from './document-pagination-rules'
import { documentImageAtom } from '../images/document-image-atom'
import { appendDocumentInlineImages, documentImageBox, documentImageObstacle, positionDocumentImage, type DocumentImageObstacle } from './document-image-layout'

export interface ParagraphGeometry {
  indent: number
  firstIndent: number
  rightIndent: number
  markerIndent?: number
  markerFont?: string
  markerHeight: number
  markerWidth: number
  width: (index: number) => number
}

export function paragraphGeometry(block: DocumentBlock, list: ParagraphListMarker | undefined, available: number, measurer: LayoutMeasurer): ParagraphGeometry {
  const legacyIndent = block.style.indent * 24
  const left = block.style.indentStart ?? legacyIndent
  let indent = left, firstIndent = left + (block.style.indentFirstLine ?? 0) - (block.style.hanging ?? 0)
  let rightIndent = block.style.indentEnd ?? 0
  let markerIndent: number | undefined, markerFont: string | undefined, markerHeight = 0, markerWidth = 0
  if (list) {
    const extra = block.style.indentStart === undefined ? legacyIndent : 0
    indent = list.inset + extra
    rightIndent = list.rightIndent
    markerIndent = list.markerIndent + extra
    markerFont = runFont(createRun('', { fontFamily: list.fontFamily, fontSize: list.fontSize, bold: list.bold, italic: list.italic }))
    markerHeight = list.picture ? list.picture.heightPx ?? list.fontSize * 4 / 3 : list.fontSize * 4 / 3
    markerWidth = list.picture ? list.picture.widthPx ?? markerHeight : measurer.measure(list.marker, markerFont)
    if (list.alignment === 'center') markerIndent -= markerWidth / 2
    else if (list.alignment === 'right' || list.alignment === 'end') markerIndent -= markerWidth
    firstIndent = Math.max(indent, markerIndent + markerWidth + (markerWidth ? list.gap : 0))
  } else if (block.type === 'checklist') {
    indent += 28 * (block.level + 1)
    firstIndent = indent
  } else if (block.type === 'quote') {
    indent += 16
    firstIndent += 16
  }
  return { indent, firstIndent, rightIndent, markerIndent, markerFont, markerHeight, markerWidth,
    width: index => Math.max(1, available - (index === 0 ? firstIndent : indent) - rightIndent) }
}

export function appendListMarker(line: LayoutLine, block: DocumentBlock, list: ParagraphListMarker | undefined, geometry: ParagraphGeometry, left: number, images: LayoutImage[]): void {
  if (!list || geometry.markerIndent === undefined) return
  if (list.picture?.src) {
    images.push({ blockId: block.id, src: list.picture.src, alt: '项目符号', x: left + geometry.markerIndent,
      y: line.y + (line.height - geometry.markerHeight) / 2, width: geometry.markerWidth, height: geometry.markerHeight, page: line.page,
      sectionId: line.sectionId, sectionFragment: line.sectionFragment, column: line.column })
  } else if (geometry.markerFont) line.markerStyle = { x: left + geometry.markerIndent, font: geometry.markerFont, color: list.color }
}

export interface DocumentParagraphLineMeasurement {
  geometry: ParagraphGeometry
  lines: RawLine[]
  heights: number[]
  height: number
}

export interface DocumentImageTextRow {
  fragments: Array<{ band: LayoutRect; raw: RawLine }>
  geometry: ParagraphGeometry; height: number; nextOffset: number; complete: boolean
}

/** Each fragment consumes the same source suffix once; all share one baseline. */
export function measureDocumentImageTextRow(block: DocumentBlock, list: ParagraphListMarker | undefined, available: number,
  bands: readonly LayoutRect[], measurer: LayoutMeasurer, display?: DocumentTextLayoutDisplay, parentCode?: DocumentCodeConfig,
  lineIndex = 0, baseOffset = 0, maximumImageHeight = Number.MAX_SAFE_INTEGER): DocumentImageTextRow {
  const geometry = paragraphGeometry(block, list, available, measurer), indent = lineIndex === 0 ? geometry.firstIndent : geometry.indent
  const fragments: DocumentImageTextRow['fragments'] = []
  let nextOffset = baseOffset, complete = false
  for (const band of bands) {
    const fragmentDisplay = fragments.length && display ? { ...display, leadingLabels: undefined } : display
    const measured = measureDocumentParagraphLines(block, list, band.width + indent + geometry.rightIndent,
      measurer, fragmentDisplay, parentCode, lineIndex, nextOffset, maximumImageHeight), first = measured.lines[0]!, following = measured.lines[1]
    fragments.push({ band, raw: { ...first, start: nextOffset + first.start, end: nextOffset + first.end,
      glyphs: first.glyphs.map(glyph => ({ ...glyph, offset: nextOffset + glyph.offset, end: nextOffset + glyph.end })) } })
    nextOffset += following?.start ?? first.end
    complete = !following
    if (complete || following!.start > first.end) break
  }
  const ascent = Math.max(0, ...fragments.map(fragment => fragment.raw.ascent ?? fragment.raw.fontHeight * 0.82))
  const descent = Math.max(0, ...fragments.map(fragment => fragment.raw.fontHeight - (fragment.raw.ascent ?? fragment.raw.fontHeight * 0.82)))
  const fontHeight = ascent + descent, height = paragraphLineHeight(Math.max(fontHeight, lineIndex === 0 ? geometry.markerHeight : 0), block.style)
  for (const fragment of fragments) fragment.raw = { ...fragment.raw, fontHeight, ascent }
  return { fragments, geometry, height, nextOffset, complete }
}

/** Pagination preflight and drawing use the same wrapping, tabs, note labels,
 * code settings and actual line heights. Only geometry is returned here. */
export function measureDocumentParagraphLines(block: DocumentBlock, list: ParagraphListMarker | undefined, available: number,
  measurer: LayoutMeasurer, display?: DocumentTextLayoutDisplay, parentCode?: DocumentCodeConfig,
  lineIndex = 0, baseOffset = 0, maximumImageHeight = Number.MAX_SAFE_INTEGER): DocumentParagraphLineMeasurement {
  const codeData = codeBlockData(block), geometry = paragraphGeometry(block, list, available, measurer)
  const code = parentCode ?? (codeData ? { language: codeData.language, showLineNumbers: codeData.lineNumbers, tabSize: codeData.tabSize ?? 4, wrap: codeData.wrap ?? true } : undefined)
  const wrapping = { ...(code ? { softWrap: code.wrap, tabSize: code.tabSize } : {}), maximumImageHeight }
  const pageTab = documentTocPageTab(block)
  const rightTabs = (index: number) => !pageTab || pageTab.offset < baseOffset ? [] : pageTab.stops.map(stop => ({ ...stop,
    offset: pageTab.offset - baseOffset,
    ...(stop.right === undefined ? {} : { right: Math.max(0, stop.right - (index + lineIndex === 0 ? geometry.firstIndent : geometry.indent)) }),
  }))
  const remaining = baseOffset === 0 ? block.runs as DocumentRun[] : sliceRuns(block.runs as DocumentRun[], baseOffset, runText(block.runs).length)
  if (!remaining.length && block.runs.length) remaining.push({ ...block.runs.at(-1)!, text: '' })
  const lines = wrapRuns(remaining, index => geometry.width(index + lineIndex), measurer, block.style.wordWrap === true, undefined, pageTab ? rightTabs : undefined,
    { references: display?.references, leading: lineIndex === 0 && baseOffset === 0 ? display?.leadingLabels?.get(block.id) : undefined }, wrapping)
  const heights = lines.map((line, index) => paragraphLineHeight(Math.max(line.fontHeight, index + lineIndex === 0 ? geometry.markerHeight : 0), block.style))
  return { geometry, lines, heights, height: heights.reduce((sum, value) => sum + value, 0) }
}

export function layoutDocumentParagraph(block: DocumentBlock, list: ParagraphListMarker | undefined, flow: DocumentSectionFlow, result: DocumentLayout, measurer: LayoutMeasurer,
  display?: DocumentTextLayoutDisplay, parentCode?: DocumentCodeConfig, trailingHeight = 0): void {
  const codeData = codeBlockData(block), text = runText(block.runs)
  const flowKey = () => `${flow.page.index}:${flow.location.sectionFragment}:${flow.column.index}`
  let floatingFrame = '', floatingCommitted = false
  const floating: Array<{ image: LayoutImage; obstacle?: DocumentImageObstacle }> = []
  const synchronizeFloatingFrame = () => {
    if (floatingCommitted || floatingFrame === flowKey()) return
    for (const placed of floating) {
      if (placed.obstacle) flow.removeImageObstacle(placed.obstacle, placed.image.page)
      const index = result.images.indexOf(placed.image)
      if (index >= 0) result.images.splice(index, 1)
    }
    floating.length = 0; floatingFrame = flowKey()
    let sourceOffset = 0
    for (const run of block.runs) {
      const atom = documentImageAtom(run)
      if (atom && (atom.image.wrappingStyle ?? 'inline') !== 'inline') {
        const image = atom.image, box = documentImageBox(image, flow.column.width, flow.fullPageContentHeight)
        const positioned: LayoutImage = { ...positionDocumentImage(image, flow.page, flow.column, flow.cursor, box), ...box,
          blockId: atom.imageId, src: image.src, alt: image.alt, page: flow.page.index, ...flow.location, objectImage: true,
          angle: image.angle, flipX: image.flipX, flipY: image.flipY, crop: image.crop, wrappingStyle: image.wrappingStyle, layer: image.layer ?? 0,
          anchor: { target: { blockId: block.id }, startOffset: sourceOffset, endOffset: sourceOffset + run.text.length } }
        result.images.push(positioned)
        const obstacle = documentImageObstacle(positioned, image)
        if (obstacle) flow.addImageObstacle(obstacle)
        floating.push({ image: positioned, obstacle })
      }
      sourceOffset += run.text.length
    }
  }
  synchronizeFloatingFrame()
  const retainBlockedFloatingAnchor = (frame: string, cursor: number, height: number) => {
    if (!floatingCommitted && floating.some(placed => placed.obstacle) && (frame !== flowKey()
      || flow.cursor > cursor + 1e-6 && flow.cursor + height > flow.column.y + flow.column.height + 1e-6)) floatingCommitted = true
  }
  let available = flow.column.width, measured = measureDocumentParagraphLines(block, list, available, measurer, display, parentCode, 0, 0, flow.fullPageContentHeight)
  let geometry = measured.geometry, lines = measured.lines, rawIndex = 0, lineIndex = 0, baseOffset = 0, fragmentEnd = -1
  let fragmentKey = '', imageBandHeightFloor = 0
  const pagination = flow.paragraphPagination(block.style)
  while (rawIndex < lines.length) {
    synchronizeFloatingFrame()
    const currentKey = flowKey()
    if (fragmentEnd >= 0 && rawIndex === fragmentEnd && currentKey === fragmentKey) { flow.nextColumn(); fragmentEnd = -1; continue }
    const raw = lines[rawIndex]!
    const wholeGeometry = paragraphGeometry(block, list, flow.column.width, measurer), wholeIndent = lineIndex === 0 ? wholeGeometry.firstIndent : wholeGeometry.indent
    const minimumHeight = paragraphLineHeight(Math.max(block.runs[0]?.fontSize ? block.runs[0].fontSize * 4 / 3 : 14.67,
      lineIndex === 0 ? wholeGeometry.markerHeight : 0), block.style)
    flow.ensureRoom(Math.max(minimumHeight, imageBandHeightFloor))
    if (currentKey !== flowKey()) { fragmentEnd = -1; continue }
    const beforeWrapCursor = flow.cursor
    const readingBands = flow.imageWrapBands(Math.max(minimumHeight, imageBandHeightFloor), wholeIndent, wholeGeometry.rightIndent)
    // A full-column float can own an otherwise textless frame. Retain that
    // anchor when wrapping sends the first text below it to a later frame.
    retainBlockedFloatingAnchor(currentKey, beforeWrapCursor, Math.max(minimumHeight, imageBandHeightFloor))
    if (currentKey !== flowKey()) { fragmentEnd = -1; continue }
    if (readingBands.length > 1) {
      const rowStart = baseOffset + raw.start, row = measureDocumentImageTextRow(block, list, flow.column.width, readingBands,
        measurer, display, parentCode, lineIndex, rowStart, flow.fullPageContentHeight)
      if (row.height > Math.max(minimumHeight, imageBandHeightFloor) + 1e-6) { imageBandHeightFloor = row.height; continue }
      const beforeFrame = flowKey(), referenceFloor = flow.referencePageFloor(row.fragments.flatMap(fragment =>
        fragment.raw.glyphs.filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id)))
      flow.advanceToPage(referenceFloor)
      if (beforeFrame !== flowKey()) { fragmentEnd = -1; continue }
      let requiredHeight = row.height + (row.complete ? trailingHeight : 0)
      if (!row.complete && flow.paginationEnabled && pagination.widowControl) {
        const nextGeometry = paragraphGeometry(block, list, flow.column.width, measurer)
        const nextY = flow.cursor + row.height, nextBands = flow.imageWrapBandsAt(nextY, row.height, nextGeometry.indent, nextGeometry.rightIndent)
        const following = measureDocumentImageTextRow(block, list, flow.column.width, nextBands, measurer, display, parentCode,
          lineIndex + 1, row.nextOffset, flow.fullPageContentHeight)
        requiredHeight += nextBands[0]!.y - nextY + following.height + (following.complete ? trailingHeight : 0)
        const nextFloor = flow.referencePageFloor(following.fragments.flatMap(fragment => fragment.raw.glyphs
          .filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id)))
        if (requiredHeight <= flow.fullPageContentHeight + 1e-6) flow.advanceToPage(nextFloor)
      }
      if (requiredHeight <= flow.fullPageContentHeight + 1e-6) flow.ensureRoom(requiredHeight)
      else flow.ensureRoom(row.height)
      if (beforeFrame !== flowKey()) { fragmentEnd = -1; continue }
      const rowKey = `${block.id}:${flow.page.index}:${flow.location.sectionFragment}:${flow.column.index}:${rowStart}`
      if (lineIndex === 0 && block.type === 'heading') result.headings.push({ id: block.id, text, level: block.level || 1, y: flow.cursor, page: flow.page.index })
      for (const [fragmentIndex, fragment] of row.fragments.entries()) {
        const first = lineIndex === 0 && fragmentIndex === 0
        const line = positionLine(fragment.raw, { blockId: block.id }, fragment.band.x, flow.cursor, fragment.band.width,
          row.height, flow.page.index, block.style.align, !row.complete, first ? list?.marker ?? (block.type === 'quote' ? '│' : undefined) : undefined)
        Object.assign(line, flow.location, { rowKey, fragmentIndex, fragmentCount: row.fragments.length })
        if (first) appendListMarker(line, block, list, wholeGeometry, fragment.band.x - wholeIndent, result.images)
        if (codeData) {
          line.codeBlock = { language: codeData.language, lineNumber: lineIndex + 1, numberVisible: fragmentIndex === 0 }
          line.x += 8; line.width = Math.max(2, line.width - 16); for (const glyph of line.glyphs) glyph.x += 8
        }
        if (first && block.type === 'checklist') line.checklist = { blockId: block.id, checked: block.checklist?.checked === true }
        appendDocumentInlineImages(line, result.images)
        result.lines.push(line)
      }
      appendNoteSeparator(block, wholeGeometry, flow, result, lineIndex, row.height)
      floatingCommitted = true
      flow.consume(row.height)
      baseOffset = row.nextOffset; lineIndex++; rawIndex = 0; fragmentEnd = -1; imageBandHeightFloor = 0
      available = flow.column.width
      measured = measureDocumentParagraphLines(block, list, available, measurer, display, parentCode, lineIndex, baseOffset, flow.fullPageContentHeight)
      geometry = measured.geometry; lines = row.complete ? [] : measured.lines
      continue
    }
    const height = paragraphLineHeight(Math.max(raw.fontHeight, lineIndex === 0 ? geometry.markerHeight : 0), block.style)
    const beforeFrame = flowKey()
    flow.advanceToPage(flow.referencePageFloor(raw.glyphs.filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id)))
    flow.ensureRoom(height)
    if (beforeFrame !== flowKey()) { fragmentEnd = -1; continue }
    imageBandHeightFloor = Math.max(imageBandHeightFloor, height)
    const beforeActualWrapCursor = flow.cursor
    const actualBands = flow.imageWrapBands(imageBandHeightFloor, wholeIndent, wholeGeometry.rightIndent), imageBand = actualBands[0]!
    retainBlockedFloatingAnchor(beforeFrame, beforeActualWrapCursor, imageBandHeightFloor)
    if (beforeFrame !== flowKey()) { fragmentEnd = -1; continue }
    if (actualBands.length > 1) continue
    const measurementWidth = imageBand.width + wholeIndent + wholeGeometry.rightIndent
    if (Math.abs(available - measurementWidth) > 1e-6) {
      // A continuous section can retain the old page's margins until overflow.
      // Rewrap only the unplaced suffix, with its original UTF-16 positions.
      baseOffset += raw.start
      available = measurementWidth
      measured = measureDocumentParagraphLines(block, list, available, measurer, display, parentCode, lineIndex, baseOffset, flow.fullPageContentHeight)
      geometry = measured.geometry; lines = measured.lines; rawIndex = 0; fragmentEnd = -1
      continue
    }
    const key = `${flow.page.index}:${flow.location.sectionFragment}:${flow.column.index}`
    if (fragmentEnd < 0 || fragmentKey !== key) {
      const plan = planDocumentParagraphFragment(measured.heights, rawIndex, flow.remainingContentHeight,
        flow.fullPageContentHeight, flow.paginationEnabled && pagination.widowControl, trailingHeight)
      if (!plan.lineCount && plan.requiredHeight <= flow.fullPageContentHeight + 1e-6
        && (!flow.atFullColumnStart || plan.requiredHeight <= flow.remainingContentHeight + 1e-6)) {
        flow.ensureRoom(plan.requiredHeight); fragmentEnd = -1; continue
      }
      let take = Math.max(1, plan.lineCount)
      // A retained note reference floor is a stronger physical boundary. Move
      // a two-line opening together when the first reference would isolate it.
      const future = lines.slice(rawIndex, rawIndex + take).map((line, index) => ({ index,
        floor: flow.referencePageFloor(line.glyphs.filter(glyph => glyph.note?.role === 'reference').map(glyph => glyph.run.id)) }))
        .find(value => value.floor > flow.page.index)
      if (future && future.index < 2 && flow.paginationEnabled && pagination.widowControl
        && (measured.heights[rawIndex]! + (measured.heights[rawIndex + 1] ?? 0)) <= flow.fullPageContentHeight + 1e-6) {
        flow.advanceToPage(future.floor); fragmentEnd = -1; continue
      }
      if (future) take = Math.min(take, Math.max(1, future.index))
      fragmentEnd = rawIndex + take; fragmentKey = key
    }
    if (lineIndex === 0 && block.type === 'heading') result.headings.push({ id: block.id, text, level: block.level || 1, y: flow.cursor, page: flow.page.index })
    const marker = lineIndex === 0 ? list?.marker ?? (block.type === 'quote' ? '│' : undefined) : undefined
    const line = positionLine({ ...raw, start: baseOffset + raw.start, end: baseOffset + raw.end,
      glyphs: raw.glyphs.map(glyph => ({ ...glyph, offset: baseOffset + glyph.offset, end: baseOffset + glyph.end })) },
    { blockId: block.id }, imageBand.x, flow.cursor, geometry.width(lineIndex), height, flow.page.index,
    block.style.align, rawIndex < lines.length - 1, marker)
    Object.assign(line, flow.location)
    if (lineIndex === 0) appendListMarker(line, block, list, geometry, imageBand.x - wholeIndent, result.images)
    if (codeData) {
      line.codeBlock = { language: codeData.language, lineNumber: lineIndex + 1 }
      line.x += 8
      line.width = Math.max(2, line.width - 16)
      for (const glyph of line.glyphs) glyph.x += 8
    }
    if (lineIndex === 0 && block.type === 'checklist') line.checklist = { blockId: block.id, checked: block.checklist?.checked === true }
    appendDocumentInlineImages(line, result.images)
    result.lines.push(line)
    appendNoteSeparator(block, geometry, flow, result, lineIndex, height)
    floatingCommitted = true
    flow.consume(height)
    imageBandHeightFloor = 0
    rawIndex++; lineIndex++
  }
}

function appendNoteSeparator(block: DocumentBlock, geometry: ParagraphGeometry, flow: DocumentSectionFlow, result: DocumentLayout, lineIndex: number, height: number): void {
  if (lineIndex !== 0 || !block.metadata.noteSeparator || typeof block.metadata.noteSeparator !== 'object') return
  const separator = block.metadata.noteSeparator, fraction = Reflect.get(separator, 'widthFraction'), color = Reflect.get(separator, 'color')
  if (Reflect.get(separator, 'version') === 1 && typeof fraction === 'number' && Number.isFinite(fraction) && fraction >= 0 && fraction <= 1) {
    result.dividers.push({ x: flow.column.x + geometry.firstIndent, y: flow.cursor + Math.max(1, height - 2), width: geometry.width(0) * fraction,
      height: 1, page: flow.page.index, ...flow.location, ...(typeof color === 'string' ? { color } : {}) })
  }
}
