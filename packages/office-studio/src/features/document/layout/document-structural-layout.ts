import type { DocumentBlock, DocumentColumnGroup } from '../../../domain/model'
import { runText } from '../model/document-model'
import { documentTextFieldParagraphs } from '../model/document-text-fields'
import { isDocumentStructuredBlock, DOCUMENT_COLUMN_HORIZONTAL_PADDING, DOCUMENT_COLUMN_MIN_HEIGHT } from '../model/document-structural-blocks'
import type { DocumentSectionFlow } from './document-section-flow'
import type { DocumentLayout, LayoutColumn, LayoutFlowLocation, LayoutPage, LayoutRect } from './document-layout-types'
import { documentQuoteLineColor } from '../model/document-structured-appearance'

type AppendChildren = (blocks: DocumentBlock[], flow: DocumentSectionFlow) => void
type Positioned = LayoutRect & Partial<LayoutFlowLocation> & { page: number }
const locationKey = (value: Positioned | (LayoutFlowLocation & { page: number })) => `${value.page}:${value.sectionFragment ?? 0}:${value.column ?? 0}`
const cursorLocation = (flow: DocumentSectionFlow) => ({ ...flow.location, page: flow.page.index, cursor: flow.cursor })
const offsets = (layout: DocumentLayout) => ({ lines: layout.lines.length, cells: layout.cells.length, images: layout.images.length,
  dividers: layout.dividers.length, fragments: layout.tableFragments?.length ?? 0, containers: layout.containers?.length ?? 0 })

export { DOCUMENT_COLUMN_HORIZONTAL_PADDING, DOCUMENT_COLUMN_MIN_HEIGHT } from '../model/document-structural-blocks'
export function documentColumnMetrics(group: DocumentColumnGroup) {
  return { horizontalPadding: group.horizontalPadding ?? DOCUMENT_COLUMN_HORIZONTAL_PADDING, minHeight: group.minHeight ?? DOCUMENT_COLUMN_MIN_HEIGHT }
}

export function documentColumnWidths(group: DocumentColumnGroup, available: number): { widths: number[]; gap: number; stacked: boolean } {
  const count = group.columns.length, total = group.columns.reduce((sum, column) => sum + column.widthRatio, 0)
  const availableWidth = Math.max(0, available), gap = group.gap, width = Math.max(0, availableWidth - gap * Math.max(0, count - 1))
  const minimums = group.columns.map(column => column.minWidth ?? 0), minimumTotal = minimums.reduce((sum, minimum) => sum + minimum, 0)
  if (group.responsive === 'stack' && minimumTotal + gap * Math.max(0, count - 1) > availableWidth) {
    return { widths: group.columns.map(() => availableWidth), gap, stacked: true }
  }
  const widths = group.columns.map((column, index) => Math.max(minimums[index]!, width * column.widthRatio / total))
  const overflow = widths.reduce((sum, value) => sum + value, 0) - width
  const shrink = widths.map((value, index) => Math.max(0, value - minimums[index]!)), totalShrink = shrink.reduce((sum, value) => sum + value, 0)
  const applied = Math.min(Math.max(0, overflow), totalShrink)
  if (applied > 0 && totalShrink > 0) for (let index = 0; index < count; index++) widths[index] = widths[index]! - applied * shrink[index]! / totalShrink
  return { widths, gap, stacked: false }
}

function appendContainerDecorations(block: DocumentBlock, flow: DocumentSectionFlow, result: DocumentLayout,
  start: ReturnType<typeof cursorLocation>, from: ReturnType<typeof offsets>): void {
  const items: Positioned[] = [...result.lines.slice(from.lines), ...result.cells.slice(from.cells), ...result.images.slice(from.images),
    ...result.dividers.slice(from.dividers), ...(result.tableFragments?.slice(from.fragments) ?? [])]
  const keys = [...new Set([locationKey(start), ...items.map(locationKey), locationKey(cursorLocation(flow))])]
  const end = cursorLocation(flow), decorations = result.containers ?? (result.containers = [])
  for (const [index, key] of keys.entries()) {
    const values = key.split(':').map(Number), page = flow.pages.find(page => page.index === values[0]), section = page?.sections.find(section => section.fragment === values[1])
    const column = section?.columns.find(column => column.index === values[2])
    if (!page || !section || !column) continue
    const top = key === locationKey(start) ? start.cursor : column.y, bottom = key === locationKey(end) ? end.cursor : column.y + column.height
    decorations.push({ blockId: block.id, kind: block.type as 'quote' | 'callout' | 'code' | 'columns', x: column.x, y: top,
      width: column.width, height: Math.max(1, bottom - top), page: page.index, sectionId: section.sectionId, sectionFragment: section.fragment,
      column: column.index, first: index === 0, last: index === keys.length - 1, ...(block.callout ? { callout: { ...block.callout } } : {}),
      ...(block.code ? { codeWrap: block.code.wrap } : {}),
      ...(block.columns ? { columnGroup: { columnIds: block.columns.columns.map(value => value.id), ...documentColumnWidths(block.columns, column.width), ...documentColumnMetrics(block.columns) } } : {}),
      ...(block.type === 'quote' ? { quoteLineColor: documentQuoteLineColor(block) } : {}) })
  }
}

/** Container decoration, text and hit targets all use the parent's physical page cursor. */
export function layoutDocumentStructuredBlock(block: DocumentBlock, flow: DocumentSectionFlow, result: DocumentLayout, append: AppendChildren): void {
  if (!isDocumentStructuredBlock(block)) throw new TypeError('结构排版需要一个真实容器。')
  if (block.type === 'columns') { layoutColumns(block, flow, result, append); return }
  const from = offsets(result), start = cursorLocation(flow), config = block.callout
  const left = config ? config.paddingLeft + (config.showIcon ? config.iconSize + config.iconGap : 0)
    : block.type === 'code' ? block.code?.showLineNumbers ? 44 : 12 : 18
  const right = config?.paddingRight ?? 12, top = config?.paddingTop ?? 8,
    bottom = config?.paddingBottom ?? (block.type === 'code' && !block.code?.wrap ? 24 : 8)
  flow.withColumnInsets(left, right, () => {
    flow.addSpace(top); append(block.children!, flow); flow.addSpace(bottom)
  })
  if (block.type === 'code') {
    const paragraphs = documentTextFieldParagraphs({ blocks: block.children! }), byId = new Map(paragraphs.map(paragraph => [paragraph.block.id, paragraph]))
    const source = paragraphs.map(paragraph => runText(paragraph.block.runs)).join('\n')
    for (const line of result.lines.slice(from.lines)) {
      if (line.codeBlock?.containerId && line.codeBlock.containerId !== block.id) continue
      const paragraph = byId.get(line.target.cellBlockId ?? line.target.blockId)
      if (!paragraph) continue
      const offset = paragraph.absoluteStart + line.start, lineNumber = 1 + (source.slice(0, offset).match(/\n/g)?.length ?? 0)
      const previousCharacter = offset > 0 ? source[offset - 1] : undefined
      line.codeBlock = { language: block.code!.language, lineNumber, containerId: block.id, sourceOffset: paragraph.absoluteStart,
        showLineNumbers: block.code!.showLineNumbers, numberVisible: offset === 0 || previousCharacter === '\n' }
    }
  }
  appendContainerDecorations(block, flow, result, start, from)
}

interface ParentFrame { page: LayoutPage; column: LayoutColumn; location: LayoutFlowLocation; top: number; x: number; width: number; height: number }

function layoutColumns(block: DocumentBlock, flow: DocumentSectionFlow, result: DocumentLayout, append: AppendChildren): void {
  const group = block.columns!, metrics = documentColumnMetrics(group)
  flow.ensureRoom(Math.min(metrics.minHeight, flow.fullPageContentHeight))
  const prepared = documentColumnWidths(group, flow.column.width), fromGroup = offsets(result), groupStart = cursorLocation(flow)
  if (prepared.stacked) {
    for (const [index, column] of group.columns.entries()) {
      if (index) flow.addSpace(group.gap)
      flow.ensureRoom(Math.min(metrics.minHeight, flow.fullPageContentHeight))
      const start = cursorLocation(flow)
      flow.withColumnInsets(metrics.horizontalPadding, metrics.horizontalPadding, () => append(column.blocks, flow))
      if (locationKey(start) === locationKey(cursorLocation(flow))) flow.consume(Math.max(0, metrics.minHeight - (flow.cursor - start.cursor)))
    }
    appendContainerDecorations(block, flow, result, groupStart, fromGroup)
    return
  }
  const frames: ParentFrame[] = [], endings: Array<{ frame: number; used: number }> = []
  const capture = () => frames.push({ page: flow.page, column: flow.column, location: { ...flow.location }, top: flow.cursor,
    x: flow.column.x, width: flow.column.width, height: Math.max(1, flow.column.y + flow.column.height - flow.cursor) })
  capture()
  const frameAt = (index: number) => {
    while (frames.length <= index) { flow.nextColumn(); capture() }
    return frames[index]!
  }
  for (const [columnIndex, column] of group.columns.entries()) {
    const from = offsets(result)
    const child = flow.fork(index => {
      const frame = frameAt(index), dimensions = documentColumnWidths(group, frame.width)
      const x = frame.x + dimensions.widths.slice(0, columnIndex).reduce((sum, width) => sum + width + dimensions.gap, 0), width = dimensions.widths[columnIndex]!
      return { ...frame.page, content: { x, y: frame.top, width, height: frame.height }, columnCount: 1, columnGap: 0,
        maximumContentHeight: flow.fullPageContentHeight }
    })
    child.withColumnInsets(metrics.horizontalPadding, metrics.horizontalPadding, () => append(column.blocks, child))
    if (child.pages.length === 1) child.consume(Math.max(0, metrics.minHeight - (child.cursor - child.page.content.y)))
    for (const page of child.pages) {
      const local = child.pages.indexOf(page), frame = frameAt(local), used = Math.max(...page.sections.map(section => section.usedHeight), 0)
      frame.column.usedHeight = Math.max(frame.column.usedHeight, frame.top - frame.column.y + used)
      const section = frame.page.sections.find(section => section.fragment === frame.location.sectionFragment)!
      section.usedHeight = Math.max(...section.columns.map(column => column.usedHeight))
      if (local === child.pages.length - 1) endings.push({ frame: local, used })
    }
    // Child section/column indices are transient; all editor addresses refer to parent geometry.
    const remap = (items: Array<Positioned & { sectionId?: string; sectionFragment?: number; column?: number }>) => {
      for (const item of items) {
        const local = child.pages.findIndex(page => page.index === item.page && item.x >= page.content.x - 1e-6 && item.x <= page.content.x + page.content.width + 1e-6
          && item.y >= page.content.y - 1e-6 && item.y <= page.content.y + page.content.height + 1e-6)
        const frame = frameAt(Math.max(0, local)); Object.assign(item, frame.location)
      }
    }
    remap(result.lines.slice(from.lines)); remap(result.cells.slice(from.cells)); remap(result.images.slice(from.images)); remap(result.dividers.slice(from.dividers))
    remap(result.tableFragments?.slice(from.fragments) ?? []); remap(result.containers?.slice(from.containers) ?? [])
  }
  const last = frames.length - 1, frame = frames[last]!, used = Math.max(0, ...endings.filter(ending => ending.frame === last).map(ending => ending.used))
  flow.cursor = frame.top; flow.consume(used)
  const containers = result.containers ?? (result.containers = [])
  for (const [index, entry] of frames.entries()) {
    const dimensions = documentColumnWidths(group, entry.width)
    containers.push({ blockId: block.id, kind: 'columns', x: entry.x, y: entry.top, width: entry.width,
      height: Math.max(1, index === last ? used : entry.height), page: entry.page.index, ...entry.location,
      first: index === 0, last: index === last,
      columnGroup: { columnIds: group.columns.map(column => column.id), ...dimensions, ...metrics } })
  }
}
