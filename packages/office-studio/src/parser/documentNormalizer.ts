import { createBlankDocument, type DocumentBlock, type DocumentModel, type DocumentSectionBreakType, type InlineRun } from '../domain/model.ts'
import type { ParseEnvelope } from './contracts.ts'
import { array, finite, positive, record, text } from './validation.ts'

const pointToPixel = (value: number): number => value * 96 / 72
function normalizeRun(value: unknown, id: string): InlineRun {
  const raw = record(value)
  const style = record(raw.style)
  const verticalAlign = normalizeVerticalAlign(style.verticalAlign ?? style.vertAlign ?? raw.verticalAlign)
  return {
    id, text: text(raw.text), bold: style.bold === true, italic: style.italic === true,
    underline: style.underline === true, strike: style.strike === true,
    ...(verticalAlign ? { verticalAlign } : {}),
    color: text(style.color, '#1f2937'), fontFamily: text(style.fontFamily, 'Aptos'), fontSize: positive(style.fontSizePt, 11),
    metadata: { source: raw, link: raw.link, highlight: style.highlight, noteReference: raw.noteReference },
  }
}
function normalizeVerticalAlign(value: unknown): InlineRun['verticalAlign'] {
  if (value === 'sup' || value === 'super' || value === 'superscript') return 'superscript'
  if (value === 'sub' || value === 'subscript') return 'subscript'
  return undefined
}
function normalizeBlock(value: unknown, fallbackId: string): DocumentBlock {
  const raw = record(value)
  const details = record(raw.details)
  const sectionBreakData = record(details.sectionBreak)
  const sectionBreakType = sectionBreakData.type ?? record(raw.sectionBreak).type
  const sectionBreak = sectionBreakType === 'nextPage' || sectionBreakType === 'continuous' || sectionBreakType === 'nextColumn' || sectionBreakType === 'evenPage' || sectionBreakType === 'oddPage'
    ? { type: sectionBreakType as DocumentSectionBreakType }
    : undefined
  const pageBreak = details.pageBreak === true || raw.pageBreak === true || raw.type === 'page_break' || raw.type === 'pageBreak'
  const columnBreak = details.columnBreak === true || raw.columnBreak === true || raw.type === 'column_break' || raw.type === 'columnBreak'
  const style = record(details.style)
  const indent = record(style.indent)
  const spacing = record(style.spacing)
  const id = text(raw.id, fallbackId)
  const runs = array(details.inlines).filter((inline) => {
    const type = record(inline).type
    return type === 'text' || type === 'form_field'
  }).map((inline, index) => normalizeRun(inline, id+'-run-'+index))
  if (!runs.length) runs.push(normalizeRun({ text: text(raw.content) }, id+'-run-0'))
  const level = Math.min(6, Math.max(0, finite(details.headingLevel, finite(style.headingLevel))))
  const numbering = record(details.numbering)
  const listLabel = text(details.listLabel).trimStart()
  const checklistLabel = listLabel.startsWith('☐') || listLabel.startsWith('☑')
  const type = sectionBreak ? 'sectionBreak' : columnBreak ? 'columnBreak' : pageBreak ? 'pageBreak' : raw.type === 'table' ? 'table' : raw.type === 'figure' ? 'image'
    : raw.type === 'heading' || raw.type === 'section_heading' || level > 0 ? 'heading'
      : checklistLabel ? 'checklist'
        : Object.keys(numbering).length ? (listLabel.match(/^[\d一二三四五六七八九十]/) ? 'numbered' : 'bullet') : 'paragraph'
  const block: DocumentBlock = {
    id, type, level, runs: ['sectionBreak', 'columnBreak', 'pageBreak'].includes(type) ? [] : runs,
    style: {
      align: style.align === 'center' || style.align === 'right' || style.align === 'justify' ? style.align : 'left',
      indent: pointToPixel(finite(indent.leftPt, finite(indent.left) / 20)),
      lineHeight: positive(spacing.lineMultiplier, 1.5),
      spaceBefore: pointToPixel(finite(spacing.beforePt, finite(spacing.before) / 20)),
      spaceAfter: pointToPixel(finite(spacing.afterPt, finite(spacing.after, 120) / 20)),
    },
    metadata: { source: raw, sourceRefs: raw.sourceRefs, style, listLabel: details.listLabel, numbering },
  }
  const numberingId = numbering.numId
  if (type === 'checklist') block.checklist = {
    listId: typeof numberingId === 'string' || typeof numberingId === 'number' ? String(numberingId) : id,
    checked: listLabel.startsWith('☑'),
  }
  if (sectionBreak) block.sectionBreak = sectionBreak
  if (type === 'image') {
    const image = record(details.image)
    const src = text(details.src, text(details.asset, text(image.src)))
    block.metadata.image = {
      src, width: positive(details.widthPx, positive(image.widthPx, 320)),
      height: positive(details.heightPx, positive(image.heightPx, 180)),
      alt: text(details.alt, text(details.description, text(raw.content))),
    }
    block.metadata.src = src
    block.metadata.width = positive(details.widthPx, 320)
    block.metadata.height = positive(details.heightPx, 180)
  }
  if (type === 'table') {
    const sourceCells: unknown[] = []
    for (const child of array(raw.children)) {
      const item = record(child)
      if (item.type === 'table_head') sourceCells.push(...array(item.children))
      else if (item.type === 'table_cell') sourceCells.push(item)
    }
    const rows = new Map<number, Array<Record<string, unknown>>>()
    for (const [cellIndex, source] of sourceCells.entries()) {
      const cell = record(source)
      const cellDetails = record(cell.details)
      const rowIndex = Math.max(0, finite(cellDetails.rowIndex))
      const row = rows.get(rowIndex) ?? []
      const cellId = text(cell.id, id+'-cell-'+cellIndex)
      let blocks = array(cell.children).map((child, index) => normalizeBlock(child, cellId+'-block-'+index))
      if (!blocks.length) blocks = [normalizeBlock({ id: cellId+'-text', content: text(cell.content), type: 'text' }, cellId+'-text')]
      row.push({
        id: cellId, blocks, colSpan: positive(cellDetails.columnSpan, 1), rowSpan: positive(cellDetails.rowSpan, 1),
        columnIndex: finite(cellDetails.columnIndex), style: record(cellDetails.style),
      })
      rows.set(rowIndex, row)
    }
    block.metadata.table = {
      rows: [...rows.entries()].sort(([a], [b]) => a-b).map(([rowIndex, cells]) => ({
        id: id+'-row-'+rowIndex, cells: cells.sort((a, b) => finite(a.columnIndex)-finite(b.columnIndex)),
      })),
      style, rowStyles: details.rowStyles,
    }
    block.runs = []
  }
  return block
}
export function normalizeDocument(envelope: ParseEnvelope): DocumentModel {
  const data = envelope.payload.data
  const result = createBlankDocument()
  const id = text(data.documentId, crypto.randomUUID())
  result.id = id
  result.name = envelope.fileName
  result.blocks = []
  const stories: Record<string, unknown[]> = {}
  for (const [index, source] of array(data.blocks).entries()) {
    const raw = record(source)
    const story = text(record(array(raw.sourceRefs)[0]).story, 'body')
    if (story !== 'body' || raw.type === 'header' || raw.type === 'footer' || raw.type === 'watermark') {
      const key = raw.type === 'header' || raw.type === 'footer' ? text(raw.type) : story
      ;(stories[key] ??= []).push(raw)
      continue
    }
    const block = normalizeBlock(raw, id+'-block-'+index)
    result.blocks.push(block)
    if (block.type !== 'table') for (const [childIndex, child] of array(raw.children).entries()) {
      if (record(child).type === 'figure') result.blocks.push(normalizeBlock(child, block.id+'-object-'+childIndex))
    }
  }
  if (!result.blocks.length) result.blocks = createBlankDocument().blocks
  result.modified = false
  result.wordCount = result.blocks.reduce((count, block) => count+block.runs.reduce((sum, run) => sum+run.text.length, 0), 0)
  result.metadata = {
    source: data, stories, format: envelope.format, parserVersion: envelope.parserVersion,
    geometryStatus: data.geometryStatus,
    geometryNote: 'StructuredDocument 目前未携带完整节/纸张几何；使用默认页面并保留原数据，待补齐解析契约。',
  }
  return result
}
