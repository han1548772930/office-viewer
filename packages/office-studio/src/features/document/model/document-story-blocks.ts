import type { DocumentBlock, DocumentDefaultStyle, DocumentPageField, DocumentSectionHeaderFooter, DocumentStoryContent, DocumentStoryParagraph } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { documentDefaultStyle } from './document-default-style'
import { createParagraph, createRun, runText, tableForBlock, type DocumentRun } from './document-model'
import { documentTableCellContentContainers } from './document-table-continuations'
import { documentTableGrid } from './document-table-grid'
import { documentStructuredChildContainers, validateDocumentStructuredBlock } from './document-structural-blocks'
import { splitDocumentRunRange } from './document-run-range'
import { documentTextFieldParagraphs, documentTextFieldRanges, setDocumentTextFields, type DocumentTextFieldRange } from './document-text-fields'
import { documentBlockHostId, type DocumentAuxiliaryBlockHost } from './document-block-host'
import { documentStoryFieldDisplayRuns } from './document-story'

const PAGE_FIELD_TYPES = new Set<string>(['PAGE', 'NUMPAGES', 'DATE', 'TIME'])

export interface DocumentStoryBlockProjection {
  text: string
  displayText: string
  content: DocumentStoryContent
  pageFields: DocumentPageField[]
  fields: DocumentTextFieldRange[]
}

function projectedParagraphs(text: string, source: ReturnType<typeof documentTextFieldParagraphs>, removed: readonly DocumentTextFieldRange[]): DocumentStoryParagraph[] {
  const toDisplay = (offset: number) => {
    let result = offset
    for (const range of removed) if (range.absoluteStart <= result) result += range.absoluteEnd - range.absoluteStart
    return result
  }
  let next = 0
  return text.split('\n').map(part => {
    const displayOffset = toDisplay(next)
    const paragraph = source.find(candidate => displayOffset >= candidate.absoluteStart && displayOffset <= candidate.absoluteEnd) ?? source.at(-1)
    const startOffset = next, endOffset = next + part.length
    next = endOffset + 1
    return { startOffset, endOffset,
      style: { ...(paragraph?.block.style ?? documentDefaultStyle().paragraph) },
      ...(paragraph ? { metadata: { ...paragraph.block.metadata, officeStoryBlockProjection: { version: 1, blockId: paragraph.block.id } } } : {}) }
  })
}

function sliceProjectionRuns(runs: readonly DocumentRun[], start: number, end: number): DocumentRun[] {
  splitDocumentRunRange(runs, start, end)
  const result: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    const from = offset, to = from + run.text.length
    offset = to
    const a = Math.max(from, start), b = Math.min(to, end)
    if (a < b) result.push({ ...run, id: a === from && b === to ? run.id : `${run.id}:projection:${a - from}:${b - from}`, text: run.text.slice(a - from, b - from) })
    else if (!run.text.length && from >= start && (from < end || start === end && from === start)) result.push(run)
  }
  return result
}

/** Ordinary field caches use the same paragraph-local UTF-16 coordinates in every host. */
export function validateDocumentStoryBlocks(blocks: readonly DocumentBlock[]): void {
  if (!Array.isArray(blocks)) throw new TypeError('内容区块树无效。')
  const blockIds = new Set<string>(), runIds = new Set<string>(), rowIds = new Set<string>(), cellIds = new Set<string>(), columnIds = new Set<string>()
  const visit = (content: readonly DocumentBlock[], depth: number) => {
    if (depth > 64) throw new RangeError('内容区嵌套深度超过 64 层。')
    for (const block of content) {
      if (!block || typeof block.id !== 'string' || !block.id || blockIds.has(block.id) || !Array.isArray(block.runs)
        || !block.metadata || typeof block.metadata !== 'object' || Array.isArray(block.metadata)) throw new TypeError('内容区块标识或结构无效。')
      blockIds.add(block.id)
      validateDocumentStructuredBlock(block)
      documentDefaultStyle({ defaultStyle: { version: 1, text: documentDefaultStyle().text, paragraph: block.style } })
      for (const run of block.runs) {
        if (!run || typeof run.id !== 'string' || !run.id || runIds.has(run.id) || typeof run.text !== 'string') throw new TypeError('内容区字符标识或文本无效。')
        runIds.add(run.id)
        documentDefaultStyle({ defaultStyle: { version: 1, text: run, paragraph: block.style } })
      }
      for (const column of block.columns?.columns ?? []) {
        if (columnIds.has(column.id)) throw new TypeError(`内容区栏标识重复：${column.id}`)
        columnIds.add(column.id)
      }
      for (const container of documentStructuredChildContainers(block)) visit(container, depth + 1)
      const table = tableForBlock(block)
      if (block.type === 'table' && !table) throw new TypeError('内容区表格模型无效。')
      if (!table) continue
      documentTableGrid(table)
      for (const row of table.rows) {
        if (rowIds.has(row.id)) throw new TypeError(`内容区表格行标识重复：${row.id}`)
        rowIds.add(row.id)
        for (const cell of row.cells) {
          if (cellIds.has(cell.id)) throw new TypeError(`内容区单元格标识重复：${cell.id}`)
          cellIds.add(cell.id)
          for (const container of documentTableCellContentContainers(cell)) visit(container, depth + 1)
        }
      }
    }
  }
  visit(blocks, 0)
  documentTextFieldRanges({ blocks: blocks as DocumentBlock[] })
}

/** Compatibility text and fields are derived from the block tree; no flattened body is saved as a second authority. */
export function projectDocumentStoryBlocks(content: DocumentStoryContent, mode: 'literal' | 'display' = 'literal'): DocumentStoryBlockProjection {
  if (content.blocks === undefined) throw new TypeError('内容区尚未建立块树。')
  validateDocumentStoryBlocks(content.blocks)
  const source = documentTextFieldParagraphs({ blocks: content.blocks })
  const runs: DocumentRun[] = []
  for (const [index, paragraph] of source.entries()) {
    if (index) runs.push({ ...createRun('\n', paragraph.block.runs[0] ?? documentDefaultStyle().text), id: `story-projection-separator:${paragraph.block.id}` })
    runs.push(...paragraph.block.runs as DocumentRun[])
  }
  const displayText = runText(runs), ranges = documentTextFieldRanges({ blocks: content.blocks })
  const removed: DocumentTextFieldRange[] = []
  if (mode === 'literal') for (const range of [...ranges].sort((a, b) => a.absoluteStart - b.absoluteStart || b.absoluteEnd - a.absoluteEnd)) {
    if (!PAGE_FIELD_TYPES.has(range.field.field.toUpperCase()) || removed.some(parent => range.absoluteStart >= parent.absoluteStart && range.absoluteEnd <= parent.absoluteEnd
      && (parent.absoluteStart < parent.absoluteEnd || parent.absoluteStart < range.absoluteStart || parent.absoluteEnd > range.absoluteEnd))) continue
    // A page field contained in a different native field remains part of that parent's cached result.
    if (ranges.some(parent => parent.field.id !== range.field.id && parent.absoluteStart <= range.absoluteStart && parent.absoluteEnd >= range.absoluteEnd
      && (parent.absoluteStart < range.absoluteStart || parent.absoluteEnd > range.absoluteEnd))) continue
    removed.push(range)
  }
  const projectedRuns: DocumentRun[] = [], pageFields: DocumentPageField[] = []
  let cursor = 0, removedLength = 0
  for (const range of removed) {
    projectedRuns.push(...sliceProjectionRuns(runs, cursor, range.absoluteStart))
    const cache = sliceProjectionRuns(runs, range.absoluteStart, range.absoluteEnd)
    const textStyle = cache[0] ? documentDefaultStyle({ defaultStyle: { version: 1, text: cache[0], paragraph: range.block.style } }).text : undefined
    pageFields.push({ story: 'header', field: range.field.field.toUpperCase() as DocumentPageField['field'],
      offset: range.absoluteStart - removedLength, instruction: range.field.instruction, cachedResult: displayText.slice(range.absoluteStart, range.absoluteEnd),
      locked: range.field.locked, containsNestedFields: range.field.containsNestedFields,
      textStyle,
      metadata: { ...range.field.metadata, officeStoryBlockField: { version: 1, blockId: range.block.id, fieldId: range.field.id },
        ...(textStyle ? { officeStoryStyleBaseline: { ...textStyle } } : {}),
        officeStoryFieldCache: { version: 1, runs: cloneModel(cache) } } })
    removedLength += range.absoluteEnd - range.absoluteStart
    cursor = range.absoluteEnd
  }
  projectedRuns.push(...sliceProjectionRuns(runs, cursor, displayText.length))
  if (!projectedRuns.length) projectedRuns.push({ ...createRun('', source[0]?.block.runs[0] ?? documentDefaultStyle().text), id: `story-projection-empty:${source[0]?.block.id ?? 'empty'}` })
  const text = runText(projectedRuns)
  return { text, displayText, content: { ...content, runs: projectedRuns, paragraphs: projectedParagraphs(text, source, removed) }, pageFields, fields: removed }
}

/** Map result coordinates to the old bridge: each supported page field occupies one input object. */
export function documentStoryDisplayToInputOffset(projection: DocumentStoryBlockProjection, offset: number, affinity: 'start' | 'end' = 'start'): number {
  let removed = 0
  for (const range of projection.fields) {
    if (offset < range.absoluteStart) break
    if (offset < range.absoluteEnd || range.absoluteStart === range.absoluteEnd && offset === range.absoluteStart) {
      const zeroFieldsAtBoundary = range.absoluteStart === range.absoluteEnd
        ? projection.fields.filter(candidate => candidate.absoluteStart === offset && candidate.absoluteEnd === offset).length : 1
      return range.absoluteStart - removed + (affinity === 'end' ? zeroFieldsAtBoundary : 0)
    }
    removed += range.absoluteEnd - range.absoluteStart - 1
  }
  return offset - removed
}

export function documentStoryInputToDisplayOffset(projection: DocumentStoryBlockProjection, offset: number, affinity: 'start' | 'end' = 'start'): number {
  let removed = 0
  for (const range of projection.fields) {
    const position = range.absoluteStart - removed
    if (offset < position) break
    if (offset === position) return affinity === 'end' ? range.absoluteEnd : range.absoluteStart
    removed += range.absoluteEnd - range.absoluteStart - 1
  }
  return offset + removed
}

/** Legacy pageFields.offset excludes cached results and input-object placeholders. */
export function documentStoryLiteralToDisplayOffset(projection: DocumentStoryBlockProjection, offset: number, affinity: 'start' | 'end' = 'start'): number {
  let removed = 0
  for (const range of projection.fields) {
    const position = range.absoluteStart - removed
    if (offset < position || offset === position && affinity === 'start') break
    removed += range.absoluteEnd - range.absoluteStart
  }
  return offset + removed
}

export function projectDocumentSectionStories(section: DocumentSectionHeaderFooter): DocumentSectionHeaderFooter {
  const stories = Object.entries(section.storyContent ?? {}).filter(([, content]) => content?.blocks !== undefined)
  if (!stories.length) return section
  const next: DocumentSectionHeaderFooter = { ...section, storyContent: { ...section.storyContent }, pageFields: [...(section.pageFields ?? [])] }
  for (const [key, content] of stories) {
    const story = key as DocumentPageField['story'], projection = projectDocumentStoryBlocks(content!)
    next[story] = projection.text
    next.storyContent![story] = projection.content
    next.pageFields = [...next.pageFields!.filter(field => field.story !== story), ...projection.pageFields.map(field => ({ ...field, story }))]
  }
  return next
}

/** Legacy stories gain stable paragraph IDs and real cached-result field ranges on first block edit. */
export function migrateDocumentStoryBlocks(content: DocumentStoryContent, fields: readonly DocumentPageField[], host: DocumentAuxiliaryBlockHost,
  defaults: DocumentDefaultStyle): DocumentStoryContent {
  if (content.blocks !== undefined) return projectDocumentStoryBlocks(content, host.kind === 'story' ? 'literal' : 'display').content
  const key = encodeURIComponent(documentBlockHostId(host)), text = runText(content.runs)
  const paragraphs = content.paragraphs.length ? content.paragraphs : [{ startOffset: 0, endOffset: text.length, style: defaults.paragraph }]
  const blocks = paragraphs.map((paragraph, index) => {
    const block = createParagraph('', defaults.text)
    block.id = `legacy-story-block:${key}:${index}`
    block.style = cloneModel(paragraph.style)
    block.metadata = { ...paragraph.metadata }
    const localFields = fields.map((field, fieldIndex) => ({ field, fieldIndex, offset: field.offset ?? text.length }))
      .filter(entry => entry.offset >= paragraph.startOffset && entry.offset <= paragraph.endOffset)
      .sort((a, b) => a.offset - b.offset || a.fieldIndex - b.fieldIndex)
    const runs: DocumentRun[] = [], definitions: import('./document-text-fields').DocumentTextField[] = []
    let cursor = paragraph.startOffset, length = 0
    for (const { field, fieldIndex, offset } of localFields) {
      const prefix = sliceProjectionRuns(content.runs, cursor, offset)
      runs.push(...prefix); length += runText(prefix).length
      const cache = documentStoryFieldDisplayRuns(field, field.cachedResult ?? '', defaults.text)
      if (runText(cache) !== (field.cachedResult ?? '')) throw new RangeError('内容区字段缓存与文本不一致。')
      const startOffset = length
      runs.push(...cache); length += runText(cache).length
      definitions.push({ id: `legacy-story-field:${key}:${fieldIndex}`, field: field.field, instruction: field.instruction,
        cachedResult: field.cachedResult ?? '', startOffset, endOffset: length, locked: field.locked, containsNestedFields: field.containsNestedFields,
        metadata: { ...field.metadata, officeStoryPageFieldSource: { version: 1, index: fieldIndex } } })
      cursor = offset
    }
    runs.push(...sliceProjectionRuns(content.runs, cursor, paragraph.endOffset))
    if (!runs.length) runs.push(createRun('', defaults.text))
    block.runs = runs.map((run, runIndex) => ({ ...run, id: `legacy-story-run:${key}:${index}:${runIndex}`,
      metadata: { ...run.metadata, officeStoryRunSourceId: run.id } }))
    setDocumentTextFields(block, definitions)
    return block
  })
  return projectDocumentStoryBlocks({ ...content, blocks }, host.kind === 'story' ? 'literal' : 'display').content
}
