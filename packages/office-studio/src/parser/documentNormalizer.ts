import { createBlankDocument, type DocumentBlock, type DocumentCommentMessage, type DocumentModel, type DocumentPageField, type DocumentSectionBreakType, type DocumentSectionHeaderFooter, type DocumentThreadComment, type InlineRun } from '../domain/model.ts'
import type { ParseEnvelope } from './contracts.ts'
import { array, finite, positive, record, text } from './validation.ts'
import { normalizeDocumentNotes } from '../features/document/notes/document-notes.ts'
import { documentNumberingIndex } from './documentNumbering.ts'

const pointToPixel = (value: number): number => value * 96 / 72
function normalizeRun(value: unknown, id: string): InlineRun {
  const raw = record(value)
  const style = record(raw.style)
  const verticalAlign = normalizeVerticalAlign(style.verticalAlign ?? style.vertAlign ?? raw.verticalAlign)
  const eastAsiaFontFamily = text(style.fontFamilyEastAsia) || text(style.resolvedFontFamilyEastAsia)
  return {
    id, text: text(raw.text), bold: style.bold === true, italic: style.italic === true,
    underline: style.underline === true, strike: style.strike === true,
    ...(verticalAlign ? { verticalAlign } : {}),
    ...(eastAsiaFontFamily ? { eastAsiaFontFamily } : {}),
    color: text(style.color, '#1f2937'), fontFamily: text(style.fontFamilyAscii) || text(style.resolvedFontFamilyAscii) || text(style.fontFamily, 'Aptos'), fontSize: positive(style.fontSizePt, 11),
    metadata: { source: raw, link: raw.link, highlight: style.highlight, noteReference: raw.noteReference },
  }
}
function normalizeVerticalAlign(value: unknown): InlineRun['verticalAlign'] {
  if (value === 'sup' || value === 'super' || value === 'superscript') return 'superscript'
  if (value === 'sub' || value === 'subscript') return 'subscript'
  return undefined
}
function normalizeBlock(value: unknown, fallbackId: string, lists: ReturnType<typeof documentNumberingIndex>): DocumentBlock {
  const raw = record(value)
  const details = record(raw.details)
  const sectionBreakData = record(details.sectionBreak)
  const sectionBreakType = sectionBreakData.type ?? record(raw.sectionBreak).type
  const sectionBreak = sectionBreakType === 'unspecified' || sectionBreakType === 'nextPage' || sectionBreakType === 'continuous' || sectionBreakType === 'nextColumn' || sectionBreakType === 'evenPage' || sectionBreakType === 'oddPage'
    ? { type: sectionBreakType as DocumentSectionBreakType }
    : undefined
  const pageBreak = details.pageBreak === true || raw.pageBreak === true || raw.type === 'page_break' || raw.type === 'pageBreak'
  const columnBreak = details.columnBreak === true || raw.columnBreak === true || raw.type === 'column_break' || raw.type === 'columnBreak'
  const style = record(details.style)
  const styleId = (text(details.styleId) || text(style.styleId)).toLowerCase()
  const titleStyle = styleId === 'title' || styleId === 'subtitle' ? styleId : undefined
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
  const listLevel = Math.min(8, Math.max(0, Math.trunc(finite(numbering.ilvl))))
  const definition = typeof numbering.numId === 'number' ? lists.get(numbering.numId) : undefined
  const format = definition?.levels.get(listLevel)?.format
  const listLabel = text(details.listLabel).trimStart()
  const checklistLabel = listLabel.startsWith('☐') || listLabel.startsWith('☑')
  const type = sectionBreak ? 'sectionBreak' : columnBreak ? 'columnBreak' : pageBreak ? 'pageBreak' : raw.type === 'table' ? 'table' : raw.type === 'figure' ? 'image'
    : titleStyle ? 'paragraph' : raw.type === 'heading' || raw.type === 'section_heading' || level > 0 ? 'heading'
      : checklistLabel ? 'checklist'
        : Object.keys(numbering).length && numbering.numId !== 0 && format !== 'none'
          ? (format ? format !== 'bullet' : /^[\d一二三四五六七八九十]|^[a-z]+[.)]/i.test(listLabel)) ? 'numbered' : 'bullet'
          : 'paragraph'
  const block: DocumentBlock = {
    id, type, level, runs: ['sectionBreak', 'columnBreak', 'pageBreak'].includes(type) ? [] : runs,
    style: {
      align: style.align === 'center' || style.align === 'right' || style.align === 'justify' ? style.align : 'left',
      indent: pointToPixel(finite(indent.leftPt, finite(indent.left) / 20)),
      lineHeight: positive(spacing.lineMultiplier, 1.5),
      spaceBefore: pointToPixel(finite(spacing.beforePt, finite(spacing.before) / 20)),
      spaceAfter: pointToPixel(finite(spacing.afterPt, finite(spacing.after, 120) / 20)),
    },
    metadata: { source: raw, sourceRefs: raw.sourceRefs, style, listLabel: details.listLabel, numbering, ...(titleStyle ? { namedStyleType: titleStyle } : {}) },
  }
  const taggedFields = array(raw.children).map((child) => record(record(child).details).field).map(record)
  const codeBookmark = array(details.bookmarks).map((bookmark) => text(record(bookmark).name)).find((name) => name.startsWith('office_code_v1_'))
  const codeTag = codeBookmark?.replace(/^office_code_v1_/, 'office-code:v1:').replaceAll('_', ':')
    ?? taggedFields.map((field) => text(field.tag)).find((tag) => tag.startsWith('office-code:v1:'))
  if (codeTag) {
    const [, , languageValue, lineNumberValue] = codeTag.split(':')
    const languages = ['plain', 'javascript', 'typescript', 'json', 'css', 'html', 'python', 'sql', 'rust']
    const language = languages.includes(languageValue ?? '') ? languageValue! : 'plain'
    const source = text(raw.content)
    block.metadata.codeBlock = { source, language, lineNumbers: lineNumberValue === '1' }
    block.runs = [{ ...block.runs[0]!, text: source, fontFamily: 'Consolas', fontSize: 10 }]
    block.style = { ...block.style, align: 'left', indent: 0, lineHeight: 1.35, spaceBefore: 8, spaceAfter: 8 }
  } else {
    for (const [inlineIndex, inline] of array(details.inlines).entries()) {
      if (record(inline).type !== 'form_field') continue
      const field = taggedFields.find((candidate) => text(candidate.tag).startsWith('office-latex:v1:'))
      if (!field) continue
      const source = text(record(inline).text)
      const mode = text(field.tag).endsWith(':block') ? 'block' : 'inline'
      const run = block.runs[inlineIndex]
      if (run) {
        run.text = '\uFFFC'
        run.fontFamily = 'Cambria Math'
        run.metadata = { ...run.metadata, latexFormula: { latex: source, mode, version: 1 } }
        if (mode === 'block') {
          block.metadata.latexFormulaBlock = true
          block.style = { ...block.style, align: 'center', spaceBefore: 8, spaceAfter: 8 }
        }
      }
      break
    }
  }
  const numberingId = numbering.numId
  if (type === 'numbered' || type === 'bullet' || type === 'checklist') block.level = listLevel
  if ((type === 'numbered' || type === 'bullet') && definition?.preset && definition.levels.has(listLevel)) {
    block.list = { id: `num-${numberingId}`, preset: definition.preset }
  }
  if (definition && (type === 'numbered' || type === 'bullet' || type === 'checklist')) {
    block.metadata.numberingDefinition = definition.definition
    block.metadata.numberingListId = `num-${numberingId}`
  }
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
      let blocks = array(cell.children).map((child, index) => normalizeBlock(child, cellId+'-block-'+index, lists))
      if (!blocks.length) blocks = [normalizeBlock({ id: cellId+'-text', content: text(cell.content), type: 'text' }, cellId+'-text', lists)]
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

function normalizeSectionHeaderFooters(data: Record<string, unknown>, blocks: DocumentBlock[]): Record<string, DocumentSectionHeaderFooter> {
  const layout = array(data.sectionLayout)
  if (!layout.length) return {}
  const parsedBlocks = array(data.blocks).map(record)
  const documentId = text(data.documentId)
  const result: Record<string, DocumentSectionHeaderFooter> = {}
  const sectionBreaks = blocks.filter((block) => block.type === 'sectionBreak')
  const storyText = (partName: string, kind: 'header' | 'footer'): string => {
    const story = parsedBlocks.find((block) => block.type === kind && array(block.sourceRefs).some((source) => text(record(source).partName) === partName))
    return text(story?.content)
  }
  for (const [index, rawSection] of layout.entries()) {
    const section = record(rawSection)
    const sectionId = index === 0
      ? 'first'
      : text(sectionBreaks[index - 1]?.id, `${documentId}-section-break-${finite(section.startNodeIndex)}`)
    const headerFields: Record<string, 'header' | 'firstPageHeader' | 'evenPageHeader'> = {
      default: 'header', first: 'firstPageHeader', even: 'evenPageHeader',
    }
    const footerFields: Record<string, 'footer' | 'firstPageFooter' | 'evenPageFooter'> = {
      default: 'footer', first: 'firstPageFooter', even: 'evenPageFooter',
    }
    const value: DocumentSectionHeaderFooter = {
      header: '', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '',
      pageFields: [],
      differentFirstPage: text(section.sectionPropertiesXml).includes('titlePg'),
      differentOddEven: data.evenAndOddHeaders === true,
      linkToPrevious: index > 0 && !array(section.headerSections).length && !array(section.footerSections).length,
    }
    const defined = new Set<DocumentPageField['story']>()
    for (const entry of array(section.headerSections)) {
      const reference = record(entry)
      const field = headerFields[text(reference.referenceType, 'default')]
      if (field) {
        defined.add(field)
        readStory(value, field, reference, storyText(text(reference.partName), 'header'))
      }
    }
    for (const entry of array(section.footerSections)) {
      const reference = record(entry)
      const field = footerFields[text(reference.referenceType, 'default')]
      if (field) {
        defined.add(field)
        readStory(value, field, reference, storyText(text(reference.partName), 'footer'))
      }
    }
    if (index > 0 && !value.linkToPrevious) {
      let prior = index - 1
      let previous = result[prior === 0 ? 'first' : sectionBreaks[prior - 1]?.id ?? '']
      while (prior > 0 && previous?.linkToPrevious) {
        prior--
        previous = result[prior === 0 ? 'first' : sectionBreaks[prior - 1]?.id ?? '']
      }
      if (previous) {
        for (const key of ['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter'] as const) {
          if (!defined.has(key)) value[key] = previous[key]
        }
        value.pageFields = [...(previous.pageFields ?? []).filter(f => !defined.has(f.story)), ...(value.pageFields ?? [])]
      }
    }
    result[sectionId] = value
  }
  return result
}

function readStory(section: DocumentSectionHeaderFooter, story: DocumentPageField['story'], reference: Record<string, unknown>, fallbackText: string) {
  const content = record(reference.storyText)
  const positioned = typeof content.text === 'string'
  section[story] = positioned ? text(content.text) : fallbackText
  const fields = array(positioned ? content.fields : reference.pageFields)
  for (const raw of fields) {
    const detail = record(raw)
    const field = text(positioned ? detail.field : raw)
    if (field !== 'PAGE' && field !== 'NUMPAGES' && field !== 'DATE' && field !== 'TIME') continue
    section.pageFields?.push(positioned ? { story, field, offset: finite(detail.offset), instruction: text(detail.instruction), cachedResult: text(detail.cachedResult) } : { story, field })
  }
}

function normalizeDocumentComments(value: unknown, blocks: DocumentBlock[]): Record<string, DocumentThreadComment> | undefined {
  const rawComments = array(value).map(record)
  if (!rawComments.length) return undefined
  const comments: Record<string, DocumentThreadComment> = {}
  for (const raw of rawComments) {
    if (raw.parentId !== undefined && raw.parentId !== null) continue
    const id = text(raw.id)
    if (!id) continue
    const anchor = array(raw.targets).flatMap(targetValue => {
      const target = record(targetValue)
      if (text(target.story, 'body') !== 'body') return []
      const nodeId = text(target.nodeId)
      const block = blocks.find(candidate => candidate.id === nodeId)
        ?? blocks.find(candidate => array(record(candidate.metadata.source).sourceRefs).some(ref => text(record(ref).nodeId) === nodeId))
      if (!block) return []
      const startOffset = Math.max(0, finite(target.start))
      const endOffset = Math.max(startOffset, finite(target.end, startOffset))
      if (endOffset === startOffset) return []
      const sourceText = block.runs.map(run => run.text).join('')
      const anchorText = sourceText.slice(startOffset, endOffset)
      if (!anchorText) return []
      return [{ blockId: block.id, startOffset, endOffset: startOffset + anchorText.length, text: anchorText }]
    })
    const rootMessage = importedCommentMessage(raw, id)
    const replies = rawComments.filter(reply => String(reply.parentId ?? '') === id).map(reply => importedCommentMessage(reply, text(reply.id)))
    const messages = [rootMessage, ...replies]
    const date = rootMessage.createdAt
    comments[`docx-comment-${id}`] = {
      id: `docx-comment-${id}`,
      anchor,
      referenceText: anchor.map(segment => segment.text).join('\n'),
      messages,
      resolved: raw.resolved === true,
      createdAt: date,
      updatedAt: messages.map(message => message.updatedAt ?? message.createdAt).sort().at(-1) ?? date,
      metadata: { importedCommentId: id },
    }
  }
  return Object.keys(comments).length ? comments : undefined
}

function importedCommentMessage(raw: Record<string, unknown>, id: string): DocumentCommentMessage {
  const authorName = text(raw.author, 'Unknown author')
  const createdAt = text(raw.date, '1970-01-01T00:00:00.000Z')
  return {
    id: `docx-message-${id}`,
    authorId: `docx-author-${authorName}`,
    authorName,
    text: text(raw.text),
    createdAt,
  }
}

export function normalizeDocument(envelope: ParseEnvelope): DocumentModel {
  const data = envelope.payload.data
  const result = createBlankDocument()
  const id = text(data.documentId, crypto.randomUUID())
  result.id = id
  result.name = envelope.fileName
  result.blocks = []
  const stories: Record<string, unknown[]> = {}
  const lists = documentNumberingIndex(data.numberingDefinitions)
  for (const [index, source] of array(data.blocks).entries()) {
    const raw = record(source)
    const story = text(record(array(raw.sourceRefs)[0]).story, 'body')
    if (story !== 'body' || raw.type === 'header' || raw.type === 'footer' || raw.type === 'watermark') {
      const key = raw.type === 'header' || raw.type === 'footer' ? text(raw.type) : story
      ;(stories[key] ??= []).push(raw)
      continue
    }
    const block = normalizeBlock(raw, id+'-block-'+index, lists)
    result.blocks.push(block)
    if (block.type !== 'table') for (const [childIndex, child] of array(raw.children).entries()) {
      if (record(child).type === 'figure') result.blocks.push(normalizeBlock(child, block.id+'-object-'+childIndex, lists))
    }
  }
  if (!result.blocks.length) result.blocks = createBlankDocument().blocks
  Object.assign(result, normalizeDocumentNotes(data, result.blocks))
  result.comments = normalizeDocumentComments(data.comments, result.blocks)
  result.sectionHeaderFooters = normalizeSectionHeaderFooters(data, result.blocks)
  const breaks = result.blocks.filter(block => block.type === 'sectionBreak')
  array(data.sectionLayout).forEach((raw, index) => {
    const section = record(raw)
    const distance = (val: unknown) => typeof val === 'number' && Number.isFinite(val) && val >= 0
    const setup = {
      ...(distance(section.marginHeader) ? { marginHeader: Number(section.marginHeader) } : {}),
      ...(distance(section.marginFooter) ? { marginFooter: Number(section.marginFooter) } : {}),
    }
    if (!Object.keys(setup).length) return
    if (index === 0) result.sectionSetup = { ...result.sectionSetup, ...setup }
    else {
      const marker = breaks[index - 1]?.sectionBreak
      if (marker) marker.setup = { ...marker.setup, ...setup }
    }
  })
  result.modified = false
  result.wordCount = result.blocks.reduce((count, block) => count+block.runs.reduce((sum, run) => sum+run.text.length, 0), 0)
  result.metadata = {
    source: data, stories, format: envelope.format, parserVersion: envelope.parserVersion,
    geometryStatus: data.geometryStatus,
    geometryNote: 'StructuredDocument 目前未携带完整节/纸张几何；使用默认页面并保留原数据，待补齐解析契约。',
  }
  return result
}
