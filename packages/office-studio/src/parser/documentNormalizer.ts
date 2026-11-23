import { createBlankDocument, type DocumentBlock, type DocumentCommentMessage, type DocumentDefaultStyle, type DocumentModel, type DocumentPageField, type DocumentParagraphStyle, type DocumentSectionBreakType, type DocumentSectionHeaderFooter, type DocumentStoryContent, type DocumentStoryLock, type DocumentThreadComment, type InlineRun } from '../domain/model.ts'
import type { ParseEnvelope } from './contracts.ts'
import { array, finite, positive, record, text } from './validation.ts'
import { normalizeDocumentHighlight } from '../features/document/model/document-highlight.ts'
import { normalizeDocumentNotes } from '../features/document/notes/document-notes.ts'
import { documentNumberingIndex } from './documentNumbering.ts'
import { officeNumberingMarkerBaseline } from '../features/document/model/document-office-numbering-marker.ts'
import { runText, sliceRuns, tableForBlock } from '../features/document/model/document-model.ts'
import { DOCUMENT_STORY_KEYS, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryTextStyle, isDocumentTextBoundary, validateDocumentStoryContent, validateDocumentStoryLocks } from '../features/document/model/document-story.ts'
import { storyBuffer } from '../features/document/services/story-editor.ts'
import { validateDocumentOfficeSource } from './documentOfficeSource.ts'
import { documentTextFields, documentTextFieldRanges } from '../features/document/model/document-text-fields.ts'

const pointToPixel = (value: number): number => value * 96 / 72
const optionalNumber = (value: unknown): number | undefined => typeof value === 'number' && Number.isFinite(value) ? value : undefined

function normalizeParagraphStyle(style: Record<string, unknown>): DocumentParagraphStyle {
  const indent = record(style.indent), spacing = record(style.spacing)
  const indentPixels = (twips: string, points: string, legacy: string): number | undefined => {
    const raw = optionalNumber(indent[twips])
    if (raw !== undefined) return raw / 15
    const pt = optionalNumber(indent[points])
    if (pt !== undefined) return pointToPixel(pt)
    const old = optionalNumber(indent[legacy])
    return old === undefined ? undefined : old / 15
  }
  const left = indentPixels('leftTwips', 'leftPt', 'left')
  const right = indentPixels('rightTwips', 'rightPt', 'right')
  const firstLine = indentPixels('firstLineTwips', 'firstLinePt', 'firstLine')
  const hanging = indentPixels('hangingTwips', 'hangingPt', 'hanging')
  const spacingRule = spacing.lineRule === 'exact' || spacing.lineRule === 'atLeast' ? spacing.lineRule : 'auto'
  const lineTwips = optionalNumber(spacing.lineTwips)
  return {
    align: style.align === 'center' || style.align === 'right' || style.align === 'justify' ? style.align : 'left',
    indent: 0,
    ...(left === undefined ? {} : { indentStart: left }),
    ...(right === undefined ? {} : { indentEnd: right }),
    ...(firstLine === undefined ? {} : { indentFirstLine: firstLine }),
    ...(hanging === undefined ? {} : { hanging }),
    spacingRule,
    lineHeight: lineTwips === undefined ? positive(spacing.lineMultiplier, spacingRule === 'auto' ? 1.5 : 20) : Math.max(spacingRule === 'auto' ? 0.01 : 1, lineTwips / (spacingRule === 'auto' ? 240 : 15)),
    spaceBefore: Math.max(0, optionalNumber(spacing.beforeTwips) !== undefined ? finite(spacing.beforeTwips) / 15 : pointToPixel(finite(spacing.beforePt, finite(spacing.before) / 20))),
    spaceAfter: Math.max(0, optionalNumber(spacing.afterTwips) !== undefined ? finite(spacing.afterTwips) / 15 : pointToPixel(finite(spacing.afterPt, finite(spacing.after, 120) / 20))),
    ...(typeof style.wordWrap === 'boolean' ? { wordWrap: style.wordWrap } : {}),
    ...(typeof style.snapToGrid === 'boolean' ? { snapToGrid: style.snapToGrid } : {}),
  }
}
function normalizeRun(value: unknown, id: string): InlineRun {
  const raw = record(value)
  const style = record(raw.style)
  const verticalAlign = normalizeVerticalAlign(style.verticalAlign ?? style.vertAlign ?? raw.verticalAlign)
  const highlight = normalizeDocumentHighlight(style.highlight, style.backgroundColor)
  const eastAsiaFontFamily = text(style.fontFamilyEastAsia) || text(style.resolvedFontFamilyEastAsia)
  const run: InlineRun = {
    id, text: text(raw.text), bold: style.bold === true, italic: style.italic === true,
    underline: style.underline === true, strike: style.strike === true,
    ...(verticalAlign ? { verticalAlign } : {}),
    ...(highlight ? { highlight } : {}),
    ...(eastAsiaFontFamily ? { eastAsiaFontFamily } : {}),
    color: text(style.color, '#1f2937'), fontFamily: text(style.fontFamilyAscii) || text(style.resolvedFontFamilyAscii) || text(style.fontFamily, 'Aptos'), fontSize: positive(style.fontSizePt, 11),
    metadata: { source: raw, link: raw.link, ...(highlight ? { highlight } : {}), noteReference: raw.noteReference },
  }
  if (raw.sourceFormat !== undefined) {
    validateStorySourceFormat(raw.sourceFormat)
    const { id: _id, text: _text, metadata: _metadata, ...baseline } = run
    run.metadata = { ...run.metadata, officeTextStyleBaseline: baseline }
  }
  return run
}
function normalizeVerticalAlign(value: unknown): InlineRun['verticalAlign'] {
  if (value === 'sup' || value === 'super' || value === 'superscript') return 'superscript'
  if (value === 'sub' || value === 'subscript') return 'subscript'
  return undefined
}

function normalizeDefaultStyle(value: unknown): DocumentDefaultStyle | undefined {
  const raw = record(value)
  if (raw.version !== 1) return undefined
  const source = record(raw.text)
  const fontFamily = text(source.fontFamilyAscii) || text(source.resolvedFontFamilyAscii) || text(source.fontFamily, 'Aptos')
  const eastAsiaFontFamily = text(source.fontFamilyEastAsia) || text(source.resolvedFontFamilyEastAsia) || fontFamily
  const verticalAlign = normalizeVerticalAlign(source.verticalAlign)
  const highlight = normalizeDocumentHighlight(source.highlight, source.backgroundColor)
  const paragraph = record(raw.paragraph)
  return {
    version: 1,
    text: {
      fontFamily, eastAsiaFontFamily, fontSize: positive(source.fontSizePt, 12), color: text(source.color, 'auto'),
      bold: source.bold === true, italic: source.italic === true, underline: source.underline === true, strike: source.strike === true,
      ...(verticalAlign ? { verticalAlign } : {}), ...(highlight ? { highlight } : {}),
    },
    // Empty Word/ODF paragraph defaults use compatibility single/zero spacing.
    // The Rust Word reader supplies its separate absent-default fallback.
    paragraph: normalizeParagraphStyle({
      ...paragraph,
      spacing: { beforePt: 0, afterPt: 0, lineMultiplier: 1, ...record(paragraph.spacing) },
      ...(typeof raw.wordWrap === 'boolean' ? { wordWrap: raw.wordWrap } : {}),
    }),
  }
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
  const id = text(raw.id, fallbackId)
  const runs = array(details.inlines).filter((inline) => {
    const type = record(inline).type
    return type === 'text' || type === 'form_field'
  }).map((inline, index) => normalizeRun(inline, id+'-run-'+index))
  if (!runs.length) runs.push(normalizeRun({ text: text(raw.content) }, id+'-run-0'))
  const level = Math.min(6, Math.max(0, finite(details.headingLevel, finite(style.headingLevel))))
  const numbering = record(details.numbering)
  const listLevel = Math.min(8, Math.max(0, Math.trunc(finite(numbering.ilvl, finite(numbering.level)))))
  const definition = typeof numbering.numId === 'number' ? lists.get(numbering.numId) : undefined
  const format = definition?.levels.get(listLevel)?.format
  const listLabel = text(details.listLabel).trimStart()
  const checklistLabel = listLabel.startsWith('☐') || listLabel.startsWith('☑')
  const type = sectionBreak ? 'sectionBreak' : columnBreak ? 'columnBreak' : pageBreak ? 'pageBreak' : raw.type === 'table' ? 'table' : raw.type === 'figure' ? 'image'
    : titleStyle ? 'paragraph' : raw.type === 'heading' || raw.type === 'section_heading' || level > 0 ? 'heading'
      : checklistLabel ? 'checklist'
        : Object.keys(numbering).length && numbering.numId !== 0
          ? (format ? format !== 'bullet' : /^[\d一二三四五六七八九十]|^[a-z]+[.)]/i.test(listLabel)) ? 'numbered' : 'bullet'
          : 'paragraph'
  const block: DocumentBlock = {
    id, type, level, runs: ['sectionBreak', 'columnBreak', 'pageBreak'].includes(type) ? [] : runs,
    style: normalizeParagraphStyle(style),
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
    const source = runText(block.runs)
    block.metadata.codeBlock = { source, language, lineNumbers: lineNumberValue === '1' }
    const plainRanges = array(details.bookmarks).map(bookmark => text(record(bookmark).name))
      .map(name => /^office_code_plain_v1_(\d+)_(\d+)$/.exec(name)).filter(match => match !== null)
      .map(match => ({ start: Number(match[1]), end: Number(match[2]) }))
    for (const range of plainRanges) if (!isDocumentTextBoundary(source, range.start) || !isDocumentTextBoundary(source, range.end) || range.start >= range.end) throw new RangeError('代码块清除格式的 DOCX 范围无效。')
    if (plainRanges.length) {
      const boundaries = [...new Set([0, source.length, ...plainRanges.flatMap(range => [range.start, range.end])])].sort((a, b) => a - b)
      block.runs = boundaries.slice(0, -1).flatMap((start, index) => {
        const end = boundaries[index + 1]!, plain = plainRanges.some(range => range.start <= start && end <= range.end)
        return sliceRuns(block.runs, start, end).map(run => plain ? { ...run, metadata: { ...run.metadata, codeSyntaxHighlight: false } } : run)
      })
    }
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
  const numberingListId = text(details.numberingListId, `num-${numberingId}`)
  const hasNumbering = Object.keys(numbering).length > 0 && numberingId !== 0 && ['paragraph', 'heading', 'numbered', 'bullet', 'checklist'].includes(type)
  if (type === 'numbered' || type === 'bullet' || type === 'checklist') block.level = listLevel
  else if (hasNumbering) block.metadata.numberingLevel = listLevel
  if (hasNumbering && type !== 'checklist' && definition?.preset && definition.levels.has(listLevel)) {
    block.list = { id: numberingListId, preset: definition.preset }
  }
  if (definition && hasNumbering) {
    block.metadata.numberingDefinition = definition.definition
    block.metadata.numberingListId = numberingListId
  }
  if (type === 'checklist') block.checklist = {
    listId: typeof numberingId === 'string' || typeof numberingId === 'number' ? String(numberingId) : id,
    checked: listLabel.startsWith('☑'),
  }
  const officeNumberingMarker = record(details.officeNumberingMarker)
  if (officeNumberingMarker.version === 1 && typeof officeNumberingMarker.xml === 'string' && typeof officeNumberingMarker.text === 'string' && typeof officeNumberingMarker.bookmark === 'string' && typeof officeNumberingMarker.bookmarkId === 'number' && Number.isSafeInteger(officeNumberingMarker.bookmarkId)) {
    block.metadata.officeNumberingMarker = { ...officeNumberingMarker, baseline: officeNumberingMarkerBaseline(block, officeNumberingMarker.text) }
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
  if (details.documentTextFields !== undefined) {
    block.metadata.documentTextFields = details.documentTextFields
    documentTextFields(block)
  }
  if (details.nativeFieldSource !== undefined) block.metadata.nativeFieldSource = details.nativeFieldSource
  return block
}

function normalizeSectionHeaderFooters(data: Record<string, unknown>, blocks: DocumentBlock[], model: DocumentModel): Record<string, DocumentSectionHeaderFooter> {
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
        readStory(value, field, reference, storyText(text(reference.partName), 'header'), model)
      }
    }
    for (const entry of array(section.footerSections)) {
      const reference = record(entry)
      const field = footerFields[text(reference.referenceType, 'default')]
      if (field) {
        defined.add(field)
        readStory(value, field, reference, storyText(text(reference.partName), 'footer'), model)
      }
    }
    // Each omitted native reference inherits that specific previous story.
    // Keep its local override empty so later edits of the previous section remain observable.
    value.storyLinkToPrevious = Object.fromEntries(DOCUMENT_STORY_KEYS.map(story => [story, index > 0 && !defined.has(story)]))
    value.linkToPrevious = index > 0 && DOCUMENT_STORY_KEYS.every(story => value.storyLinkToPrevious![story])
    result[sectionId] = value
  }
  return result
}

function readStory(section: DocumentSectionHeaderFooter, story: DocumentPageField['story'], reference: Record<string, unknown>, fallbackText: string, model: DocumentModel) {
  const content = record(reference.storyText)
  const positioned = typeof content.text === 'string'
  section[story] = positioned ? text(content.text) : fallbackText
  const fields = array(positioned ? content.fields : reference.pageFields)
  for (const [fieldIndex, raw] of fields.entries()) {
    const detail = record(raw)
    const field = text(positioned ? detail.field : raw)
    if (field !== 'PAGE' && field !== 'NUMPAGES' && field !== 'DATE' && field !== 'TIME') continue
    if (positioned && typeof detail.offset !== 'number') throw new RangeError('页眉页脚字段偏移无效。')
    validateStorySourceFormat(detail.sourceFormat)
    const textStyle = detail.textStyle && typeof detail.textStyle === 'object'
      ? documentStoryTextStyle(normalizeRun({ text: '', style: detail.textStyle }, story+'-field'), documentStoryDefaultStyle(model).text) : undefined
    const cachedRuns = detail.cachedRuns === undefined ? undefined : normalizeStoryFieldCachedRuns(detail.cachedRuns, story+'-field-'+fieldIndex+'-cache', model)
    for (const key of ['locked', 'containsNestedFields'] as const) if (detail[key] !== undefined && typeof detail[key] !== 'boolean') throw new TypeError(`页眉页脚字段 ${key} 必须为布尔值。`)
    const normalizedField: DocumentPageField = positioned ? { story, field, offset: finite(detail.offset), instruction: text(detail.instruction), cachedResult: text(detail.cachedResult),
      ...(textStyle ? { textStyle } : {}),
      ...(typeof detail.locked === 'boolean' ? { locked: detail.locked } : {}), ...(typeof detail.containsNestedFields === 'boolean' ? { containsNestedFields: detail.containsNestedFields } : {}),
      ...(detail.sourceFormat !== undefined || cachedRuns !== undefined ? { metadata: { source: detail,
        ...(textStyle ? { officeStoryStyleBaseline: { ...textStyle } } : {}), ...(cachedRuns ? { officeStoryFieldCache: { version: 1, runs: cachedRuns } } : {}) } } : {}) } : { story, field }
    documentStoryFieldCache(normalizedField)
    section.pageFields?.push(normalizedField)
  }
  if (Array.isArray(content.runs) && Array.isArray(content.paragraphs)) {
    const rich: DocumentStoryContent = {
      runs: content.runs.map((raw, index) => {
        const source = record(raw), run = normalizeRun(raw, story+'-run-'+index)
        validateStorySourceFormat(source.sourceFormat)
        if (source.sourceFormat !== undefined) run.metadata = { ...run.metadata, officeStoryStyleBaseline: documentStoryTextStyle(run, documentStoryDefaultStyle(model).text) }
        return run
      }),
      paragraphs: content.paragraphs.map(raw => {
        const paragraph = record(raw)
        if (typeof paragraph.startOffset !== 'number' || typeof paragraph.endOffset !== 'number') throw new RangeError('页眉页脚段落偏移无效。')
        validateStorySourceFormat(paragraph.sourceFormat)
        const style = normalizeParagraphStyle({ ...record(paragraph.style), ...(typeof paragraph.wordWrap === 'boolean' ? { wordWrap: paragraph.wordWrap } : {}) })
        return { startOffset: paragraph.startOffset, endOffset: paragraph.endOffset,
          style,
          metadata: { source: paragraph, ...(paragraph.sourceFormat !== undefined ? { officeStoryParagraphStyleBaseline: { ...style } } : {}) } }
      }),
    }
    if (content.locks !== undefined) {
      if (!Array.isArray(content.locks)) throw new RangeError('页眉页脚内容锁无效。')
      rich.locks = array(content.locks).map(raw => {
        const lock = record(raw)
        if (typeof lock.id !== 'string' || typeof lock.startOffset !== 'number' || typeof lock.endOffset !== 'number'
          || lock.lock !== 'contentLocked' && lock.lock !== 'sdtLocked' && lock.lock !== 'sdtContentLocked') throw new RangeError('页眉页脚内容锁无效。')
        return { id: lock.id, startOffset: lock.startOffset, endOffset: lock.endOffset, lock: lock.lock } satisfies DocumentStoryLock
      })
    }
    validateDocumentStoryContent(rich, section[story])
    validateDocumentStoryLocks(rich.locks, storyBuffer(section, story).value)
    const baseline = {
      text: section[story],
      runs: rich.runs.map(({ id: _id, metadata, ...run }) => ({ ...run, ...(typeof metadata?.link === 'string' ? { href: metadata.link } : {}) })),
      paragraphs: rich.paragraphs.map(({ metadata: _metadata, ...paragraph }) => paragraph),
      fields: (section.pageFields ?? []).filter(field => field.story === story).map(({ story: _story, metadata, ...field }) => ({ ...field,
        ...(metadata?.officeStoryFieldCache !== undefined ? { cachedRuns: metadata.officeStoryFieldCache } : {}) })),
      locks: rich.locks ?? [],
    }
    if (typeof content.sourceXml === 'string') rich.metadata = { officeSource: { version: 1, xml: content.sourceXml,
      ...(typeof content.relationshipsXml === 'string' ? { relationshipsXml: content.relationshipsXml } : {}),
      unsupportedFeatures: array(content.unsupportedFeatures).filter(feature => typeof feature === 'string'), baseline } }
    section.storyContent = { ...section.storyContent, [story]: rich }
  }
}

function normalizeStoryFieldCachedRuns(value: unknown, id: string, model: DocumentModel): InlineRun[] {
  if (!Array.isArray(value)) throw new TypeError('页眉页脚字段缓存运行必须为数组。')
  return value.map((raw, index) => {
    const source = record(raw)
    if (typeof source.text !== 'string') throw new TypeError('页眉页脚字段缓存文本无效。')
    validateStorySourceFormat(source.sourceFormat)
    const run = normalizeRun(raw, id+'-'+index)
    if (source.sourceFormat !== undefined) run.metadata = { ...run.metadata, officeStoryStyleBaseline: documentStoryTextStyle(run, documentStoryDefaultStyle(model).text) }
    return run
  })
}

function validateStorySourceFormat(value: unknown): void {
  if (value === undefined) return
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('页眉页脚源格式片段无效。')
  const format = record(value)
  if (format.version !== 1
    || format.propertiesXml !== undefined && typeof format.propertiesXml !== 'string'
    || format.paragraphStyleId !== undefined && (typeof format.paragraphStyleId !== 'string' || !format.paragraphStyleId)
    || !format.namespaces || typeof format.namespaces !== 'object' || Array.isArray(format.namespaces)
    || Object.values(record(format.namespaces)).some(namespace => typeof namespace !== 'string')) throw new TypeError('页眉页脚源格式片段版本或命名空间无效。')
}

function normalizeDocumentComments(value: unknown, blocks: DocumentBlock[]): Record<string, DocumentThreadComment> | undefined {
  const rawComments = array(value).map(record)
  if (!rawComments.length) return undefined
  const paragraphs: DocumentBlock[] = []
  const visit = (items: DocumentBlock[]): void => {
    for (const block of items) {
      const table = tableForBlock(block)
      if (table) { for (const row of table.rows) for (const cell of row.cells) visit(cell.blocks) }
      else paragraphs.push(block)
    }
  }
  visit(blocks)
  const validBoundary = (text: string, offset: number): boolean => {
    const previous = text.charCodeAt(offset - 1), next = text.charCodeAt(offset)
    return !(previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF)
  }
  const comments: Record<string, DocumentThreadComment> = {}
  for (const raw of rawComments) {
    if (raw.parentId !== undefined && raw.parentId !== null) continue
    const id = text(raw.id)
    if (!id) continue
    const anchor = array(raw.targets).flatMap(targetValue => {
      const target = record(targetValue)
      if (text(target.story, 'body') !== 'body') return []
      const nodeId = text(target.nodeId)
      const block = paragraphs.find(candidate => candidate.id === nodeId)
        ?? paragraphs.find(candidate => array(record(candidate.metadata.source).sourceRefs).some(ref => text(record(ref).nodeId) === nodeId))
      if (!block) return []
      const startOffset = optionalNumber(target.start), end = optionalNumber(target.end)
      if (startOffset === undefined || end === undefined || !Number.isSafeInteger(startOffset) || !Number.isSafeInteger(end) || startOffset < 0 || end < startOffset) return []
      const sourceText = block.runs.map(run => run.text).join('')
      const endOffset = Math.min(end, sourceText.length)
      if (startOffset > sourceText.length || !validBoundary(sourceText, startOffset) || !validBoundary(sourceText, endOffset)) return []
      const anchorText = sourceText.slice(startOffset, endOffset)
      if (!anchorText && end !== startOffset) return []
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
  validateDocumentOfficeSource(data.officeSource, envelope.format)
  if (data.defaultStyle !== undefined && data.officeSource?.resourceErrors?.length) {
    throw new TypeError('文档默认样式含有未解决的源资源错误。')
  }
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
  documentTextFieldRanges(result)
  Object.assign(result, normalizeDocumentNotes(data, result.blocks))
  result.comments = normalizeDocumentComments(data.comments, result.blocks)
  const defaultStyle = normalizeDefaultStyle(data.defaultStyle)
  if (defaultStyle) result.defaultStyle = defaultStyle
  else delete result.defaultStyle
  result.sectionHeaderFooters = normalizeSectionHeaderFooters(data, result.blocks, result)
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
