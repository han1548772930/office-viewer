import type { DocumentBlock, DocumentDefaultStyle, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { applyNamedParagraphStyle, createParagraph, createRun, runText, DOCUMENT_NAMED_PARAGRAPH_STYLES, type DocumentNamedParagraphStyle } from '../features/document/model/document-model'
import { DOCUMENT_HORIZONTAL_LINE_BORDER, documentHorizontalLineBorder } from '../features/document/model/document-horizontal-line'
import { blockHasNumbering, blockNumberingDefinition, blockNumberingId, blockNumberingLevelIndex, clearBlockNumbering, consecutiveNumberingGroup, numberingDefinitionFromParts, numberingRecord, type DocumentNumberingDefinition } from '../features/document/model/document-numbering'
import { blockListPreset, isListPreset, presetNumberingDefinition } from '../features/document/model/document-lists'
import { bodyBoolean, bodyColor, bodyLength, bodyNumber, bodyRecord, documentParagraphStyleFromBody, documentParagraphStyleToBody, documentRunStyleFromBody, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentBodyArray, documentBodyIdentity, type DocumentBodyDecodeOptions } from './document-body-types'
import { applyDocumentBodyBlockSource, documentBodyBlockSource } from './document-body-source'
import { documentBodyListInteger, documentBodyListLevels, documentBodyListRecord, documentBodyMissingListLevel,
  isDocumentBodyMissingListLevel } from './document-body-list-levels'

const namedStyles: DocumentNamedParagraphStyle[] = ['normal', 'normal', 'title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5']
const numberFormats = ['bullet', 'none', 'decimal', 'decimalZero', 'upperLetter', 'lowerLetter', 'upperRoman', 'lowerRoman',
  'ordinal', 'cardinalText', 'ordinalText', 'hex', 'chicago', 'ideographDigital', 'japaneseCounting', 'aiueo', 'iroha', 'decimalFullWidth',
  'decimalHalfWidth', 'japaneseLegal', 'japaneseDigitalTenThousand', 'decimalEnclosedCircle', 'decimalFullWidth2', 'aiueoFullWidth', 'irohaFullWidth',
  'ganada', 'chosung', 'decimalEnclosedFullstop', 'decimalEnclosedParen', 'decimalEnclosedCircleChinese', 'ideographEnclosedCircle',
  'ideographTraditional', 'ideographZodiac', 'ideographZodiacTraditional', 'taiwaneseCounting', 'ideographLegalTraditional',
  'taiwaneseCountingThousand', 'taiwaneseDigital', 'chineseCounting', 'chineseLegalSimplified', 'chineseCountingThousand',
  'koreanDigital', 'koreanCounting', 'koreanLegal', 'koreanDigital2', 'vietnameseCounting', 'russianLower', 'russianUpper', 'numberInDash',
  'hebrew1', 'hebrew2', 'arabicAlpha', 'arabicAbjad', 'hindiVowels', 'hindiConsonants', 'hindiNumbers', 'hindiCounting',
  'thaiLetters', 'thaiNumbers', 'thaiCounting', 'custom']

/** Named style chains are resolved once at the boundary; cycles are input errors. */
export function documentBodyNamedStyle(id: unknown, styles: DocumentBodyDecodeOptions['styles']): UniverBodyRecord {
  if (id === undefined || id === null || !styles) return {}
  const currentId = documentBodyIdentity(id, '文档样式'), visited = new Set<string>()
  const resolve = (key: string, depth: number): UniverBodyRecord => {
    if (depth > 64 || visited.has(key)) throw new RangeError('文档命名样式继承存在循环。')
    const raw = styles[key]
    if (!raw) return {}
    visited.add(key)
    const parent = typeof raw.basedOn === 'string' ? resolve(raw.basedOn, depth + 1) : {}
    const bullet = mergeBodyBullet(parent.bullet, raw.bullet)
    return { ...parent, ...raw, textStyle: { ...bodyRecord(parent.textStyle), ...bodyRecord(raw.textStyle) },
      paragraphStyle: { ...bodyRecord(parent.paragraphStyle), ...bodyRecord(raw.paragraphStyle) }, ...(bullet === undefined ? {} : { bullet }) }
  }
  return resolve(currentId, 0)
}

export function documentParagraphFromBody(raw: UniverBodyRecord, runs: InlineRun[], options: DocumentBodyDecodeOptions,
  lists: Record<string, UniverBodyRecord> = {}): DocumentBlock {
  const inherited = documentBodyNamedStyle(raw.styleId ?? options.defaultParagraphStyleId, options.styles), paragraphStyle = { ...bodyRecord(inherited.paragraphStyle), ...bodyRecord(raw.paragraphStyle) }
  const block = createParagraph('', options.defaults.text)
  block.style = { ...options.defaults.paragraph }
  if (raw.paragraphId !== undefined) block.id = documentBodyIdentity(raw.paragraphId, '段落')
  const named = bodyNumber(paragraphStyle.namedStyleType)
  if (named !== undefined && named !== 0) {
    if (!Number.isSafeInteger(named) || !namedStyles[named]) throw new RangeError('段落命名样式无效。')
    applyNamedParagraphStyle(block, namedStyles[named]!)
  }
  block.style = documentParagraphStyleFromBody(paragraphStyle, { ...options.defaults, paragraph: block.style })
  block.runs = runs.length ? runs : [createRun('', { ...options.defaults.text, ...documentRunStyleFromBody({ ...bodyRecord(inherited.textStyle), ...bodyRecord(paragraphStyle.textStyle) }) })]
  const outline = bodyNumber(paragraphStyle.outlineLevel)
  if (outline !== undefined) {
    if (!Number.isSafeInteger(outline) || outline < 0 || outline > 9) throw new RangeError('段落大纲级别无效。')
    if (outline < 9) { block.type = 'heading'; block.level = outline + 1 }
  }
  block.metadata = { ...block.metadata, univerBodyParagraphSource: { version: 1, paragraph: cloneModel(raw) },
    univerParagraphPaginationResolution: { version: 1, inherited: cloneModel(bodyRecord(inherited.paragraphStyle)),
      baseline: Object.fromEntries(['keepLines', 'keepNext', 'widowControl', 'pageBreakBefore'].map(key => [key, Reflect.get(block.style, key)])) } }
  if (bodyBoolean(raw.headingCollapsed) !== undefined) block.metadata.headingCollapsed = bodyBoolean(raw.headingCollapsed)
  const paragraphType = block.type, paragraphLevel = block.level
  applyDocumentBodyBlockSource(block, raw)
  const hasText = runText(block.runs).length > 0
  if (block.type === 'divider' && hasText) { block.type = paragraphType; block.level = paragraphLevel }
  const sourceDefinition = blockNumberingDefinition(block), contract = bodyRecord(bodyRecord(raw.officeStudio).numberingContract)
  // Public bullets are current data. The extension contributes provenance,
  // never an obsolete list/check state after a public edit or bullet:null.
  clearBlockNumbering(block); delete block.checklist
  if (paragraphType === 'heading' || ['bullet', 'numbered', 'checklist'].includes(block.type) || block.type === 'heading' && paragraphType !== 'heading') {
    block.type = paragraphType; block.level = paragraphLevel
  }
  const bullet = mergeBodyBullet(inherited.bullet, raw.bullet), effective = bodyRecord(bullet)
  if (bullet !== null && Object.keys(effective).length && effective.listId !== '0') {
    const listType = effective.listType ?? Object.keys(lists).find(key => {
      const counters = lists[key]?.counterIds
      return Array.isArray(counters) && counters.includes(effective.listId)
    }) ?? effective.listId
    const current = { ...effective, listType }
    const preserve = sourceDefinition && (contract.listType === listType && contract.listId === (effective.listId ?? listType)
      || !Object.keys(contract).length && (sourceDefinition.abstract.univerBodyPublicListType === listType || documentBodyListType(sourceDefinition, blockNumberingId(block) ?? String(effective.listId ?? listType)) === listType))
    applyBullet(block, current, lists, options.defaults, preserve ? sourceDefinition : undefined)
  }
  const border = bodyRecord(paragraphStyle.borderBottom)
  if (!hasText && (block.type === 'divider' || Object.keys(border).length)) {
    const width = bodyNumber(border.width) ?? DOCUMENT_HORIZONTAL_LINE_BORDER.width
    const padding = bodyNumber(border.padding) ?? DOCUMENT_HORIZONTAL_LINE_BORDER.padding
    const dashStyle = bodyNumber(border.dashStyle) ?? 1
    if (width < 0 || padding < 0 || ![0, 1, 2, 3].includes(dashStyle)) throw new RangeError('分隔线边框的宽度、间距或线型无效。')
    if (bullet !== undefined) bodyRecord(bodyRecord(block.metadata.univerBodyParagraphSource).paragraph).bullet = cloneModel(bullet)
    block.type = 'divider'
    delete block.checklist
    block.runs = []
    block.metadata.documentDividerBorder = { width, padding, color: bodyColor(border.color) ?? DOCUMENT_HORIZONTAL_LINE_BORDER.color,
      style: dashStyle === 2 ? 'dot' : dashStyle === 3 ? 'dash' : 'solid' }
  }
  return block
}

export function documentBodyParagraphMarks(raw: UniverBodyRecord, options: DocumentBodyDecodeOptions): UniverBodyRecord {
  const inherited = documentBodyNamedStyle(raw.styleId ?? options.defaultParagraphStyleId, options.styles)
  const paragraph = { ...bodyRecord(inherited.paragraphStyle), ...bodyRecord(raw.paragraphStyle) }, named = bodyNumber(paragraph.namedStyleType)
  const preset = named !== undefined && named > 0 && namedStyles[named] ? DOCUMENT_NAMED_PARAGRAPH_STYLES[namedStyles[named]!] : undefined
  return { ...(preset ? { fs: preset.fontSize, bl: Number(preset.bold), ...(preset.color ? { cl: { rgb: preset.color } } : {}) } : {}),
    ...bodyRecord(inherited.textStyle), ...bodyRecord(bodyRecord(inherited.paragraphStyle).textStyle), ...bodyRecord(bodyRecord(raw.paragraphStyle).textStyle) }
}

function applyBullet(block: DocumentBlock, bullet: UniverBodyRecord, lists: Record<string, UniverBodyRecord>, defaults: DocumentDefaultStyle, native?: DocumentNumberingDefinition): void {
  const listType = documentBodyIdentity(bullet.listType, '列表类型'), listId = documentBodyIdentity(bullet.listId ?? listType, '列表')
  const level = documentBodyListInteger(bullet.nestingLevel, 0, 0, 8, '列表层级')
  const raw = lists[listType], checklist = listType === 'CHECK_LIST' || listType === 'CHECK_LIST_CHECKED'
  let definition: DocumentNumberingDefinition
  if (raw) {
    const levels = documentBodyListLevels(raw.nestingLevel).flatMap((source, ilvl) => {
      if (ilvl !== level && isDocumentBodyMissingListLevel(source, ilvl)) return []
      const formatIndex = typeof source.glyphType === 'string' ? undefined : bodyNumber(source.glyphType), format = source.glyphSymbol !== undefined ? 'bullet' : typeof source.glyphType === 'string' ? source.glyphType : numberFormats[formatIndex ?? 2]
      if (!format || ilvl > 8) throw new RangeError('列表编号格式或层级无效。')
      const start = documentBodyListInteger(source.startNumber, 0, -1, Number.MAX_SAFE_INTEGER - 1, '列表起始编号')
      const style = documentRunStyleFromBody(source.textStyle), properties = bodyRecord(source.paragraphProperties), twips = (value: unknown) => {
        const length = bodyLength(value, defaults.text.fontSize); return length === undefined ? undefined : length * 15
      }
      return [{ ilvl, start: start + 1, format, text: typeof source.glyphSymbol === 'string' ? source.glyphSymbol : typeof source.glyphFormat === 'string' ? source.glyphFormat : `%${ilvl + 1}.`,
        suffix: source.followNumberWith === 2 ? 'nothing' as const : source.followNumberWith === 1 ? 'space' as const : 'tab' as const,
        alignment: source.bulletAlignment === 2 ? 'end' as const : source.bulletAlignment === 1 ? 'center' as const : 'start' as const,
        isLgl: bodyBoolean(source.isLegal), indent: { leftTwips: twips(properties.indentStart), rightTwips: twips(properties.indentEnd),
          hangingTwips: twips(properties.hanging), firstLineTwips: twips(properties.indentFirstLine) },
        bulletFontFamily: style.fontFamily, bulletColor: style.color ?? bodyColor(bodyRecord(source.textStyle).cl),
        runStyle: { fontSizePt: style.fontSize, bold: style.bold, italic: style.italic, color: style.color, fontFamily: style.fontFamily } }]
    })
    if (!levels.some(item => item.ilvl === level)) throw new RangeError('列表缺少当前层级定义。')
    if (raw.counterIds !== undefined && (!Array.isArray(raw.counterIds) || raw.counterIds.some(id => typeof id !== 'string' || !id) || new Set(raw.counterIds).size !== raw.counterIds.length)) throw new TypeError('共享列表计数器标识无效。')
    definition = { abstract: { abstractNumId: listType, levels, univerBodyListSource: cloneModel(raw) },
      instance: { numId: listId, ...(Array.isArray(raw.counterIds) && raw.counterIds.includes(listId) ? { consecutiveListId: raw.counterIds[0] } : {}) }, levels }
  } else if (checklist) {
    definition = checklistNumberingDefinition(listType === 'CHECK_LIST_CHECKED')
  } else if (isListPreset(listType)) {
    definition = presetNumberingDefinition(listType)
    block.list = { id: listId, preset: listType }
  } else throw new RangeError(`列表 ${listType} 缺少资源定义。`)
  definition.abstract.univerBodyPublicListType = listType
  if (Array.isArray(raw?.counterIds) && raw.counterIds.includes(listId)) definition.abstract.univerBodyPublicCounterId = raw.counterIds[0]
  definition.instance.numId = listId
  const counterStarts = documentBodyListRecord(raw?.counterStarts, '列表重启编号'), starts: Record<string, number> = {}
  for (const [id, entry] of Object.entries(counterStarts)) {
    documentBodyIdentity(id, '列表重启')
    for (const [key, value] of Object.entries(documentBodyListRecord(entry, '列表实例重启编号'))) {
      if (!/^[0-8]$/.test(key) || value === undefined) throw new RangeError('列表重启编号无效。')
      const start = documentBodyListInteger(value, 0, -1, Number.MAX_SAFE_INTEGER - 1, '列表重启编号')
      if (id === listId) starts[key] = start + 1
    }
  }
  if (bullet.startNumber !== undefined) {
    const start = documentBodyListInteger(bullet.startNumber, 0, -1, Number.MAX_SAFE_INTEGER - 1, '段落起始编号')
    starts[level] = start + 1
  }
  if (Object.keys(starts).length) definition.instance.levelStartOverrides = starts
  if (native) {
    const levels = definition.levels.map(level => {
      const prior = native.levels.find(prior => prior.ilvl === level.ilvl)
      const alignment = (value: string | undefined) => value === 'right' || value === 'end' ? 2 : value === 'center' ? 1 : 0
      return { ...prior, ...level, ...(prior?.alignment !== undefined && alignment(prior.alignment) === alignment(level.alignment) ? { alignment: prior.alignment } : {}) }
    })
    const instance = { ...native.instance, ...definition.instance }
    if (definition.instance.consecutiveListId === undefined) delete instance.consecutiveListId
    if (definition.instance.levelStartOverrides === undefined) delete instance.levelStartOverrides
    if (Array.isArray(native.instance.levelOverrides)) instance.levelOverrides = native.instance.levelOverrides.map(value => {
      const prior = numberingRecord(value), current = levels.find(level => level.ilvl === prior.ilvl)
      return current ? { ...prior, ...current } : prior
    })
    definition = { abstract: { ...native.abstract, ...definition.abstract, abstractNumId: native.abstract.abstractNumId ?? listType, levels }, instance, levels }
  }
  const numbered = definition.levels.find(item => item.ilvl === level)?.format !== 'bullet'
  if (block.type !== 'heading') { block.type = numbered ? 'numbered' : 'bullet'; block.level = level }
  else block.metadata.numberingLevel = level
  block.metadata.numberingListId = listId; block.metadata.numberingDefinition = definition
  const sourcePreset = bodyRecord(raw?.officeStudio).preset
  if (isListPreset(listType)) block.list = { id: listId, preset: listType }
  else if (isListPreset(sourcePreset) && (sourcePreset.startsWith('BULLET_') ? !numbered : numbered)) block.list = { id: listId, preset: sourcePreset }
  if (checklist) { block.type = 'checklist'; block.level = level; block.checklist = { listId, checked: listType === 'CHECK_LIST_CHECKED' } }
}

export function documentParagraphToBody(block: DocumentBlock, end: number, instanceId?: string): UniverBodyRecord {
  const source = bodyRecord(block.metadata.univerBodyParagraphSource), result = cloneModel(bodyRecord(source.paragraph))
  result.startIndex = end; result.paragraphId = block.id
  result.paragraphStyle = documentParagraphStyleToBody(block.style, result.paragraphStyle, block.metadata.univerParagraphPaginationResolution)
  const style = bodyRecord(result.paragraphStyle)
  if (block.type === 'divider') {
    const border = documentHorizontalLineBorder(block)
    style.borderBottom = { ...bodyRecord(style.borderBottom), width: border.width, padding: border.padding,
      color: { rgb: border.color }, dashStyle: border.style === 'dot' ? 2 : border.style === 'dash' ? 3 : 1 }
  }
  if (typeof block.metadata.documentCodeFontFamily === 'string') {
    style.textStyle = { ...bodyRecord(style.textStyle), ff: block.metadata.documentCodeFontFamily }
  }
  if (block.type === 'heading') { style.outlineLevel = Math.max(0, block.level - 1); style.namedStyleType = block.level <= 5 ? block.level + 3 : 0 }
  else { delete style.outlineLevel; style.namedStyleType = namedStyles.indexOf(block.metadata.namedStyleType as DocumentNamedParagraphStyle) > 0 ? namedStyles.indexOf(block.metadata.namedStyleType as DocumentNamedParagraphStyle) : 1 }
  if (block.metadata.headingCollapsed !== undefined) result.headingCollapsed = block.metadata.headingCollapsed
  const hasNumbering = blockHasNumbering(block)
  if (hasNumbering) {
    const listId = instanceId ?? blockNumberingId(block) ?? block.id
    result.bullet = { ...bodyRecord(result.bullet), listId, listType: documentBodyListTypeForBlock(block), nestingLevel: blockNumberingLevelIndex(block) }
    delete bodyRecord(result.bullet).startNumber
  }
  if (!hasNumbering && block.type !== 'divider') { result.bullet = null; delete result.sourceBullet; delete result.directBullet }
  result.officeStudio = { ...documentBodyBlockSource(block), ...(result.bullet ? { numberingContract: cloneModel(result.bullet) } : {}) }
  return result
}

export interface DocumentBodyListResource { id: string; value: UniverBodyRecord; listId: string; counterId?: string }
export function documentBodyListForBlock(block: DocumentBlock, instanceId?: string): DocumentBodyListResource | undefined {
  if (!blockHasNumbering(block)) return undefined
  const preset = blockListPreset(block)
  const definition = block.type === 'checklist' ? checklistNumberingDefinition(block.checklist?.checked === true)
    : blockNumberingDefinition(block) ?? (preset ? presetNumberingDefinition(preset) : undefined)
  if (!definition) return undefined
  const listId = instanceId ?? blockNumberingId(block) ?? block.id, id = documentBodyListTypeForBlock(block)
  const result = cloneModel(bodyRecord(definition.abstract.univerBodyListSource))
  result.listType = id; delete result.counterIds; delete result.counterStarts
  result.officeStudio = { ...bodyRecord(result.officeStudio), version: 1, ...(block.list?.preset ? { preset: block.list.preset } : {}) }
  const baseLevels = numberingDefinitionFromParts(definition.abstract, { ...definition.instance, levelStartOverrides: undefined }).levels
  const sourceLevels = documentBodyListLevels(result.nestingLevel), writtenLevels = new Map<number, UniverBodyRecord>()
  for (const effective of definition.levels) {
    if (!Number.isSafeInteger(effective.ilvl) || effective.ilvl < 0 || effective.ilvl > 8 || writtenLevels.has(effective.ilvl)) throw new RangeError('原生列表层级无效或重复。')
    const level = baseLevels.find(level => level.ilvl === effective.ilvl) ?? effective
    const prior = sourceLevels[level.ilvl] ?? {}, rawStyle = bodyRecord(prior.textStyle)
    const style = createRun('', { fontFamily: level.bulletFontFamily ?? 'Arial', color: level.bulletColor ?? 'auto',
      fontSize: typeof level.runStyle?.fontSizePt === 'number' ? level.runStyle.fontSizePt : 11, bold: level.runStyle?.bold === true, italic: level.runStyle?.italic === true, strike: level.runStyle?.strike === true })
    const format = numberFormats.indexOf(level.format ?? 'decimal')
    const indent: UniverBodyRecord = {}
    for (const [key, target] of [['leftTwips', 'indentStart'], ['rightTwips', 'indentEnd'], ['hangingTwips', 'hanging'], ['firstLineTwips', 'indentFirstLine']] as const) {
      const value = level.indent?.[key]
      if (typeof value === 'number') indent[target] = { v: value / 15, u: 3 }
    }
    const start = documentBodyListInteger(level.start, 1, 0, Number.MAX_SAFE_INTEGER, '原生列表起始编号')
    const value: UniverBodyRecord = { ...prior, glyphType: format < 0 ? level.format : format, glyphFormat: level.text, ...(level.format === 'bullet' ? { glyphSymbol: level.text ?? '•' } : { glyphSymbol: undefined }),
      startNumber: start - 1, followNumberWith: level.suffix === 'nothing' ? 2 : level.suffix === 'space' ? 1 : 0,
      bulletAlignment: level.alignment === 'right' || level.alignment === 'end' ? 2 : level.alignment === 'center' ? 1 : 0,
      textStyle: { ...rawStyle, ...documentRunStyleToBody(style) }, isLegal: level.isLgl,
      paragraphProperties: { ...bodyRecord(prior.paragraphProperties), ...indent,
        ...(block.type === 'checklist' && block.checklist?.checked ? { textStyle: { ...bodyRecord(bodyRecord(prior.paragraphProperties).textStyle), st: { s: 1 } } } : {}) } }
    delete bodyRecord(value.officeStudio).nativeMissingLevel
    writtenLevels.set(level.ilvl, value)
  }
  result.nestingLevel = Array.from({ length: Math.max(-1, ...writtenLevels.keys()) + 1 }, (_value, ilvl) => writtenLevels.get(ilvl) ?? documentBodyMissingListLevel(ilvl))
  const starts: UniverBodyRecord = {}
  for (const [key, value] of Object.entries(numberingRecord(definition.instance.levelStartOverrides))) {
    if (!/^[0-8]$/.test(key) || typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) throw new RangeError('原生列表重启编号无效。')
    starts[key] = value - 1
  }
  if (Object.keys(starts).length) result.counterStarts = { [listId]: starts }
  const group = consecutiveNumberingGroup(definition, listId)
  return { id, value: result, listId, ...(definition.instance.consecutiveListId === undefined ? {} : { counterId: group.id }) }
}

function documentBodyListType(definition: DocumentNumberingDefinition, listId: string): string {
  const saved = definition.abstract.univerBodyPublicListType
  return typeof saved === 'string' && saved ? saved : `office-list:${String(definition.abstract.abstractNumId ?? 'default')}:${encodeURIComponent(listId)}`
}
function documentBodyListTypeForBlock(block: DocumentBlock): string {
  if (block.type === 'checklist') return block.checklist?.checked ? 'CHECK_LIST_CHECKED' : 'CHECK_LIST'
  const definition = blockNumberingDefinition(block), listId = blockNumberingId(block) ?? block.id
  if (definition && block.list?.preset && definition.instance.consecutiveListId !== undefined) {
    const group = consecutiveNumberingGroup(definition, listId)
    const saved = definition.abstract.univerBodyPublicListType
    if (typeof saved === 'string' && definition.abstract.univerBodyPublicCounterId === group.id) return saved
    return `office-counter:${block.list.preset}:${encodeURIComponent(group.id)}`
  }
  return block.list?.preset ?? (definition ? documentBodyListType(definition, listId) : blockListPreset(block) ?? listId)
}
function checklistNumberingDefinition(checked: boolean): DocumentNumberingDefinition {
  const levels = Array.from({ length: 9 }, (_, ilvl) => ({ ilvl, start: 1, format: 'bullet', text: checked ? '☑' : '☐', suffix: 'tab' as const,
    indent: { leftTwips: 315 * (ilvl + 1), hangingTwips: 315 }, bulletFontFamily: 'Arial', ...(checked ? { runStyle: { strike: true } } : {}) }))
  return { abstract: { abstractNumId: checked ? 'CHECK_LIST_CHECKED' : 'CHECK_LIST', levels }, instance: {}, levels }
}

function mergeBodyBullet(previous: unknown, next: unknown): UniverBodyRecord | null | undefined {
  if (next === null) return null
  if (next === undefined) return previous === null ? null : previous === undefined ? undefined : bodyRecord(previous)
  if (!next || typeof next !== 'object' || Array.isArray(next)) throw new TypeError('段落编号覆盖必须为对象或 null。')
  const inherited = { ...bodyRecord(previous) }, direct = bodyRecord(next)
  if (direct.listId !== undefined && direct.listId !== inherited.listId && direct.listType == null) delete inherited.listType
  return { ...inherited, ...direct }
}
