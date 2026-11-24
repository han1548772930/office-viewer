import type { DocumentBlock, DocumentDefaultStyle, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { applyNamedParagraphStyle, createParagraph, createRun, DOCUMENT_NAMED_PARAGRAPH_STYLES, type DocumentNamedParagraphStyle } from '../features/document/model/document-model'
import { blockHasNumbering, blockNumberingDefinition, blockNumberingId, blockNumberingLevelIndex, type DocumentNumberingDefinition } from '../features/document/model/document-numbering'
import { isListPreset, presetNumberingDefinition } from '../features/document/model/document-lists'
import { bodyBoolean, bodyColor, bodyLength, bodyNumber, bodyRecord, documentParagraphStyleFromBody, documentParagraphStyleToBody, documentRunStyleFromBody, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentBodyArray, documentBodyIdentity, type DocumentBodyDecodeOptions } from './document-body-types'

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
    return { ...parent, ...raw, textStyle: { ...bodyRecord(parent.textStyle), ...bodyRecord(raw.textStyle) },
      paragraphStyle: { ...bodyRecord(parent.paragraphStyle), ...bodyRecord(raw.paragraphStyle) } }
  }
  return resolve(currentId, 0)
}

export function documentParagraphFromBody(raw: UniverBodyRecord, runs: InlineRun[], options: DocumentBodyDecodeOptions,
  lists: Record<string, UniverBodyRecord> = {}): DocumentBlock {
  const inherited = documentBodyNamedStyle(raw.styleId, options.styles), paragraphStyle = { ...bodyRecord(inherited.paragraphStyle), ...bodyRecord(raw.paragraphStyle) }
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
  block.metadata = { ...block.metadata, univerBodyParagraphSource: { version: 1, paragraph: cloneModel(raw) } }
  if (bodyBoolean(raw.headingCollapsed) !== undefined) block.metadata.headingCollapsed = bodyBoolean(raw.headingCollapsed)
  const bullet = raw.bullet === undefined ? inherited.bullet : raw.bullet
  if (bullet !== null && bullet !== undefined) applyBullet(block, bodyRecord(bullet), lists, options.defaults)
  return block
}

export function documentBodyParagraphMarks(raw: UniverBodyRecord, options: DocumentBodyDecodeOptions): UniverBodyRecord {
  const inherited = documentBodyNamedStyle(raw.styleId, options.styles)
  const paragraph = { ...bodyRecord(inherited.paragraphStyle), ...bodyRecord(raw.paragraphStyle) }, named = bodyNumber(paragraph.namedStyleType)
  const preset = named !== undefined && named > 0 && namedStyles[named] ? DOCUMENT_NAMED_PARAGRAPH_STYLES[namedStyles[named]!] : undefined
  return { ...(preset ? { fs: preset.fontSize, bl: Number(preset.bold), ...(preset.color ? { cl: { rgb: preset.color } } : {}) } : {}),
    ...bodyRecord(inherited.textStyle), ...bodyRecord(bodyRecord(inherited.paragraphStyle).textStyle), ...bodyRecord(bodyRecord(raw.paragraphStyle).textStyle) }
}

function applyBullet(block: DocumentBlock, bullet: UniverBodyRecord, lists: Record<string, UniverBodyRecord>, defaults: DocumentDefaultStyle): void {
  const listType = documentBodyIdentity(bullet.listType, '列表类型'), listId = documentBodyIdentity(bullet.listId ?? listType, '列表'), level = bodyNumber(bullet.nestingLevel) ?? 0
  if (!Number.isSafeInteger(level) || level < 0 || level > 8) throw new RangeError('列表层级无效。')
  const raw = lists[listType]
  let definition: DocumentNumberingDefinition
  if (raw) {
    const levels = documentBodyArray(raw.nestingLevel, '列表层级').map((source, ilvl) => {
      const formatIndex = typeof source.glyphType === 'string' ? undefined : bodyNumber(source.glyphType), format = source.glyphSymbol !== undefined ? 'bullet' : typeof source.glyphType === 'string' ? source.glyphType : numberFormats[formatIndex ?? 2]
      if (!format || ilvl > 8) throw new RangeError('列表编号格式或层级无效。')
      const start = bodyNumber(source.startNumber) ?? 0
      if (!Number.isSafeInteger(start) || start < 0 || start >= Number.MAX_SAFE_INTEGER) throw new RangeError('列表起始编号无效。')
      const style = documentRunStyleFromBody(source.textStyle), properties = bodyRecord(source.paragraphProperties), twips = (value: unknown) => {
        const length = bodyLength(value, defaults.text.fontSize); return length === undefined ? undefined : length * 15
      }
      return { ilvl, start: start + 1, format, text: typeof source.glyphSymbol === 'string' ? source.glyphSymbol : typeof source.glyphFormat === 'string' ? source.glyphFormat : `%${ilvl + 1}.`,
        suffix: source.followNumberWith === 2 ? 'nothing' as const : source.followNumberWith === 1 ? 'space' as const : 'tab' as const,
        isLgl: bodyBoolean(source.isLegal), indent: { leftTwips: twips(properties.indentStart), rightTwips: twips(properties.indentEnd),
          hangingTwips: twips(properties.hanging), firstLineTwips: twips(properties.indentFirstLine) },
        bulletFontFamily: style.fontFamily, bulletColor: style.color ?? bodyColor(bodyRecord(source.textStyle).cl),
        runStyle: { fontSizePt: style.fontSize, bold: style.bold, italic: style.italic, color: style.color, fontFamily: style.fontFamily } }
    })
    if (!levels.length || !levels[level]) throw new RangeError('列表缺少当前层级定义。')
    if (raw.counterIds !== undefined && (!Array.isArray(raw.counterIds) || raw.counterIds.some(id => typeof id !== 'string' || !id))) throw new TypeError('共享列表计数器标识无效。')
    definition = { abstract: { abstractNumId: listType, levels, univerBodyListSource: cloneModel(raw) },
      instance: { numId: listId, ...(Array.isArray(raw.counterIds) && raw.counterIds.includes(listId) ? { consecutiveListId: raw.counterIds[0] } : {}) }, levels }
    const restarts = bodyRecord(bodyRecord(raw.counterStarts)[listId]), ownStart = bodyNumber(bullet.startNumber)
    const starts: Record<string, number> = {}
    for (const [key, value] of Object.entries(restarts)) {
      const start = bodyNumber(value)!
      if (!/^[0-8]$/.test(key) || !Number.isSafeInteger(start) || start < 0 || start >= Number.MAX_SAFE_INTEGER) throw new RangeError('列表重启编号无效。')
      starts[key] = start + 1
    }
    if (ownStart !== undefined) {
      if (!Number.isSafeInteger(ownStart) || ownStart < 0 || ownStart >= Number.MAX_SAFE_INTEGER) throw new RangeError('段落起始编号无效。')
      starts[level] = ownStart + 1
    }
    if (Object.keys(starts).length) definition.instance.levelStartOverrides = starts
  } else if (isListPreset(listType)) {
    definition = presetNumberingDefinition(listType)
    block.list = { id: listId, preset: listType }
  } else throw new RangeError(`列表 ${listType} 缺少资源定义。`)
  const numbered = definition.levels.find(item => item.ilvl === level)?.format !== 'bullet'
  if (block.type !== 'heading') { block.type = numbered ? 'numbered' : 'bullet'; block.level = level }
  else block.metadata.numberingLevel = level
  block.metadata.numberingListId = listId; block.metadata.numberingDefinition = definition
}

export function documentParagraphToBody(block: DocumentBlock, end: number): UniverBodyRecord {
  const source = bodyRecord(block.metadata.univerBodyParagraphSource), result = cloneModel(bodyRecord(source.paragraph))
  result.startIndex = end; result.paragraphId = block.id
  result.paragraphStyle = documentParagraphStyleToBody(block.style, result.paragraphStyle)
  const style = bodyRecord(result.paragraphStyle)
  if (block.type === 'heading') { style.outlineLevel = Math.max(0, block.level - 1); style.namedStyleType = block.level <= 5 ? block.level + 3 : 0 }
  else { delete style.outlineLevel; style.namedStyleType = namedStyles.indexOf(block.metadata.namedStyleType as DocumentNamedParagraphStyle) > 0 ? namedStyles.indexOf(block.metadata.namedStyleType as DocumentNamedParagraphStyle) : 1 }
  if (block.metadata.headingCollapsed !== undefined) result.headingCollapsed = block.metadata.headingCollapsed
  if (blockHasNumbering(block)) {
    const listId = blockNumberingId(block) ?? block.id, definition = blockNumberingDefinition(block)
    result.bullet = { ...bodyRecord(result.bullet), listId, listType: block.list?.preset ?? String(definition?.abstract.abstractNumId ?? listId), nestingLevel: blockNumberingLevelIndex(block) }
  } else { result.bullet = null; delete result.sourceBullet; delete result.directBullet }
  return result
}

export function documentBodyListForBlock(block: DocumentBlock): { id: string; value: UniverBodyRecord } | undefined {
  const definition = blockNumberingDefinition(block)
  if (!definition) return undefined
  const id = block.list?.preset ?? String(definition.abstract.abstractNumId ?? blockNumberingId(block) ?? block.id)
  const result = cloneModel(bodyRecord(definition.abstract.univerBodyListSource))
  result.listType = id
  result.nestingLevel = definition.levels.map(level => {
    const prior = documentBodyArray(result.nestingLevel, '列表层级')[level.ilvl] ?? {}, rawStyle = bodyRecord(prior.textStyle)
    const style = createRun('', { fontFamily: level.bulletFontFamily ?? 'Arial', color: level.bulletColor ?? 'auto',
      fontSize: typeof level.runStyle?.fontSizePt === 'number' ? level.runStyle.fontSizePt : 11, bold: level.runStyle?.bold === true, italic: level.runStyle?.italic === true })
    const format = numberFormats.indexOf(level.format ?? 'decimal')
    return { ...prior, glyphType: format < 0 ? 60 : format, glyphFormat: level.text, ...(level.format === 'bullet' ? { glyphSymbol: level.text ?? '•' } : { glyphSymbol: undefined }),
      startNumber: Math.max(0, (level.start ?? 1) - 1), followNumberWith: level.suffix === 'nothing' ? 2 : level.suffix === 'space' ? 1 : 0,
      textStyle: { ...rawStyle, ...documentRunStyleToBody(style) }, isLegal: level.isLgl,
      paragraphProperties: { ...bodyRecord(prior.paragraphProperties), ...Object.fromEntries([
        ['leftTwips', 'indentStart'], ['rightTwips', 'indentEnd'], ['hangingTwips', 'hanging'], ['firstLineTwips', 'indentFirstLine'],
      ].flatMap(([key, target]) => typeof level.indent?.[key!] === 'number' ? [[target!, { v: Number(level.indent[key!]) / 15, u: 3 }]] : [])) } }
  })
  return { id, value: result }
}
