import type { DocumentModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import { applyRunPatch, documentId, findDocumentBlock, type DocumentRun } from '../features/document/model/document-model'
import type { DocumentTextStylePatch } from '../features/document/services/document-facade'
import { assertDocumentBodySerializable } from './document-body-codec'
import { bodyRecord, documentRunStyleFromBody, documentRunStyleSource, documentRunStyleToBody, type UniverBodyRecord } from './document-body-style'
import { documentBodyLocalOffset } from './document-body-types'
import { assertDocumentPublicRange, type DocumentPublicBodyIndex } from './document-public-body'

export interface UniverDocumentColorStyle extends UniverBodyRecord { rgb?: string | null; th?: number; auto?: 0 | 1 }
export interface UniverDocumentTextDecoration extends UniverBodyRecord { s: 0 | 1; c?: 0 | 1; cl?: UniverDocumentColorStyle; t?: number }
/** Fixed ITextStyle keys. fs/sc/pos use points; source text remains unchanged. */
export interface UniverDocumentTextStyle extends UniverBodyRecord {
  ff?: string | null; eastAsiaFontFamily?: string | null; fs?: number; bl?: 0 | 1; it?: 0 | 1
  ul?: UniverDocumentTextDecoration; st?: UniverDocumentTextDecoration; bbl?: UniverDocumentTextDecoration; ol?: UniverDocumentTextDecoration
  cl?: UniverDocumentColorStyle | null; bg?: UniverDocumentColorStyle | null; va?: 1 | 2 | 3 | null
  hidden?: boolean; caps?: boolean; smallCaps?: boolean; sc?: number; pos?: number; sa?: number; kerning?: number
  lang?: string; altLang?: string; bidiLang?: string; fontHint?: 'default' | 'eastAsia' | 'cs'; eastAsiaFontTheme?: 'major' | 'minor' | null
  textFill?: UniverBodyRecord; textOutline?: UniverBodyRecord; glow?: UniverBodyRecord; outerShadow?: UniverBodyRecord
  symbol?: { font: string; code: number }; positionedTab?: UniverBodyRecord
}
export interface UniverDocumentTextStyleRun { startOffset: number; endOffset: number; textStyle: UniverDocumentTextStyle }
export interface UniverDocumentTextRangeDescription {
  startOffset: number; endOffset: number; segmentId: string; text: string; length: number
  explicitTextStyleRuns: UniverDocumentTextStyleRun[]; commonExplicitTextStyle: UniverDocumentTextStyle
}
const LOCAL_KEYS = ['bold', 'italic', 'underline', 'strike', 'verticalAlign', 'color', 'fontFamily', 'fontSize', 'highlight', 'href']
export function isLocalDocumentTextStyle(style: UniverDocumentTextStyle | DocumentTextStylePatch): style is DocumentTextStylePatch {
  if (!style || typeof style !== 'object' || Array.isArray(style)) throw new TypeError('文字格式必须为对象。')
  const local = Object.keys(style).filter(key => LOCAL_KEYS.includes(key))
  if (local.length && Object.keys(style).some(key => !LOCAL_KEYS.includes(key))) throw new TypeError('文字格式不能混用本地字段与公开字段。')
  return local.length > 0
}
export function validatePublicDocumentTextStyle(style: UniverDocumentTextStyle): UniverDocumentTextStyle {
  assertDocumentBodySerializable(style)
  if (!style || typeof style !== 'object' || Array.isArray(style)) throw new TypeError('公开文字格式必须为对象。')
  const result = cloneModel(style)
  for (const key of Object.keys(result)) if (result[key] === undefined) delete result[key]
  documentRunStyleFromBody(result)
  for (const key of ['hidden', 'caps', 'smallCaps'] as const) if (result[key] !== undefined && typeof result[key] !== 'boolean') throw new TypeError(`公开文字 ${key} 必须为布尔值。`)
  for (const key of ['sc', 'pos', 'sa', 'kerning'] as const) {
    const value = result[key]
    if (value !== undefined && (typeof value !== 'number' || !Number.isFinite(value) || (key === 'sa' ? value <= 0 : key === 'kerning' && value < 0))) throw new RangeError(`公开文字 ${key} 无效。`)
  }
  for (const key of ['lang', 'altLang', 'bidiLang'] as const) if (result[key] !== undefined && (typeof result[key] !== 'string' || !result[key]!.trim())) throw new TypeError('公开文字语言无效。')
  for (const key of ['cl', 'bg'] as const) if (result[key] !== undefined && result[key] !== null) {
    const color = result[key]!
    if (!color || typeof color !== 'object' || Array.isArray(color) || color.rgb !== undefined && color.rgb !== null && typeof color.rgb !== 'string'
      || color.th !== undefined && (typeof color.th !== 'number' || !Number.isSafeInteger(color.th))) throw new TypeError('公开文字颜色无效。')
  }
  for (const key of ['ul', 'st', 'bbl', 'ol'] as const) if (result[key] !== undefined) {
    const decoration = result[key]!
    if (!decoration || typeof decoration !== 'object' || Array.isArray(decoration) || ![0, 1].includes(decoration.s)
      || decoration.c !== undefined && ![0, 1].includes(decoration.c) || decoration.t !== undefined && (!Number.isInteger(decoration.t) || decoration.t < 0 || decoration.t > 19)) throw new TypeError('公开文字装饰无效。')
  }
  if (result.fontHint !== undefined && !['default', 'eastAsia', 'cs'].includes(result.fontHint)) throw new TypeError('公开文字字体提示无效。')
  if (result.eastAsiaFontTheme !== undefined && result.eastAsiaFontTheme !== null && !['major', 'minor'].includes(result.eastAsiaFontTheme)) throw new TypeError('公开东亚字体主题无效。')
  if (result.symbol !== undefined && (!result.symbol || typeof result.symbol.font !== 'string' || !result.symbol.font.trim()
    || !Number.isSafeInteger(result.symbol.code) || result.symbol.code < 0 || result.symbol.code > 0xFFFF)) throw new TypeError('公开字体符号无效。')
  return result
}

/** Explicit gaps remain gaps; inherited fonts and colors are not reported as
 * direct formatting. Returned records are detached and document-relative. */
export function publicDocumentExplicitTextStyleRuns(index: DocumentPublicBodyIndex, start: number, end: number): UniverDocumentTextStyleRun[] {
  assertDocumentPublicRange(index, start, end)
  return (index.data.body.textRuns ?? []).filter(run => run.st < end && run.ed > start).map(run => ({
    startOffset: Math.max(start, run.st), endOffset: Math.min(end, run.ed), textStyle: cloneModel(bodyRecord(run.ts)) as UniverDocumentTextStyle,
  }))
}
export function commonPublicDocumentExplicitTextStyle(index: DocumentPublicBodyIndex, start: number, end: number): UniverDocumentTextStyle {
  const runs = publicDocumentExplicitTextStyleRuns(index, start, end)
  if (start === end || !runs.length || runs[0]!.startOffset > start) return {}
  let offset = start
  const common = cloneModel(runs[0]!.textStyle)
  for (const run of runs) {
    if (run.startOffset > offset) return {}
    for (const key of Object.keys(common)) if (JSON.stringify(common[key]) !== JSON.stringify(run.textStyle[key])) delete common[key]
    offset = Math.max(offset, run.endOffset)
  }
  return offset < end ? {} : common
}

function patchedPublicRun(run: DocumentRun, patch: UniverDocumentTextStyle, model: DocumentModel): DocumentRun {
  const explicit = { ...documentRunStyleToBody(run), ...patch }, inherited = bodyRecord(bodyRecord(run.metadata?.univerTextStyleResolution).inherited)
  const resolved = documentRunStyleFromBody({ ...inherited, ...explicit }), defaults = documentDefaultStyle(model).text
  delete resolved.metadata
  const updated = applyRunPatch(run, { ...defaults, ...resolved, highlight: resolved.highlight ?? defaults.highlight,
    eastAsiaFontFamily: resolved.eastAsiaFontFamily ?? defaults.eastAsiaFontFamily,
    verticalAlign: resolved.verticalAlign ?? defaults.verticalAlign ?? 'baseline' })
  updated.metadata = { ...run.metadata, ...documentRunStyleSource(updated, explicit, inherited) }
  return updated
}
function formatRuns(model: DocumentModel, runs: readonly DocumentRun[], start: number, end: number, patch: UniverDocumentTextStyle): DocumentRun[] {
  const result: DocumentRun[] = []; let offset = 0
  for (const run of runs) {
    const from = Math.max(0, start - offset), to = Math.min(run.text.length, end - offset)
    offset += run.text.length
    if (to <= from) { result.push(run); continue }
    const updated = patchedPublicRun(run, patch, model)
    if (JSON.stringify(updated) === JSON.stringify(run)) { result.push(run); continue }
    if (from > 0) result.push({ ...run, text: run.text.slice(0, from) })
    result.push({ ...updated, id: from > 0 ? documentId('run') : run.id, text: run.text.slice(from, to) })
    if (to < run.text.length) result.push({ ...run, id: documentId('run'), text: run.text.slice(to) })
  }
  return result
}
/** Works on the same detached transaction candidate as local formatting. A
 * delimiter-only range does not expand into adjacent characters. */
export function setPublicDocumentTextStyle(model: DocumentModel, index: DocumentPublicBodyIndex, start: number, end: number,
  patch: UniverDocumentTextStyle): void {
  assertDocumentPublicRange(index, start, end)
  if (start === end || !Object.keys(patch).length) return
  for (const position of index.positions) {
    const block = findDocumentBlock(model, position.block.id)
    if (!block) throw new RangeError('公开文字格式目标段落已不存在。')
    const from = Math.max(start, position.start), to = Math.min(end, position.end)
    if (to > from && position.spans.some(span => span.targetEnd > span.targetStart && span.sourceStart < to && span.sourceEnd > from)) {
      block.runs = formatRuns(model, block.runs, documentBodyLocalOffset(position, from), documentBodyLocalOffset(position, to, 'end'), patch)
    }
  }
  for (const paragraph of index.data.body.paragraphs ?? []) if (paragraph.startIndex >= start && paragraph.startIndex < end && paragraph.paragraphId) {
    const block = findDocumentBlock(model, paragraph.paragraphId)
    if (!block) throw new RangeError('公开文字格式段落标记已不存在。')
    block.metadata.univerBodyParagraphMarkTextStyle = { ...bodyRecord(block.metadata.univerBodyParagraphMarkTextStyle
      ?? bodyRecord(paragraph.paragraphStyle).paragraphMarkTextStyle), ...patch }
  }
}
