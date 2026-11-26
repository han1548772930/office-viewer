import type { CellRange, SparklineSpec } from '../../domain/model.ts'
import type { SparklineBrokenSourceRange, SparklineSourceReference } from '../../domain/sparkline-source.ts'
import { parseSparklineReference, shiftSparklineReference, type SparklineReference } from './sparkline-reference.ts'

export interface SparklineSourceSheet { id: string; name: string }
export interface SparklineSourceWorkbook<S extends SparklineSourceSheet = SparklineSourceSheet> {
  sheets: readonly S[]
  metadata?: Record<string, unknown>
}
export type SparklineSource =
  | { kind: 'a1'; text: string; reference: SparklineReference }
  | { kind: 'name'; text: string; name: string; sheetName?: string }
  | { kind: 'external'; text: string; qualifier?: string; addressText?: string; unitId?: string; sheetName?: string }
  | { kind: 'three-dimensional'; text: string; qualifier: string; addressText?: string }
  | { kind: 'broken'; text: string }
export type SparklineSourceResolution<S extends SparklineSourceSheet = SparklineSourceSheet> =
  | { status: 'resolved'; source: SparklineSource; sheet: S; range: CellRange; reference: SparklineReference; viaName?: string }
  | { status: 'broken' | 'unresolved'; source: SparklineSource; reason: 'deleted-range' | 'deleted-sheet' | 'external-workbook' | 'three-dimensional'
      | 'name-not-found' | 'name-ambiguous' | 'name-cycle' | 'name-definition-not-range' }
export interface SparklinePublicSourceRange {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  sheetId?: string
  unitId?: string
  rangeType?: 0 | 1 | 2 | 3
  officeStudioSourceReference: string
  officeStudioSourceMetadata?: SparklineSourceReference
}
export type SparklineSourceInput = string | Pick<SparklineSpec, 'sourceRange' | 'sourceReference'>
const encoder = new TextEncoder()
const object = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === 'object' && !Array.isArray(value))
const SENTINEL_KEYS = new Set(['startRow', 'endRow', 'startColumn', 'endColumn', 'sheetId', 'unitId', 'rangeType'])

export function normalizeSparklineSourceReference(value: unknown): SparklineSourceReference | undefined {
  if (value === undefined) return undefined
  if (!object(value) || value.version !== 1 || Object.keys(value).some(key => key !== 'version' && key !== 'invalidRange') || !object(value.invalidRange)) {
    throw new TypeError('迷你图断开来源的版本或结构无效。')
  }
  const raw = value.invalidRange
  if (Object.keys(raw).some(key => !SENTINEL_KEYS.has(key)) || [raw.startRow, raw.endRow, raw.startColumn, raw.endColumn].some(value => value !== -9999)) {
    throw new TypeError('迷你图断开来源必须保留固定版的 -9999 范围标记。')
  }
  const invalidRange: SparklineBrokenSourceRange = { startRow: -9999, endRow: -9999, startColumn: -9999, endColumn: -9999 }
  for (const key of ['sheetId', 'unitId'] as const) if (raw[key] !== undefined) {
    const id = raw[key]
    if (typeof id !== 'string' || !id.trim() || id.length > 4096 || encoder.encode(id).byteLength > 4096 || /[\u0000-\u001f\u007f-\u009f]/.test(id)) throw new TypeError('迷你图断开来源的单元标识无效。')
    invalidRange[key] = id
  }
  if (raw.rangeType !== undefined) {
    if (raw.rangeType !== 0 && raw.rangeType !== 1 && raw.rangeType !== 2 && raw.rangeType !== 3) throw new TypeError('迷你图断开来源的范围类型无效。')
    invalidRange.rangeType = raw.rangeType
  }
  return { version: 1, invalidRange }
}

function name(value: string): boolean {
  return value.length <= 255 && /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(value)
    && !/^[rc]$/i.test(value) && !/^r\d+c\d+$/i.test(value) && !parseSparklineReference(value)
}
function unquote(value: string): string | undefined {
  if (value.startsWith("'") && value.endsWith("'")) {
    const inner = value.slice(1, -1)
    return inner.replaceAll("''", '').includes("'") ? undefined : inner.replaceAll("''", "'")
  }
  return value && !value.includes("'") && value.trim() === value ? value : undefined
}
function splitSource(text: string): { qualifier?: string; body: string; broken: boolean } | undefined {
  let quoted = false, brackets = 0, delimiter = -1, broken = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'") {
      if (quoted && text[index + 1] === "'") { index += 1; continue }
      quoted = !quoted
    } else if (!quoted) {
      if (character === '[') brackets += 1
      else if (character === ']') { if (!brackets) return undefined; brackets -= 1 }
      else if (!brackets && text.slice(index, index + 5).toUpperCase() === '#REF!') { broken = true; index += 4 }
      else if (character === '!' && !brackets) { if (delimiter !== -1) return undefined; delimiter = index }
    }
  }
  if (quoted || brackets) return undefined
  return { ...(delimiter === -1 ? {} : { qualifier: text.slice(0, delimiter) }), body: text.slice(delimiter + 1), broken }
}

/** Admit known reference syntax while preserving named, external and broken source text. */
export function parseSparklineSource(input: string): SparklineSource {
  if (typeof input !== 'string' || !input.trim() || input.length > 4096 || encoder.encode(input).byteLength > 4096 || /[\u0000-\u001f\u007f-\u009f]/.test(input)) {
    throw new TypeError('迷你图数据来源为空、含控制字符或超出引用预算。')
  }
  const text = input.trim().replace(/^=/, '').trim(), reference = parseSparklineReference(text)
  if (reference) return { kind: 'a1', text: input, reference }
  const split = splitSource(text)
  if (!split) throw new TypeError('迷你图来源的工作表引号或范围结构无效。')
  if (split.broken) return { kind: 'broken', text: input }
  if (split.qualifier !== undefined) {
    const qualifier = unquote(split.qualifier)
    if (!qualifier) throw new TypeError('迷你图来源的工作表限定无效。')
    const bodyReference = parseSparklineReference(split.body), bodyName = name(split.body)
    if (!bodyReference && !bodyName) throw new TypeError('迷你图来源必须是有效范围或名称。')
    if (qualifier.includes('[')) {
      if (!/\[[^\[\]]+\]/.test(qualifier) || /[\[\]]/.test(qualifier.replace(/\[[^\[\]]+\]/, ''))) throw new TypeError('迷你图外部来源的工作簿限定无效。')
      const bracket = /\[([^\[\]]+)\]/.exec(qualifier)!, sheetName = qualifier.slice(bracket.index + bracket[0].length)
      return { kind: 'external', text: input, qualifier: split.qualifier, unitId: bracket[1], ...(sheetName ? { sheetName } : {}), ...(bodyReference ? { addressText: split.body } : {}) }
    }
    if (qualifier.includes(':')) {
      const sheets = qualifier.split(':')
      if (sheets.length !== 2 || sheets.some(sheet => !sheet || /[\[\]\\/?*]/.test(sheet))) throw new TypeError('迷你图三维来源的工作表边界无效。')
      return { kind: 'three-dimensional', text: input, qualifier: split.qualifier, ...(bodyReference ? { addressText: split.body } : {}) }
    }
    if (/[\[\]:\\/?*]/.test(qualifier) || !bodyName) throw new TypeError('迷你图命名来源的工作表限定无效。')
    return { kind: 'name', text: input, name: split.body, sheetName: qualifier }
  }
  const externalName = /^\[([^\[\]]+)\]/.exec(text)
  if (externalName && name(text.slice(externalName[0].length))) return { kind: 'external', text: input, unitId: externalName[1] }
  if (name(text)) return { kind: 'name', text: input, name: text }
  throw new TypeError('迷你图来源必须是有效范围、名称或保留的断开来源。')
}

export function normalizeSparklineSourceText(input: string): string {
  const source = parseSparklineSource(input)
  return source.kind === 'a1' ? source.reference.normalized : source.text
}

/** Source expressions may contain quoted sheets or an external workbook with separators. */
export function parseSparklineSourceList(input: string): string[] {
  if (typeof input !== 'string') throw new TypeError('请填写迷你图数据来源。')
  const result: string[] = []
  let quoted = false, brackets = 0, start = 0
  for (let index = 0; index < input.length; index += 1) {
    const character = input[index]
    if (character === "'") {
      if (quoted && input[index + 1] === "'") { index += 1; continue }
      quoted = !quoted
    } else if (!quoted) {
      if (character === '[') brackets += 1
      else if (character === ']') { if (!brackets) throw new TypeError('迷你图外部来源的括号未匹配。'); brackets -= 1 }
      else if (!brackets && (character === ',' || character === ';' || character === '\n' || character === '\r')) {
        const item = input.slice(start, index).trim()
        if (item) result.push(normalizeSparklineSourceText(item))
        start = index + 1
      }
    }
  }
  if (quoted || brackets) throw new TypeError('迷你图来源的引号或括号未闭合。')
  const last = input.slice(start).trim()
  if (last) result.push(normalizeSparklineSourceText(last))
  if (!result.length) throw new TypeError('请填写迷你图数据来源。')
  return result
}
function findSheet<S extends SparklineSourceSheet>(sheets: readonly S[], identity: string): S | undefined {
  return sheets.find(sheet => sheet.id === identity) ?? sheets.find(sheet => sheet.name.toLocaleLowerCase() === identity.toLocaleLowerCase())
}
function resolve<S extends SparklineSourceSheet>(source: SparklineSource, host: S, workbook: SparklineSourceWorkbook<S>, visited: Set<string>): SparklineSourceResolution<S> {
  if (source.kind === 'broken') return { status: 'broken', source, reason: 'deleted-range' }
  if (source.kind === 'external' || source.kind === 'three-dimensional') return { status: 'unresolved', source, reason: source.kind === 'external' ? 'external-workbook' : 'three-dimensional' }
  const sheetName = source.kind === 'a1' ? source.reference.sheetName : source.sheetName
  const sheet = sheetName === undefined ? host : findSheet(workbook.sheets, sheetName)
  if (!sheet) return { status: 'broken', source, reason: 'deleted-sheet' }
  if (source.kind === 'a1') return { status: 'resolved', source, sheet, range: source.reference.range, reference: source.reference }
  const identity = `${sheet.id.toLocaleLowerCase()}:${source.name.toLocaleLowerCase()}`
  if (visited.has(identity) || visited.size >= 32) return { status: 'unresolved', source, reason: 'name-cycle' }
  visited.add(identity)
  const matching = (items: unknown): Record<string, unknown>[] => Array.isArray(items) ? items.filter((item: unknown): item is Record<string, unknown> => object(item) && typeof item.name === 'string'
    && item.name.toLocaleLowerCase() === source.name.toLocaleLowerCase()) : []
  const local = matching(workbook.metadata?.localDefinedNames).filter(item => object(item) && typeof item.sheet === 'string'
    && (item.sheet === sheet.id || item.sheet.toLocaleLowerCase() === sheet.name.toLocaleLowerCase()))
  const globals = matching(workbook.metadata?.definedNames)
  const definitions = local.length ? local : globals
  if (!definitions.length) return { status: 'unresolved', source, reason: 'name-not-found' }
  if (definitions.length !== 1) return { status: 'unresolved', source, reason: 'name-ambiguous' }
  const definition = definitions[0]
  if (!object(definition) || typeof definition.refersTo !== 'string') return { status: 'unresolved', source, reason: 'name-definition-not-range' }
  let target: SparklineSource
  try { target = parseSparklineSource(definition.refersTo) } catch { return { status: 'unresolved', source, reason: 'name-definition-not-range' } }
  const result = resolve(target, sheet, workbook, visited)
  return result.status === 'resolved' ? { ...result, source, viaName: source.name } : { ...result, source }
}

/** One source resolver serves Canvas, Facade, source-change events and permission footprints. */
export function resolveSparklineSource<S extends SparklineSourceSheet>(input: SparklineSourceInput, host: S, workbook: SparklineSourceWorkbook<S>): SparklineSourceResolution<S> {
  const source = parseSparklineSource(typeof input === 'string' ? input : input.sourceRange)
  if (typeof input !== 'string' && input.sourceReference !== undefined) {
    normalizeSparklineSourceReference(input.sourceReference)
    return { status: 'broken', source, reason: 'deleted-range' }
  }
  return resolve(source, host, workbook, new Set())
}

export function sparklineSourceToPublicRange<S extends SparklineSourceSheet>(input: SparklineSourceInput, host: S, workbook: SparklineSourceWorkbook<S>): SparklinePublicSourceRange {
  const text = typeof input === 'string' ? input : input.sourceRange, resolved = resolveSparklineSource(input, host, workbook)
  const retained = typeof input === 'string' ? undefined : normalizeSparklineSourceReference(input.sourceReference)
  let range: Omit<SparklinePublicSourceRange, 'officeStudioSourceReference' | 'officeStudioSourceMetadata'>
  if (retained) range = { ...retained.invalidRange }
  else if (resolved.status === 'resolved') range = { startRow: resolved.range.startRow, endRow: resolved.range.endRow,
    startColumn: resolved.range.startCol, endColumn: resolved.range.endCol,
    ...(resolved.sheet.id === host.id && resolved.source.kind === 'a1' && resolved.source.reference.sheetName === undefined ? {} : { sheetId: resolved.sheet.id }) }
  else if (resolved.source.kind === 'a1') range = { startRow: resolved.source.reference.range.startRow, endRow: resolved.source.reference.range.endRow,
    startColumn: resolved.source.reference.range.startCol, endColumn: resolved.source.reference.range.endCol,
    ...(resolved.source.reference.sheetName === undefined ? {} : { sheetId: resolved.source.reference.sheetName }) }
  else if ((resolved.source.kind === 'external' || resolved.source.kind === 'three-dimensional') && resolved.source.addressText) {
    const reference = parseSparklineReference(resolved.source.addressText)!
    range = { startRow: reference.range.startRow, endRow: reference.range.endRow, startColumn: reference.range.startCol, endColumn: reference.range.endCol,
      ...(resolved.source.kind === 'external' ? { ...(resolved.source.unitId ? { unitId: resolved.source.unitId } : {}), ...(resolved.source.sheetName ? { sheetId: resolved.source.sheetName } : {}) }
        : { sheetId: unquote(resolved.source.qualifier) }) }
  }
  else range = { startRow: -9999, endRow: -9999, startColumn: -9999, endColumn: -9999, rangeType: 0 }
  return { ...range, officeStudioSourceReference: text, ...(retained ? { officeStudioSourceMetadata: retained } : {}) }
}

/** Named sources keep their identity; foreign/3D A1 anchors move without losing their qualifier. */
export function shiftSparklineSource(input: string, rowDelta: number, columnDelta: number): string {
  if (!Number.isSafeInteger(rowDelta) || !Number.isSafeInteger(columnDelta)) throw new TypeError('迷你图移动距离无效。')
  const source = parseSparklineSource(input)
  if (source.kind === 'a1') return shiftSparklineReference(input, rowDelta, columnDelta)
  if ((source.kind === 'external' || source.kind === 'three-dimensional') && source.addressText && source.qualifier) {
    return `${source.qualifier}!${shiftSparklineReference(source.addressText, rowDelta, columnDelta)}`
  }
  return source.text
}
