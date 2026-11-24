import type { DocumentDefaultStyle, DocumentParagraphStyle, InlineRun } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import type { RunPatch } from '../features/document/model/document-model'

export type UniverBodyRecord = Record<string, unknown>
export function bodyRecord(value: unknown): UniverBodyRecord {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as UniverBodyRecord : {}
}
export function bodyNumber(value: unknown): number | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new TypeError('文档内容数值必须有限。')
  return value
}
export function bodyBoolean(value: unknown): boolean | undefined {
  if (value === undefined || value === null) return undefined
  if (value === true || value === 1) return true
  if (value === false || value === 0) return false
  throw new TypeError('文档内容开关必须是布尔值或 0/1。')
}
export function bodyColor(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value === 'string') return value
  const raw = bodyRecord(value)
  if (typeof raw.rgb === 'string') return /^[0-9a-f]{6}$/i.test(raw.rgb) ? `#${raw.rgb}` : raw.rgb
  if (raw.auto === true || raw.auto === 1) return 'auto'
  return undefined
}

/** NumberUnitType from fixed core/types/interfaces/i-document-data.ts.
 * Domain geometry uses CSS pixels; font sizes remain points. */
export function bodyLength(value: unknown, fontSize: number, reference = fontSize * 4 / 3): number | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value === 'number') return bodyNumber(value)! * 4 / 3
  const raw = bodyRecord(value), number = bodyNumber(raw.v)
  if (number === undefined) throw new TypeError('文档长度缺少 NumberUnit.v。')
  switch (raw.u ?? 0) {
    case 0: return number * 4 / 3
    case 1: return number * reference
    case 2: return number * fontSize * 4 / 3
    case 3: return number
    case 4: return number * reference / 100
    default: throw new RangeError('文档长度单位无效。')
  }
}

/** Map public run keys while retaining unmodeled authored fields as source data. */
export function documentRunStyleFromBody(value: unknown): RunPatch {
  const raw = bodyRecord(value), result: RunPatch = {}
  for (const [source, target] of [['bl', 'bold'], ['it', 'italic']] as const) {
    const current = bodyBoolean(raw[source]); if (current !== undefined) result[target] = current
  }
  for (const [source, target] of [['ul', 'underline'], ['st', 'strike']] as const) {
    if (raw[source] === undefined || raw[source] === null) continue
    const decoration = bodyRecord(raw[source]), enabled = bodyBoolean(decoration.s)
    if (enabled !== undefined) result[target] = enabled && decoration.t !== 11
  }
  for (const [source, target] of [['ff', 'fontFamily'], ['eastAsiaFontFamily', 'eastAsiaFontFamily']] as const) {
    const current = raw[source]
    if (current !== undefined && current !== null) {
      if (typeof current !== 'string' || !current) throw new TypeError('文档字体名无效。')
      result[target] = current
    }
  }
  const size = bodyNumber(raw.fs)
  if (size !== undefined) { if (size <= 0) throw new RangeError('文档字号必须为正。'); result.fontSize = size }
  const color = bodyColor(raw.cl), highlight = bodyColor(raw.bg)
  if (color !== undefined) result.color = color
  if (highlight !== undefined) result.highlight = highlight
  if (raw.va !== undefined && raw.va !== null) {
    const align = bodyNumber(raw.va)!
    if (![1, 2, 3].includes(align)) throw new RangeError('文字上下标无效。')
    result.verticalAlign = align === 2 ? 'subscript' : align === 3 ? 'superscript' : 'baseline'
  }
  result.metadata = { univerTextStyle: cloneModel(raw) }
  return result
}

export function documentRunStyleToBody(run: InlineRun): UniverBodyRecord {
  const raw = cloneModel(bodyRecord(run.metadata?.univerTextStyle))
  return { ...raw, ff: run.fontFamily, ...(run.eastAsiaFontFamily ? { eastAsiaFontFamily: run.eastAsiaFontFamily } : {}),
    fs: run.fontSize, bl: Number(run.bold), it: Number(run.italic),
    ul: { ...bodyRecord(raw.ul), s: Number(run.underline), t: run.underline && bodyRecord(raw.ul).t !== 11 ? bodyRecord(raw.ul).t ?? 12 : 11 },
    st: { ...bodyRecord(raw.st), s: Number(run.strike) }, cl: run.color === 'auto' ? { auto: 1 } : { rgb: run.color },
    ...(run.highlight === undefined ? { bg: undefined } : { bg: { rgb: run.highlight } }),
    va: run.verticalAlign === 'subscript' ? 2 : run.verticalAlign === 'superscript' ? 3 : 1 }
}

export function documentParagraphStyleFromBody(value: unknown, defaults: DocumentDefaultStyle): DocumentParagraphStyle {
  const raw = bodyRecord(value), style = { ...defaults.paragraph }
  const align = bodyNumber(raw.horizontalAlign)
  if (align !== undefined && align !== 0) {
    if (![1, 2, 3, 4, 5, 6].includes(align)) throw new RangeError('段落对齐方式无效。')
    style.align = align === 2 ? 'center' : align === 3 ? 'right' : align >= 4 ? 'justify' : 'left'
  }
  for (const [source, target] of [['indentStart', 'indentStart'], ['indentEnd', 'indentEnd'], ['indentFirstLine', 'indentFirstLine'],
    ['hanging', 'hanging'], ['spaceAbove', 'spaceBefore'], ['spaceBelow', 'spaceAfter']] as const) {
    const current = bodyLength(raw[source], defaults.text.fontSize)
    if (current !== undefined) style[target] = current
  }
  const spacing = bodyNumber(raw.lineSpacing)
  const rule = bodyNumber(raw.spacingRule)
  if (rule !== undefined && ![0, 1, 2].includes(rule)) throw new RangeError('段落行距规则无效。')
  if (rule !== undefined) style.spacingRule = rule === 1 ? 'atLeast' : rule === 2 ? 'exact' : 'auto'
  if (spacing !== undefined) {
    if (spacing <= 0) throw new RangeError('段落行距必须为正。')
    style.lineHeight = !style.spacingRule || style.spacingRule === 'auto' ? spacing : spacing * 4 / 3
  }
  for (const key of ['wordWrap', 'snapToGrid'] as const) { const current = bodyBoolean(raw[key]); if (current !== undefined) style[key] = current }
  return style
}

export function documentParagraphStyleToBody(style: DocumentParagraphStyle, source?: unknown): UniverBodyRecord {
  const raw = cloneModel(bodyRecord(source)), px = (v: number) => ({ v, u: 3 })
  return { ...raw, horizontalAlign: style.align === 'center' ? 2 : style.align === 'right' ? 3 : style.align === 'justify' ? 4 : 1,
    indentStart: px(style.indentStart ?? style.indent), indentEnd: px(style.indentEnd ?? 0),
    indentFirstLine: px(style.indentFirstLine ?? 0), hanging: px(style.hanging ?? 0), spaceAbove: px(style.spaceBefore), spaceBelow: px(style.spaceAfter),
    spacingRule: style.spacingRule === 'atLeast' ? 1 : style.spacingRule === 'exact' ? 2 : 0,
    lineSpacing: style.spacingRule && style.spacingRule !== 'auto' ? style.lineHeight * 3 / 4 : style.lineHeight,
    ...(style.wordWrap === undefined ? {} : { wordWrap: Number(style.wordWrap) }), ...(style.snapToGrid === undefined ? {} : { snapToGrid: Number(style.snapToGrid) }) }
}
