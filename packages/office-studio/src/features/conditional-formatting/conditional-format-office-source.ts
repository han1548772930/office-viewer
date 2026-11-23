import type { SheetConditionalFormatOfficeBaseline, SheetConditionalFormatOfficeSource, SheetConditionalFormatRule } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { conditionalFormatModelByteLimit, conditionalFormatOfficeFragmentByteLimit, conditionalFormatOfficeSourceXmlByteLimit,
  conditionalFormatSerializedSize, conditionalFormatUtf8Size } from './conditional-format-source-budget.ts'

export const conditionalFormatRuleTypes: readonly SheetConditionalFormatRule['type'][] = ['cellIs', 'highlightCell', 'top10', 'average',
  'duplicateValues', 'uniqueValues', 'timePeriod', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'expression', 'dataBar', 'colorScale', 'iconSet', 'office']
const ruleFields = ['type', 'subType', 'ranges', 'formulaAnchor', 'operator', 'formula1', 'formula2', 'value', 'isBottom', 'isPercent', 'config', 'style', 'priority', 'stopIfTrue']
const styleFields = ['fill', 'fontColor', 'bold', 'italic', 'underline', 'strike']
const configFields = ['min', 'max', 'isGradient', 'positiveColor', 'nativeColor', 'thresholds', 'icons', 'isShowValue', 'iconSet', 'reverse',
  'minLength', 'maxLength', 'border', 'borderColor', 'negativeBorderColor', 'axisColor', 'negativeBarColorSameAsPositive', 'negativeBarBorderColorSameAsPositive', 'direction', 'axisPosition']

function object(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}
function selected(value: unknown, fields: readonly string[]): Record<string, unknown> {
  return Object.fromEntries(Object.entries(object(value)).filter(([key, item]) => fields.includes(key) && item !== undefined && item !== null))
}

function baselineProjection(rule: SheetConditionalFormatOfficeBaseline): Record<string, unknown> {
  const baseline = selected(rule, ruleFields)
  baseline.operator ??= ''
  baseline.formula1 ??= ''
  baseline.priority ??= 1
  baseline.stopIfTrue ??= false
  if (baseline.isBottom === false) delete baseline.isBottom
  if (baseline.isPercent === false) delete baseline.isPercent
  baseline.style = selected(rule.style, styleFields)
  const config = selected(rule.config, configFields)
  for (const field of ['min', 'max']) if (config[field]) config[field] = selected(config[field], ['type', 'value'])
  if (Array.isArray(config.thresholds)) config.thresholds = config.thresholds.map(item => ({
    ...selected(item, ['color']), value: selected(object(item).value, ['type', 'value']),
  }))
  if (Array.isArray(config.icons)) config.icons = config.icons.map(item => {
    const entry = selected(item, ['operator', 'iconType', 'iconId'])
    if (typeof entry.iconId === 'number') entry.iconId = String(entry.iconId)
    return { ...entry, value: selected(object(item).value, ['type', 'value']) }
  })
  if (Object.keys(config).length) baseline.config = config
  else delete baseline.config
  return baseline
}

/** Source metadata changes cannot masquerade as edits to Office rule fields. */
export function conditionalFormatOfficeBaseline(rule: SheetConditionalFormatOfficeBaseline): SheetConditionalFormatOfficeBaseline {
  return cloneModel(baselineProjection(rule)) as unknown as SheetConditionalFormatOfficeBaseline
}

function validRanges(value: unknown): boolean {
  return Array.isArray(value) && value.length > 0 && value.length <= 100_000 && value.every(item => {
    const range = object(item)
    const [sr, sc, er, ec] = [range.startRow, range.startCol, range.endRow, range.endCol]
    return [sr, sc, er, ec].every(item => typeof item === 'number' && Number.isSafeInteger(item))
      && (sr as number) >= 0 && (sc as number) >= 0 && (er as number) >= (sr as number) && (ec as number) >= (sc as number)
      && (er as number) <= 1_048_575 && (ec as number) <= 16_383
  })
}

function validColor(value: unknown): boolean { return typeof value === 'string' && /^#(?:[0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value) }
function validScalar(value: unknown): boolean { return typeof value === 'string' || typeof value === 'number' && Number.isFinite(value) }
function validThresholdShape(value: unknown): boolean {
  const threshold = object(value)
  return typeof threshold.type === 'string' && (threshold.value === undefined || validScalar(threshold.value))
}

/** Shared primitive checks keep both commands and Office baselines inspectable. */
export function assertConditionalFormatKnownFields(value: unknown): void {
  const rule = object(value), style = object(rule.style), config = object(rule.config)
  if (!conditionalFormatRuleTypes.includes(rule.type as SheetConditionalFormatRule['type']) || !validRanges(rule.ranges)
    || !rule.style || typeof rule.style !== 'object' || Array.isArray(rule.style)) throw new TypeError('条件格式的已知字段无效。')
  if (rule.subType !== undefined && (typeof rule.subType !== 'string' || !['number', 'text', 'formula', 'timePeriod', 'rank', 'average', 'duplicateValues', 'uniqueValues'].includes(rule.subType))) throw new TypeError('条件格式的条件类型无效。')
  for (const field of ['operator', 'formula1', 'formula2']) if (rule[field] !== undefined && typeof rule[field] !== 'string') throw new TypeError('条件格式的条件字段无效。')
  if (rule.priority !== undefined && (typeof rule.priority !== 'number' || !Number.isSafeInteger(rule.priority) || rule.priority < 1 || rule.priority > 2_147_483_647)) throw new TypeError('条件格式的优先级无效。')
  const anchor = object(rule.formulaAnchor)
  if (rule.formulaAnchor !== undefined && (![anchor.row, anchor.col].every(item => typeof item === 'number' && Number.isSafeInteger(item))
    || (anchor.row as number) < 0 || (anchor.row as number) > 1_048_575 || (anchor.col as number) < 0 || (anchor.col as number) > 16_383)) throw new TypeError('条件格式的公式基准无效。')
  for (const field of ['stopIfTrue', 'isBottom', 'isPercent']) if (rule[field] !== undefined && typeof rule[field] !== 'boolean') throw new TypeError('条件格式的开关无效。')
  if (rule.value !== undefined && !validScalar(rule.value) && !(Array.isArray(rule.value) && rule.value.length === 2 && rule.value.every(value => typeof value === 'number' && Number.isFinite(value)))) throw new TypeError('条件格式的匹配值无效。')
  for (const field of ['fill', 'fontColor']) if (style[field] !== undefined && !validColor(style[field])) throw new TypeError('条件格式的样式颜色无效。')
  for (const field of ['bold', 'italic', 'underline', 'strike']) if (style[field] !== undefined && typeof style[field] !== 'boolean') throw new TypeError('条件格式的字体开关无效。')
  if (rule.config === undefined) return
  if (!rule.config || typeof rule.config !== 'object' || Array.isArray(rule.config)) throw new TypeError('条件格式的配置无效。')
  for (const field of ['isGradient', 'isShowValue', 'reverse', 'border', 'negativeBarColorSameAsPositive', 'negativeBarBorderColorSameAsPositive']) {
    if (config[field] !== undefined && typeof config[field] !== 'boolean') throw new TypeError('条件格式的显示开关无效。')
  }
  for (const field of ['positiveColor', 'nativeColor', 'borderColor', 'negativeBorderColor', 'axisColor']) if (config[field] !== undefined && !validColor(config[field])) throw new TypeError('条件格式的数据条颜色无效。')
  for (const field of ['minLength', 'maxLength']) if (config[field] !== undefined && (typeof config[field] !== 'number' || !Number.isSafeInteger(config[field]) || (config[field] as number) < 0 || (config[field] as number) > 100)) throw new TypeError('条件格式的数据条长度无效。')
  const minimum = typeof config.minLength === 'number' ? config.minLength : 0
  const maximum = typeof config.maxLength === 'number' ? config.maxLength : 100
  if (minimum > maximum) throw new TypeError('条件格式的数据条最小长度不能大于最大长度。')
  if (config.direction !== undefined && (typeof config.direction !== 'string' || !['context', 'leftToRight', 'rightToLeft'].includes(config.direction))
    || config.axisPosition !== undefined && (typeof config.axisPosition !== 'string' || !['automatic', 'middle', 'none'].includes(config.axisPosition))
    || config.iconSet !== undefined && typeof config.iconSet !== 'string') throw new TypeError('条件格式的方向、坐标轴或图标集无效。')
  for (const field of ['min', 'max']) if (config[field] !== undefined && !validThresholdShape(config[field])) throw new TypeError('条件格式的数据条阈值无效。')
  if (config.thresholds !== undefined && (!Array.isArray(config.thresholds) || config.thresholds.length > 100_000 || config.thresholds.some(item => !validThresholdShape(object(item).value) || !validColor(object(item).color)))) throw new TypeError('条件格式的色阶停止点无效。')
  if (config.icons !== undefined && (!Array.isArray(config.icons) || config.icons.length > 100_000 || config.icons.some(item => {
    const entry = object(item)
    return !validThresholdShape(entry.value) || ['operator', 'iconType'].some(field => entry[field] !== undefined && typeof entry[field] !== 'string')
      || entry.iconId !== undefined && typeof entry.iconId !== 'string' && !(typeof entry.iconId === 'number' && Number.isSafeInteger(entry.iconId) && entry.iconId >= 0)
  }))) throw new TypeError('条件格式的图标配置无效。')
}

/** XML parsing stays in the Rust import/export boundary; this validates the model envelope. */
export function normalizeConditionalFormatOfficeSource(value: unknown): SheetConditionalFormatOfficeSource | undefined {
  if (value === undefined) return undefined
  const source = object(value), baseline = object(source.baseline)
  if (source.version !== 1 || source.format !== 'xlsx' || !['base', 'x14'].includes(String(source.storage))
    || typeof source.originalId !== 'string' || !source.originalId.trim() || conditionalFormatUtf8Size(source.originalId, 4096) > 4096 || typeof source.ruleXml !== 'string' || !source.ruleXml.trim()
    || !conditionalFormatRuleTypes.includes(baseline.type as SheetConditionalFormatRule['type']) || !validRanges(baseline.ranges)
    || !baseline.style || typeof baseline.style !== 'object' || Array.isArray(baseline.style)) throw new TypeError('条件格式的 Office 来源记录无效。')
  let xmlBytes = 0
  for (const field of ['ruleXml', 'groupXml', 'differentialFormatXml', 'x14Xml', 'themeXml']) {
    if (source[field] === undefined) continue
    if (typeof source[field] !== 'string') throw new TypeError('条件格式的 Office 来源片段无效。')
    const bytes = conditionalFormatUtf8Size(source[field] as string, conditionalFormatOfficeFragmentByteLimit)
    if (bytes > conditionalFormatOfficeFragmentByteLimit) throw new RangeError('CF_SOURCE_FRAGMENT_LIMIT: 条件格式的 Office 来源片段超过字节预算。')
    xmlBytes += bytes
  }
  if (xmlBytes > conditionalFormatOfficeSourceXmlByteLimit) throw new RangeError('CF_SOURCE_TOTAL_LIMIT: 条件格式的 Office 来源片段超过累计字节预算。')
  const anchor = object(baseline.formulaAnchor)
  if (baseline.formulaAnchor !== undefined && (![anchor.row, anchor.col].every(item => typeof item === 'number' && Number.isSafeInteger(item))
    || (anchor.row as number) < 0 || (anchor.row as number) > 1_048_575 || (anchor.col as number) < 0 || (anchor.col as number) > 16_383)) throw new TypeError('条件格式的 Office 来源公式基准无效。')
  for (const field of ['operator', 'formula1', 'formula2']) if (baseline[field] !== undefined && typeof baseline[field] !== 'string') throw new TypeError('条件格式的 Office 来源条件无效。')
  if (baseline.priority !== undefined && (typeof baseline.priority !== 'number' || !Number.isSafeInteger(baseline.priority) || baseline.priority < 1)) throw new TypeError('条件格式的 Office 来源优先级无效。')
  for (const field of ['stopIfTrue', 'isBottom', 'isPercent']) if (baseline[field] !== undefined && typeof baseline[field] !== 'boolean') throw new TypeError('条件格式的 Office 来源开关无效。')
  assertConditionalFormatKnownFields(baseline)
  const normalized = { ...source, baseline: baselineProjection(baseline as unknown as SheetConditionalFormatOfficeBaseline) }
  conditionalFormatSerializedSize(normalized, conditionalFormatModelByteLimit)
  return cloneModel(normalized) as unknown as SheetConditionalFormatOfficeSource
}
