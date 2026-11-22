import type { SheetConditionalFormatOfficeBaseline, SheetConditionalFormatOfficeSource, SheetConditionalFormatRule } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'

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

/** Source metadata changes cannot masquerade as edits to Office rule fields. */
export function conditionalFormatOfficeBaseline(rule: SheetConditionalFormatOfficeBaseline): SheetConditionalFormatOfficeBaseline {
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
  return cloneModel(baseline) as unknown as SheetConditionalFormatOfficeBaseline
}

function validRanges(value: unknown): boolean {
  return Array.isArray(value) && value.length > 0 && value.every(item => {
    const range = object(item)
    const [sr, sc, er, ec] = [range.startRow, range.startCol, range.endRow, range.endCol]
    return [sr, sc, er, ec].every(item => typeof item === 'number' && Number.isSafeInteger(item))
      && (sr as number) >= 0 && (sc as number) >= 0 && (er as number) >= (sr as number) && (ec as number) >= (sc as number)
      && (er as number) <= 1_048_575 && (ec as number) <= 16_383
  })
}

/** XML parsing stays in the Rust import/export boundary; this validates the model envelope. */
export function normalizeConditionalFormatOfficeSource(value: unknown): SheetConditionalFormatOfficeSource | undefined {
  if (value === undefined) return undefined
  const source = object(value), baseline = object(source.baseline)
  if (source.version !== 1 || source.format !== 'xlsx' || !['base', 'x14'].includes(String(source.storage))
    || typeof source.originalId !== 'string' || !source.originalId.trim() || typeof source.ruleXml !== 'string' || !source.ruleXml.trim()
    || !conditionalFormatRuleTypes.includes(baseline.type as SheetConditionalFormatRule['type']) || !validRanges(baseline.ranges)
    || !baseline.style || typeof baseline.style !== 'object' || Array.isArray(baseline.style)) throw new TypeError('条件格式的 Office 来源记录无效。')
  for (const field of ['ruleXml', 'groupXml', 'differentialFormatXml', 'x14Xml', 'themeXml']) {
    if (source[field] !== undefined && (typeof source[field] !== 'string' || (source[field] as string).length > 16 * 1024 * 1024)) throw new TypeError('条件格式的 Office 来源片段无效。')
  }
  const anchor = object(baseline.formulaAnchor)
  if (baseline.formulaAnchor !== undefined && (![anchor.row, anchor.col].every(item => typeof item === 'number' && Number.isSafeInteger(item))
    || (anchor.row as number) < 0 || (anchor.row as number) > 1_048_575 || (anchor.col as number) < 0 || (anchor.col as number) > 16_383)) throw new TypeError('条件格式的 Office 来源公式基准无效。')
  for (const field of ['operator', 'formula1', 'formula2']) if (baseline[field] !== undefined && typeof baseline[field] !== 'string') throw new TypeError('条件格式的 Office 来源条件无效。')
  if (baseline.priority !== undefined && (typeof baseline.priority !== 'number' || !Number.isSafeInteger(baseline.priority) || baseline.priority < 1)) throw new TypeError('条件格式的 Office 来源优先级无效。')
  for (const field of ['stopIfTrue', 'isBottom', 'isPercent']) if (baseline[field] !== undefined && typeof baseline[field] !== 'boolean') throw new TypeError('条件格式的 Office 来源开关无效。')
  return cloneModel({ ...source, baseline: conditionalFormatOfficeBaseline(baseline as unknown as SheetConditionalFormatOfficeBaseline) }) as unknown as SheetConditionalFormatOfficeSource
}
