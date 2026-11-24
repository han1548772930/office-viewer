import type { ChartLabelOptions } from '../../domain/model'

/** Fixed LabelContentType flags, in the fixed formatter's visible order. */
export const CHART_LABEL_CONTENT_CHOICES = [
  ['series', '系列', 4], ['category', '分类', 2], ['value', '数值', 8], ['percentage', '百分比', 16],
  ['bubbleSize', '气泡大小', 32],
] as const
export type ChartLabelContent = typeof CHART_LABEL_CONTENT_CHOICES[number][0]
export const CHART_LABEL_POSITION_CHOICES: ReadonlyArray<readonly [NonNullable<ChartLabelOptions['position']>, string]> = [
  ['auto', '自动'], ['inside', '内部'], ['insideBase', '内部基线'], ['insideEnd', '内部末端'],
  ['outside', '外部'], ['top', '上方'], ['bottom', '下方'], ['left', '左侧'], ['right', '右侧'], ['center', '中心'],
]

/** Omission keeps the historical Value default. Zero is an authored empty
 * label; unsupported flag bits are rejected instead of being silently lost. */
export function chartLabelContentMask(content: unknown, allowBubbleSize = true, name = '标签内容'): number {
  if (content === undefined) return 8
  if (typeof content === 'string') {
    const entry = CHART_LABEL_CONTENT_CHOICES.find(([key]) => key === content)
    if (entry && (allowBubbleSize || entry[2] !== 32) && entry[0] !== 'bubbleSize') return entry[2]
  } else if (typeof content === 'number' && Number.isSafeInteger(content) && content >= 0 && content <= 62) {
    const supported = allowBubbleSize ? 62 : 30
    if ((content & ~supported) === 0) return content
  }
  throw new TypeError(name + '包含无效的内容或组合标志。')
}

export function chartLabelContentValues(content: ChartLabelOptions['contentType'], allowBubbleSize = true): ChartLabelContent[] {
  const mask = chartLabelContentMask(content, allowBubbleSize)
  return CHART_LABEL_CONTENT_CHOICES.filter(([_key, _label, flag]) => (mask & flag) !== 0).map(([key]) => key)
}
export function chartLabelContentFromValues(values: readonly string[], allowBubbleSize = true): number {
  if (!Array.isArray(values)) throw new TypeError('标签内容须为选项数组。')
  let mask = 0
  const selected = new Set<string>()
  for (const value of values) {
    const entry = CHART_LABEL_CONTENT_CHOICES.find(([key]) => key === value)
    if (!entry || !allowBubbleSize && entry[2] === 32 || selected.has(value)) throw new TypeError('标签内容包含无效或重复选项。')
    selected.add(value); mask |= entry[2]
  }
  return mask
}

export interface ChartPointLabelLocation { x: number; y: number; baseX?: number; baseY?: number; value?: number }
export interface ChartPointLabelAnchor { x: number; y: number; align: 'left' | 'center' | 'right'; baseline: 'middle' }
/** Baseline/endpoint placement follows the actual projected direction, so
 * negative marks, stacked baselines and reversed value axes share one rule. */
export function chartPointLabelAnchor(position: ChartLabelOptions['position'], point: ChartPointLabelLocation,
  horizontal = false, fontSize = 11): ChartPointLabelAnchor {
  const gap = Math.max(9, fontSize * 0.8), anchor: ChartPointLabelAnchor = { x: point.x, y: point.y, align: 'center', baseline: 'middle' }
  const start = horizontal ? point.baseX : point.baseY, end = horizontal ? point.x : point.y
  const distance = start === undefined ? 0 : Math.abs(end - start)
  const direction = start !== undefined && end !== start ? Math.sign(end - start) : horizontal ? (point.value ?? 0) < 0 ? -1 : 1 : (point.value ?? 0) < 0 ? 1 : -1
  const current = position === 'auto' ? start === undefined ? 'top' : 'outside' : position ?? 'top'
  if (current === 'inside' || current === 'center') {
    if (horizontal) anchor.x = start === undefined ? point.x : start + (end - start) / 2
    else anchor.y = start === undefined ? point.y : start + (end - start) / 2
  } else if (current === 'insideBase' || current === 'insideEnd') {
    const offset = Math.min(gap, distance / 2)
    const coordinate = start === undefined ? end : current === 'insideBase' ? start + direction * offset : end - direction * offset
    if (horizontal) { anchor.x = coordinate; anchor.align = distance <= gap * 2 ? 'center' : (current === 'insideBase' ? direction : -direction) > 0 ? 'left' : 'right' }
    else anchor.y = coordinate
  } else if (current === 'outside') {
    if (horizontal) { anchor.x += direction * gap; anchor.align = direction > 0 ? 'left' : 'right' }
    else anchor.y += direction * gap
  } else if (current === 'left') { anchor.x -= gap; anchor.align = 'right' }
  else if (current === 'right') { anchor.x += gap; anchor.align = 'left' }
  else anchor.y += current === 'bottom' ? Math.max(10, gap) : -gap
  return anchor
}
