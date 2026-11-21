import type { SheetConditionalFormatVisualConfig, SheetConditionalFormatValueConfig } from '../../domain/model.ts'
import { conditionalIconPresets } from './conditional-icon-presets.ts'

export type ConditionalIconConfig = NonNullable<SheetConditionalFormatVisualConfig['icons']>[number]
export const EMPTY_CONDITIONAL_ICON_TYPE = 'EMPTY_ICON_TYPE'

export function createConditionalIconConfig(iconType: string): ConditionalIconConfig[] {
  const icons = conditionalIconPresets[iconType]
  if (!icons) throw new TypeError('请选择有效的图标集。')
  return icons.map((_icon, index) => ({
    operator: index === icons.length - 1 ? 'lessThanOrEqual' : 'greaterThanOrEqual',
    value: { type: 'percent', value: Math.round((icons.length - 1 - index) * 100 / icons.length) },
    iconType, iconId: String(index),
  }))
}

/** Mirror the complete selected sequence, including mixed sets and hidden icons. */
export function conditionalIconSequence(config: SheetConditionalFormatVisualConfig): ConditionalIconConfig[] {
  const icons = config.icons ?? []
  return icons.map((item, index) => {
    const source = config.reverse ? icons[icons.length - 1 - index]! : item
    return { ...item, value: { ...item.value }, iconType: source.iconType, iconId: source.iconId }
  })
}

export function validateConditionalThreshold(threshold: SheetConditionalFormatValueConfig): void {
  if (threshold.type === 'min' || threshold.type === 'max') return
  if (threshold.type === 'formula') {
    if (!String(threshold.value ?? '').trim().replace(/^=/, '').trim()) throw new TypeError('阈值公式不能为空。')
    return
  }
  if (!['num', 'percent', 'percentile'].includes(threshold.type)) throw new TypeError('阈值类型无效。')
  if (threshold.value === undefined || String(threshold.value).trim() === '' || !Number.isFinite(Number(threshold.value))) throw new TypeError('请输入有效的阈值。')
  if (['percent', 'percentile'].includes(threshold.type) && (Number(threshold.value) < 0 || Number(threshold.value) > 100)) throw new RangeError('百分比和百分位必须在 0 到 100 之间。')
}

export function validateConditionalIconConfig(config: SheetConditionalFormatVisualConfig): void {
  const icons = config.icons ?? []
  if (icons.length < 3 || icons.length > 5) throw new RangeError('图标集必须包含 3 到 5 个图标。')
  for (const [index, icon] of icons.entries()) {
    if (icon.iconType !== EMPTY_CONDITIONAL_ICON_TYPE) {
      const preset = icon.iconType ? conditionalIconPresets[icon.iconType] : undefined
      if (!preset || icon.iconId === undefined || !/^\d+$/.test(icon.iconId) || !preset[Number(icon.iconId)]) throw new TypeError('请选择有效的单元格图标。')
    }
    if (index === icons.length - 1) continue
    if (icon.operator !== 'greaterThan' && icon.operator !== 'greaterThanOrEqual') throw new TypeError('图标阈值比较条件必须为大于或大于或等于。')
    if (icon.value.type === 'min' || icon.value.type === 'max') throw new TypeError('图标阈值需要数字、百分比、百分位或公式。')
    validateConditionalThreshold(icon.value)
    const previous = icons[index - 1]
    if (!previous || previous.value.type !== icon.value.type || icon.value.type === 'formula') continue
    const previousValue = Number(previous.value.value), value = Number(icon.value.value)
    if (value > previousValue || value === previousValue && previous.operator === 'greaterThanOrEqual') throw new RangeError('图标阈值必须按从高到低的顺序排列。')
  }
}
