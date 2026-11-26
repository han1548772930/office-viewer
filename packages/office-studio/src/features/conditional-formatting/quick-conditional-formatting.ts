import type { CellRange, SheetConditionalFormatRule, SheetConditionalFormatType, SheetConditionalFormatVisualConfig } from '../../domain/model.ts'
import { sheetRangeContainsCell, sheetRangesIntersect, subtractSheetCommandRange, type SheetRangeCommandInput } from '../sheet-workbench/sheet-range-command.ts'
import { conditionalFormatRuleTypes } from './conditional-format-office-source.ts'
import { reanchorConditionalFormatRule } from './conditional-format-formulas.ts'
import { prependConditionalFormatRuleDraft } from './conditional-format-rule-editor.ts'
import { conditionalIconPresets } from './conditional-icon-presets.ts'
import { createConditionalIconConfig, EMPTY_CONDITIONAL_ICON_TYPE, type ConditionalIconConfig } from './icon-set-config.ts'

export const QUICK_CONDITIONAL_FORMAT_COMMANDS = {
  dataBar: 'sheet.command.add-data-bar-conditional-rule',
  iconSet: 'sheet.command.add-icon-set-conditional-rule',
  clear: 'sheet.command.clear-range-conditional-rule',
} as const
export type QuickConditionalFormatCommandId = typeof QUICK_CONDITIONAL_FORMAT_COMMANDS[keyof typeof QUICK_CONDITIONAL_FORMAT_COMMANDS]
export type QuickConditionalFormatType = 'dataBar' | 'iconSet'
export const QUICK_CONDITIONAL_ICON_SETS = ['3Stars', '3TrafficLights1', '3Arrows'] as const
type QuickConditionalIconSet = typeof QUICK_CONDITIONAL_ICON_SETS[number]
export interface QuickConditionalFormatSelection {
  hasDataBar: boolean
  hasIconSet: boolean
  activeIconSet?: QuickConditionalIconSet
  canClearDataBar: boolean
  canClearIconSet: boolean
}

export const QUICK_CONDITIONAL_DATA_BAR = {
  min: { type: 'min' }, max: { type: 'max' },
  positiveColor: '#ffbe38', nativeColor: '#abd91a', isGradient: false, isShowValue: true,
} as const

export interface QuickConditionalFormatInput extends SheetRangeCommandInput,
  Pick<SheetConditionalFormatVisualConfig, 'min' | 'max' | 'positiveColor' | 'nativeColor' | 'isGradient' | 'isShowValue' | 'iconSet'> {
  config?: ConditionalIconConfig[]
  types?: readonly SheetConditionalFormatType[]
  stopIfTrue?: boolean
  expectedRevision?: number
  host?: 'workbench'
}

export function quickConditionalFormatRules(rules: readonly SheetConditionalFormatRule[], ranges: readonly CellRange[],
  id: QuickConditionalFormatCommandId, input: QuickConditionalFormatInput): readonly SheetConditionalFormatRule[] {
  if (!ranges.length) return rules
  if (id === QUICK_CONDITIONAL_FORMAT_COMMANDS.clear) {
    const types = input.types
    if (types !== undefined && (!Array.isArray(types) || types.some(type => !conditionalFormatRuleTypes.includes(type)))) {
      throw new TypeError('条件格式的清除类型无效。')
    }
    let work = 0
    return rules.flatMap(rule => {
      if (types?.length && !types.includes(rule.type)) return [rule]
      let remaining = rule.ranges, changed = false
      for (const cut of ranges) {
        work += remaining.length
        if (work > 5_000_000) throw new RangeError('条件格式清除超出当前范围计算预算。')
        remaining = remaining.flatMap(range => {
          if (!sheetRangesIntersect(range, cut)) return [range]
          changed = true
          return subtractSheetCommandRange(range, cut)
        })
        if (remaining.length > 100_000) throw new RangeError('条件格式清除后的范围数量过多。')
        if (!remaining.length) break
      }
      return remaining.length ? [changed ? reanchorConditionalFormatRule(rule, remaining) : rule] : []
    })
  }

  const type = id === QUICK_CONDITIONAL_FORMAT_COMMANDS.dataBar ? 'dataBar' : 'iconSet'
  const config: SheetConditionalFormatVisualConfig = type === 'dataBar'
    ? { min: input.min ?? QUICK_CONDITIONAL_DATA_BAR.min, max: input.max ?? QUICK_CONDITIONAL_DATA_BAR.max,
      positiveColor: input.positiveColor ?? QUICK_CONDITIONAL_DATA_BAR.positiveColor,
      nativeColor: input.nativeColor ?? QUICK_CONDITIONAL_DATA_BAR.nativeColor,
      isGradient: input.isGradient ?? QUICK_CONDITIONAL_DATA_BAR.isGradient, isShowValue: input.isShowValue ?? QUICK_CONDITIONAL_DATA_BAR.isShowValue }
    : quickIconSetConfig(input)
  return prependConditionalFormatRuleDraft(rules, {
    id: crypto.randomUUID(), type, ranges: [...ranges], operator: '', formula1: '', style: {}, priority: 1,
    stopIfTrue: input.stopIfTrue ?? false, config,
  })
}

function quickIconSetConfig(input: QuickConditionalFormatInput): SheetConditionalFormatVisualConfig {
  if (input.config !== undefined && !Array.isArray(input.config)) throw new TypeError('条件格式图标配置无效。')
  const icons = input.config?.map((icon, index, items) => {
    // Univer's last icon is an unconditional fallback. Its sentinel has no percent-domain meaning.
    return index === items.length - 1 && icon?.value?.type === 'percent' && icon.value.value === Number.MAX_SAFE_INTEGER
      ? { ...icon, value: { type: 'num', value: Number.MAX_SAFE_INTEGER } } : icon
  })
  const iconSet = input.iconSet ?? icons?.find(icon => icon?.iconType && icon.iconType !== EMPTY_CONDITIONAL_ICON_TYPE)?.iconType ?? '3Stars'
  if (!Object.hasOwn(conditionalIconPresets, iconSet)) throw new TypeError('请选择有效的单元格图标集。')
  if (icons) return { iconSet, icons, isShowValue: input.isShowValue ?? true }
  if (!QUICK_CONDITIONAL_ICON_SETS.some(preset => preset === iconSet)) throw new TypeError('请选择有效的快捷图标集。')
  const config = createConditionalIconConfig(iconSet)
  config[0] = { ...config[0]!, operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 } }
  config[1] = { ...config[1]!, operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 } }
  return { iconSet, icons: config, isShowValue: input.isShowValue ?? true }
}

export function quickConditionalFormatSelection(rules: readonly SheetConditionalFormatRule[], ranges: readonly CellRange[],
  primary?: { row: number; col: number }, canEditRule?: (ranges: readonly CellRange[]) => boolean): QuickConditionalFormatSelection {
  const relevant = rules.filter(rule => (rule.type === 'dataBar' || rule.type === 'iconSet')
    && rule.ranges.some(scope => ranges.some(range => sheetRangesIntersect(scope, range))))
  const active = relevant.filter(rule => primary && rule.ranges.some(scope => sheetRangeContainsCell(scope, primary.row, primary.col)))
    .sort((left, right) => left.priority - right.priority)
  const activeIconRule = active.find(rule => rule.type === 'iconSet')
  const config = activeIconRule?.config
  const icons = config?.icons
  const activeIconSet = !config?.reverse && icons?.length === 3
    ? QUICK_CONDITIONAL_ICON_SETS.find(preset => icons.every((icon, index) => icon.iconType === preset && icon.iconId === String(index))) : undefined
  const canClear = (type: QuickConditionalFormatType) => {
    const matched = relevant.filter(rule => rule.type === type)
    return matched.length > 0 && (!canEditRule || matched.every(rule => canEditRule(rule.ranges)))
  }
  return {
    hasDataBar: active.some(rule => rule.type === 'dataBar'), hasIconSet: activeIconRule !== undefined, activeIconSet,
    canClearDataBar: canClear('dataBar'), canClearIconSet: canClear('iconSet'),
  }
}
