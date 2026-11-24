import type { ChartColor, ChartLabelPattern, ChartTreemapLabelOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { cloneChartOptions } from './chart-options'
import { chartLabelContentFromValues, chartLabelContentMask, chartLabelContentValues } from './chart-labels'
import { CHART_FILL_PATTERN_PRESETS } from './chart-fill-pattern'
import { chartTreemapLabelBackground, chartTreemapParentLayout, resolveChartTreemapLabel } from './chart-treemap'
import { chartFontControls, colorControl, multiselectControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'

const LABEL_KEY = 'treemap.label'
function label(config: ChartConfig): ChartTreemapLabelOptions { return optionGroup(config, 'treemap').label ??= {} }
function leafVisible(config: ChartConfig): boolean { return resolveChartTreemapLabel(config).visible }

export function chartTreemapParentControls(): ChartSettingControl[] {
  return [{ ...selectControl('treemap.parentLabelLayout', '父级标签布局', [['banner', '独立标题带'], ['overlapping', '叠放'], ['none', '隐藏']],
    config => chartTreemapParentLayout(config.options?.treemap), (config, value) => { optionGroup(config, 'treemap').parentLabelLayout = value }, 'banner'),
    canReset: config => config.options?.treemap?.parentLabelLayout !== undefined,
    reset: config => { if (config.options?.treemap) config.options.treemap.parentLabelLayout = undefined } }]
}

/** Fixed to exposes only CategoryName and Value. Other imported flags stay
 * authored so editing these active flags does not lose a future type's data. */
export function chartTreemapLeafControls(): ChartSettingControl[] {
  const contents = multiselectControl(`${LABEL_KEY}.contentType`, '叶级标签内容', [['category', '分类名称'], ['value', '数值']],
    config => chartLabelContentValues(resolveChartTreemapLabel(config).contentType).filter(value => value === 'category' || value === 'value'),
    (config, values) => { label(config).contentType = chartLabelContentFromValues(values) | chartLabelContentMask(config.options?.treemap?.label?.contentType ?? 'category') & ~10 })
  const format = textControl(`${LABEL_KEY}.format`, '叶级标签数字格式', config => config.options?.treemap?.label?.format, (config, value) => { label(config).format = value })
  return [
    switchControl(`${LABEL_KEY}.visible`, '显示叶级标签', config => resolveChartTreemapLabel(config).visible, (config, value) => { label(config).visible = value }, true),
    { ...contents, visible: leafVisible },
    { ...format, visible: config => leafVisible(config) && Boolean(chartLabelContentMask(resolveChartTreemapLabel(config).contentType) & 8) },
    ...chartFontControls(LABEL_KEY, config => resolveChartTreemapLabel(config), label).map(control => ({ ...control, visible: leafVisible })),
  ]
}

export type ChartTreemapBackgroundMode = 'color' | 'pattern'
export function chartTreemapBackgroundMode(config: ChartConfig): ChartTreemapBackgroundMode {
  const value = config.options?.treemap?.label?.backgroundColor
  return value && typeof value === 'object' && value.type === 'pattern' ? 'pattern' : 'color'
}
export function chartTreemapPattern(config: ChartConfig): ChartLabelPattern | undefined {
  const value = config.options?.treemap?.label?.backgroundColor
  return value && typeof value === 'object' && value.type === 'pattern' ? value : undefined
}
export function setChartTreemapBackground(config: ChartConfig, value: ChartTreemapLabelOptions['backgroundColor']): void {
  label(config).backgroundColor = value === undefined ? undefined : cloneChartOptions({ treemap: { label: { backgroundColor: value } } })!.treemap!.label!.backgroundColor
}
export function setChartTreemapBackgroundMode(config: ChartConfig, mode: ChartTreemapBackgroundMode): void {
  if (mode === chartTreemapBackgroundMode(config)) return
  const value = chartTreemapLabelBackground(config.options?.treemap?.label?.backgroundColor)
  setChartTreemapBackground(config, mode === 'pattern' ? { type: 'pattern', preset: 'pct50', foregroundColor: typeof value === 'string' ? value : '#000000', backgroundColor: '#ffffff' } : value)
}
export function chartTreemapColorBackground(config: ChartConfig): ChartColor | undefined {
  return chartTreemapBackgroundMode(config) === 'color' ? config.options?.treemap?.label?.backgroundColor as ChartColor | undefined : undefined
}
function pattern(config: ChartConfig): ChartLabelPattern {
  const current = chartTreemapPattern(config)
  if (current) return current
  label(config).backgroundColor = { type: 'pattern', preset: 'pct50', foregroundColor: '#000000', backgroundColor: '#ffffff' }
  return label(config).backgroundColor as ChartLabelPattern
}
export function chartTreemapPatternControls(config: ChartConfig): ChartSettingControl[] {
  const current = chartTreemapPattern(config)?.preset, choices: Array<readonly [string, string]> = [...CHART_FILL_PATTERN_PRESETS]
  if (current && !choices.some(([value]) => value === current)) choices.push([current, current])
  return [
    selectControl(`${LABEL_KEY}.backgroundColor.preset`, '标签背景图案', choices, item => chartTreemapPattern(item)?.preset, (item, value) => { pattern(item).preset = value ?? 'pct50' }, 'pct50'),
    colorControl(`${LABEL_KEY}.backgroundColor.foregroundColor`, '标签图案前景色', item => chartTreemapPattern(item)?.foregroundColor, (item, value) => { pattern(item).foregroundColor = value ?? '#000000' }),
    colorControl(`${LABEL_KEY}.backgroundColor.backgroundColor`, '标签图案背景色', item => chartTreemapPattern(item)?.backgroundColor, (item, value) => { pattern(item).backgroundColor = value ?? '#ffffff' }),
  ]
}
export function clearChartTreemapSettings(config: ChartConfig, target: 'label' | 'all'): void {
  if (!config.options?.treemap) return
  if (target === 'all') config.options.treemap = undefined
  else config.options.treemap.label = undefined
}
