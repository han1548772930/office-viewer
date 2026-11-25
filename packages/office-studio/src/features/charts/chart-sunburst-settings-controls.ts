import type { ChartColor, ChartLabelPattern, ChartSunburstCenterLabelOptions, ChartSunburstLabelOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { cloneChartOptions } from './chart-options'
import { chartLabelContentFromValues, chartLabelContentMask, chartLabelContentValues } from './chart-labels'
import { CHART_FILL_PATTERN_PRESETS } from './chart-fill-pattern'
import { chartLabelBackground } from './chart-label-background'
import { resolveChartSunburstCenterLabel, resolveChartSunburstLabel } from './chart-sunburst'
import { chartFontControls, colorControl, multiselectControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'

const LABEL_KEY = 'sunburst.label', CENTER_KEY = 'sunburst.centerLabel'
function label(config: ChartConfig): ChartSunburstLabelOptions { return optionGroup(config, 'sunburst').label ??= {} }
function center(config: ChartConfig): ChartSunburstCenterLabelOptions { return optionGroup(config, 'sunburst').centerLabel ??= {} }
function nodeVisible(config: ChartConfig): boolean { return resolveChartSunburstLabel(config).visible }
function centerVisible(config: ChartConfig): boolean { return resolveChartSunburstCenterLabel(config).visible }

export function chartSunburstNodeLabelControls(): ChartSettingControl[] {
  const contents = multiselectControl(`${LABEL_KEY}.contentType`, '节点标签内容', [['category', '分类名称'], ['value', '数值']],
    config => chartLabelContentValues(resolveChartSunburstLabel(config).contentType).filter(value => value === 'category' || value === 'value'),
    (config, values) => { label(config).contentType = chartLabelContentFromValues(values) | chartLabelContentMask(config.options?.sunburst?.label?.contentType ?? 'category') & ~10 })
  const position = selectControl(`${LABEL_KEY}.position`, '节点标签位置', [['auto', '自动'], ['inside', '内部'], ['outside', '外部']],
    config => resolveChartSunburstLabel(config).position ?? 'auto', (config, value) => { label(config).position = value === 'inside' || value === 'outside' ? value : undefined }, 'auto')
  return [
    switchControl(`${LABEL_KEY}.visible`, '显示节点标签', config => resolveChartSunburstLabel(config).visible, (config, value) => { label(config).visible = value }, true),
    { ...contents, visible: nodeVisible }, { ...position, visible: nodeVisible },
    { ...textControl(`${LABEL_KEY}.format`, '节点标签数字格式', config => config.options?.sunburst?.label?.format, (config, value) => { label(config).format = value }),
      visible: config => nodeVisible(config) && Boolean(chartLabelContentMask(resolveChartSunburstLabel(config).contentType) & 8) },
    ...chartFontControls(LABEL_KEY, config => resolveChartSunburstLabel(config), label).map(control => ({ ...control, label: `节点标签${control.label}`, visible: nodeVisible })),
  ]
}

export function chartSunburstCenterLabelControls(): ChartSettingControl[] {
  return [
    switchControl(`${CENTER_KEY}.visible`, '显示中心标签', config => resolveChartSunburstCenterLabel(config).visible, (config, value) => { center(config).visible = value }),
    { ...textControl(`${CENTER_KEY}.format`, '中心标签数字格式', config => config.options?.sunburst?.centerLabel?.format, (config, value) => { center(config).format = value }), visible: centerVisible },
    ...chartFontControls(CENTER_KEY, config => resolveChartSunburstCenterLabel(config), center).map(control => ({ ...control, label: `中心标签${control.label}`, visible: centerVisible })),
  ]
}

export type ChartSunburstBackgroundMode = 'color' | 'pattern'
export function chartSunburstBackgroundMode(config: ChartConfig): ChartSunburstBackgroundMode {
  const value = config.options?.sunburst?.label?.backgroundColor
  return value && typeof value === 'object' && value.type === 'pattern' ? 'pattern' : 'color'
}
export function chartSunburstPattern(config: ChartConfig): ChartLabelPattern | undefined {
  const value = config.options?.sunburst?.label?.backgroundColor
  return value && typeof value === 'object' && value.type === 'pattern' ? value : undefined
}
export function setChartSunburstBackground(config: ChartConfig, value: ChartSunburstLabelOptions['backgroundColor']): void {
  const next = value === undefined ? undefined : cloneChartOptions({ sunburst: { label: { backgroundColor: value } } })!.sunburst!.label!.backgroundColor
  label(config).backgroundColor = next
}
export function setChartSunburstBackgroundMode(config: ChartConfig, mode: ChartSunburstBackgroundMode): void {
  if (mode === chartSunburstBackgroundMode(config)) return
  const value = chartLabelBackground(config.options?.sunburst?.label?.backgroundColor)
  setChartSunburstBackground(config, mode === 'pattern' ? { type: 'pattern', preset: 'pct50', foregroundColor: typeof value === 'string' ? value : '#000000', backgroundColor: '#ffffff' } : value)
}
export function chartSunburstColorBackground(config: ChartConfig): ChartColor | undefined {
  const value = config.options?.sunburst?.label?.backgroundColor
  return value && typeof value === 'object' && value.type === 'pattern' ? undefined : value
}
function pattern(config: ChartConfig): ChartLabelPattern {
  const current = chartSunburstPattern(config)
  if (current) return current
  const next: ChartLabelPattern = { type: 'pattern', preset: 'pct50', foregroundColor: '#000000', backgroundColor: '#ffffff' }
  label(config).backgroundColor = next
  return next
}
export function chartSunburstPatternControls(config: ChartConfig): ChartSettingControl[] {
  const current = chartSunburstPattern(config)?.preset, choices: Array<readonly [string, string]> = [...CHART_FILL_PATTERN_PRESETS]
  if (current && !choices.some(([value]) => value === current)) choices.push([current, current])
  return [
    selectControl(`${LABEL_KEY}.backgroundColor.preset`, '节点标签背景图案', choices, item => chartSunburstPattern(item)?.preset, (item, value) => { pattern(item).preset = value ?? 'pct50' }, 'pct50'),
    colorControl(`${LABEL_KEY}.backgroundColor.foregroundColor`, '节点标签图案前景色', item => chartSunburstPattern(item)?.foregroundColor, (item, value) => { pattern(item).foregroundColor = value ?? '#000000' }),
    colorControl(`${LABEL_KEY}.backgroundColor.backgroundColor`, '节点标签图案背景色', item => chartSunburstPattern(item)?.backgroundColor, (item, value) => { pattern(item).backgroundColor = value ?? '#ffffff' }),
  ]
}
export function clearChartSunburstSettings(config: ChartConfig, target: 'label' | 'centerLabel' | 'all'): void {
  if (!config.options?.sunburst) return
  if (target === 'all') config.options.sunburst = undefined
  else config.options.sunburst[target] = undefined
}
