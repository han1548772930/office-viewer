import type { ChartConfig } from './chart-engine'
import { chartAreaLineStylePatch, resolveChartAreaLineStyle, supportsChartAreaLineStyle } from './chart-line-and-area'
import { selectControl, type ChartSettingControl } from './chart-settings-controls'

function inheritsLegacyCombination(config: ChartConfig): boolean {
  return config.type === 'combination' && config.options?.area?.lineStyle === undefined && config.options?.area?.smooth === true
}

/** Rendered by the existing antdv-next Form.Item / Select adapter. */
export function chartLineAndAreaSettingsControls(config: ChartConfig): ChartSettingControl[] {
  if (!supportsChartAreaLineStyle(config.type)) return []
  const control = selectControl('area.lineStyle', '折线与面积线型', [['line', '直线'], ['smooth', '平滑'], ['step', '阶梯']],
    item => inheritsLegacyCombination(item) ? undefined : resolveChartAreaLineStyle(item), (item, value) => {
      const patch = chartAreaLineStylePatch(item.type, value)
      const area = (item.options ??= {}).area ??= {}
      area.lineStyle = patch.area?.lineStyle
      area.smooth = undefined
    }, inheritsLegacyCombination(config) ? undefined : 'line')
  return [{ ...control, extra: inheritsLegacyCombination(config)
    ? '继承旧配置：面积系列平滑，折线系列直线。选择线型后统一生效。'
    : config.type === 'combination' ? '作用于组合图中的折线和面积系列。' : undefined,
    canReset: item => item.options?.area?.lineStyle !== undefined || item.options?.area?.smooth !== undefined,
    reset: item => { if (item.options?.area) { item.options.area.lineStyle = undefined; item.options.area.smooth = undefined } } }]
}
