import type { ChartConfig } from './chart-engine'
import { chartAutoGradientFill, supportsChartAutoGradientControl } from './chart-auto-gradient'
import { resolveChartPlotAreaForEditing, type ChartPlotSourceSeries } from './chart-layout'
import { supportsChartPlotArea } from './chart-plot-area'
import { numberControl, switchControl, type ChartSettingControl } from './chart-settings-controls'

export function chartAutoGradientControls(): ChartSettingControl[] {
  const control = switchControl('autoGradientFill', '自动渐变填充', chartAutoGradientFill,
    (config, value) => { (config.options ??= {}).autoGradientFill = value })
  control.visible = config => supportsChartAutoGradientControl(config.type)
  control.extra = '面积图默认开启，其他图型默认关闭；纹理填充优先显示。'
  control.resetLabel = '按图型恢复默认'
  control.canReset = config => config.options?.autoGradientFill !== undefined
  control.reset = config => { (config.options ??= {}).autoGradientFill = undefined }
  return [control]
}

export function chartPlotAreaControls(sourceSeries?: ChartPlotSourceSeries): ChartSettingControl[] {
  const automatic = (config: ChartConfig) => resolveChartPlotAreaForEditing({ ...config, options: { ...config.options, plotArea: undefined } }, sourceSeries)!
  const manual = switchControl('plotArea.manual', '手动设置绘图区', config => Boolean(config.options?.plotArea), (config, value) => {
    const options = config.options ??= {}
    options.plotArea = value ? options.plotArea ? { ...options.plotArea } : automatic(config) : undefined
  })
  manual.visible = config => supportsChartPlotArea(config.type)
  manual.extra = '按图表外框的百分比设置；左侧位置加宽度、顶部位置加高度须不超过 100%。'
  const coordinates = (['left', 'top', 'width', 'height'] as const).map((key, index) => {
    const control = numberControl(`plotArea.${key}`, ['左侧位置 (%)', '顶部位置 (%)', '宽度 (%)', '高度 (%)'][index]!,
      config => config.options?.plotArea ? config.options.plotArea[key] * 100 : undefined,
      (config, value) => { (config.options ??= {}).plotArea = { ...(config.options?.plotArea ?? automatic(config)), [key]: value === undefined ? automatic(config)[key] : value / 100 } },
      key === 'width' || key === 'height' ? 0.0001 : 0, 100)
    control.visible = config => supportsChartPlotArea(config.type) && Boolean(config.options?.plotArea)
    if (control.kind === 'number') control.step = 0.5
    return control
  })
  return [manual, ...coordinates]
}
