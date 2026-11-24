import type { ChartRadarAxisLabelOptions, ChartRadarOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { chartAxisControls, chartFontControls, numberControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'
import { chartRadarAxisLabel } from './chart-radar'

function resetRadarField(control: ChartSettingControl, field: keyof ChartRadarOptions): ChartSettingControl {
  return { ...control, canReset: config => config.options?.radar?.[field] !== undefined,
    reset: config => { if (config.options?.radar) config.options.radar[field] = undefined } }
}
function resetRadarLabelField(control: ChartSettingControl, field: keyof ChartRadarAxisLabelOptions): ChartSettingControl {
  return { ...control, canReset: config => config.options?.radar?.axisLabel?.[field] !== undefined,
    reset: config => { if (config.options?.radar?.axisLabel) config.options.radar.axisLabel[field] = undefined } }
}

export function chartRadarSettingsControls(config: ChartConfig): ChartSettingControl[] {
  const read = (item: ChartConfig) => chartRadarAxisLabel(item)
  const write = (item: ChartConfig) => optionGroup(item, 'radar').axisLabel ??= {}
  const controls: ChartSettingControl[] = [
    resetRadarField(selectControl('radar.shape', '雷达网格', [['polygon', '多边形'], ['circle', '圆形']], config => config.options?.radar?.shape, (config, value) => { optionGroup(config, 'radar').shape = value }, 'polygon'), 'shape'),
    resetRadarField(switchControl('radar.fill', '填充雷达区域', config => config.options?.radar?.fill, (config, value) => { optionGroup(config, 'radar').fill = value }), 'fill'),
    resetRadarField(switchControl('radar.axisLabel.visible', '显示数值刻度标签', config => read(config).visible, (config, value) => { write(config).visible = value }), 'axisLabel'),
  ]
  if (read(config).visible) controls.push(
    resetRadarLabelField(textControl('radar.axisLabel.format', '刻度数字格式', item => read(item).format, (item, value) => { write(item).format = value }), 'format'),
    ...chartFontControls('radar.axisLabel', read, write).map(control => resetRadarLabelField(control, control.key.slice('radar.axisLabel.'.length) as keyof ChartRadarAxisLabelOptions)),
  )
  return controls
}

export function chartRadarCategoryAxisControls(): ChartSettingControl[] {
  return chartAxisControls('xAxis', false).filter(control => ['xAxis.title', 'xAxis.titleAlignment', 'xAxis.reverse'].includes(control.key)
    || control.key.startsWith('xAxis.titleStyle.') || control.key.startsWith('xAxis.label.'))
}

export function chartRadarValueAxisControls(): ChartSettingControl[] {
  return [
    ...chartAxisControls('yAxis', false).filter(control => !control.key.startsWith('yAxis.label.')),
    numberControl('yAxis.min', '最小刻度', config => config.yAxis?.min, (config, value) => { (config.yAxis ??= {}).min = value }),
    numberControl('yAxis.max', '最大刻度', config => config.yAxis?.max, (config, value) => { (config.yAxis ??= {}).max = value }),
    numberControl('yAxis.label.rotate', '刻度标签旋转', config => config.yAxis?.label?.rotate, (config, value) => { ((config.yAxis ??= {}).label ??= {}).rotate = value }, -90, 90),
  ]
}
