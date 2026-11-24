import type { ChartLabelOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { chartFontControls, numberControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'

type Heatmap = NonNullable<NonNullable<ChartConfig['options']>['heatmap']>
type CellLabel = NonNullable<Heatmap['label']>
const label = (config: ChartConfig): CellLabel => optionGroup(config, 'heatmap').label ??= {}
const visible = (config: ChartConfig): boolean => (config.options?.heatmap?.label?.visible ?? config.options?.allSeriesStyle?.label?.visible) === true
const owned = (config: ChartConfig, key: keyof Heatmap): boolean => config.options?.heatmap?.[key] !== undefined
function resetField(control: ChartSettingControl, field: keyof Heatmap): ChartSettingControl {
  return { ...control, canReset: config => owned(config, field), reset: config => { if (config.options?.heatmap) config.options.heatmap[field] = undefined } }
}
function resetLabelField(control: ChartSettingControl, field: keyof ChartLabelOptions): ChartSettingControl {
  return { ...control, visible, canReset: config => config.options?.heatmap?.label?.[field] !== undefined,
    reset: config => { if (config.options?.heatmap?.label) config.options.heatmap.label[field] = undefined } }
}

/** Fixed ml/hl order: min/max, unit, label toggle; visible label exposes
 * inside/top/bottom then its font. Format is a functional public-API extension;
 * the other chart types' content flags are not offered here. */
export function chartHeatmapSettingsControls(): ChartSettingControl[] {
  const controls: ChartSettingControl[] = [
    resetField(numberControl('heatmap.min', '色阶最小值', config => config.options?.heatmap?.min,
      (config, value) => { optionGroup(config, 'heatmap').min = value }), 'min'),
    resetField(numberControl('heatmap.max', '色阶最大值', config => config.options?.heatmap?.max,
      (config, value) => { optionGroup(config, 'heatmap').max = value }), 'max'),
    resetField(textControl('heatmap.unit', '数值单位', config => config.options?.heatmap?.unit,
      (config, value) => { optionGroup(config, 'heatmap').unit = value }), 'unit'),
    resetField(switchControl('heatmap.label.visible', '显示单元格标签', config => config.options?.heatmap?.label?.visible ?? config.options?.allSeriesStyle?.label?.visible,
      (config, value) => { label(config).visible = value }), 'label'),
    resetLabelField(selectControl('heatmap.label.position', '单元格标签位置', [['inside', '内部'], ['top', '上方'], ['bottom', '下方']],
      config => config.options?.heatmap?.label?.position === 'top' ? 'top' : config.options?.heatmap?.label?.position === 'bottom' ? 'bottom' : 'inside',
      (config, value) => { label(config).position = value }, 'inside'), 'position'),
  ]
  for (const control of chartFontControls('heatmap.label', config => config.options?.heatmap?.label, label)) {
    const field = control.key.slice('heatmap.label.'.length) as keyof ChartLabelOptions
    const reset = resetLabelField(control, field)
    if (reset.kind === 'number' && field === 'fontSize') reset.fallback = 12
    controls.push(reset)
  }
  controls.push(resetLabelField(textControl('heatmap.label.format', '单元格标签数字格式', config => config.options?.heatmap?.label?.format,
    (config, value) => { label(config).format = value }), 'format'))
  return controls
}
