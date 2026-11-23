import type { ChartFontOptions } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { chartGaugePresetOptions, resolveChartGaugeOptions, type ChartGaugePreset } from './chart-gauge'
import { chartFontControls, colorControl, numberControl, optionGroup, selectControl, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'
import { chartSeriesStyle } from './chart-options'

function gaugeFontControls(key: string, title: string, read: (config: ChartConfig) => ChartFontOptions | undefined, write: (config: ChartConfig) => ChartFontOptions): ChartSettingControl[] {
  return chartFontControls(key, read, write).map(control => ({ ...control, label: `${title}${control.label}` }))
}

export function applyChartGaugePreset(config: ChartConfig, preset: ChartGaugePreset): void {
  const gauge = chartGaugePresetOptions(preset)
  ;(config.options ??= {}).gauge = gauge
}

/** Undefined must survive draft cloning: omitted option groups mean retain on update. */
export function resetChartGaugeSettings(config: ChartConfig): void { (config.options ??= {}).gauge = undefined }

/** Explicit label omission is meaningful; a single-field source uses its field title. */
export function chartGaugeSettingsMapping(config: ChartConfig, fieldCount?: number): { valueIndex: number; labelIndex?: number } {
  const authored = config.options?.mapping?.gauge
  if (authored) return { ...authored }
  const range = config.source.range
  const count = fieldCount ?? (config.source.orientation === 'columns' ? range.endCol - range.startCol + 1 : range.endRow - range.startRow + 1)
  return { valueIndex: config.valueFields?.[0] ?? (count > 1 ? 1 : 0), labelIndex: count > 1 ? config.categoryField ?? 0 : undefined }
}

/** UI defaults follow the selected semantic mode. Each control writes only authored model values. */
export function chartGaugeSettingsControls(config: ChartConfig, fieldCount?: number, sourceName?: string): ChartSettingControl[] {
  const resolved = (item: ChartConfig) => {
    const field = chartGaugeSettingsMapping(item, fieldCount).valueIndex
    return resolveChartGaugeOptions(item.options?.gauge, 100, item.options?.appearance?.font,
      chartSeriesStyle(item, field, 0, undefined, sourceName).label)
  }
  const defaults = resolved(config), legacy = defaults.mode === 'pointer' || defaults.mode === 'progress'
  const axisLine = (item: ChartConfig) => optionGroup(item, 'gauge').axisLine ??= {}
  const axisTick = (item: ChartConfig) => optionGroup(item, 'gauge').axisTick ??= {}
  const splitLine = (item: ChartConfig) => optionGroup(item, 'gauge').splitLine ??= {}
  const axisLabel = (item: ChartConfig) => optionGroup(item, 'gauge').axisLabel ??= {}
  const pointer = (item: ChartConfig) => optionGroup(item, 'gauge').pointer ??= {}
  const anchor = (item: ChartConfig) => optionGroup(item, 'gauge').anchor ??= {}
  const progress = (item: ChartConfig) => optionGroup(item, 'gauge').progress ??= {}
  const title = (item: ChartConfig) => optionGroup(item, 'gauge').title ??= {}
  const detail = (item: ChartConfig) => optionGroup(item, 'gauge').detail ??= {}
  return [
    selectControl('gauge.mode', '仪表模式', [['modern', '现代'], ['classic', '经典'], ['pointer', '指针样式'], ['progress', '进度样式']], item => item.options?.gauge?.mode, (item, value) => { optionGroup(item, 'gauge').mode = value }, 'modern'),
    numberControl('gauge.min', '仪表最小刻度', item => item.options?.gauge?.min, (item, value) => { optionGroup(item, 'gauge').min = value }, undefined, undefined, 0),
    numberControl('gauge.max', '仪表最大刻度', item => item.options?.gauge?.max, (item, value) => { optionGroup(item, 'gauge').max = value }, undefined, undefined, 100),
    numberControl('gauge.startAngle', '仪表起始角', item => item.options?.gauge?.startAngle, (item, value) => { optionGroup(item, 'gauge').startAngle = value }, -360, 360, 225),
    numberControl('gauge.endAngle', '仪表终止角', item => item.options?.gauge?.endAngle, (item, value) => { optionGroup(item, 'gauge').endAngle = value }, -360, 360, -45),
    numberControl('gauge.splitNumber', '主刻度分段数', item => item.options?.gauge?.splitNumber, (item, value) => { optionGroup(item, 'gauge').splitNumber = value }, 1, 100, 5),
    switchControl('gauge.axisLine.visible', '显示仪表轴线', item => item.options?.gauge?.axisLine?.visible, (item, value) => { axisLine(item).visible = value }, defaults.axisLine.visible),
    numberControl('gauge.axisLine.width', '仪表轴线宽度', item => item.options?.gauge?.axisLine?.width, (item, value) => { axisLine(item).width = value }, 0, 100, legacy ? undefined : defaults.axisLine.width),
    colorControl('gauge.axisLine.color', '仪表轴线颜色', item => item.options?.gauge?.axisLine?.color, (item, value) => { axisLine(item).color = value }),
    switchControl('gauge.axisTick.visible', '显示次刻度', item => item.options?.gauge?.axisTick?.visible, (item, value) => { axisTick(item).visible = value }, defaults.axisTick.visible),
    numberControl('gauge.axisTick.splitNumber', '每段次刻度分段数', item => item.options?.gauge?.axisTick?.splitNumber, (item, value) => { axisTick(item).splitNumber = value }, 1, 100, defaults.axisTick.splitNumber),
    numberControl('gauge.axisTick.length', '次刻度长度', item => item.options?.gauge?.axisTick?.length, (item, value) => { axisTick(item).length = value }, 0, 100, defaults.axisTick.length),
    numberControl('gauge.axisTick.width', '次刻度线宽', item => item.options?.gauge?.axisTick?.width, (item, value) => { axisTick(item).width = value }, 0, 100, defaults.axisTick.width),
    colorControl('gauge.axisTick.color', '次刻度颜色', item => item.options?.gauge?.axisTick?.color, (item, value) => { axisTick(item).color = value }),
    switchControl('gauge.splitLine.visible', '显示主刻度', item => item.options?.gauge?.splitLine?.visible, (item, value) => { splitLine(item).visible = value }, defaults.splitLine.visible),
    numberControl('gauge.splitLine.length', '主刻度长度', item => item.options?.gauge?.splitLine?.length, (item, value) => { splitLine(item).length = value }, 0, 100, legacy ? undefined : defaults.splitLine.length),
    numberControl('gauge.splitLine.width', '主刻度线宽', item => item.options?.gauge?.splitLine?.width, (item, value) => { splitLine(item).width = value }, 0, 100, defaults.splitLine.width),
    colorControl('gauge.splitLine.color', '主刻度颜色', item => item.options?.gauge?.splitLine?.color, (item, value) => { splitLine(item).color = value }),
    switchControl('gauge.axisLabel.visible', '显示刻度文字', item => item.options?.gauge?.axisLabel?.visible, (item, value) => { axisLabel(item).visible = value }, defaults.axisLabel.visible),
    numberControl('gauge.axisLabel.distance', '刻度文字距离', item => item.options?.gauge?.axisLabel?.distance, (item, value) => { axisLabel(item).distance = value }, -100, 100, legacy ? undefined : defaults.axisLabel.distance),
    textControl('gauge.axisLabel.format', '刻度数字格式', item => item.options?.gauge?.axisLabel?.format, (item, value) => { axisLabel(item).format = value }),
    ...gaugeFontControls('gauge.axisLabel', '刻度文字', item => resolved(item).axisLabel, axisLabel),
    switchControl('gauge.pointer.visible', '显示指针', item => item.options?.gauge?.pointer?.visible, (item, value) => { pointer(item).visible = value }, defaults.pointer.visible),
    selectControl('gauge.pointer.shape', '指针形状', [['needle', '尖针'], ['slim', '细针'], ['arrow', '箭头']], item => item.options?.gauge?.pointer?.shape, (item, value) => { pointer(item).shape = value }, 'needle'),
    numberControl('gauge.pointer.length', '指针长度比例', item => item.options?.gauge?.pointer?.length, (item, value) => { pointer(item).length = value }, 0, 1, defaults.pointer.length),
    numberControl('gauge.pointer.width', '指针宽度', item => item.options?.gauge?.pointer?.width, (item, value) => { pointer(item).width = value }, 0, 100, defaults.pointer.width),
    colorControl('gauge.pointer.color', '指针颜色', item => item.options?.gauge?.pointer?.color, (item, value) => { pointer(item).color = value }),
    switchControl('gauge.anchor.visible', '显示中心锚点', item => item.options?.gauge?.anchor?.visible, (item, value) => { anchor(item).visible = value }, defaults.anchor.visible),
    numberControl('gauge.anchor.size', '中心锚点直径', item => item.options?.gauge?.anchor?.size, (item, value) => { anchor(item).size = value }, 0, 256, defaults.anchor.size),
    colorControl('gauge.anchor.color', '中心锚点颜色', item => item.options?.gauge?.anchor?.color, (item, value) => { anchor(item).color = value }),
    colorControl('gauge.anchor.borderColor', '中心锚点边框颜色', item => item.options?.gauge?.anchor?.borderColor, (item, value) => { anchor(item).borderColor = value }),
    numberControl('gauge.anchor.borderWidth', '中心锚点边框宽度', item => item.options?.gauge?.anchor?.borderWidth, (item, value) => { anchor(item).borderWidth = value }, 0, 100, defaults.anchor.borderWidth),
    switchControl('gauge.progress.visible', '显示进度弧', item => item.options?.gauge?.progress?.visible, (item, value) => { progress(item).visible = value }, defaults.progress.visible),
    numberControl('gauge.progress.width', '进度弧宽度', item => item.options?.gauge?.progress?.width, (item, value) => { progress(item).width = value }, 0, 100, legacy ? undefined : defaults.progress.width),
    colorControl('gauge.progress.color', '进度弧颜色', item => item.options?.gauge?.progress?.color, (item, value) => { progress(item).color = value }),
    switchControl('gauge.progress.rounded', '进度弧圆端', item => item.options?.gauge?.progress?.rounded, (item, value) => { progress(item).rounded = value }, defaults.progress.rounded),
    switchControl('gauge.title.visible', '显示观测名称', item => resolved(item).title.visible, (item, value) => { title(item).visible = value }),
    numberControl('gauge.title.offsetY', '观测名称纵向比例', item => item.options?.gauge?.title?.offsetY, (item, value) => { title(item).offsetY = value }, 0, 1, defaults.title.offsetY),
    ...gaugeFontControls('gauge.title', '观测名称', item => resolved(item).title, title),
    switchControl('gauge.detail.visible', '显示观测数值', item => resolved(item).detail.visible, (item, value) => { detail(item).visible = value }),
    numberControl('gauge.detail.offsetY', '观测数值纵向比例', item => item.options?.gauge?.detail?.offsetY, (item, value) => { detail(item).offsetY = value }, 0, 1, defaults.detail.offsetY),
    textControl('gauge.detail.format', '观测数值数字格式', item => resolved(item).detail.format, (item, value) => { detail(item).format = value }),
    ...gaugeFontControls('gauge.detail', '观测数值', item => resolved(item).detail, detail),
  ]
}
