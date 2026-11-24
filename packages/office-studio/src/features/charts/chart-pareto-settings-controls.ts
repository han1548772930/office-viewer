import type { ChartLabelOptions, ChartParetoBarStyleOptions, ChartParetoLineStyleOptions, ChartSeriesOptions } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import { chartColorControl, chartFillPatternControls, chartLabelControls, chartLineControls, colorControl, numberControl, optionGroup, selectControl, type ChartSettingControl } from './chart-settings-controls'
import { chartPalette } from './chart-options'
import { chartPointStyle } from './chart-derived'
import { chartParetoBarStyle, chartParetoCumulativePointStyle, chartParetoCumulativeStyle } from './chart-pareto'

export type ChartParetoStyleGroup = 'barStyle' | 'cumulativeLineStyle'
type SourceSeries = ChartData['series'][number]
const shapes = [['circle', '圆形'], ['rect', '方形'], ['triangle', '三角形'], ['diamond', '菱形'], ['none', '无标记']] as const
const borderLabels = ['显示边框', '边框颜色', '边框线宽', '边框不透明度', '边框线型'] as const

/** Fixed cc panel keeps authored label fields while their effective visibility
 * is off. The local bar panel additionally reflects generic style inheritance. */
function paretoLabelControls(key: string, read: (config: ChartConfig) => ChartLabelOptions | undefined,
  write: (config: ChartConfig) => ChartLabelOptions, effective: (config: ChartConfig) => ChartLabelOptions | undefined): ChartSettingControl[] {
  const controls = chartLabelControls(key, read, write), toggle = controls[0]
  if (toggle?.kind === 'switch') toggle.read = config => read(config)?.visible ?? effective(config)?.visible
  return controls.map(control => control.key.endsWith('.visible') ? control
    : { ...control, visible: (config: ChartConfig) => read(config)?.visible ?? effective(config)?.visible ?? false })
}

export function chartParetoBarControls(source?: SourceSeries): ChartSettingControl[] {
  const read = (config: ChartConfig): ChartParetoBarStyleOptions | undefined => config.options?.pareto?.barStyle
  const write = (config: ChartConfig): ChartParetoBarStyleOptions => optionGroup(config, 'pareto').barStyle ??= {}
  const effective = (config: ChartConfig) => chartParetoBarStyle(config, source?.field ?? config.valueFields?.[0] ?? 1, 0,
    chartPalette(config), source?.sourceName ?? source?.name)
  return [
    chartColorControl('pareto.barStyle.color', '帕累托柱形颜色', config => read(config)?.color, (config, value) => { write(config).color = value }, 'series', config => effective(config).color),
    numberControl('pareto.barStyle.opacity', '柱形不透明度', config => read(config)?.opacity ?? effective(config).fillOpacity, (config, value) => { write(config).opacity = value }, 0, 1, 1),
    ...chartLineControls('pareto.barStyle.border', config => read(config)?.border, config => write(config).border ??= {})
      .map((control, index) => ({ ...control, label: `柱形${borderLabels[index]}` })),
    ...paretoLabelControls('pareto.barStyle.label', config => read(config)?.label, config => write(config).label ??= {}, config => effective(config).label),
  ]
}

export function chartParetoLineControls(): ChartSettingControl[] {
  const read = (config: ChartConfig): ChartParetoLineStyleOptions | undefined => config.options?.pareto?.cumulativeLineStyle
  const write = (config: ChartConfig): ChartParetoLineStyleOptions => optionGroup(config, 'pareto').cumulativeLineStyle ??= {}
  const width = numberControl('pareto.cumulativeLineStyle.width', '累计线宽', config => { const authored = read(config)?.width; return authored !== undefined && authored > 0 ? authored : undefined },
    (config, value) => { if (value === undefined || value > 0) write(config).width = value }, 0, 100, 2)
  width.extra = '非正线宽保留当前值；清空恢复 2。'
  const labels = paretoLabelControls('pareto.cumulativeLineStyle.label', config => read(config)?.label,
    config => write(config).label ??= {}, config => chartParetoCumulativeStyle(config, chartPalette(config)).label)
    .map(control => control.kind === 'select' && control.key.endsWith('.position') ? { ...control, choices: control.choices.filter(choice => choice.value !== 'outside') } : control)
  return [
    chartColorControl('pareto.cumulativeLineStyle.color', '帕累托累计线颜色', config => read(config)?.color, (config, value) => { write(config).color = value }, 'series', config => chartParetoCumulativeStyle(config, chartPalette(config)).color),
    numberControl('pareto.cumulativeLineStyle.opacity', '累计线不透明度', config => read(config)?.opacity, (config, value) => { write(config).opacity = value }, 0, 1, 1), width,
    selectControl('pareto.cumulativeLineStyle.dashType', '累计线型', [['solid', '实线'], ['dashed', '虚线'], ['dotted', '点线']], config => read(config)?.dashType, (config, value) => { write(config).dashType = value }, 'solid'),
    selectControl('pareto.cumulativeLineStyle.point.shape', '累计线标记形状', shapes, config => read(config)?.point?.shape, (config, value) => { (write(config).point ??= {}).shape = value }, 'circle'),
    numberControl('pareto.cumulativeLineStyle.point.size', '累计线标记大小', config => read(config)?.point?.size, (config, value) => { (write(config).point ??= {}).size = value }, 0, 256, 7),
    colorControl('pareto.cumulativeLineStyle.point.color', '累计线标记颜色', config => read(config)?.point?.color, (config, value) => { (write(config).point ??= {}).color = value }), ...labels,
  ]
}

export function chartParetoDataPointControls(group: ChartParetoStyleGroup, index: number, source?: SourceSeries): ChartSettingControl[] {
  const title = group === 'barStyle' ? '柱形项目' : '累计线项目', key = `pareto.${group}.dataPoints.${index}`
  const read = (config: ChartConfig) => config.options?.pareto?.[group]?.dataPoints?.[index]
  const write = (config: ChartConfig) => ((optionGroup(config, 'pareto')[group] ??= {}).dataPoints ??= {})[index] ??= {}
  const effective = (config: ChartConfig): ChartSeriesOptions => group === 'cumulativeLineStyle'
    ? chartParetoCumulativePointStyle(chartParetoCumulativeStyle(config, chartPalette(config)), index)
    : chartPointStyle(chartParetoBarStyle(config, source?.field ?? config.valueFields?.[0] ?? 1, 0, chartPalette(config), source?.sourceName ?? source?.name), index, 'pareto')
  const border = chartLineControls(`${key}.border`, config => read(config)?.border, config => write(config).border ??= {})
    .map((control, position) => ({ ...control, label: `${title}${borderLabels[position]}` }))
  return [
    chartColorControl(`${key}.color`, `${title}颜色`, config => read(config)?.color, (config, value) => { const point = write(config); point.color = value; point.fillPattern = null }, 'dataPoint', config => effective(config).color),
    numberControl(`${key}.fillOpacity`, `${title}不透明度`, config => read(config)?.fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1),
    ...chartFillPatternControls(key, read, write, effective, `${title}`),
    ...(group === 'cumulativeLineStyle' ? [
      numberControl(`${key}.size`, `${title}标记大小`, config => read(config)?.size, (config, value) => { write(config).size = value }, 0, 256),
      selectControl(`${key}.shape`, `${title}标记形状`, shapes, config => read(config)?.shape, (config, value) => { write(config).shape = value }),
    ] : []), ...border,
  ]
}
