import type { ChartLineOptions, ChartPieAggregateStyleOptions, ChartPieSliceStyleOptions } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import { chartPalette } from './chart-options'
import { chartPieAggregateStyle, chartPieSliceStyle } from './chart-pie'
import { chartColorControl, chartFillPatternControls, chartLineControls, numberControl, optionGroup, textControl,
  type ChartSettingControl } from './chart-settings-controls'

export function chartPieAggregateControls(source?: Pick<ChartData['series'][number], 'field' | 'name' | 'sourceName'>, aggregateIndex = 0): ChartSettingControl[] {
  const read = (config: ChartConfig) => config.options?.pie?.aggregateStyle
  const write = (config: ChartConfig): ChartPieAggregateStyleOptions => optionGroup(config, 'pie').aggregateStyle ??= {}
  const effective = (config: ChartConfig) => chartPieAggregateStyle(config,
    { field: source?.field ?? config.valueFields?.[0] ?? 1, name: source?.name ?? '', sourceName: source?.sourceName, values: [] }, aggregateIndex, chartPalette(config))
  return [
    textControl('pie.aggregateLabel', '主图合计标签', config => config.options?.pie?.aggregateLabel, (config, value) => { optionGroup(config, 'pie').aggregateLabel = value }),
    chartColorControl('pie.aggregateStyle.color', '合计扇区颜色', config => read(config)?.color,
      (config, value) => { const style = write(config); style.color = value; style.fillPattern = null }, 'dataPoint', config => effective(config).color),
    ...chartFillPatternControls('pie.aggregateStyle', read, write, effective, '合计扇区'),
    numberControl('pie.aggregateStyle.fillOpacity', '合计扇区不透明度', config => read(config)?.fillOpacity,
      (config, value) => { write(config).fillOpacity = value }, 0, 1, 1),
    numberControl('pie.aggregateStyle.explosion', '合计扇区分离（%）', config => read(config)?.explosion,
      (config, value) => { write(config).explosion = value }, 0, 100),
    ...chartLineControls('pie.aggregateStyle.border', config => read(config)?.border, config => write(config).border ??= {}),
  ]
}

export function chartPieConnectorControls(): ChartSettingControl[] {
  const read = (config: ChartConfig) => config.options?.pie?.secondaryPlot?.connector
  const write = (config: ChartConfig): ChartLineOptions => (optionGroup(config, 'pie').secondaryPlot ??= {}).connector ??= {}
  const controls = chartLineControls('pie.secondaryPlot.connector', read, write, true)
  const labels = ['显示主副图连接线', '主副图连接线颜色', '主副图连接线宽', '主副图连接线不透明度', '主副图连接线型']
  return controls.map((control, index) => ({ ...control, label: labels[index] ?? control.label,
    ...(index === 0 ? {} : { visible: (config: ChartConfig) => read(config)?.visible !== false }) }))
}

export function chartPieSliceControls(sourcePointIndex: number, source?: Pick<ChartData['series'][number], 'field' | 'name' | 'sourceName'>, allowExplosion = true): ChartSettingControl[] {
  const key = `pie.sliceStyles.${sourcePointIndex}`
  const read = (config: ChartConfig) => config.options?.pie?.sliceStyles?.[sourcePointIndex]
  const write = (config: ChartConfig): ChartPieSliceStyleOptions => (optionGroup(config, 'pie').sliceStyles ??= {})[sourcePointIndex] ??= {}
  const effective = (config: ChartConfig) => chartPieSliceStyle(config,
    { field: source?.field ?? config.valueFields?.[0] ?? 1, name: source?.name ?? '', sourceName: source?.sourceName, values: [] }, sourcePointIndex, chartPalette(config))
  return [
    chartColorControl(`${key}.color`, '扇区颜色', config => read(config)?.color,
      (config, value) => { const style = write(config); style.color = value; style.fillPattern = null }, 'dataPoint', config => effective(config).color),
    ...chartFillPatternControls(key, read, write, effective, '扇区'),
    numberControl(`${key}.fillOpacity`, '扇区不透明度', config => read(config)?.fillOpacity,
      (config, value) => { write(config).fillOpacity = value }, 0, 1, 1),
    ...(allowExplosion ? [numberControl(`${key}.explosion`, '扇区分离（%）', config => read(config)?.explosion ?? config.options?.pie?.explosion,
      (config, value) => { write(config).explosion = value }, 0, 100)] : []),
    ...chartLineControls(`${key}.border`, config => read(config)?.border, config => write(config).border ??= {}),
  ]
}
