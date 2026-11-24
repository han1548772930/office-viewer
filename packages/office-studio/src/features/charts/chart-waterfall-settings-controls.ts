import type { ChartSeriesSelector, ChartWaterfallConnectorOptions, ChartWaterfallSeriesStyle, ChartWaterfallStyleTarget } from '../../domain/model'
import type { ChartConfig, ChartData } from './chart-engine'
import type { ChartSettingsField } from './chart-settings-draft'
import { chartSeriesOverride } from './chart-series-selector'
import { chartPalette } from './chart-options'
import { chartPointStyle } from './chart-derived'
import { chartWaterfallConnector, chartWaterfallRole, chartWaterfallRoleStyle } from './chart-waterfall'
import { chartColorControl, chartFillPatternControls, chartLabelControls, chartLineControls, colorControl, numberControl, optionGroup, selectControl, seriesStyleGroup, switchControl, textControl, type ChartSettingControl } from './chart-settings-controls'

type SourceSeries = Pick<ChartData['series'][number], 'field' | 'name' | 'sourceName'> & { position?: number }

export const CHART_WATERFALL_POINT_CHOICE_LIMIT = 16_384

export function chartWaterfallLayoutControls(): ChartSettingControl[] {
  const stack = selectControl('waterfall.stackType', '瀑布图排列方式', [['arrangement', '按来源依次排列'], ['stacked', '按分类堆积']],
    config => config.options?.waterfall?.stackType, (config, value) => {
      const waterfall = optionGroup(config, 'waterfall')
      if (value === 'stacked' && waterfall.stackType !== 'stacked') delete waterfall.pointRoles
      waterfall.stackType = value
    }, 'arrangement')
  stack.extra = '切换为堆积会清除源项目的合计标记。'
  stack.reset = config => { optionGroup(config, 'waterfall').stackType = undefined }
  stack.canReset = config => config.options?.waterfall?.stackType !== undefined
  const subtotal = switchControl('waterfall.useSubtotal', '显示自动合计', config => config.options?.waterfall?.useSubtotal,
    (config, value) => { optionGroup(config, 'waterfall').useSubtotal = value }, true)
  subtotal.extra = '排列时在各来源末尾显示整图最终合计；堆积时只显示一个全局合计。'
  subtotal.reset = config => { optionGroup(config, 'waterfall').useSubtotal = undefined }
  subtotal.canReset = config => config.options?.waterfall?.useSubtotal !== undefined
  return [stack, subtotal]
}

/** The fixed resolver builds a Map from the authored array, so duplicate
 * identities take their final role without rewriting the saved collection. */
export function chartWaterfallPointRole(config: ChartConfig, field: number, index: number): 'change' | 'total' {
  const roles = config.options?.waterfall?.pointRoles ?? []
  for (let position = roles.length - 1; position >= 0; position--) {
    const role = roles[position]!
    if (role.seriesIndex === field && role.dataPointIndex === index) return role.role
  }
  return 'change'
}

export function chartWaterfallPointRoleControls(field: number, index: number): ChartSettingControl[] {
  if (!Number.isSafeInteger(field) || field < 0 || field > 1_048_575 || !Number.isSafeInteger(index) || index < 0 || index > 1_048_575) {
    throw new RangeError('瀑布图项目须使用有效的来源字段和项目索引。')
  }
  const write = (config: ChartConfig, value: 'change' | 'total' | undefined): void => {
    if (config.options?.waterfall?.stackType === 'stacked') throw new TypeError('堆积瀑布图不能指定源项目含义。')
    const previous = config.options?.waterfall?.pointRoles ?? []
    const retained = previous.filter(role => role.seriesIndex !== field || role.dataPointIndex !== index)
    const roles = value === undefined ? retained : [...retained, { seriesIndex: field, dataPointIndex: index, role: value }]
    optionGroup(config, 'waterfall').pointRoles = roles.length ? roles : undefined
  }
  const control = selectControl(`waterfall.pointRoles.${field}.${index}`, '源项目含义', [['change', '变化'], ['total', '合计']],
    config => chartWaterfallPointRole(config, field, index), write, 'change')
  control.extra = '“合计”从零绘制并重设后续累计；“变化”接续前一项目累计。'
  control.reset = config => write(config, undefined)
  control.canReset = config => Boolean(config.options?.waterfall?.pointRoles?.some(role => role.seriesIndex === field && role.dataPointIndex === index))
  control.resetLabel = '恢复默认变化'
  return [control]
}

function pointSources(selector: ChartSeriesSelector | undefined, data: ChartData): ChartData['series'] {
  return selector === undefined ? data.series : data.series.filter(series => typeof selector === 'number'
    ? series.field === selector : (series.sourceName ?? series.name) === selector.name)
}

export function chartWaterfallStylePointIsActive(config: ChartConfig, target: ChartWaterfallStyleTarget, index: number,
  selector: ChartSeriesSelector | undefined, data: ChartData | undefined): boolean {
  if (!data || !Number.isSafeInteger(index) || index < 0 || index > data.categories.length) return false
  const stacked = config.options?.waterfall?.stackType === 'stacked', sources = pointSources(selector, data)
  if (index === data.categories.length) return data.categories.length > 0 && sources.length > 0 && target === 'subtotal'
    && config.options?.waterfall?.useSubtotal !== false && (!stacked || selector === undefined)
  if (stacked && target === 'subtotal') return false
  return sources.some(series => {
    const value = series.values[index], role = stacked ? 'change' : chartWaterfallPointRole(config, series.field, index)
    return chartWaterfallRole(role, value != null && Number.isFinite(value) ? value : 0) === target
  })
}

/** Large sources use a direct one-based index input; this bounded list only
 * supplies ordinary Select options, including the real synthetic identity. */
export function chartWaterfallStylePoints(config: ChartConfig, target: ChartWaterfallStyleTarget,
  selector: ChartSeriesSelector | undefined, data: ChartData | undefined): ChartSettingsField[] {
  if (!data) return []
  const sources = pointSources(selector, data), roles = new Map<string, 'change' | 'total'>()
  for (const role of config.options?.waterfall?.pointRoles ?? []) roles.set(`${role.seriesIndex}:${role.dataPointIndex}`, role.role)
  const stacked = config.options?.waterfall?.stackType === 'stacked', points: ChartSettingsField[] = []
  if (!stacked || target !== 'subtotal') for (let value = 0; value < Math.min(data.categories.length, CHART_WATERFALL_POINT_CHOICE_LIMIT); value++) {
    if (!sources.some(series => {
      const source = series.values[value], role = stacked ? 'change' : roles.get(`${series.field}:${value}`) ?? 'change'
      return chartWaterfallRole(role, source != null && Number.isFinite(source) ? source : 0) === target
    })) continue
    points.push({ value, label: `${value + 1} · ${data.categories[value]}` })
  }
  if (chartWaterfallStylePointIsActive(config, target, data.categories.length, selector, data)) points.push({ value: data.categories.length,
    label: stacked ? '自动全局合计' : selector === undefined ? '各来源的自动合计' : '该来源的自动合计' })
  return points
}

function effectiveField(config: ChartConfig, target: ChartWaterfallStyleTarget, selector: ChartSeriesSelector | undefined, source: SourceSeries | undefined): number | undefined {
  if (selector === undefined && target === 'subtotal' && config.options?.waterfall?.stackType === 'stacked') return undefined
  return typeof selector === 'number' ? selector : source?.field ?? config.valueFields?.[0] ?? 1
}

export function chartWaterfallStyleGroup(config: ChartConfig, target: ChartWaterfallStyleTarget, selector?: ChartSeriesSelector): ChartWaterfallSeriesStyle {
  if (selector === undefined) return (optionGroup(config, 'waterfall').styles ??= {})[target] ??= {}
  return (seriesStyleGroup(config, selector).waterfallStyles ??= {})[target] ??= {}
}

export function chartWaterfallStyleControls(target: ChartWaterfallStyleTarget, selector?: ChartSeriesSelector, source?: SourceSeries): ChartSettingControl[] {
  const prefix = selector === undefined ? 'waterfall.styles' : typeof selector === 'number' ? `series.${selector}.waterfallStyles` : `namedSeries.${JSON.stringify(selector.name)}.waterfallStyles`
  const key = `${prefix}.${target}`
  const read = (config: ChartConfig): ChartWaterfallSeriesStyle | undefined => selector === undefined
    ? config.options?.waterfall?.styles?.[target] : chartSeriesOverride(config.options, selector)?.waterfallStyles?.[target]
  const write = (config: ChartConfig) => chartWaterfallStyleGroup(config, target, selector)
  const effective = (config: ChartConfig) => chartWaterfallRoleStyle(config, effectiveField(config, target, selector, source),
    source?.position ?? 0, target, chartPalette(config), typeof selector === 'object' ? selector.name : source?.sourceName ?? source?.name)
  const controls: ChartSettingControl[] = [
    textControl(`${key}.name`, '角色显示名称', config => read(config)?.name, (config, value) => { write(config).name = value }),
    chartColorControl(`${key}.color`, '角色填充颜色', config => read(config)?.color, (config, value) => { const style = write(config); style.color = value; style.fillPattern = null }, 'series', config => effective(config).color),
    ...chartFillPatternControls(key, read, write, effective, '角色'),
    numberControl(`${key}.fillOpacity`, '角色填充不透明度', config => read(config)?.fillOpacity ?? effective(config).fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1, 1),
    switchControl(`${key}.rightAxis`, '角色使用右侧数值轴', config => read(config)?.rightAxis ?? effective(config).rightAxis, (config, value) => { write(config).rightAxis = value }),
    ...chartLineControls(`${key}.border`, config => read(config)?.border, config => write(config).border ??= {}),
    ...chartLabelControls(`${key}.label`, config => read(config)?.label, config => write(config).label ??= {}),
  ]
  return controls
}

export function chartWaterfallPointControls(target: ChartWaterfallStyleTarget, index: number, selector?: ChartSeriesSelector, source?: SourceSeries): ChartSettingControl[] {
  const key = `waterfall.${target}.point.${index}`
  const read = (config: ChartConfig) => (selector === undefined ? config.options?.waterfall?.styles?.[target]
    : chartSeriesOverride(config.options, selector)?.waterfallStyles?.[target])?.dataPoints?.[index]
  const write = (config: ChartConfig) => (chartWaterfallStyleGroup(config, target, selector).dataPoints ??= {})[index] ??= {}
  const effective = (config: ChartConfig) => chartPointStyle(chartWaterfallRoleStyle(config,
    effectiveField(config, target, selector, source), source?.position ?? 0, target,
    chartPalette(config), typeof selector === 'object' ? selector.name : source?.sourceName ?? source?.name), index, 'waterfall')
  return [
    chartColorControl(`${key}.color`, '角色项目填充颜色', config => read(config)?.color, (config, value) => { const point = write(config); point.color = value; point.fillPattern = null }, 'dataPoint', config => effective(config).color),
    ...chartFillPatternControls(key, read, write, effective, '角色项目'),
    numberControl(`${key}.fillOpacity`, '角色项目不透明度', config => read(config)?.fillOpacity, (config, value) => { write(config).fillOpacity = value }, 0, 1),
    ...chartLineControls(`${key}.border`, config => read(config)?.border, config => write(config).border ??= {}),
  ]
}

function connectorGroup(config: ChartConfig): ChartWaterfallConnectorOptions {
  const waterfall = optionGroup(config, 'waterfall')
  if (typeof waterfall.connector !== 'object') waterfall.connector = {}
  return waterfall.connector
}

export function chartWaterfallConnectorControls(): ChartSettingControl[] {
  const read = (config: ChartConfig): ChartWaterfallConnectorOptions | undefined => typeof config.options?.waterfall?.connector === 'object' ? config.options.waterfall.connector : undefined
  const effective = (config: ChartConfig) => chartWaterfallConnector(config)
  const toggle = switchControl('waterfall.connector', '显示累计连接线', config => Boolean(effective(config)), (config, value) => {
    const previous = config.options?.waterfall?.connector
    optionGroup(config, 'waterfall').connector = value ? typeof previous === 'object' ? previous : true : false
  })
  toggle.reset = config => { optionGroup(config, 'waterfall').connector = undefined }
  toggle.resetLabel = '恢复默认隐藏'
  toggle.canReset = config => config.options?.waterfall?.connector !== undefined
  const controls: ChartSettingControl[] = [
    colorControl('waterfall.connector.color', '连接线颜色', config => read(config)?.color, (config, value) => { connectorGroup(config).color = value }),
    numberControl('waterfall.connector.width', '连接线宽', config => read(config)?.width, (config, value) => { connectorGroup(config).width = value }, 0, 100, 1),
    numberControl('waterfall.connector.opacity', '连接线不透明度', config => read(config)?.opacity, (config, value) => { connectorGroup(config).opacity = value }, 0, 1, 1),
    selectControl('waterfall.connector.dashType', '连接线型', [['solid', '实线'], ['dashed', '虚线'], ['dotted', '点线']], config => read(config)?.dashType,
      (config, value) => { connectorGroup(config).dashType = value }, 'dashed'),
  ]
  return [toggle, ...controls.map(control => ({ ...control, visible: (config: ChartConfig) => Boolean(effective(config)) }))]
}
