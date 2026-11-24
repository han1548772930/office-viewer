import type { ChartAxisOptions, ChartLegendOptions, ChartSemanticOptions, ChartSeriesOptions, ChartType } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { chartAutoGradientFill } from './chart-auto-gradient'
import { resolveChartType } from './chart-catalog'
import { chartSeriesColorIsAuthored, chartSeriesStyle, cloneChartOptions, cloneOptionValue } from './chart-options'
import { validateChartSeriesSelector } from './chart-series-selector'
import { chartWaterfallAuthoredStyles } from './chart-waterfall'

/** Description names use canonical catalog IDs; getInfo() keeps local aliases. */
export type ChartDescriptionType = Exclude<ChartType, 'bar-horizontal' | 'doughnut' | 'stacked-column'> | 'donut'
export type ChartDescriptionAxis = 'category' | 'secondaryCategory' | 'primaryValue' | 'secondaryValue'
export interface ChartAxisBindingDescription { dimension: 'horizontal' | 'vertical'; position: 'top' | 'left' | 'right' | 'bottom' }
export interface ChartSeriesDescription extends Omit<ChartSeriesOptions, 'name'> { name: string }
export interface ChartDescription extends Omit<ChartSemanticOptions, 'series' | 'autoGradientFill'> {
  id: string
  type: ChartDescriptionType
  autoGradientFill: boolean
  title?: string
  subtitle?: string
  palette?: string[]
  legend?: Partial<ChartLegendOptions>
  axes?: { x?: Partial<ChartAxisOptions>; y?: Partial<ChartAxisOptions>; rightY?: ChartAxisOptions }
  /** Stable source-field indexes, never positions in a rendered/derived series array. */
  series: Partial<Record<number, ChartSeriesDescription>>
  category?: { name: string }
  axisBindings?: Partial<Record<ChartDescriptionAxis, ChartAxisBindingDescription>>
}
export type ChartDescriptionConfig = Pick<ChartConfig, 'type' | 'options' | 'categoryField' | 'valueFields' | 'title' | 'subtitle' | 'palette' | 'legend' | 'xAxis' | 'yAxis'>
/** The host can provide header/role metadata already known to its parser or source
 * service. No workbook, cells, records, values, layout or rendering is accepted. */
export interface ChartDescriptionContext {
  id: string
  fields?: Readonly<Record<number, string>>
  seriesFields?: readonly number[]
  /** Optional known source identities; a display-name override remains a style. */
  seriesNames?: Readonly<Record<number, string>>
  categoryName?: string
}

const SERIES_FIELDS = ['color', 'fillOpacity', 'fillPattern', 'border', 'point', 'rightAxis', 'label', 'dataPoints'] as const
const AXIS_FAMILIES = ['column', 'bar', 'line', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'pareto', 'heatmap', 'boxplot', 'candlestick', 'histogram']

/** Projects detached semantic authoring and data-free source metadata. An omitted
 * auto plot stays omitted: resolving an editing seed belongs to chart-layout. */
export function describeChartConfiguration(config: ChartDescriptionConfig, context: ChartDescriptionContext): ChartDescription {
  if (!context || typeof context.id !== 'string' || !context.id.trim()) throw new TypeError('图表描述需要有效的对象或草稿标识。')
  const current = descriptionConfig(config), options = current.options ?? {}
  const active = checkedFields(context.seriesFields ?? configuredSeriesFields(current)), activeFields = new Set(active), all = descriptionSeriesFields(current, active)
  const series: ChartDescription['series'] = {}, activeStyles: ChartSeriesOptions[] = []
  for (const [position, field] of all.entries()) {
    const title = metadataName(context.fields, field), knownName = metadataName(context.seriesNames, field)
      ?? (title === undefined ? undefined : title || `Series ${field + 1}`)
    const sourceName = knownName ?? `Series ${field + 1}`
    const style = chartSeriesStyle(current, field, position, config.palette, knownName)
    const entry: ChartSeriesDescription = { name: style.name ?? sourceName }
    for (const key of SERIES_FIELDS) {
      const value = style[key]
      if (value === undefined || ['border', 'point', 'label'].includes(key) && value && typeof value === 'object' && !Object.keys(value).length) continue
      Object.assign(entry, { [key]: cloneOptionValue(value) })
    }
    // Descriptions contain semantic colors, rather than a theme-derived renderer palette.
    if (!config.palette?.length && !chartSeriesColorIsAuthored(current, field, knownName)) delete entry.color
    const type = style.type ?? descriptionSeriesType(config.type, position)
    if (type !== undefined) entry.type = type
    const waterfallStyles = chartWaterfallAuthoredStyles(current, field, knownName)
    if (waterfallStyles !== undefined) entry.waterfallStyles = waterfallStyles
    series[field] = entry
    if (activeFields.has(field)) activeStyles.push(style)
  }
  const description: ChartDescription = { ...options, id: context.id, type: resolveChartType(config.type).id as ChartDescriptionType,
    autoGradientFill: chartAutoGradientFill(current), series }
  for (const key of ['title', 'subtitle', 'palette', 'legend'] as const) if (config[key] !== undefined) Object.assign(description, { [key]: cloneOptionValue(config[key]) })
  if (config.xAxis !== undefined || config.yAxis !== undefined || options.rightYAxis !== undefined) {
    description.axes = { ...(config.xAxis === undefined ? {} : { x: cloneOptionValue(config.xAxis) as Partial<ChartAxisOptions> }),
      ...(config.yAxis === undefined ? {} : { y: cloneOptionValue(config.yAxis) as Partial<ChartAxisOptions> }),
      ...(options.rightYAxis === undefined ? {} : { rightY: cloneOptionValue(options.rightYAxis) as ChartAxisOptions }) }
  }
  const category = descriptionCategoryName(current, context)
  if (category !== undefined) description.category = { name: category }
  const bindings = descriptionAxisBindings(current, activeStyles)
  if (bindings) description.axisBindings = bindings
  return description
}

/** Lets a host fetch just the necessary source titles. Numeric role inference must
 * be supplied as metadata; this function never samples or scans source values. */
export function chartDescriptionFieldIndexes(config: ChartDescriptionConfig, seriesFields?: readonly number[]): number[] {
  const current = descriptionConfig(config)
  return [...new Set([...configuredCategoryFields(current), ...descriptionSeriesFields(current, checkedFields(seriesFields ?? configuredSeriesFields(current)))])]
}

function descriptionConfig(config: ChartDescriptionConfig): ChartDescriptionConfig {
  const categoryField = config.categoryField
  if (categoryField !== undefined) checkedFields([categoryField])
  return { type: config.type, options: cloneChartOptions(config.options), categoryField,
    valueFields: config.valueFields === undefined ? undefined : checkedFields(config.valueFields), palette: config.palette }
}

function configuredSeriesFields(config: ChartDescriptionConfig): number[] {
  const descriptor = resolveChartType(config.type), mapping = config.options?.mapping, fields = config.valueFields ?? []
  if (descriptor.family === 'bubble') return mapping?.bubble ? [mapping.bubble.yIndex] : fields.slice(0, 1)
  if (descriptor.family === 'gauge') return mapping?.gauge ? [mapping.gauge.valueIndex] : fields.slice(0, 1)
  if (descriptor.family === 'candlestick') return mapping?.candlestick ? [mapping.candlestick.closeIndex] : fields.length >= 4 ? [fields[3]!] : fields.length >= 3 ? [fields[2]!] : []
  if (descriptor.dataShape === 'edge') return mapping?.edge ? [mapping.edge.valueIndex ?? mapping.edge.targetIndex] : fields.length ? [fields[1] ?? fields[0]!] : []
  if (descriptor.dataShape === 'hierarchy' || ['pie', 'funnel', 'wordCloud', 'pareto', 'histogram'].includes(descriptor.family)) return fields.slice(0, 1)
  return fields
}

function configuredCategoryFields(config: ChartDescriptionConfig): number[] {
  const descriptor = resolveChartType(config.type), mapping = config.options?.mapping
  if (descriptor.dataShape === 'distribution') return []
  if (descriptor.family === 'bubble' && mapping?.bubble) return [mapping.bubble.xIndex]
  if (descriptor.family === 'gauge' && mapping?.gauge) return mapping.gauge.labelIndex === undefined ? [] : [mapping.gauge.labelIndex]
  if (descriptor.dataShape === 'edge' && mapping?.edge) return checkedFields([mapping.edge.sourceIndex, mapping.edge.targetIndex])
  return checkedFields(descriptor.dataShape === 'hierarchy' ? mapping?.hierarchyIndexes ?? [config.categoryField ?? 0] : mapping?.categoryIndexes ?? [config.categoryField ?? 0])
}

function descriptionSeriesFields(config: ChartDescriptionConfig, active: readonly number[]): number[] {
  // Inactive authored field overrides survive a source/type switch. Named
  // overrides remain in namedSeries until source identities are known.
  const authored = Object.keys(config.options?.series ?? {}).map(Number)
  return [...new Set([...active, ...checkedFields(authored)])]
}

function checkedFields(fields: readonly number[]): number[] {
  if (!Array.isArray(fields)) throw new TypeError('图表描述系列字段需要密集索引数组。')
  const checked: number[] = [], seen = new Set<number>()
  for (let index = 0; index < fields.length; index++) {
    if (!Object.hasOwn(fields, index)) throw new TypeError('图表描述系列字段不能含稀疏位置。')
    const field = fields[index]!
    if (typeof field !== 'number') throw new TypeError('图表描述源字段需要数值索引。')
    validateChartSeriesSelector(field)
    if (seen.has(field)) throw new RangeError('图表描述源字段索引不能重复。')
    checked.push(field); seen.add(field)
  }
  return checked
}

function metadataName(names: Readonly<Record<number, string>> | undefined, field: number): string | undefined {
  if (names === undefined) return undefined
  if (!names || typeof names !== 'object') throw new TypeError('图表描述字段名称需要索引记录。')
  if (!Object.hasOwn(names, field)) return undefined
  const name = names[field]
  if (typeof name !== 'string') throw new TypeError('图表描述字段名称需要字符串。')
  return name
}

function descriptionCategoryName(config: ChartDescriptionConfig, context: ChartDescriptionContext): string | undefined {
  if (context.categoryName !== undefined) {
    if (typeof context.categoryName !== 'string') throw new TypeError('图表描述分类名称需要字符串。')
    return context.categoryName
  }
  const fields = configuredCategoryFields(config), names = fields.map(field => metadataName(context.fields, field))
  return fields.length && names.every((name): name is string => name !== undefined) ? names.join(' / ') : undefined
}

function descriptionSeriesType(type: ChartType, position: number): ChartSeriesOptions['type'] {
  const descriptor = resolveChartType(type)
  if (descriptor.family === 'combination') return position === 0 ? 'column' : 'line'
  if (['column', 'bar', 'histogram', 'pareto', 'waterfall'].includes(descriptor.family)) return descriptor.stacked ? 'columnStacked' : 'column'
  if (descriptor.family === 'area') return descriptor.stacked ? 'areaStacked' : 'area'
  return descriptor.family === 'line' ? 'line' : undefined
}

function descriptionAxisBindings(config: ChartDescriptionConfig, active: readonly ChartSeriesOptions[]): ChartDescription['axisBindings'] {
  const family = resolveChartType(config.type).family
  // Radar axes are radial; an invented horizontal/vertical edge would be misleading.
  if (!AXIS_FAMILIES.includes(family)) return undefined
  if (family === 'heatmap') return { category: { dimension: 'horizontal', position: 'bottom' }, secondaryCategory: { dimension: 'vertical', position: 'left' } }
  const horizontal = family === 'bar'
  const bindings: ChartDescription['axisBindings'] = { category: { dimension: horizontal ? 'vertical' : 'horizontal', position: horizontal ? 'left' : 'bottom' },
    primaryValue: { dimension: horizontal ? 'horizontal' : 'vertical', position: horizontal ? 'bottom' : 'left' } }
  if (family === 'pareto' || active.some(style => style.rightAxis)) bindings.secondaryValue = { dimension: horizontal ? 'horizontal' : 'vertical', position: horizontal ? 'top' : 'right' }
  return bindings
}
