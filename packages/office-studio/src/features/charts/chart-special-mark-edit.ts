import type { ChartLineOptions, ChartModel, ChartParetoBarStyleOptions, ChartParetoLineStyleOptions,
  ChartPieAggregateStyleOptions, ChartSeriesOptions, ChartWaterfallSeriesStyle } from '../../domain/model'
import type { ChartConfig, ChartUpdatePatch } from './chart-engine'
import type { ChartCanvasMarkDirectEditTarget } from './chart-element-interaction'
import { supportsChartSeriesStyleMark } from './chart-element-interaction'
import { chartSeriesOverride } from './chart-series-selector'
import { chartParetoBarControls, chartParetoLineControls } from './chart-pareto-settings-controls'
import { chartWaterfallStyleControls, chartWaterfallStyleGroup } from './chart-waterfall-settings-controls'
import { CHART_WATERFALL_ROLE_NAMES, patchChartWaterfallStyle } from './chart-waterfall'
import { chartPieAggregateControls, chartPieConnectorControls, chartPieSliceControls } from './chart-pie-settings-controls'
import { chartPieSliceAllowsExplosion, isChartCompoundPie } from './chart-pie'
import { cloneChartDraftConfig } from './chart-settings-draft'
import { chartLabelControls, chartSeriesControls, optionGroup, seriesStyleGroup, type ChartSettingControl } from './chart-settings-controls'
import { patchChartSeriesOverride } from './chart-cartesian-api'
import { resolveChartType } from './chart-catalog'

const paretoBarFields = ['color', 'opacity', 'border', 'label'] as const
const paretoLineFields = ['color', 'opacity', 'width', 'dashType', 'point', 'label'] as const
const waterfallFields = ['name', 'color', 'fillOpacity', 'fillPattern', 'border', 'rightAxis', 'label'] as const
const aggregateFields = ['color', 'fillOpacity', 'fillPattern', 'border', 'explosion'] as const
const connectorFields = ['visible', 'color', 'width', 'opacity', 'dashType'] as const
const seriesFields = ['color', 'fillOpacity', 'fillPattern', 'border', 'point', 'label'] as const

function fields<T extends object, K extends keyof T>(value: T | undefined, keys: readonly K[]): Pick<T, K> {
  return Object.fromEntries(keys.map(key => [key, value?.[key]])) as Pick<T, K>
}
function sliceFields(config: ChartConfig, target: Extract<ChartCanvasMarkDirectEditTarget, { kind: 'pieSliceStyle' }>) {
  return chartPieSliceAllowsExplosion(config, target.sourcePointIndex, undefined, target.plot) ? aggregateFields : aggregateFields.filter(key => key !== 'explosion')
}
export function chartSpecialMarkCompatible(config: ChartConfig, target: ChartCanvasMarkDirectEditTarget): boolean {
  if (target.kind === 'seriesStyle') return supportsChartSeriesStyleMark(config.type)
  if (target.kind === 'pieStyle' || target.kind === 'pieSliceStyle') {
    const descriptor = resolveChartType(config.type)
    return descriptor.family === 'pie' && !descriptor.threeD
  }
  const sameType = target.kind === 'paretoStyle' ? config.type === 'pareto'
    : target.kind === 'waterfallStyle' ? config.type === 'waterfall' : isChartCompoundPie(config.type)
  return sameType && (target.kind !== 'waterfallStyle' || !!target.binding
    || target.part === 'subtotal' && config.options?.waterfall?.stackType === 'stacked' && config.options?.waterfall?.useSubtotal !== false)
}
function assertType(config: ChartConfig, target: ChartCanvasMarkDirectEditTarget): void {
  if (target.kind === 'waterfallStyle' && !target.binding
    && (target.part !== 'subtotal' || config.options?.waterfall?.stackType !== 'stacked' || config.options?.waterfall?.useSubtotal === false))
    throw new TypeError('全局瀑布图元仅适用于堆积瀑布图的自动合计。')
  if (!chartSpecialMarkCompatible(config, target)) throw new TypeError('所选专属图元不属于当前图表类型。')
}
function waterfallStyle(config: ChartConfig, target: Extract<ChartCanvasMarkDirectEditTarget, { kind: 'waterfallStyle' }>): ChartWaterfallSeriesStyle | undefined {
  return target.binding ? chartSeriesOverride(config.options, target.binding.selector)?.waterfallStyles?.[target.part]
    : config.options?.waterfall?.styles?.[target.part]
}

export function chartSpecialMarkLabel(target: ChartCanvasMarkDirectEditTarget): string {
  if (target.kind === 'seriesStyle') return `系列标记 · ${target.binding.sourceName}`
  if (target.kind === 'pieStyle') return '饼图数据标签'
  if (target.kind === 'pieSliceStyle') return `饼图来源扇区 ${target.sourcePointIndex + 1} · ${target.sourceName}`
  if (target.kind === 'paretoStyle') return target.part === 'bar' ? '帕累托柱形' : '帕累托累计线'
  if (target.kind === 'waterfallStyle') return target.binding ? `瀑布图${CHART_WATERFALL_ROLE_NAMES[target.part]} · ${target.binding.sourceName}` : '瀑布图全局自动合计'
  return target.kind === 'pieAggregateStyle' ? '复合饼图合计扇区' : '复合饼图主副图连接线'
}
export function chartSpecialMarkScopeNote(target: ChartCanvasMarkDirectEditTarget): string {
  if (target.kind === 'seriesStyle') return '设置应用于所选来源系列的标记、边框和数据标签；独立项目格式保留。'
  if (target.kind === 'pieStyle') return '这些设置应用于全部扇区的数据标签。'
  if (target.kind === 'pieSliceStyle') return '设置应用于所选来源扇区；主图合计与其他扇区格式保留。'
  if (target.kind === 'paretoStyle') return target.part === 'bar' ? '设置应用于帕累托柱形及其数据标签。' : '设置应用于累计百分比线、标记及数据标签。'
  if (target.kind === 'waterfallStyle') return target.binding ? `设置应用于所选来源的“${CHART_WATERFALL_ROLE_NAMES[target.part]}”角色及其数据标签。` : '设置应用于自动生成的全局合计。'
  return target.kind === 'pieAggregateStyle' ? '设置应用于主图合计扇区；合计值由副图数据生成。' : '设置应用于主图与副图之间的两条连接线。'
}
export function chartSpecialMarkControls(target: ChartCanvasMarkDirectEditTarget, config: ChartConfig): ChartSettingControl[] {
  assertType(config, target)
  if (target.kind === 'seriesStyle') {
    const { binding } = target
    return chartSeriesControls(binding.selector, false, false, false,
      { field: binding.field, position: binding.sourcePosition, name: binding.sourceName }, config.type === 'bubble')
      .filter(control => !control.key.endsWith('.name') && !control.key.endsWith('.type') && !control.key.endsWith('.rightAxis'))
      .map(control => !control.key.includes('.point.') ? control : { ...control,
        visible: (item: ChartConfig) => control.visible?.(item) !== false && (item.type === 'boxplot' ? item.options?.boxplot?.showOutliers !== false
          : ['line', 'area', 'scatter', 'bubble', 'radar', 'combination'].includes(resolveChartType(item.type).family)) })
  }
  if (target.kind === 'pieStyle') return chartLabelControls('pie.label', item => item.options?.pie?.label, item => optionGroup(item, 'pie').label ??= {}, true)
  if (target.kind === 'pieSliceStyle') return chartPieSliceControls(target.sourcePointIndex,
    { field: target.sourceField, name: target.sourceName, sourceName: target.sourceName }, chartPieSliceAllowsExplosion(config, target.sourcePointIndex, undefined, target.plot))
  if (target.kind === 'paretoStyle') return target.part === 'bar' ? chartParetoBarControls(target.sourceField === undefined ? undefined
    : { field: target.sourceField, name: target.sourceName ?? '', sourceName: target.sourceName, values: [] }) : chartParetoLineControls()
  if (target.kind === 'waterfallStyle') {
    const binding = target.binding
    return chartWaterfallStyleControls(target.part, binding?.selector, binding ? {
      field: binding.field, name: binding.sourceName, sourceName: binding.sourceName, position: binding.sourcePosition,
    } : undefined)
  }
  return target.kind === 'pieAggregateStyle' ? chartPieAggregateControls({ field: target.sourceField, name: target.sourceName, sourceName: target.sourceName }, target.aggregateIndex)
    : chartPieConnectorControls()
}
export function chartSpecialMarkAuthorSignature(config: ChartConfig, target: ChartCanvasMarkDirectEditTarget): string {
  if (target.kind === 'seriesStyle') return JSON.stringify(fields<ChartSeriesOptions, typeof seriesFields[number]>(chartSeriesOverride(config.options, target.binding.selector), seriesFields))
  if (target.kind === 'pieStyle') return JSON.stringify(config.options?.pie?.label) ?? ''
  if (target.kind === 'pieSliceStyle') return JSON.stringify(fields(config.options?.pie?.sliceStyles?.[target.sourcePointIndex], sliceFields(config, target)))
  if (target.kind === 'paretoStyle') return JSON.stringify(target.part === 'bar'
    ? fields<ChartParetoBarStyleOptions, typeof paretoBarFields[number]>(config.options?.pareto?.barStyle, paretoBarFields)
    : fields<ChartParetoLineStyleOptions, typeof paretoLineFields[number]>(config.options?.pareto?.cumulativeLineStyle, paretoLineFields))
  if (target.kind === 'waterfallStyle') return JSON.stringify(fields(waterfallStyle(config, target), waterfallFields))
  if (target.kind === 'pieAggregateStyle') return JSON.stringify([
    fields<ChartPieAggregateStyleOptions, typeof aggregateFields[number]>(config.options?.pie?.aggregateStyle, aggregateFields), config.options?.pie?.aggregateLabel,
  ])
  return JSON.stringify(fields<ChartLineOptions, typeof connectorFields[number]>(config.options?.pie?.secondaryPlot?.connector, connectorFields))
}

/** Only the displayed author group is copied. In particular, generic series
 * overrides, sibling roles, source values and per-point maps are not replaced
 * by an unrelated dirty field in the draft. */
export function chartSpecialMarkPatch(original: ChartModel, draft: ChartConfig, target: ChartCanvasMarkDirectEditTarget): ChartUpdatePatch {
  assertType(original, target)
  if (draft.type !== original.type) throw new TypeError('图表类型已变更，请重新选择图元。')
  if (chartSpecialMarkAuthorSignature(original, target) === chartSpecialMarkAuthorSignature(draft, target)) return {}
  if (target.kind === 'seriesStyle') return { options: patchChartSeriesOverride(original.options, target.binding.selector,
    fields<ChartSeriesOptions, typeof seriesFields[number]>(chartSeriesOverride(draft.options, target.binding.selector), seriesFields)) }
  if (target.kind === 'pieStyle') return { options: { pie: { label: draft.options?.pie?.label } } }
  if (target.kind === 'pieSliceStyle') return { options: { pie: { sliceStyles: {
    [target.sourcePointIndex]: draft.options?.pie?.sliceStyles?.[target.sourcePointIndex] === undefined ? undefined
      : fields(draft.options.pie.sliceStyles[target.sourcePointIndex], sliceFields(original, target)),
  } } } }
  if (target.kind === 'paretoStyle') return { options: { pareto: target.part === 'bar'
    ? { barStyle: fields<ChartParetoBarStyleOptions, typeof paretoBarFields[number]>(draft.options?.pareto?.barStyle, paretoBarFields) }
    : { cumulativeLineStyle: fields<ChartParetoLineStyleOptions, typeof paretoLineFields[number]>(draft.options?.pareto?.cumulativeLineStyle, paretoLineFields) } } }
  if (target.kind === 'waterfallStyle') {
    const values = fields(waterfallStyle(draft, target), waterfallFields)
    return { options: target.binding ? patchChartWaterfallStyle(original.options, target.binding.selector, target.part, values)
      : { waterfall: { styles: { [target.part]: values } } } }
  }
  if (target.kind === 'pieAggregateStyle') {
    const before = fields<ChartPieAggregateStyleOptions, typeof aggregateFields[number]>(original.options?.pie?.aggregateStyle, aggregateFields)
    const after = fields<ChartPieAggregateStyleOptions, typeof aggregateFields[number]>(draft.options?.pie?.aggregateStyle, aggregateFields)
    return { options: { pie: {
      ...(JSON.stringify(before) !== JSON.stringify(after) ? { aggregateStyle: draft.options?.pie?.aggregateStyle === undefined ? undefined : after } : {}),
      ...(original.options?.pie?.aggregateLabel !== draft.options?.pie?.aggregateLabel ? { aggregateLabel: draft.options?.pie?.aggregateLabel } : {}),
    } } }
  }
  return { options: { pie: { secondaryPlot: { connector: draft.options?.pie?.secondaryPlot?.connector === undefined ? undefined
    : fields<ChartLineOptions, typeof connectorFields[number]>(draft.options.pie.secondaryPlot.connector, connectorFields) } } } }
}
export function resetChartSpecialMarkDraft(config: ChartConfig, target: ChartCanvasMarkDirectEditTarget): ChartConfig {
  assertType(config, target)
  const draft = cloneChartDraftConfig(config)
  if (target.kind === 'seriesStyle') {
    const style = seriesStyleGroup(draft, target.binding.selector)
    for (const key of seriesFields) Object.assign(style, { [key]: undefined })
  } else if (target.kind === 'pieStyle') optionGroup(draft, 'pie').label = undefined
  else if (target.kind === 'pieSliceStyle') {
    const style = draft.options?.pie?.sliceStyles?.[target.sourcePointIndex]
    if (style) for (const key of sliceFields(config, target)) Object.assign(style, { [key]: undefined })
  }
  else if (target.kind === 'paretoStyle') {
    const pareto = optionGroup(draft, 'pareto'), style = target.part === 'bar' ? pareto.barStyle ??= {} : pareto.cumulativeLineStyle ??= {}
    for (const key of target.part === 'bar' ? paretoBarFields : paretoLineFields) Object.assign(style, { [key]: undefined })
  } else if (target.kind === 'waterfallStyle') {
    const style = chartWaterfallStyleGroup(draft, target.part, target.binding?.selector)
    for (const key of waterfallFields) if (key !== 'name') Object.assign(style, { [key]: undefined })
  } else if (target.kind === 'pieAggregateStyle') optionGroup(draft, 'pie').aggregateStyle = undefined
  else if (draft.options?.pie?.secondaryPlot) draft.options.pie.secondaryPlot.connector = undefined
  return draft
}
export function deleteChartSpecialMarkDraft(config: ChartConfig, target: ChartCanvasMarkDirectEditTarget): ChartConfig | undefined {
  assertType(config, target)
  if (target.kind !== 'pieConnectorStyle' && target.kind !== 'pieStyle') return undefined
  const draft = cloneChartDraftConfig(config)
  if (target.kind === 'pieStyle') { (optionGroup(draft, 'pie').label ??= {}).visible = false; return draft }
  ;((optionGroup(draft, 'pie').secondaryPlot ??= {}).connector ??= {}).visible = false
  return draft
}
export function chartSpecialMarkHidden(config: ChartConfig, target: ChartCanvasMarkDirectEditTarget): boolean {
  if (target.kind === 'pieStyle') return config.options?.pie?.label?.visible === false
  const connector = config.options?.pie?.secondaryPlot?.connector
  return target.kind === 'pieConnectorStyle' && (connector?.visible === false || connector?.width === 0 || connector?.opacity === 0)
}
