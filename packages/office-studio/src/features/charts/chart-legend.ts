import type { ChartFontOptions, ChartLegendOptions, ChartLegendSelectMode, ChartModel, ChartType } from '../../domain/model'
import type { ChartData, ChartLayout } from './chart-engine'
import type { ChartDerivedData } from './chart-derived'
import type { ChartSankeyScene } from './chart-sankey-draw'
import type { ChartRelationScene } from './chart-relation-draw'
import { resolveChartType } from './chart-catalog'
import { chartColorOpacity } from './chart-geometry'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartPalette, chartSeriesStyle, definedChartStyleFields, validateChartFont } from './chart-options'
import { chartPointStyle } from './chart-derived'
import { chartBoxplotStyle } from './chart-boxplot'
import { chartParetoBarStyle, chartParetoCumulativeStyle, CHART_PARETO_CUMULATIVE_NAME } from './chart-pareto'
import { chartWaterfallRoleStyles } from './chart-waterfall'
import { CHART_WATERFALL_GLOBAL_SERIES } from './chart-statistics'
import { chartTitleGeometry, type ChartTitleGeometrySource } from './chart-titles'

export const ChartLegendSelectModes = Object.freeze({ Single: 'single', Multiple: 'multiple', Close: 'close' } as const)
export const ChartLegendPositions = Object.freeze({ Top: 'top', Left: 'left', Right: 'right', Bottom: 'bottom', Hide: 'hide' } as const)
/** Upstream accepts flat font fields; nested font remains a local authoring form. */
export type ChartLegendSpec = boolean | (Omit<Partial<ChartLegendOptions>, 'position'> & ChartFontOptions & { position?: ChartLegendOptions['position'] | 'hide' })
export type ChartLegendSelection = Readonly<Record<string, boolean>>

const NO_LEGEND_TYPES: readonly ChartType[] = ['treemap', 'sunburst', 'chord', 'wordCloud']
export function chartHasLegend(type: ChartType): boolean { return !NO_LEGEND_TYPES.includes(type) }
export function chartHasLegendSelection(type: ChartType): boolean {
  return chartHasLegend(type) && type !== 'heatmap' && type !== 'pieOfPie' && type !== 'barOfPie'
}
export function chartLegendSelectMode(chart: Pick<ChartModel, 'type' | 'legend'>): ChartLegendSelectMode {
  return chartHasLegendSelection(chart.type) ? chart.legend.selectMode ?? 'close' : 'close'
}

/** Validate before replacing an author candidate; callers never retain a font alias. */
export function chartLegendPatch(current: Partial<ChartLegendOptions> | undefined, value: ChartLegendSpec): Partial<ChartLegendOptions> {
  if (typeof value !== 'boolean' && (!value || typeof value !== 'object' || Array.isArray(value))) throw new TypeError('图例配置须为布尔值或配置对象。')
  const patch: Exclude<ChartLegendSpec, boolean> = typeof value === 'boolean' ? { visible: value } : value
  if (patch.visible !== undefined && typeof patch.visible !== 'boolean') throw new TypeError('图例显隐须为布尔值。')
  if (patch.position !== undefined && !Object.values(ChartLegendPositions).includes(patch.position)) throw new TypeError('图例位置无效。')
  if (patch.selectMode !== undefined && !Object.values(ChartLegendSelectModes).includes(patch.selectMode)) throw new TypeError('图例选择模式无效。')
  if (patch.wrap !== undefined && typeof patch.wrap !== 'boolean') throw new TypeError('图例换行须为布尔值。')
  const { position, font, color, fontSize, bold, italic, family, ...settings } = patch
  validateChartFont(font, '图例字体')
  const flatFont = definedChartStyleFields({ color, fontSize, bold, italic, family })
  validateChartFont(flatFont, '图例字体')
  const result: Partial<ChartLegendOptions> = { ...current, ...settings,
    ...(position === 'hide' ? { visible: false } : position === undefined ? {} : { position }) }
  // A nested font replaces its local object. Flat fields patch that object and
  // take precedence when both forms are supplied; only normalized font persists.
  if (Object.hasOwn(patch, 'font')) result.font = font
  if (Object.keys(flatFont).length) result.font = { ...result.font, ...flatFont }
  if (result.font) result.font = { ...result.font }
  return result
}

/** Selection is name-based, as in the fixed upstream legend. New names start visible;
 * removed names are discarded. Single mode always has exactly one active name. */
export function reconcileChartLegendSelection(names: readonly string[], mode: ChartLegendSelectMode, previous: ChartLegendSelection = {}): Record<string, boolean> {
  const unique = [...new Set(names.filter(name => name !== '' && name !== '\n'))]
  const active = mode === 'single' ? unique.find(name => Object.hasOwn(previous, name) && previous[name] === true) ?? unique[0] : undefined
  return Object.fromEntries(unique.map(name => [name, mode === 'close' || mode === 'single' ? mode === 'close' || name === active
    : !Object.hasOwn(previous, name) || previous[name] !== false]))
}

export function toggleChartLegendSelection(names: readonly string[], mode: ChartLegendSelectMode, previous: ChartLegendSelection, name: string): Record<string, boolean> {
  const selected = reconcileChartLegendSelection(names, mode, previous)
  if (!names.includes(name) || mode === 'close') return selected
  return mode === 'single' ? Object.fromEntries(Object.keys(selected).map(key => [key, key === name]))
    : Object.fromEntries(Object.entries(selected).map(([key, enabled]) => [key, key === name ? !enabled : enabled]))
}

export interface ChartLegendItem {
  name: string
  color: string
  pattern?: ReturnType<typeof chartFillPatternPaint>
}

/** Shared item semantics for Canvas commands and the native legend adapter. */
export function chartLegendItems(derived: ChartDerivedData, chart: ChartModel, layout: ChartLayout, palette: string[] = chartPalette(chart)): ChartLegendItem[] {
  if (!chartHasLegend(chart.type) || chart.type === 'heatmap') return []
  const data = derived.data, family = resolveChartType(chart.type).family
  if (family === 'relation') return (data.graph?.categories ?? []).map((name, index) => {
    const source = data.series[0]
    if (!source) return { name, color: palette[index % palette.length]! }
    const style = chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name)
    return { name, color: chartColorOpacity(style.point?.color ?? style.color!, style.fillOpacity ?? 1), pattern: chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1) }
  })
  if (['pie', 'funnel', 'gauge', 'sankey'].includes(family)) return data.categories.flatMap((name, index) => {
    const source = data.series[0]
    if (!source || family === 'gauge' && !Number.isFinite(source.values[index])) return []
    const style = chartPointStyle(chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name), index, chart.type)
    return [{ name, color: chartColorOpacity(style.color!, style.fillOpacity ?? 1), pattern: chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1) }]
  })
  if (chart.type === 'boxplot') {
    const box = derived.boxes?.[0], source = data.series[0]
    const style = box ? chartBoxplotStyle(chart, box, 0, palette) : chartSeriesStyle(chart, source?.field ?? 0, 0, palette, source?.sourceName ?? source?.name)
    const result: ChartLegendItem[] = [{ name: '样本分布', color: layout.series[0]?.color ?? palette[0]!, pattern: chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 0.45) }]
    if (chart.options?.boxplot?.showOutliers !== false) result.push({ name: '异常值', color: style.point?.color ?? result[0]!.color })
    return result
  }
  if (chart.type === 'waterfall') {
    const totals = new Set(derived.waterfallLayout?.points.filter(point => point.role === 'total' && !point.synthetic).map(point => point.seriesPosition))
    const result = data.series.flatMap((source, index) => {
      const roles = chartWaterfallRoleStyles(chart, source.field, index, palette, source.sourceName ?? source.name)
      const targets = derived.waterfallLayout?.stackType !== 'stacked' && (derived.waterfallLayout?.useSubtotal || totals.has(index))
        ? ['positive', 'negative', 'subtotal'] as const : ['positive', 'negative'] as const
      return targets.map(target => ({ name: roles[target].name ?? source.name, color: chartColorOpacity(roles[target].color!, roles[target].fillOpacity ?? 1), pattern: chartFillPatternPaint(roles[target].fillPattern, roles[target].fillOpacity ?? 1) }))
    })
    if (derived.waterfallLayout?.stackType === 'stacked' && derived.waterfallLayout.useSubtotal) {
      const style = chartWaterfallRoleStyles(chart, undefined, 0, palette).subtotal
      result.push({ name: style.name ?? '合计', color: chartColorOpacity(style.color!, style.fillOpacity ?? 1), pattern: chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1) })
    }
    return result
  }
  return layout.series.map((series, index) => {
    const source = data.series[index], style = source ? chart.type === 'pareto' ? chartParetoBarStyle(chart, source.field, index, palette, source.sourceName ?? source.name)
      : chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name) : chart.type === 'pareto' ? chartParetoCumulativeStyle(chart, palette) : undefined
    return { name: series.name, color: chartColorOpacity(series.color, style?.fillOpacity ?? 1), pattern: chartFillPatternPaint(style?.fillPattern, style?.fillOpacity ?? 1) }
  })
}

interface ChartLegendGeometrySource extends ChartTitleGeometrySource {
  type: ChartType
}
/** Bounded slots make native text truncation and pure Canvas layout agree without
 * DOM/font access in Worker geometry. Overflow keeps every entry reachable by paging. */
export function chartLegendGeometry(chart: ChartLegendGeometrySource, count: number) {
  const horizontal = !chart.legend?.position || chart.legend.position === 'top' || chart.legend.position === 'bottom'
  const fontSize = chart.legend?.font?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 11
  const titles = chartTitleGeometry(chart).bands
  const availableWidth = Math.max(1, chart.size.width - 24 - titles.left - titles.right), itemWidth = Math.max(72, Math.min(140, availableWidth / Math.max(1, count)))
  const perRow = Math.max(1, Math.floor(availableWidth / itemWidth)), rowHeight = fontSize + 8
  const requestedRows = Math.max(1, Math.ceil(count / perRow)), maximumBand = Math.max(rowHeight + 8, chart.size.height * 0.4)
  const wrap = horizontal && chart.type !== 'heatmap' && chart.legend?.wrap === true && requestedRows * rowHeight + 8 <= maximumBand
  return { horizontal, fontSize, itemWidth, perRow, rowHeight, rows: wrap ? requestedRows : 1, wrap, availableWidth, startX: 12 + titles.left,
    band: wrap ? requestedRows * rowHeight + 8 : Math.max(24, rowHeight + 8), textWidth: horizontal ? itemWidth - 23 : Math.max(1, Math.min(96, chart.size.width * 0.24) - 23) }
}

export function chartLegendDataCount(data: ChartData, chart: Pick<ChartModel, 'type'>): number {
  const family = resolveChartType(chart.type).family
  if (family === 'relation') return data.graph?.categories?.length ?? 0
  if (['pie', 'funnel', 'gauge', 'sankey'].includes(family)) return data.categories.length
  return chart.type === 'boxplot' ? 2 : chart.type === 'pareto' ? data.series.length + 1 : chart.type === 'waterfall' ? data.series.length * 3 : data.series.length
}

/** Mask derived views; raw records, source identities, prepared statistics and
 * author options remain owned by their original task/document. No compute job is started. */
export function chartLegendViewData(derived: ChartDerivedData, chart: ChartModel, selected: ChartLegendSelection, palette: string[]): ChartDerivedData {
  if (!Object.values(selected).some(value => value === false)) return derived
  const data = derived.data, family = resolveChartType(chart.type).family
  if (family === 'relation' || family === 'sankey') return derived
  if (chart.type === 'waterfall') return { ...derived, legendSelection: selected }
  const category = ['pie', 'funnel', 'gauge'].includes(family)
  const seriesSelected = data.series.map((source, index) => selected[(chart.type === 'pareto' ? chartParetoBarStyle(chart, source.field, index, palette, source.sourceName ?? source.name)
    : chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name)).name ?? source.name] !== false)
  const view: ChartDerivedData = { ...derived, data: { ...data, series: data.series.map((source, index) => ({ ...source,
    values: source.values.map((value, point) => category ? selected[data.categories[point] ?? ''] === false ? null : value
      : chart.type === 'boxplot' ? selected['样本分布'] === false ? null : value : !seriesSelected[index] ? null : value) })) },
    trends: derived.trends.filter(trend => seriesSelected[trend.seriesPosition]), errors: derived.errors.filter(error => seriesSelected[error.seriesPosition]), legendSelection: selected }
  if (chart.type === 'candlestick' && !seriesSelected[0] && data.candles) view.data.candles = data.candles.map(() => null)
  if (chart.type === 'pareto' && selected[CHART_PARETO_CUMULATIVE_NAME] === false) view.pareto = undefined
  return view
}

/** Waterfall accumulation remains the authoritative cached Rust result. The
 * view removes role marks/adjacent connectors without erasing source axis labels. */
export function chartLegendViewLayout(layout: ChartLayout, derived: ChartDerivedData, chart: ChartModel, selected: ChartLegendSelection, palette: string[]): ChartLayout {
  if (!layout.waterfall || !Object.values(selected).some(value => value === false)) return layout
  const data = derived.data, roles = data.series.map((source, index) => chartWaterfallRoleStyles(chart, source.field, index, palette, source.sourceName ?? source.name))
  const global = chartWaterfallRoleStyles(chart, undefined, 0, palette)
  const points = layout.waterfall.points.filter(point => {
    const group = point.seriesPosition === CHART_WATERFALL_GLOBAL_SERIES ? global : roles[point.seriesPosition]!
    const name = group[point.waterfallRole].name ?? (point.seriesPosition === CHART_WATERFALL_GLOBAL_SERIES ? '合计' : data.series[point.seriesPosition]!.name)
    return selected[name] !== false
  })
  const visible = new Set<object>(points), columns = new Set(points.map(point => point.x))
  return { ...layout, series: layout.series.map(series => ({ ...series, points: series.points.filter(point => visible.has(point)) })),
    waterfall: { points, connectors: layout.waterfall.connectors.filter(connector => columns.has(connector.from.x) && columns.has(connector.to.x)) } }
}

export function chartLegendSankeyScene(scene: ChartSankeyScene, selected: ChartLegendSelection): ChartSankeyScene {
  if (!Object.values(selected).some(value => value === false)) return scene
  const nodes = new Set(scene.layout.nodes.filter(node => selected[node.name] !== false).map(node => node.index))
  const links = new Set(scene.layout.links.filter(link => nodes.has(link.sourceIndex) && nodes.has(link.targetIndex)).map(link => link.index))
  const marks = scene.marks.filter(mark => (mark.kind === 'node' ? nodes : links).has(mark.index))
  return { ...scene, marks, commands: marks.flatMap(mark => mark.commands), layout: { ...scene.layout,
    nodes: scene.layout.nodes.map(node => nodes.has(node.index) ? node : { ...node, width: 0, height: 0 }),
    links: scene.layout.links.map(link => links.has(link.index) ? link : { ...link, width: 0, band: [] }) },
    paint: { ...scene.paint, nodes: scene.paint.nodes.map((paint, index) => nodes.has(index) ? paint : { ...paint, label: { ...paint.label, visible: false } }) } }
}

export function chartLegendRelationScene(scene: ChartRelationScene, selected: ChartLegendSelection): ChartRelationScene {
  if (!Object.values(selected).some(value => value === false)) return scene
  const nodes = new Set(scene.layout.nodes.filter(node => !node.category || selected[node.category] !== false).map(node => node.index))
  const links = new Set(scene.layout.links.filter(link => nodes.has(link.sourceIndex) && nodes.has(link.targetIndex)).map(link => link.index))
  const marks = scene.marks.filter(mark => (mark.kind === 'node' ? nodes : links).has(mark.index))
  return { ...scene, marks, commands: marks.flatMap(mark => mark.commands) }
}
