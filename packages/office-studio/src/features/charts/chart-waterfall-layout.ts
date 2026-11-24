import type { ChartAxisOptions, ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout, ChartSeriesLayout, ChartWaterfallPointLayout } from './chart-engine'
import { chartColorRepresentative } from './chart-color'
import { chartAxisExtent, chartAxisRatio } from './chart-geometry'
import { chartSeriesStyle } from './chart-options'
import { CHART_WATERFALL_GLOBAL_SERIES, type WaterfallLayoutData } from './chart-statistics'
import { chartWaterfallRole, chartWaterfallRoleStyles } from './chart-waterfall'

/** Rust owns accumulation and visible/source identities. This adapter only
 * expands labels, resolves axes and maps authoritative tuples to Canvas. */
export function layoutChartWaterfall(data: ChartData, chart: ChartModel, computed: WaterfallLayoutData,
  frame: RenderRect, plot: RenderRect, palette: readonly string[]): ChartLayout {
  const sourceStyles = data.series.map((source, index) => chartSeriesStyle(chart, source.field, index, palette, source.sourceName ?? source.name))
  const roles = data.series.map((source, index) => chartWaterfallRoleStyles(chart, source.field, index, palette, source.sourceName ?? source.name))
  const globalRoles = chartWaterfallRoleStyles(chart, undefined, 0, palette)
  const visibleCount = !data.categories.length ? 0 : computed.stackType === 'stacked' ? data.categories.length + Number(computed.useSubtotal)
    : data.series.length * (data.categories.length + Number(computed.useSubtotal))
  const categories = Array<string>(visibleCount)
  const primaryValues: number[] = [], rightValues: number[] = []
  const rightAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const stylesForPoint = (point: WaterfallLayoutData['points'][number]) => point.seriesPosition === CHART_WATERFALL_GLOBAL_SERIES
    ? globalRoles : roles[point.seriesPosition]!
  const drawable = (start: number, end: number, axis: ChartAxisOptions): boolean => Number.isFinite(start) && Number.isFinite(end)
    && (axis.scaleType !== 'log' || start > 0 && end > 0)
  for (const point of computed.points) {
    const style = stylesForPoint(point)[chartWaterfallRole(point.role, point.value)]
    categories[point.categoryIndex] = point.synthetic ? stylesForPoint(point).subtotal.name ?? '合计' : data.categories[point.dataPointIndex]!
    const axis = style.rightAxis ? rightAxis : chart.yAxis
    if (drawable(point.start, point.end, axis)) (style.rightAxis ? rightValues : primaryValues).push(point.start, point.end)
  }
  const extent = chartAxisExtent(primaryValues, chart.yAxis), rightExtent = chartAxisExtent(rightValues, rightAxis)
  const position = (categoryIndex: number): number => {
    const ratio = (categoryIndex + 0.5) / Math.max(1, categories.length)
    return plot.x + (chart.xAxis.reverse ? 1 - ratio : ratio) * plot.width
  }
  const valuePosition = (value: number, usesRightAxis: boolean): number => {
    const axis = usesRightAxis ? rightAxis : chart.yAxis, bounds = usesRightAxis ? rightExtent : extent
    return plot.y + (1 - chartAxisRatio(value, bounds.min, bounds.max, axis)) * plot.height
  }
  const series: ChartSeriesLayout[] = data.series.map((source, index) => ({ field: source.field, name: sourceStyles[index]!.name ?? source.name,
    color: chartColorRepresentative(roles[index]!.positive.color!), type: 'column', stacked: computed.stackType === 'stacked',
    rightAxis: sourceStyles[index]!.rightAxis, points: [] }))
  const points: ChartWaterfallPointLayout[] = [], actualByCategory = new Map<number, ChartWaterfallPointLayout>()
  for (const source of computed.points) {
    const waterfallRole = chartWaterfallRole(source.role, source.value), style = stylesForPoint(source)[waterfallRole]
    const usesRightAxis = style.rightAxis ?? false, axis = usesRightAxis ? rightAxis : chart.yAxis
    if (!drawable(source.start, source.end, axis)) continue
    const point: ChartWaterfallPointLayout = { x: position(source.categoryIndex), y: valuePosition(source.end, usesRightAxis),
      baseY: valuePosition(source.start, usesRightAxis), baseValue: source.start, value: source.value,
      category: categories[source.categoryIndex]!, categoryIndex: source.categoryIndex, dataPointIndex: source.dataPointIndex,
      seriesPosition: source.seriesPosition, synthetic: source.synthetic, rightAxis: usesRightAxis, waterfallRole }
    points.push(point)
    if (source.seriesPosition !== CHART_WATERFALL_GLOBAL_SERIES) series[source.seriesPosition]!.points.push(point)
    if (!source.synthetic && computed.stackType === 'arrangement') actualByCategory.set(source.categoryIndex, point)
  }
  const connectors: NonNullable<ChartLayout['waterfall']>['connectors'] = []
  for (const source of computed.connectors) {
    if (computed.stackType === 'stacked') {
      // The fixed stacked connector is on the primary axis even when a role
      // uses a secondary axis. Its height is the Rust category cumulative.
      if (!drawable(source.from.value, source.to.value, chart.yAxis)) continue
      connectors.push({ from: { x: position(source.from.categoryIndex), y: valuePosition(source.from.value, false) },
        to: { x: position(source.to.categoryIndex), y: valuePosition(source.to.value, false) } })
    } else {
      const from = actualByCategory.get(source.from.categoryIndex), to = actualByCategory.get(source.to.categoryIndex)
      if (!from || !to) continue
      const fromAxis = from.rightAxis ? rightAxis : chart.yAxis, toAxis = to.rightAxis ? rightAxis : chart.yAxis
      if (!drawable(source.from.value, source.from.value, fromAxis) || !drawable(source.to.value, source.to.value, toAxis)) continue
      connectors.push({ from: { x: from.x, y: valuePosition(source.from.value, from.rightAxis) },
        to: { x: to.x, y: valuePosition(source.to.value, to.rightAxis) } })
    }
  }
  return { frame, plot, ...extent, categories, series, horizontal: false,
    ...(rightValues.length ? { rightMin: rightExtent.min, rightMax: rightExtent.max } : {}), waterfall: { points, connectors } }
}
