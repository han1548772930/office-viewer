import type { ChartAxisOptions, ChartModel } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartData, ChartLayout, ChartPoint, ChartSeriesLayout } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { deriveChartData, chartSeriesKind, type ChartComputedStatistics, type ChartDerivedData } from './chart-derived'
import { chartAxisExtent, chartAxisRatio, chartPolar } from './chart-geometry'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chart3DScene } from './chart-3d'
import { chartChordLayout, chartRelationLayout, chartSankeyLayout, chartSunburstLayout, chartTreemapLayout, chartWordCloudLayout } from './chart-special-layout'

export function chartPlotFrame(chart: ChartModel): { frame: RenderRect; plot: RenderRect } {
  const descriptor = resolveChartType(chart.type)
  const axes = ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'pareto', 'heatmap', 'boxplot', 'candlestick', 'histogram', 'radar'].includes(descriptor.family)
  const frame = { x: chart.position.x, y: chart.position.y, width: chart.size.width, height: chart.size.height }
  const topLegend = chart.legend.visible && chart.legend.position === 'top' ? 24 : 0, bottomLegend = chart.legend.visible && chart.legend.position === 'bottom' ? 24 : 0
  const title = chart.title ? Math.max(28, (chart.options?.titleStyle?.fontSize ?? 14) + 14) : 0
  const subtitle = chart.subtitle ? Math.max(18, (chart.options?.subtitleStyle?.fontSize ?? 11) + 7) : 0
  const sideLegend = chart.legend.visible && ['left', 'right'].includes(chart.legend.position) ? Math.min(96, frame.width * 0.24) : 0
  const secondary = chart.type === 'pareto' || Object.values(chart.options?.series ?? {}).some(style => style.rightAxis) || chart.options?.allSeriesStyle?.rightAxis
  const left = (axes ? 54 + (chart.yAxis.title ? 18 : 0) : 14) + (chart.legend.position === 'left' ? sideLegend : 0)
  const right = (secondary ? 54 : 12) + (chart.legend.position === 'right' ? sideLegend : 0)
  const top = 12 + title + subtitle + topLegend, bottom = (axes ? 40 + (chart.xAxis.title ? 16 : 0) : 14) + bottomLegend
  return { frame, plot: { x: frame.x + left, y: frame.y + top, width: Math.max(1, frame.width - left - right), height: Math.max(1, frame.height - top - bottom) } }
}

export function layoutChartData(data: ChartData, chart: ChartModel, computed?: ChartComputedStatistics, palette = chartPalette(chart)): ChartLayout {
  return layoutDerivedChartData(deriveChartData(data, chart, computed), chart, palette)
}

export function layoutDerivedChartData(derived: ChartDerivedData, chart: ChartModel, palette = chartPalette(chart)): ChartLayout {
  const data = derived.data, descriptor = resolveChartType(chart.type), { frame, plot } = chartPlotFrame(chart)
  const horizontal = descriptor.family === 'bar', numericX = descriptor.dataShape === 'xy' || descriptor.dataShape === 'bubble'
  const sourceStyles = data.series.map((series, index) => chartSeriesStyle(chart, series.field, index, palette))
  const kinds = data.series.map((series, index) => chartSeriesKind(chart, series.field, index))
  const positive = new Map<string, number[]>(), negative = new Map<string, number[]>()
  const rawPoints = data.series.map((series, seriesIndex) => series.values.flatMap((value, categoryIndex) => {
    const axis = sourceStyles[seriesIndex]!.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
    const x = numericX ? series.xValues?.[categoryIndex] : categoryIndex
    if (value === null || !Number.isFinite(value) || axis.scaleType === 'log' && value <= 0 || x == null || numericX && (!Number.isFinite(x) || chart.xAxis.scaleType === 'log' && x <= 0)) return []
    const kind = kinds[seriesIndex]!, key = `${sourceStyles[seriesIndex]!.rightAxis ? 'right' : 'left'}:${kind.type}`
    let base = axis.scaleType === 'log' ? axis.min ?? 1 : 0
    if (kind.stacked) {
      const stacks = value >= 0 ? positive : negative, values = stacks.get(key) ?? Array(data.categories.length).fill(0) as number[]
      base = values[categoryIndex]!; values[categoryIndex] = base + value; stacks.set(key, values)
    }
    return [{ value, x, categoryIndex, base, end: kind.stacked ? base + value : value, size: series.sizeValues?.[categoryIndex] }]
  }))
  const primaryValues = rawPoints.flatMap((points, index) => sourceStyles[index]!.rightAxis ? [] : points.flatMap(point => kinds[index]!.stacked ? [point.base, point.end] : [point.end]))
  const rightValues = rawPoints.flatMap((points, index) => !sourceStyles[index]!.rightAxis ? [] : points.flatMap(point => kinds[index]!.stacked ? [point.base, point.end] : [point.end]))
  if (derived.boxes) for (const box of derived.boxes) if (box.statistics) primaryValues.push(box.statistics.min, box.statistics.max, ...(chart.options?.boxplot?.showOutliers === false ? [] : box.statistics.outliers))
  if (data.candles) for (const candle of data.candles) if (candle) primaryValues.push(candle.low, candle.high)
  if (derived.waterfalls) derived.waterfalls.forEach((points, index) => (sourceStyles[index]?.rightAxis ? rightValues : primaryValues).push(...points.flatMap(point => [point.start, point.end])))
  for (const trend of derived.trends) {
    const index = data.series.findIndex(series => series.field === trend.field)
    ;(sourceStyles[index]?.rightAxis ? rightValues : primaryValues).push(...trend.statistics.points.map(point => point[1]))
  }
  for (const error of derived.errors) {
    if (error.config.direction !== 'y') continue
    const index = data.series.findIndex(series => series.field === error.field)
    ;(sourceStyles[index]?.rightAxis ? rightValues : primaryValues).push(...error.intervals.flatMap(point => [point.y - point.minus, point.y + point.plus]))
  }
  const extent = chartAxisExtent(primaryValues, chart.yAxis), rightAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const rightExtent = derived.pareto ? { min: rightAxis.min ?? 0, max: rightAxis.max ?? 100 } : chartAxisExtent(rightValues, rightAxis)
  const xValues = rawPoints.flatMap(points => points.map(point => point.x))
  for (const error of derived.errors) if (error.config.direction === 'x') xValues.push(...error.intervals.flatMap(point => [point.x - point.minus, point.x + point.plus]))
  const xExtent = chartAxisExtent(xValues, chart.xAxis)
  const categoryPosition = (index: number): number => {
    const ratio = (index + 0.5) / Math.max(1, data.categories.length)
    return chart.xAxis.reverse ? 1 - ratio : ratio
  }
  const series: ChartSeriesLayout[] = data.series.map((source, index) => {
    const style = sourceStyles[index]!, axis = style.rightAxis ? rightAxis : chart.yAxis, min = style.rightAxis ? rightExtent.min : extent.min, max = style.rightAxis ? rightExtent.max : extent.max
    const points: ChartPoint[] = rawPoints[index]!.map(point => {
      const valueRatio = chartAxisRatio(point.end, min, max, axis), baseRatio = chartAxisRatio(axis.scaleType === 'log' ? Math.max(min, point.base) : point.base, min, max, axis)
      const category = numericX ? chartAxisRatio(point.x, xExtent.min, xExtent.max, chart.xAxis) : categoryPosition(point.categoryIndex)
      return {
        x: plot.x + (horizontal ? valueRatio : category) * plot.width, y: plot.y + (horizontal ? category : 1 - valueRatio) * plot.height,
        baseY: plot.y + (horizontal ? category : 1 - baseRatio) * plot.height, baseX: plot.x + (horizontal ? baseRatio : category) * plot.width, baseValue: point.base,
        value: point.value, category: data.categories[point.categoryIndex] ?? '', categoryIndex: point.categoryIndex,
        ...(numericX ? { xValue: point.x } : {}), ...(point.size == null ? {} : { size: point.size }),
      }
    })
    return { field: source.field, name: style.name ?? source.name, color: style.color ?? palette[index % palette.length]!, ...kinds[index]!, rightAxis: style.rightAxis, points }
  })
  const layout: ChartLayout = { frame, plot, ...extent, categories: data.categories, series, horizontal,
    ...(numericX ? { xMin: xExtent.min, xMax: xExtent.max } : {}), ...(rightValues.length || derived.pareto ? { rightMin: rightExtent.min, rightMax: rightExtent.max } : {}) }
  if (descriptor.threeD) {
    const scene = chart3DScene(data, chart, layout, undefined, palette)
    layout.series = layout.series.map((series, index) => ({ ...series, points: scene.points[index]! }))
    return layout
  }
  if (descriptor.family === 'radar') {
    layout.series = series.map((series, index) => ({ ...series, points: data.series[index]!.values.flatMap((value, categoryIndex) => value === null ? [] : [{
      ...chartPolar({ x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }, Math.min(plot.width, plot.height) * 0.36 * (value - extent.min) / (extent.max - extent.min), -Math.PI / 2 + categoryIndex * Math.PI * 2 / Math.max(1, data.categories.length)),
      value, category: data.categories[categoryIndex]!, categoryIndex,
    }]) }))
  } else if (descriptor.family === 'relation' || descriptor.family === 'sankey') {
    const graph = descriptor.family === 'sankey' ? chartSankeyLayout(data.graph!, plot) : chartRelationLayout(data.graph!, chart, plot)
    layout.series[0]!.points = graph.nodes.map((node, categoryIndex) => ({ x: node.x + (descriptor.family === 'sankey' ? node.width / 2 : 0), y: node.y + (descriptor.family === 'sankey' ? node.height / 2 : 0), value: node.value, category: node.name, categoryIndex }))
  } else if (descriptor.family === 'chord') {
    const chord = chartChordLayout(data.graph!, chart, plot)
    layout.series[0]!.points = chord.nodes.map((node, categoryIndex) => ({ ...chartPolar(chord.center, (chord.innerRadius + chord.outerRadius) / 2, (node.start + node.end) / 2), value: node.value, category: node.name, categoryIndex }))
  } else if (descriptor.family === 'wordCloud') {
    const words = chartWordCloudLayout(data, chart.options?.wordCloud, plot)
    layout.series[0]!.points = words.map(word => ({ x: word.x + word.width / 2, y: word.y + word.height / 2, value: word.value, category: word.text, categoryIndex: word.categoryIndex, size: word.fontSize }))
  } else if (descriptor.family === 'treemap') {
    const rectangles = chartTreemapLayout(data.hierarchy!, plot, chart.options?.treemap?.parentLabelLayout !== 'overlapping')
    layout.categories = rectangles.map(rect => rect.path.join(' / '))
    layout.series[0]!.points = rectangles.map((rect, categoryIndex) => ({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, value: rect.node.value, category: layout.categories[categoryIndex]!, categoryIndex }))
  } else if (descriptor.family === 'sunburst') {
    const sectors = chartSunburstLayout(data.hierarchy!, plot), center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
    layout.categories = sectors.map(sector => sector.node.name)
    layout.series[0]!.points = sectors.map((sector, categoryIndex) => ({ ...chartPolar(center, (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2), value: sector.node.value, category: sector.node.name, categoryIndex }))
  } else if (descriptor.family === 'waterfall') {
    layout.series = layout.series.map((series, seriesIndex) => ({ ...series, points: derived.waterfalls![seriesIndex]!.map(point => ({
      x: plot.x + categoryPosition(point.categoryIndex) * plot.width, y: plot.y + (1 - chartAxisRatio(point.end, layout.min, layout.max, chart.yAxis)) * plot.height,
      baseY: plot.y + (1 - chartAxisRatio(point.start, layout.min, layout.max, chart.yAxis)) * plot.height, baseValue: point.start,
      category: data.categories[point.categoryIndex]!, categoryIndex: point.categoryIndex, value: point.value,
    })) }))
  }
  return layout
}
