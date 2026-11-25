import type { ChartAxisOptions, ChartModel, ChartPlotAreaOptions } from '../../domain/model'
import type { RenderRect } from '../../rendering/contracts'
import type { ChartConfig, ChartData, ChartLayout, ChartPoint, ChartSeriesLayout } from './chart-engine'
import { chartHierarchyItems } from './chart-data'
import { resolveChartType } from './chart-catalog'
import { deriveChartData, chartSeriesKind, type ChartComputedStatistics, type ChartDerivedData } from './chart-derived'
import { chartAxisExtent, chartAxisRatio, chartPolar } from './chart-geometry'
import { chartPalette, chartSeriesStyle } from './chart-options'
import { chartColorRepresentative } from './chart-color'
import { CHART_PARETO_CUMULATIVE_NAME, chartParetoBarStyle, chartParetoCumulativeExtent, chartParetoCumulativeStyle } from './chart-pareto'
import { finiteChartValue } from './chart-math'
import { chart3DScene } from './chart-3d'
import { chartFunnelLayout, chartGaugeTiles, chartPieLayout } from './chart-pie-layout'
import { chartGaugeRatio, chartGaugeSweep, resolveChartGaugeOptions } from './chart-gauge'
import { chartRadarExtent, chartRadarGeometry } from './chart-radar'
import { chartBoxplotAxis, chartBoxplotGeometry, chartBoxplotStyle, chartBoxplotSummaryDrawable, chartBoxplotValueDrawable } from './chart-boxplot'
import { chartAxisTitleSpace } from './chart-axis-title'
import { chartManualPlotFrame, chartPlotAreaFromFrame, supportsChartPlotArea, validateChartPlotArea } from './chart-plot-area'
import { chartDateAxis, chartDateCoordinate, chartDateSerial, chartDateSeriesOrder, chartDateTicks } from './chart-date-axis'
import { chartWaterfallUsesSecondaryAxis } from './chart-waterfall'
import { layoutChartWaterfall } from './chart-waterfall-layout'
import { chartHeatmapCellRect, chartHeatmapExtent, chartHeatmapFieldNames, chartHeatmapLegendBand } from './chart-heatmap'
import { chartChordLayout, chartRelationLayout, chartSankeyLayout, chartSunburstLayout, chartTreemapLayout, chartWordCloudLayout } from './chart-special-layout'

export type ChartPlotFrameSource = Pick<ChartConfig, 'type' | 'position' | 'size' | 'title' | 'subtitle' | 'options' | 'legend' | 'xAxis' | 'yAxis'>
export type ChartPlotSourceSeries = readonly { field: number; name: string; sourceName?: string }[]

export function chartPlotFrame(chart: ChartPlotFrameSource, sourceSeries?: ChartPlotSourceSeries): { frame: RenderRect; plot: RenderRect } {
  const descriptor = resolveChartType(chart.type)
  const axes = ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'waterfall', 'pareto', 'heatmap', 'boxplot', 'candlestick', 'histogram', 'radar'].includes(descriptor.family)
  const frame = { x: chart.position.x, y: chart.position.y, width: chart.size.width, height: chart.size.height }
  const manual = chartManualPlotFrame(chart.type, frame, chart.options?.plotArea)
  if (manual) return { frame, plot: manual }
  const legend = { visible: true, position: 'bottom', ...chart.legend, ...(['treemap', 'sunburst', 'chord'].includes(chart.type) ? { visible: false } : {}) }
  const legendBand = chart.type === 'heatmap' ? chartHeatmapLegendBand(frame.width, chart.options?.heatmap?.visualMapType ?? 'continuous', chart.legend?.font?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 11) : 24
  const topLegend = legend.visible && legend.position === 'top' ? legendBand : 0, bottomLegend = legend.visible && legend.position === 'bottom' ? legendBand : 0
  const inheritedSize = chart.options?.appearance?.font?.fontSize
  const title = chart.title ? Math.max(28, (chart.options?.titleStyle?.fontSize ?? inheritedSize ?? 14) + 14) : 0
  const subtitle = chart.subtitle ? Math.max(18, (chart.options?.subtitleStyle?.fontSize ?? inheritedSize ?? 11) + 7) : 0
  const sideLegend = legend.visible && ['left', 'right'].includes(legend.position) ? chart.type === 'heatmap' ? Math.min(132, frame.width * 0.28) : Math.min(96, frame.width * 0.24) : 0
  const secondary = chart.type !== 'heatmap' && (chart.type === 'pareto' || chartWaterfallUsesSecondaryAxis(chart, sourceSeries) || (sourceSeries
    ? sourceSeries.some((series, index) => chartSeriesStyle(chart, series.field, index, undefined, series.sourceName ?? series.name).rightAxis)
    : Object.values(chart.options?.series ?? {}).some(style => style.rightAxis) || chart.options?.namedSeries?.some(entry => entry.style.rightAxis) || chart.options?.allSeriesStyle?.rightAxis))
  const horizontal = descriptor.family === 'bar', sideAxis = (horizontal ? chart.xAxis : chart.yAxis) ?? {}, bottomAxis = (horizontal ? chart.yAxis : chart.xAxis) ?? {}
  const rightAxis = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const left = (axes ? 54 + chartAxisTitleSpace(chart, sideAxis) : 14) + (legend.position === 'left' ? sideLegend : 0)
  const right = (secondary && !horizontal ? 54 + chartAxisTitleSpace(chart, rightAxis) : 12) + (legend.position === 'right' ? sideLegend : 0)
  const top = 12 + title + subtitle + topLegend + (secondary && horizontal ? 40 + chartAxisTitleSpace(chart, rightAxis) : 0)
  const bottom = (axes ? 40 + chartAxisTitleSpace(chart, bottomAxis) : 14) + bottomLegend
  return { frame, plot: { x: frame.x + left, y: frame.y + top, width: Math.max(1, frame.width - left - right), height: Math.max(1, frame.height - top - bottom) } }
}

/** Preserve authored rectangles; automatic seeds include real series/secondary axes. */
export function resolveChartPlotAreaForEditing(chart: ChartPlotFrameSource, sourceSeries?: ChartPlotSourceSeries): ChartPlotAreaOptions | undefined {
  if (!supportsChartPlotArea(chart.type)) return undefined
  if (chart.options?.plotArea) {
    validateChartPlotArea(chart.options.plotArea)
    return { ...chart.options.plotArea }
  }
  const { frame, plot } = chartPlotFrame(chart, sourceSeries)
  return chartPlotAreaFromFrame(frame, plot)
}

export function layoutChartData(data: ChartData, chart: ChartModel, computed?: ChartComputedStatistics, palette = chartPalette(chart)): ChartLayout {
  return layoutDerivedChartData(deriveChartData(data, chart, computed), chart, palette)
}

export function layoutDerivedChartData(derived: ChartDerivedData, chart: ChartModel, palette = chartPalette(chart)): ChartLayout {
  const data = derived.data, descriptor = resolveChartType(chart.type), { frame, plot } = chartPlotFrame(chart, derived.boxes ?? data.series)
  if (derived.waterfallLayout) return layoutChartWaterfall(data, chart, derived.waterfallLayout, frame, plot, palette)
  if (descriptor.family === 'chord') {
    const chord = data.chord ? chartChordLayout(data.chord, chart, plot) : undefined, source = data.series[0]
    if (data.categories.length && source && !chord) throw new TypeError('和弦布局需要同一任务的聚合与角度结果。')
    return { frame, plot, min: 0, max: 1, categories: data.categories, ...(chord ? { chord } : {}),
      series: source ? [{ field: source.field, name: source.name, color: palette[0] ?? '#5470c6',
        points: (chord?.nodes ?? []).map(node => ({ ...chartPolar(chord!.center, (chord!.innerRadius + chord!.outerRadius) / 2, (node.start + node.end) / 2),
          value: node.value, category: node.name, categoryIndex: node.index })) }] : [] }
  }
  if (chart.type === 'heatmap') {
    const extent = chartHeatmapExtent(data, chart), names = chartHeatmapFieldNames(data, chart)
    return { frame, plot, ...extent, categories: data.categories, series: data.series.map((source, seriesPosition) => {
      const style = chartSeriesStyle(chart, source.field, seriesPosition, palette, source.sourceName ?? source.name)
      return { field: source.field, name: names[seriesPosition]!, color: chartColorRepresentative(style.color!), points: source.values.flatMap((value, categoryIndex) => {
        if (value === null || !Number.isFinite(value)) return []
        const rect = chartHeatmapCellRect(data, chart, plot, seriesPosition, categoryIndex)
        return rect ? [{ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, value, category: data.categories[categoryIndex]!, categoryIndex }] : []
      }) }
    }) }
  }
  const dateAxis = chartDateAxis(data, chart)
  const dateOrder = dateAxis ? chartDateSeriesOrder(dateAxis, chart) : undefined
  const dateSequence = dateAxis ? Array<number>(data.categories.length).fill(-1) : undefined
  dateOrder?.forEach((categoryIndex, index) => { dateSequence![categoryIndex] = index })
  const horizontal = descriptor.family === 'bar', numericX = descriptor.dataShape === 'xy' || descriptor.dataShape === 'bubble' || Boolean(dateAxis)
  const sourceStyles = data.series.map((series, index) => chart.type === 'pareto' ? chartParetoBarStyle(chart, series.field, index, palette, series.sourceName ?? series.name)
    : chartSeriesStyle(chart, series.field, index, palette, series.sourceName ?? series.name))
  const kinds = data.series.map((series, index) => chartSeriesKind(chart, series.field, index, series.sourceName ?? series.name))
  const positive = new Map<string, number[]>(), negative = new Map<string, number[]>()
  const rawPoints = data.series.map((series, seriesIndex) => {
    const append = (value: number | null, categoryIndex: number) => {
      const axis = sourceStyles[seriesIndex]!.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
      const x = dateAxis ? dateAxis.coordinates[categoryIndex] : numericX ? series.xValues?.[categoryIndex] : categoryIndex
      if (value === null || !Number.isFinite(value) || axis.scaleType === 'log' && value <= 0 || x == null || numericX && (!Number.isFinite(x) || chart.xAxis.scaleType === 'log' && x <= 0)) return []
      const kind = kinds[seriesIndex]!, key = `${sourceStyles[seriesIndex]!.rightAxis ? 'right' : 'left'}:${kind.type}`
      let base = axis.scaleType === 'log' ? axis.min ?? 1 : 0
      if (kind.stacked) {
        const stacks = value >= 0 ? positive : negative, values = stacks.get(key) ?? Array(data.categories.length).fill(0) as number[]
        base = values[categoryIndex]!; values[categoryIndex] = finiteChartValue(base + value, '堆积合计'); stacks.set(key, values)
      }
      return [{ value, x, categoryIndex, base, end: kind.stacked ? base + value : value, size: series.sizeValues?.[categoryIndex] }]
    }
    return dateOrder ? dateOrder.flatMap(categoryIndex => append(series.values[categoryIndex] ?? null, categoryIndex)) : series.values.flatMap(append)
  })
  const primaryValues = rawPoints.flatMap((points, index) => sourceStyles[index]!.rightAxis ? [] : points.flatMap(point => kinds[index]!.stacked ? [point.base, point.end] : [point.end]))
  const rightValues = rawPoints.flatMap((points, index) => !sourceStyles[index]!.rightAxis ? [] : points.flatMap(point => kinds[index]!.stacked ? [point.base, point.end] : [point.end]))
  if (derived.boxes) {
    primaryValues.length = 0; rightValues.length = 0
    for (const [categoryIndex, box] of derived.boxes.entries()) if (box.statistics) {
      const style = chartBoxplotStyle(chart, box, categoryIndex, palette), values = style.rightAxis ? rightValues : primaryValues, axis = chartBoxplotAxis(chart, style)
      if (chartBoxplotSummaryDrawable(box.statistics, axis)) values.push(box.statistics.min, box.statistics.q1, box.statistics.median, box.statistics.q3, box.statistics.max)
      if (chart.options?.boxplot?.showOutliers !== false) for (const value of box.statistics.outliers) if (chartBoxplotValueDrawable(value, axis)) values.push(value)
    }
  }
  if (data.candles) for (const candle of data.candles) if (candle) primaryValues.push(candle.low, candle.high)
  for (const trend of derived.trends) {
    const values = sourceStyles[trend.seriesPosition]?.rightAxis ? rightValues : primaryValues
    for (const point of trend.statistics.points) values.push(point[1])
  }
  for (const error of derived.errors) {
    if (error.config.direction !== 'y') continue
    const values = sourceStyles[error.seriesPosition]?.rightAxis ? rightValues : primaryValues
    for (const point of error.intervals) values.push(finiteChartValue(point.y - point.minus, '误差下界'), finiteChartValue(point.y + point.plus, '误差上界'))
  }
  const extent = descriptor.family === 'radar' ? chartRadarExtent(data, chart.yAxis) : chartAxisExtent(primaryValues, chart.yAxis), rightAxis: ChartAxisOptions = chart.options?.rightYAxis ?? { scaleType: 'linear' }
  const rightExtent = derived.pareto ? chartParetoCumulativeExtent(chart, derived.pareto.cumulativePercent) : chartAxisExtent(rightValues, rightAxis)
  const xValues = rawPoints.flatMap(points => points.map(point => point.x))
  const dateExtraValues: number[] = []
  for (const error of derived.errors) if (error.config.direction === 'x') {
    for (const point of error.intervals) {
      const lower = finiteChartValue(point.x - point.minus, 'X 误差下界'), upper = finiteChartValue(point.x + point.plus, 'X 误差上界')
      if (dateAxis) dateExtraValues.push(lower, upper)
      else xValues.push(lower, upper)
    }
  }
  const xExtent = dateAxis ? { min: dateAxis.coordinateMin, max: dateAxis.coordinateMax } : chartAxisExtent(xValues, chart.xAxis)
  const dateSourceExtent = dateAxis ? { min: dateAxis.sourceMin, max: dateAxis.sourceMax } : undefined
  if (dateAxis) {
    const includeDate = (value: number): void => {
      if (chartDateSerial(value, data.date1904) === null) throw new RangeError('日期轴的趋势或 X 误差外延超出当前工作簿日期系统。')
      const coordinate = chartDateCoordinate(value, chart.xAxis.baseTimeUnit, dateAxis.date1904)!
      const halfBucket = chart.xAxis.baseTimeUnit === 'months' || chart.xAxis.baseTimeUnit === 'years' ? 0.5 : 0
      if (chart.xAxis.min === undefined) { xExtent.min = Math.min(xExtent.min, coordinate - halfBucket); dateSourceExtent!.min = Math.min(dateSourceExtent!.min, value) }
      if (chart.xAxis.max === undefined) { xExtent.max = Math.max(xExtent.max, coordinate + halfBucket); dateSourceExtent!.max = Math.max(dateSourceExtent!.max, value) }
    }
    for (const value of dateExtraValues) includeDate(value)
    for (const trend of derived.trends) for (const [value] of trend.statistics.points) includeDate(value)
  }
  const categoryPosition = (index: number): number => {
    if (dateAxis) {
      const value = dateAxis.coordinates[index]
      return value == null ? NaN : chartAxisRatio(value, xExtent.min, xExtent.max, chart.xAxis)
    }
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
        ...(numericX ? { xValue: dateAxis ? dateAxis.values[point.categoryIndex]! : point.x } : {}),
        ...(dateSequence ? { sequenceIndex: dateSequence[point.categoryIndex] } : {}), ...(point.size == null ? {} : { size: point.size }),
      }
    })
    return { field: source.field, name: style.name ?? source.name, color: chartColorRepresentative(style.color ?? palette[index % palette.length]!), ...kinds[index]!, rightAxis: style.rightAxis, points }
  })
  const layout: ChartLayout = { frame, plot, ...extent, categories: data.categories, series, horizontal,
    ...(numericX ? { xMin: xExtent.min, xMax: xExtent.max } : {}),
    ...(dateAxis ? { dateAxis: true, date1904: dateAxis.date1904, categoryPositions: data.categories.map((_category, index) => dateAxis.values[index] == null ? null : categoryPosition(index)),
      categoryBandwidth: Math.min(1, dateAxis.coordinateStep / (xExtent.max - xExtent.min)) * (horizontal ? plot.height : plot.width),
      dateOrder, dateSourceMin: dateSourceExtent!.min, dateSourceMax: dateSourceExtent!.max,
      dateTicks: xExtent.min === dateAxis.coordinateMin && xExtent.max === dateAxis.coordinateMax ? dateAxis.ticks
        : chartDateTicks({ ...xExtent, sourceMin: dateSourceExtent!.min, sourceMax: dateSourceExtent!.max, date1904: dateAxis.date1904 }, chart.xAxis) } : {}),
    ...(rightValues.length || derived.pareto ? { rightMin: rightExtent.min, rightMax: rightExtent.max } : {}) }
  if (!data.categories.length || !data.series.length) return layout
  if (derived.boxes) {
    layout.series = derived.boxes.map((box, categoryIndex): ChartSeriesLayout => {
      const style = chartBoxplotStyle(chart, box, categoryIndex, palette), geometry = chartBoxplotGeometry(chart, layout, style, categoryIndex, derived.boxes!.length)
      const summary = geometry.summary(box.statistics)
      return { field: box.field, name: style.name ?? box.name, color: chartColorRepresentative(style.color ?? palette[box.seriesPosition % palette.length]!), type: 'line', stacked: false, rightAxis: style.rightAxis,
        points: box.statistics && summary ? [{ x: geometry.x, y: summary.median, value: box.statistics.median, category: box.name, categoryIndex }] : [] }
    })
    return layout
  }
  if (descriptor.threeD) {
    const scene = chart3DScene(data, chart, layout, undefined, palette, false)
    layout.series = layout.series.map((series, index) => ({ ...series, points: scene.points[index]! }))
    return layout
  }
  if (descriptor.family === 'pie') {
    const pie = chartPieLayout(data, chart, plot), aggregate = pie.sectors.find(sector => sector.part.index >= data.categories.length)
    layout.categories = aggregate ? [...data.categories, aggregate.part.category] : data.categories
    layout.series[0]!.points = [...pie.sectors.map(sector => ({ ...chartPolar(sector.center, (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2),
      value: sector.part.value, category: sector.part.category, categoryIndex: sector.part.index })),
    ...pie.bars.map(bar => ({ x: bar.rect.x + bar.rect.width / 2, y: bar.rect.y + bar.rect.height / 2, value: bar.part.value, category: bar.part.category, categoryIndex: bar.part.index }))]
  } else if (descriptor.family === 'funnel') {
    layout.series[0]!.points = chartFunnelLayout(data, chart, plot).map(step => ({ ...step.center, value: step.value, category: step.category, categoryIndex: step.categoryIndex }))
  } else if (descriptor.family === 'gauge') {
    layout.series[0]!.points = chartGaugeTiles(data, plot).flatMap(tile => {
      const value = data.series[0]!.values[tile.categoryIndex]
      if (value == null || !Number.isFinite(value)) return []
      const options = resolveChartGaugeOptions(chart.options?.gauge, tile.radius), sweep = chartGaugeSweep(options)
      const angle = sweep.start + (sweep.end - sweep.start) * chartGaugeRatio(value, options.min, options.max)
      const pointer = options.pointer.visible && options.pointer.width > 0 && options.pointer.length > 0
      const progress = options.progress.visible && options.progress.width > 0 && sweep.end > sweep.start && chartGaugeRatio(value, options.min, options.max) > 0
      const radius = pointer ? tile.radius * options.pointer.length : progress ? tile.radius - options.progress.width / 2 : 0
      return [{ ...chartPolar(tile.center, radius, angle), value, category: data.categories[tile.categoryIndex]!, categoryIndex: tile.categoryIndex }]
    })
  } else if (descriptor.family === 'radar') {
    const radar = chartRadarGeometry(data, chart, plot)
    layout.series = series.map((series, index) => ({ ...series, points: data.series[index]!.values.flatMap((value, categoryIndex) => value === null ? [] : [{
      ...radar.point(categoryIndex, value),
      value, category: data.categories[categoryIndex]!, categoryIndex,
    }]) }))
  } else if (descriptor.family === 'sankey') {
    const graph = chartSankeyLayout(data.graph!, plot)
    layout.sankey = graph
    layout.series[0]!.points = graph.nodes.map(node => ({ x: node.x + node.width / 2, y: node.y + node.height / 2,
      value: node.value, category: node.name, categoryIndex: node.index }))
  } else if (descriptor.family === 'relation') {
    const graph = chartRelationLayout(data.graph!, chart, plot)
    layout.series[0]!.points = graph.nodes.map(node => ({ x: node.x, y: node.y, value: node.value, category: node.name, categoryIndex: node.index }))
  } else if (descriptor.family === 'wordCloud') {
    const words = chartWordCloudLayout(data, chart.options?.wordCloud, plot)
    layout.series[0]!.points = words.map(word => ({ x: word.x + word.width / 2, y: word.y + word.height / 2, value: word.value, category: word.text, categoryIndex: word.categoryIndex, size: word.fontSize }))
  } else if (descriptor.family === 'treemap') {
    const rectangles = chartTreemapLayout(data.hierarchy!, plot, { parentLabelLayout: chart.options?.treemap?.parentLabelLayout,
      fontSize: chart.options?.treemap?.label?.fontSize ?? chart.options?.appearance?.font?.fontSize ?? 12 })
    layout.treemap = rectangles
    layout.categories = chartHierarchyItems(data.hierarchy!).map(item => item.path.join(' / '))
    layout.series[0]!.points = rectangles.filter(rect => !rect.renderNode.isTreemapResidual).map(rect => ({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, value: rect.node.value, category: layout.categories[rect.pointIndex]!, categoryIndex: rect.pointIndex }))
  } else if (descriptor.family === 'sunburst') {
    const sectors = chartSunburstLayout(data.hierarchy!, plot), center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 }
    layout.sunburst = sectors
    layout.categories = chartHierarchyItems(data.hierarchy!).map(item => item.path.join(' / '))
    layout.series[0]!.points = sectors.filter(sector => !sector.isSunburstResidual).map(sector => ({ ...chartPolar(center, (sector.inner + sector.outer) / 2, (sector.start + sector.end) / 2), value: sector.node.value, category: layout.categories[sector.pointIndex]!, categoryIndex: sector.pointIndex }))
  }
  if (derived.pareto) {
    const style = chartParetoCumulativeStyle(chart, palette)
    layout.series.push({ name: CHART_PARETO_CUMULATIVE_NAME, color: chartColorRepresentative(style.color!), type: 'line', stacked: false, rightAxis: true,
      points: (dateOrder ?? derived.pareto.cumulativePercent.map((_value, index) => index)).flatMap(categoryIndex => {
        const value = derived.pareto!.cumulativePercent[categoryIndex]!
        if (rightAxis.scaleType === 'log' && value <= 0) return []
        const position = categoryPosition(categoryIndex)
        return Number.isFinite(position) ? [{ x: plot.x + position * plot.width,
          y: plot.y + (1 - chartAxisRatio(value, rightExtent.min, rightExtent.max, rightAxis)) * plot.height, value, category: data.categories[categoryIndex]!, categoryIndex,
          ...(dateSequence ? { sequenceIndex: dateSequence[categoryIndex], xValue: dateAxis!.values[categoryIndex]! } : {}) }] : []
      }) })
  }
  return layout
}
