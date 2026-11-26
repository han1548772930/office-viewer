import type { ChartModel, ChartSemanticOptions, ChartSeriesSelector, ChartWaterfallStyleTarget } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import type { ChartData } from './chart-engine'
import type { ChartDateAxis } from './chart-date-axis'
import { chartDateSeriesOrder } from './chart-date-axis'
import { chartEditPanelSeriesId, resolveChartEditPanelRoute, type ChartEditPanelAxis, type ChartEditPanelRoute } from './chart-edit-panel-route'
import { chartSeriesStyle } from './chart-options'
import { chartNamedSeriesStyle } from './chart-series-selector'
import { resolveChartType } from './chart-catalog'
import type { ChartDerivedData } from './chart-derived'
import { chartBoxplotAxis, chartBoxplotOutliers, chartBoxplotOutlierStyle, chartBoxplotStyle, chartBoxplotValueDrawable } from './chart-boxplot'
import type { ChartVector } from './chart-geometry'

/** Runtime identities are explicitly mapped to the existing field/name author
 * selectors. Neither an ECharts index nor a renamed display label is an ID. */
export interface ChartCanvasSeriesBinding {
  seriesId: string
  selector: ChartSeriesSelector
  field: number
  sourcePosition: number
  sourceName: string
}
/** Special marks address their dedicated author group. Aggregate identities
 * describe a generated geometry; they never become source dataPoint targets. */
export type ChartCanvasMarkDirectEditTarget =
  | { kind: 'seriesStyle'; binding: ChartCanvasSeriesBinding }
  | { kind: 'pieStyle' }
  | { kind: 'pieSliceStyle'; sourceField: number; sourceName: string; sourcePointIndex: number; plot?: 'primary' | 'secondary' }
  | { kind: 'paretoStyle'; part: 'bar' | 'line'; sourceField?: number; sourceName?: string }
  | { kind: 'waterfallStyle'; part: ChartWaterfallStyleTarget; binding?: ChartCanvasSeriesBinding }
  | { kind: 'pieAggregateStyle'; sourceField: number; sourceName: string; aggregateIndex: number }
  | { kind: 'pieConnectorStyle' }
export type ChartCanvasElement =
  | { kind: 'axisTitle'; axisId: ChartEditPanelAxis }
  | { kind: 'axis'; axisId: ChartEditPanelAxis; part: 'line' | 'label' }
  | { kind: 'gridLine'; axisId: ChartEditPanelAxis }
  | { kind: 'tick'; axisId: ChartEditPanelAxis }
  | { kind: 'legend'; name?: string }
  | { kind: 'series'; binding: ChartCanvasSeriesBinding }
  | { kind: 'dataPoint'; binding: ChartCanvasSeriesBinding; dataIndex: number }
  | { kind: 'dataLabel'; binding: ChartCanvasSeriesBinding; dataIndex: number }
  | { kind: 'mark'; directEdit: ChartCanvasMarkDirectEditTarget }
interface ChartCanvasNativeSourceSeriesBinding {
  runtimeId: string
  binding: ChartCanvasSeriesBinding
  /** Native raw input index -> prepared semantic point index. This includes
   * null slots and preserves the order used to create the actual option. */
  dataIndices: number[]
  seriesOnly?: boolean
  seriesMark?: Extract<ChartCanvasMarkDirectEditTarget, { kind: 'seriesStyle' }>
  pieLabels?: boolean
  /** Prepared source index -> independent author scope. Native raw indices
   * still pass through dataIndices before they address this map. */
  pieSliceMarks?: Partial<Record<number, Extract<ChartCanvasMarkDirectEditTarget, { kind: 'pieSliceStyle' }>>>
}
export type ChartCanvasNativeSeriesBinding = ChartCanvasNativeSourceSeriesBinding
  | { runtimeId: string; directEdit: ChartCanvasMarkDirectEditTarget }
  | { runtimeId: string; items: Array<ChartCanvasElement | undefined> }
export interface ChartCanvasElementRuntime {
  series: ChartCanvasNativeSeriesBinding[]
  axes: Array<{ component: 'xAxis' | 'yAxis'; index: number; axisId: ChartEditPanelAxis }>
}
export interface ChartCanvasElementTarget {
  element: ChartCanvasElement
  key: string
  bounds: RenderRect
}
export type ChartCanvasElementCommand = DrawCommand & { chartCanvasElement: ChartCanvasElement }

export function chartCanvasSeriesBindings(data: Pick<ChartData, 'series'>, options?: ChartSemanticOptions): Array<ChartCanvasSeriesBinding | undefined> {
  const fields = new Map<number, number>(), names = new Map<string, number>()
  for (const source of data.series) {
    fields.set(source.field, (fields.get(source.field) ?? 0) + 1)
    const name = source.sourceName ?? source.name
    names.set(name, (names.get(name) ?? 0) + 1)
  }
  return data.series.map((source, sourcePosition) => {
    const sourceName = source.sourceName ?? source.name
    // An existing exact-name style is the renderer's highest author layer.
    // Direct edits must address it; a field patch below it can have no effect.
    const selector: ChartSeriesSelector | undefined = names.get(sourceName) === 1 && chartNamedSeriesStyle(options, sourceName) ? { name: sourceName }
      : fields.get(source.field) === 1 ? source.field
      : names.get(sourceName) === 1 ? { name: sourceName } : undefined
    return selector === undefined ? undefined : { selector, seriesId: chartEditPanelSeriesId(selector), field: source.field, sourcePosition, sourceName }
  })
}

export function chartCanvasRuntimeSeriesId(chartId: string, sourcePosition: number): string {
  return `office-chart-series:${JSON.stringify([chartId, sourcePosition])}`
}
export function chartCanvasParetoLineRuntimeId(chartId: string): string {
  return `office-chart-pareto-line:${JSON.stringify(chartId)}`
}
export function chartCanvasParetoMark(part: 'bar' | 'line', source?: Pick<ChartData['series'][number], 'field' | 'name' | 'sourceName'>): Extract<ChartCanvasElement, { kind: 'mark' }> {
  return { kind: 'mark', directEdit: { kind: 'paretoStyle', part,
    ...(part === 'bar' && source ? { sourceField: source.field, sourceName: source.sourceName ?? source.name } : {}) } }
}
export function supportsChartSeriesStyleMark(type: ChartModel['type']): boolean {
  const descriptor = resolveChartType(type)
  return !descriptor.threeD && ['line', 'column', 'bar', 'area', 'scatter', 'bubble', 'combination', 'radar', 'boxplot'].includes(descriptor.family)
}
export function chartCanvasSeriesMark(binding: ChartCanvasSeriesBinding): Extract<ChartCanvasElement, { kind: 'mark' }> {
  return { kind: 'mark', directEdit: { kind: 'seriesStyle', binding } }
}
export function chartCanvasPieSliceMark(source: Pick<ChartData['series'][number], 'field' | 'name' | 'sourceName'>,
  sourcePointIndex: number, plot: 'primary' | 'secondary' = 'primary'): Extract<ChartCanvasElement, { kind: 'mark' }> & {
    directEdit: Extract<ChartCanvasMarkDirectEditTarget, { kind: 'pieSliceStyle' }>
  } {
  return { kind: 'mark', directEdit: { kind: 'pieSliceStyle', sourceField: source.field,
    sourceName: source.sourceName ?? source.name, sourcePointIndex, plot } }
}
export function chartCanvasBoxplotRuntimeId(chartId: string, part: 'summary' | 'outliers'): string {
  return `office-chart-boxplot:${JSON.stringify([chartId, part])}`
}
export function chartCanvasBoxplotBindings(boxes: NonNullable<ChartDerivedData['boxes']>, options?: ChartSemanticOptions) {
  return chartCanvasSeriesBindings({ series: boxes.map(box => ({ field: box.field, name: box.name, sourceName: box.sourceName, values: [] })) }, options)
    .map((binding, index) => binding ? { ...binding, sourcePosition: boxes[index]!.seriesPosition } : undefined)
}

/** Shared with the scatter option builder. Both callers use the same
 * filtering predicate and order for painting and runtime point identities. */
export function chartCanvasNativePointOrder(data: ChartData, chart: ChartModel, sourcePosition: number, dates?: ChartDateAxis): number[] {
  const source = data.series[sourcePosition]
  if (!source) return []
  const order = dates ? chartDateSeriesOrder(dates, chart) : source.values.map((_value, index) => index)
  if (chart.type !== 'scatter' && chart.type !== 'bubble') return order
  const style = chartSeriesStyle(chart, source.field, sourcePosition, undefined, source.sourceName ?? source.name)
  const axis = style.rightAxis ? chart.options?.rightYAxis ?? { scaleType: 'linear' } : chart.yAxis
  return order.filter(index => {
    const y = source.values[index], x = dates ? dates.coordinates[index] : source.xValues?.[index], size = source.sizeValues?.[index]
    return y != null && x != null && Number.isFinite(y) && Number.isFinite(x)
      && !(chart.xAxis.scaleType === 'log' && x <= 0) && !(axis.scaleType === 'log' && y <= 0)
      && !(chart.type === 'bubble' && (size == null || size <= 0))
  })
}

export function chartCanvasNativeElementRuntime(data: ChartData, chart: ChartModel, dates?: ChartDateAxis,
  axes?: { xAxis?: unknown; yAxis?: unknown }, derived?: ChartDerivedData): ChartCanvasElementRuntime {
  const bindings = chartCanvasSeriesBindings(data, chart.options), series: ChartCanvasNativeSeriesBinding[] = []
  if (chart.type === 'boxplot' && derived?.boxes) {
    const boxes = derived.boxes, owners = chartCanvasBoxplotBindings(boxes, chart.options), items: Array<ChartCanvasElement | undefined> = []
    series.push({ runtimeId: chartCanvasBoxplotRuntimeId(chart.id, 'summary'), items: owners.map(binding => binding ? chartCanvasSeriesMark(binding) : undefined) })
    if (chart.options?.boxplot?.showOutliers !== false) for (const [index, box] of boxes.entries()) {
      const style = chartBoxplotStyle(chart, box, index, []), axis = chartBoxplotAxis(chart, style), binding = owners[index]
      for (const outlier of chartBoxplotOutliers(box)) {
        const point = chartBoxplotOutlierStyle(style, outlier.sourceIndex)
        if (!chartBoxplotValueDrawable(outlier.value, axis) || (point.point?.size ?? 4) <= 0 || point.point?.shape === 'none') continue
        items.push(binding ? { kind: 'dataPoint', binding, dataIndex: outlier.sourceIndex } : undefined)
      }
    }
    series.push({ runtimeId: chartCanvasBoxplotRuntimeId(chart.id, 'outliers'), items })
  } else if (chart.type === 'pareto') {
    // Both native views use the same dedicated groups as graphic fallback.
    for (let sourcePosition = 0; sourcePosition < data.series.length; sourcePosition++) series.push({
      runtimeId: chartCanvasRuntimeSeriesId(chart.id, sourcePosition), directEdit: chartCanvasParetoMark('bar', data.series[sourcePosition]).directEdit,
    })
    series.push({ runtimeId: chartCanvasParetoLineRuntimeId(chart.id), directEdit: { kind: 'paretoStyle', part: 'line' } })
  } else for (const [sourcePosition, binding] of bindings.entries()) if (binding) {
    const dataIndices = chartCanvasNativePointOrder(data, chart, sourcePosition, dates)
    const descriptor = resolveChartType(chart.type), pieLabels = descriptor.family === 'pie' && !descriptor.threeD
    const pieSliceMarks: ChartCanvasNativeSourceSeriesBinding['pieSliceMarks'] = {}
    if (pieLabels && sourcePosition === 0) for (const sourcePointIndex of dataIndices) {
      if (sourcePointIndex < data.categories.length && chart.options?.pie?.sliceStyles?.[sourcePointIndex] !== undefined)
        pieSliceMarks[sourcePointIndex] = chartCanvasPieSliceMark(data.series[sourcePosition]!, sourcePointIndex).directEdit
    }
    series.push({ runtimeId: chartCanvasRuntimeSeriesId(chart.id, sourcePosition), binding, dataIndices, seriesOnly: chart.type === 'radar',
      ...(supportsChartSeriesStyleMark(chart.type) ? { seriesMark: { kind: 'seriesStyle', binding } as const } : {}),
      ...(pieLabels ? { pieLabels: true } : {}),
      ...(Object.keys(pieSliceMarks).length ? { pieSliceMarks } : {}),
    })
  }
  const result: ChartCanvasElementRuntime = { series, axes: [] }
  const horizontal = chart.type === 'bar' || chart.type.startsWith('barStacked') || chart.type.startsWith('barPercentStacked')
  for (const component of ['xAxis', 'yAxis'] as const) {
    const option = axes?.[component]
    if (!option) continue
    const count = Array.isArray(option) ? option.length : 1
    for (let index = 0; index < count; index++) result.axes.push({ component, index,
      axisId: index === 1 ? 'rightYAxis' : component === 'xAxis' ? horizontal ? 'yAxis' : 'xAxis' : horizontal ? 'xAxis' : 'yAxis' })
  }
  return result
}

export function chartCanvasElementKey(element: ChartCanvasElement): string {
  if (element.kind === 'axisTitle' || element.kind === 'gridLine' || element.kind === 'tick') return `${element.kind}:${element.axisId}`
  if (element.kind === 'axis') return `axis:${element.axisId}:${element.part}`
  if (element.kind === 'legend') return 'legend'
  if (element.kind === 'mark') {
    const target = element.directEdit
    if (target.kind === 'seriesStyle') return `mark:seriesStyle:${target.binding.seriesId}`
    if (target.kind === 'pieSliceStyle') return `mark:pieSliceStyle:${target.sourcePointIndex}`
    if (target.kind === 'paretoStyle') return `mark:pareto:${target.part}`
    if (target.kind === 'waterfallStyle') return `mark:waterfall:${target.part}:${target.binding?.seriesId ?? 'global'}`
    return `mark:${target.kind}`
  }
  return `${element.kind}:${element.binding.seriesId}${element.kind === 'series' ? '' : `:${element.dataIndex}`}`
}
function record(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}
function index(value: unknown): value is number { return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 }
function axisId(value: unknown): value is ChartEditPanelAxis { return value === 'xAxis' || value === 'yAxis' || value === 'rightYAxis' }
function readChartCanvasSeriesBinding(source: unknown): ChartCanvasSeriesBinding | undefined {
  if (!record(source) || !index(source.field) || source.field > 1_048_575 || !index(source.sourcePosition)
    || typeof source.sourceName !== 'string' || typeof source.seriesId !== 'string') return undefined
  const incoming = source.selector
  const selector: ChartSeriesSelector | undefined = index(incoming) && incoming === source.field ? incoming
    : record(incoming) && Object.keys(incoming).length === 1 && typeof incoming.name === 'string' && incoming.name === source.sourceName ? { name: incoming.name } : undefined
  if (selector === undefined || source.seriesId !== chartEditPanelSeriesId(selector)) return undefined
  return { selector, seriesId: source.seriesId, field: source.field, sourcePosition: source.sourcePosition, sourceName: source.sourceName }
}
function readChartCanvasMark(value: unknown): ChartCanvasMarkDirectEditTarget | undefined {
  if (!record(value)) return undefined
  if (value.kind === 'seriesStyle') { const binding = readChartCanvasSeriesBinding(value.binding); return binding ? { kind: value.kind, binding } : undefined }
  if (value.kind === 'pieStyle') return { kind: value.kind }
  if (value.kind === 'pieSliceStyle' && index(value.sourceField) && value.sourceField <= 1_048_575
    && typeof value.sourceName === 'string' && index(value.sourcePointIndex)
    && (value.plot === undefined || value.plot === 'primary' || value.plot === 'secondary'))
    return { kind: value.kind, sourceField: value.sourceField, sourceName: value.sourceName, sourcePointIndex: value.sourcePointIndex,
      ...(value.plot === undefined ? {} : { plot: value.plot }) }
  if (value.kind === 'paretoStyle' && (value.part === 'bar' || value.part === 'line')) {
    if (value.sourceField === undefined && value.sourceName === undefined) return { kind: 'paretoStyle', part: value.part }
    return index(value.sourceField) && value.sourceField <= 1_048_575 && typeof value.sourceName === 'string'
      ? { kind: 'paretoStyle', part: value.part, sourceField: value.sourceField, sourceName: value.sourceName } : undefined
  }
  if (value.kind === 'waterfallStyle' && (value.part === 'positive' || value.part === 'negative' || value.part === 'subtotal')) {
    const binding = readChartCanvasSeriesBinding(value.binding)
    // Only the Rust global synthetic subtotal lacks a business source.
    if (!binding && (value.binding !== undefined || value.part !== 'subtotal')) return undefined
    return { kind: 'waterfallStyle', part: value.part, ...(binding ? { binding } : {}) }
  }
  if (value.kind === 'pieAggregateStyle' && index(value.sourceField) && value.sourceField <= 1_048_575
    && typeof value.sourceName === 'string' && index(value.aggregateIndex) && value.aggregateIndex > 0)
    return { kind: 'pieAggregateStyle', sourceField: value.sourceField, sourceName: value.sourceName, aggregateIndex: value.aggregateIndex }
  return value.kind === 'pieConnectorStyle' ? { kind: 'pieConnectorStyle' } : undefined
}
/** Reconstruct runtime metadata at the renderer boundary. Text, raw series
 * indices and malformed selectors never become author targets. */
export function readChartCanvasElement(value: unknown): ChartCanvasElement | undefined {
  if (!record(value)) return undefined
  const kind = value.kind
  if (kind === 'axisTitle' || kind === 'gridLine' || kind === 'tick') return axisId(value.axisId) ? { kind, axisId: value.axisId } : undefined
  if (kind === 'axis') return axisId(value.axisId) && (value.part === 'line' || value.part === 'label')
    ? { kind, axisId: value.axisId, part: value.part } : undefined
  if (kind === 'legend') return value.name === undefined || typeof value.name === 'string' ? { kind, ...(value.name === undefined ? {} : { name: value.name }) } : undefined
  if (kind === 'mark') { const directEdit = readChartCanvasMark(value.directEdit); return directEdit ? { kind, directEdit } : undefined }
  if (kind !== 'series' && kind !== 'dataPoint' && kind !== 'dataLabel') return undefined
  const binding = readChartCanvasSeriesBinding(value.binding)
  if (!binding) return undefined
  return kind === 'series' ? { kind, binding } : index(value.dataIndex) ? { kind, binding, dataIndex: value.dataIndex } : undefined
}

/** Convex clipping of transformed display bounds against the actual frame or
 * clip-path bounds. Handles rotated rectangles without expanding their hit. */
export function clipChartCanvasPolygon(subject: readonly ChartVector[], clip: readonly ChartVector[]): ChartVector[] {
  if (clip.length < 3 || subject.length === 0) return []
  const area = clip.reduce((sum, point, i) => { const next = clip[(i + 1) % clip.length]!; return sum + point.x * next.y - next.x * point.y }, 0)
  const sign = area < 0 ? -1 : 1
  let result = subject.map(point => ({ ...point }))
  for (let i = 0; i < clip.length && result.length; i++) {
    const start = clip[i]!, end = clip[(i + 1) % clip.length]!, input = result
    const distance = (point: ChartVector) => sign * ((end.x - start.x) * (point.y - start.y) - (end.y - start.y) * (point.x - start.x))
    result = []
    let previous = input[input.length - 1]!, before = distance(previous)
    for (const current of input) {
      const after = distance(current), inside = after >= -1e-8, wasInside = before >= -1e-8
      if (inside !== wasInside) {
        const ratio = before / (before - after)
        result.push({ x: previous.x + (current.x - previous.x) * ratio, y: previous.y + (current.y - previous.y) * ratio })
      }
      if (inside) result.push(current)
      previous = current; before = after
    }
  }
  return result
}
export function chartCanvasElementAsSeries(element: ChartCanvasElement): ChartCanvasElement {
  if (element.kind === 'mark' && element.directEdit.kind === 'seriesStyle') return { kind: 'series', binding: element.directEdit.binding }
  return element.kind === 'dataPoint' || element.kind === 'dataLabel' ? { kind: 'series', binding: element.binding } : element
}
export function chartCanvasElementRoute(element: ChartCanvasElement, type?: ChartModel['type']): ChartEditPanelRoute {
  if (element.kind === 'mark') {
    const target = element.directEdit
    if (target.kind === 'seriesStyle') return { tab: 'style', section: 'series', seriesId: target.binding.seriesId }
    if (target.kind === 'pieStyle') return { tab: 'style', section: 'pieStyle', focus: { kind: 'pieLabel' } }
    if (target.kind === 'pieSliceStyle') return { tab: 'style', section: 'pieStyle', focus: { kind: 'pieSlice', sourcePointIndex: target.sourcePointIndex } }
    if (target.kind === 'paretoStyle') return resolveChartEditPanelRoute({ scope: 'mark', directEdit: target })!
    if (target.kind === 'waterfallStyle') return { ...resolveChartEditPanelRoute({ scope: 'mark', panel: 'waterfallSeries', seriesId: target.binding?.seriesId })!,
      focus: { kind: 'waterfallRole', role: target.part } }
    return { tab: 'style', section: 'pieStyle', focus: { kind: target.kind === 'pieAggregateStyle' ? 'pieAggregate' : 'pieConnector' } }
  }
  if (element.kind === 'axisTitle') return resolveChartEditPanelRoute({ scope: 'axisTitle', axisId: element.axisId,
    titleKey: element.axisId === 'xAxis' ? 'xAxisTitle' : element.axisId === 'yAxis' ? 'yAxisTitle' : 'rightYAxisTitle' })!
  if (element.kind === 'axis' || element.kind === 'gridLine' || element.kind === 'tick') return resolveChartEditPanelRoute({ scope: element.kind, axisId: element.axisId })!
  if (element.kind === 'legend') return { tab: 'style', section: 'legend' }
  if (element.kind === 'dataLabel' && type && resolveChartType(type).family === 'pie') return { tab: 'style', section: 'pieStyle' }
  return resolveChartEditPanelRoute({ scope: element.kind, seriesId: element.binding.seriesId })!
}
export function withChartCanvasElement<T extends DrawCommand>(command: T, element: ChartCanvasElement | undefined): T {
  if (element) Object.assign(command, { chartCanvasElement: element })
  return command
}
export function tagChartCanvasElementCommands(commands: DrawCommand[], start: number, element: ChartCanvasElement | undefined): void {
  if (!element) return
  for (let index = start; index < commands.length; index++) withChartCanvasElement(commands[index]!, element)
}
export function chartCanvasCommandElement(command: DrawCommand): ChartCanvasElement | undefined {
  return 'chartCanvasElement' in command ? readChartCanvasElement(command.chartCanvasElement) : undefined
}
export function chartCanvasElementOutline(target: Pick<ChartCanvasElementTarget, 'bounds'>, position: ChartVector, color: string, width = 1.5): DrawCommand[] {
  const { bounds } = target, x = bounds.x + position.x, y = bounds.y + position.y
  return [{ kind: 'line', x1: x, y1: y, x2: x + bounds.width, y2: y, color, width },
    { kind: 'line', x1: x + bounds.width, y1: y, x2: x + bounds.width, y2: y + bounds.height, color, width },
    { kind: 'line', x1: x + bounds.width, y1: y + bounds.height, x2: x, y2: y + bounds.height, color, width },
    { kind: 'line', x1: x, y1: y + bounds.height, x2: x, y2: y, color, width }]
}
