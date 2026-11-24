import type { CellScalar, ChartModel, ChartSemanticOptions } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { finiteChartValue } from './chart-math'
import { assertChartPreparationMemory, CHART_PREPARATION_CHUNK, consumeChartPreparationSync,
  type ChartComputePreparationContext } from './chart-compute-lifecycle'
import { chartMapSteps, chartNumericFieldSteps, chartProjectionSourceBytes } from './chart-data-preparation'
import { chartHierarchyChartDataSteps, chartHierarchyResultSteps, chartHierarchyValueField } from './chart-hierarchy-data'
export { chartHierarchyData, chartHierarchyDataSteps, chartHierarchyValueField } from './chart-hierarchy-data'

export interface ChartRecords {
  fields: string[]
  records: CellScalar[][]
  /** Optional display text aligned to the exact source record/field, separate from numeric inputs. */
  formattedRecords?: Array<Array<string | null>>
  /** Runtime date-system owner; this is not persisted in chart options. */
  date1904?: boolean
  sourceCategoryFields?: number[]
  sourceValueFields?: number[]
  /** Detected numeric columns in the current source records, independent of
   * pivot-declared roles. Prepared settings do not need to rescan all rows. */
  sourceNumericFields?: number[]
  /** Original field indices allowed by plotVisibleOnly; record axes are already filtered. */
  visibleFields?: number[]
  /** Pivot values placed on rows retain the legacy value-field zero shorthand. */
  syntheticValues?: boolean
}
export type ChartDataBinding = Pick<ChartModel, 'type' | 'categoryField' | 'valueFields' | 'seriesToCategory' | 'options'>
export interface ChartEdge { source: string; target: string; value: number }
export interface ChartGraphData { nodes: Array<{ name: string; value: number }>; links: ChartEdge[]; cyclic: boolean }
export interface ChartHierarchyNode {
  id: string; name: string; path: string[]; value: number; selfValue: number
  formattedValue?: string; formattedSelfValue?: string; children?: ChartHierarchyNode[]
  /** Runtime root palette identity, independent of weight-sorted display order. */
  paletteIndex?: number
}
export interface ChartHierarchyItem { node: ChartHierarchyNode; pointIndex: number; rootIndex: number; depth: number; path: string[] }

/** Source hierarchy order gives every root and descendant one stable point index. */
export function chartHierarchyItems(nodes: ChartHierarchyNode[]): ChartHierarchyItem[] {
  const items: ChartHierarchyItem[] = [], stack = nodes.map((node, rootIndex) => ({ node, rootIndex: node.paletteIndex ?? rootIndex, depth: 0, path: node.path ?? [node.name] })).reverse()
  while (stack.length) {
    const current = stack.pop()!
    items.push({ ...current, pointIndex: items.length })
    for (const child of [...(current.node.children ?? [])].reverse()) stack.push({ node: child, rootIndex: current.rootIndex, depth: current.depth + 1, path: [...current.path, child.name] })
  }
  return items
}
export interface ChartCandle { open?: number; high: number; low: number; close: number; categoryIndex: number }

export function chartNumber(value: CellScalar | undefined): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'string' && value.trim() && Number.isFinite(Number(value))) return Number(value)
  return null
}

export function chartField(recordSet: ChartRecords, field: number, label: string): void {
  if (!Number.isSafeInteger(field) || field < 0 || field >= recordSet.fields.length) throw new RangeError(`${label}必须位于图表数据区域内。`)
}

/** Pivot output declares value roles explicitly: numeric-looking category IDs
 * remain category fields. Ordinary sources preserve null/invalid sample rules. */
export function chartValueField(recordSet: ChartRecords, field: number, label: string): void {
  chartField(recordSet, field, label)
  if (recordSet.sourceValueFields !== undefined && !recordSet.sourceValueFields.includes(field)) throw new RangeError(`${label}须使用透视数据源声明的数值字段。`)
}

export function readSpecialChartData(source: ChartRecords, chart: ChartDataBinding): ChartData | undefined {
  return consumeChartPreparationSync(readSpecialChartDataSteps(source, chart))
}

export function* readSpecialChartDataSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartData | undefined, void> {
  const shape = resolveChartType(chart.type).dataShape
  if (shape === 'bubble') {
    const mapping = chartBubbleMapping(source, chart)
    const groups = new Map<string, { field: number; name: string; values: Array<number | null>; xValues: Array<number | null>; sizeValues: Array<number | null> }>()
    const categories = yield* chartMapSteps(source.records, (record, index) => String(record[mapping.labelIndex ?? mapping.xIndex] ?? index + 1))
    for (let index = 0; index < source.records.length; index++) {
      const record = source.records[index]!
      const group = mapping.seriesIndex === undefined ? source.fields[mapping.yIndex]! : String(record[mapping.seriesIndex] ?? '')
      let series = groups.get(group)
      if (!series) {
        // Grouped bubbles use aligned series arrays. Reject their actual
        // quadratic peak before allocating a new group, even for sparse rows.
        assertChartPreparationMemory(chartProjectionSourceBytes(source.records.length, source.fields.length)
          + (groups.size + 1) * (512 + source.records.length * 48), context)
        const values = yield* chartMapSteps(source.records, (): number | null => null)
        const xValues = yield* chartMapSteps(source.records, (): number | null => null)
        const sizeValues = yield* chartMapSteps(source.records, (): number | null => null)
        series = { field: mapping.yIndex, name: group, values, xValues, sizeValues }
        groups.set(group, series)
      }
      series.values[index] = chartNumber(record[mapping.yIndex])
      series.xValues[index] = chartNumber(record[mapping.xIndex])
      series.sizeValues[index] = mapping.sizeIndex === undefined ? 1 : chartNumber(record[mapping.sizeIndex])
      if ((index + 1) % CHART_PREPARATION_CHUNK === 0) yield
    }
    const categoryValues = yield* chartMapSteps(source.records, record => record[mapping.xIndex] ?? null)
    return { ...source, categories, categoryValues, series: [...groups.values()] }
  }
  if (shape === 'edge') {
    const graph = yield* chartGraphDataSteps(source, chart, context)
    const mapping = chartEdgeMapping(source, chart), field = mapping.valueIndex ?? mapping.targetIndex
    const categories = yield* chartMapSteps(graph.nodes, node => node.name)
    const values = yield* chartMapSteps(graph.nodes, node => node.value)
    return { ...source, categories, series: [{ field,
      name: mapping.valueIndex === undefined ? '关系次数' : source.fields[field] ?? '关系权重', values }], graph }
  }
  if (shape === 'ohlc') {
    const candleFields = chartCandlestickMapping(source, chart)
    const candles = yield* chartCandlestickDataSteps(source, chart)
    const categories = yield* chartMapSteps(source.records, row => String(row[chart.categoryField] ?? ''))
    const categoryValues = yield* chartMapSteps(source.records, row => row[chart.categoryField] ?? null)
    const values = yield* chartMapSteps(candles, candle => candle?.close ?? null)
    const fallbackName = candleFields.openIndex === undefined ? '最高 / 最低 / 收盘' : '开盘 / 最高 / 最低 / 收盘'
    const name = source.fields[candleFields.closeIndex] || fallbackName
    return { ...source, categories, categoryValues, series: [{ field: candleFields.closeIndex, name, sourceName: name, values }], candles, candleFields: { ...candleFields } }
  }
  if (shape === 'hierarchy') {
    const valueField = chartHierarchyValueField(source, chart)
    const result = yield* chartHierarchyResultSteps(source, chart, context)
    return yield* chartHierarchyChartDataSteps(source, result, valueField)
  }
  if (shape === 'gauge') {
    const mapping = chart.options?.mapping?.gauge ?? { valueIndex: chart.valueFields[0] ?? (source.fields.length > 1 ? 1 : 0), labelIndex: source.fields.length > 1 ? chart.categoryField : undefined }
    chartValueField(source, mapping.valueIndex, '仪表数值字段')
    if (mapping.labelIndex !== undefined) chartField(source, mapping.labelIndex, '仪表标签字段')
    const categories = yield* chartMapSteps(source.records, (row, index) => String(mapping.labelIndex === undefined ? source.fields[mapping.valueIndex] : row[mapping.labelIndex] ?? index + 1))
    const values = yield* chartMapSteps(source.records, row => chartNumber(row[mapping.valueIndex]))
    return { ...source, categories, series: [{ field: mapping.valueIndex, name: source.fields[mapping.valueIndex]!, values }] }
  }
  if (shape === 'distribution') {
    const numericFields = source.sourceNumericFields ?? (yield* chartNumericFieldSteps(source.records, source.fields.length, chartNumber))
    const fields = chart.valueFields.length ? chart.valueFields : chart.type === 'histogram' ? numericFields.slice(0, 1) : numericFields
    fields.forEach(field => chartValueField(source, field, '样本字段'))
    const categories = yield* chartMapSteps(source.records, (_row, index) => String(index + 1))
    const series: ChartData['series'] = []
    for (const field of fields) {
      const values = yield* chartMapSteps(source.records, row => chartNumber(row[field]))
      series.push({ field, name: source.fields[field]!, values })
    }
    return { ...source, categories, series }
  }
  return undefined
}

export function chartBubbleMapping(source: ChartRecords, chart: ChartDataBinding): NonNullable<NonNullable<ChartSemanticOptions['mapping']>['bubble']> {
  const values = chart.valueFields.length ? chart.valueFields : source.fields.map((_field, index) => index).filter(index => index !== chart.categoryField)
  const mapping = chart.options?.mapping?.bubble ?? { xIndex: chart.categoryField, yIndex: values[0] ?? 1, sizeIndex: values[1] }
  for (const [name, field] of Object.entries(mapping)) if (field !== undefined) chartField(source, field, `气泡${name}`)
  chartValueField(source, mapping.xIndex, '气泡 X 字段'); chartValueField(source, mapping.yIndex, '气泡 Y 字段')
  if (mapping.sizeIndex !== undefined) chartValueField(source, mapping.sizeIndex, '气泡大小字段')
  return mapping
}

export function chartEdgeMapping(source: ChartRecords, chart: ChartDataBinding): NonNullable<NonNullable<ChartSemanticOptions['mapping']>['edge']> {
  const mapping = chart.options?.mapping?.edge ?? { sourceIndex: chart.categoryField, targetIndex: chart.valueFields[0] ?? 1, valueIndex: chart.valueFields[1] ?? (source.fields.length > 2 ? 2 : undefined) }
  chartField(source, mapping.sourceIndex, '来源字段'); chartField(source, mapping.targetIndex, '目标字段')
  if (mapping.sourceIndex === mapping.targetIndex) throw new RangeError('来源和目标须使用不同字段。')
  if (mapping.valueIndex !== undefined) chartValueField(source, mapping.valueIndex, '权重字段')
  return mapping
}

export function chartGraphData(source: ChartRecords, chart: ChartDataBinding): ChartGraphData {
  return consumeChartPreparationSync(chartGraphDataSteps(source, chart))
}

export function* chartGraphDataSteps(source: ChartRecords, chart: ChartDataBinding,
  context: ChartComputePreparationContext = {}): Generator<void, ChartGraphData, void> {
  const mapping = chartEdgeMapping(source, chart)
  const weights = new Map<string, number>(), links = new Map<string, ChartEdge>()
  const flows = chart.type === 'sankey' ? new Map<string, { incoming: number; outgoing: number }>() : undefined
  const undirected = chart.type === 'chord'
  const sourceBytes = chartProjectionSourceBytes(source.records.length, source.fields.length)
  let work = 0, nameBytes = 0
  for (const row of source.records) {
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
    const first = String(row[mapping.sourceIndex] ?? '').trim(), second = String(row[mapping.targetIndex] ?? '').trim()
    if (!first || !second) continue
    const value = mapping.valueIndex === undefined ? 1 : chartNumber(row[mapping.valueIndex])
    if (value === null || value === 0) continue
    if (value < 0) throw new RangeError('关系图的权重不能为负数。')
    if (!weights.has(first)) nameBytes += first.length * 2
    if (!weights.has(second)) nameBytes += second.length * 2
    const [sourceName, targetName] = undirected && first > second ? [second, first] : [first, second]
    const key = JSON.stringify([sourceName, targetName])
    const link = links.get(key) ?? { source: sourceName, target: targetName, value: 0 }
    link.value = finiteChartValue(link.value + value, '关系边权重合计'); links.set(key, link)
    if (flows) {
      const sourceFlow = flows.get(first) ?? { incoming: 0, outgoing: 0 }
      sourceFlow.outgoing = finiteChartValue(sourceFlow.outgoing + value, '桑基节点出流合计'); flows.set(first, sourceFlow)
      const targetFlow = flows.get(second) ?? { incoming: 0, outgoing: 0 }
      targetFlow.incoming = finiteChartValue(targetFlow.incoming + value, '桑基节点入流合计'); flows.set(second, targetFlow)
      // A transit node carries one flow. Adding its input and output would
      // count the same quantity twice and may overflow a balanced finite flow.
      weights.set(first, Math.max(sourceFlow.incoming, sourceFlow.outgoing))
      weights.set(second, Math.max(targetFlow.incoming, targetFlow.outgoing))
    } else {
      weights.set(first, finiteChartValue((weights.get(first) ?? 0) + value, '关系节点权重合计'))
      weights.set(second, finiteChartValue((weights.get(second) ?? 0) + value, '关系节点权重合计'))
    }
    assertChartPreparationMemory(sourceBytes + weights.size * 384 + links.size * 256 + nameBytes, context)
  }
  const nodes: ChartGraphData['nodes'] = [], edges: ChartEdge[] = []
  const degrees = new Map<string, number>(), outgoing = new Map<string, string[]>()
  for (const [name, value] of weights) {
    nodes.push({ name, value }); degrees.set(name, 0)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  for (const edge of links.values()) { edges.push(edge); if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  for (const link of edges) {
    degrees.set(link.target, (degrees.get(link.target) ?? 0) + 1)
    const list = outgoing.get(link.source) ?? []; list.push(link.target); outgoing.set(link.source, list)
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  const queue: string[] = []
  for (const node of nodes) { if (degrees.get(node.name) === 0) queue.push(node.name); if (++work % CHART_PREPARATION_CHUNK === 0) yield }
  let visited = 0
  for (let index = 0; index < queue.length; index++) {
    visited++
    for (const target of outgoing.get(queue[index]!) ?? []) {
      const degree = degrees.get(target)! - 1; degrees.set(target, degree); if (degree === 0) queue.push(target)
      if (++work % CHART_PREPARATION_CHUNK === 0) yield
    }
    if (++work % CHART_PREPARATION_CHUNK === 0) yield
  }
  return { nodes, links: edges, cyclic: visited !== nodes.length }
}

export function chartCandlestickMapping(source: ChartRecords, chart: ChartDataBinding): NonNullable<NonNullable<ChartSemanticOptions['mapping']>['candlestick']> {
  const fields = chart.valueFields.length ? chart.valueFields : source.fields.map((_name, index) => index).filter(index => index !== chart.categoryField)
  const mapping = chart.options?.mapping?.candlestick ?? (fields.length >= 4
    ? { openIndex: fields[0], highIndex: fields[1]!, lowIndex: fields[2]!, closeIndex: fields[3]! }
    : { highIndex: fields[0]!, lowIndex: fields[1]!, closeIndex: fields[2]! })
  for (const [name, field] of Object.entries(mapping)) if (field !== undefined) chartValueField(source, field, name)
  const roles = Object.values(mapping).filter((field): field is number => field !== undefined)
  if (roles.length < 3 || new Set(roles).size !== roles.length) throw new RangeError('K 线须选择不同的最高、最低、收盘字段，开盘字段可选。')
  return mapping
}

export function chartCandlestickData(source: ChartRecords, chart: ChartDataBinding): Array<ChartCandle | null> {
  return consumeChartPreparationSync(chartCandlestickDataSteps(source, chart))
}

export function* chartCandlestickDataSteps(source: ChartRecords, chart: ChartDataBinding): Generator<void, Array<ChartCandle | null>, void> {
  const mapping = chartCandlestickMapping(source, chart)
  return yield* chartMapSteps(source.records, (row, categoryIndex): ChartCandle | null => {
    const high = chartNumber(row[mapping.highIndex]), low = chartNumber(row[mapping.lowIndex]), close = chartNumber(row[mapping.closeIndex])
    const open = mapping.openIndex === undefined ? undefined : chartNumber(row[mapping.openIndex])
    if (high === null || low === null || close === null || open === null) return null
    if (high < low || close < low || close > high || (open !== undefined && (open < low || open > high))) throw new RangeError(`K 线第 ${categoryIndex + 1} 条数据的最高、最低、开盘或收盘关系无效。`)
    return { ...(open === undefined ? {} : { open }), high, low, close, categoryIndex }
  })
}

