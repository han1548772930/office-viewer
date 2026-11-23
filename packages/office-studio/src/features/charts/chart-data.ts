import type { CellScalar, ChartModel, ChartSemanticOptions } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { finiteChartValue } from './chart-math'

export interface ChartRecords {
  fields: string[]
  records: CellScalar[][]
  /** Runtime date-system owner; this is not persisted in chart options. */
  date1904?: boolean
  sourceCategoryFields?: number[]
  sourceValueFields?: number[]
  /** Original field indices allowed by plotVisibleOnly; record axes are already filtered. */
  visibleFields?: number[]
  /** Pivot values placed on rows retain the legacy value-field zero shorthand. */
  syntheticValues?: boolean
}
export type ChartDataBinding = Pick<ChartModel, 'type' | 'categoryField' | 'valueFields' | 'seriesToCategory' | 'options'>
export interface ChartEdge { source: string; target: string; value: number }
export interface ChartGraphData { nodes: Array<{ name: string; value: number }>; links: ChartEdge[]; cyclic: boolean }
export interface ChartHierarchyNode { name: string; value: number; children?: ChartHierarchyNode[] }
export interface ChartHierarchyItem { node: ChartHierarchyNode; pointIndex: number; rootIndex: number; depth: number; path: string[] }

/** Source hierarchy order gives every root and descendant one stable point index. */
export function chartHierarchyItems(nodes: ChartHierarchyNode[]): ChartHierarchyItem[] {
  const items: ChartHierarchyItem[] = [], stack = nodes.map((node, rootIndex) => ({ node, rootIndex, depth: 0, path: [node.name] })).reverse()
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
  const shape = resolveChartType(chart.type).dataShape
  if (shape === 'bubble') {
    const mapping = chartBubbleMapping(source, chart)
    const groups = new Map<string, { field: number; name: string; values: Array<number | null>; xValues: Array<number | null>; sizeValues: Array<number | null> }>()
    const categories = source.records.map((record, index) => String(record[mapping.labelIndex ?? mapping.xIndex] ?? index + 1))
    for (let index = 0; index < source.records.length; index++) {
      const record = source.records[index]!
      const group = mapping.seriesIndex === undefined ? source.fields[mapping.yIndex]! : String(record[mapping.seriesIndex] ?? '')
      let series = groups.get(group)
      if (!series) {
        series = { field: mapping.yIndex, name: group, values: Array(source.records.length).fill(null), xValues: Array(source.records.length).fill(null), sizeValues: Array(source.records.length).fill(null) }
        groups.set(group, series)
      }
      series.values[index] = chartNumber(record[mapping.yIndex])
      series.xValues[index] = chartNumber(record[mapping.xIndex])
      series.sizeValues[index] = mapping.sizeIndex === undefined ? 1 : chartNumber(record[mapping.sizeIndex])
    }
    return { ...source, categories, categoryValues: source.records.map(record => record[mapping.xIndex] ?? null), series: [...groups.values()] }
  }
  if (shape === 'edge') {
    const graph = chartGraphData(source, chart)
    const mapping = chartEdgeMapping(source, chart), field = mapping.valueIndex ?? mapping.targetIndex
    return { ...source, categories: graph.nodes.map(node => node.name), series: [{ field,
      name: mapping.valueIndex === undefined ? '关系次数' : source.fields[field] ?? '关系权重', values: graph.nodes.map(node => node.value) }], graph }
  }
  if (shape === 'ohlc') {
    const candles = chartCandlestickData(source, chart)
    return { ...source, categories: source.records.map(row => String(row[chart.categoryField] ?? '')), categoryValues: source.records.map(row => row[chart.categoryField] ?? null), series: [{ field: chartCandlestickMapping(source, chart).closeIndex, name: '最高 / 最低 / 收盘', values: candles.map(candle => candle?.close ?? null) }], candles }
  }
  if (shape === 'hierarchy') {
    const hierarchy = chartHierarchyData(source, chart)
    const valueField = chartHierarchyValueField(source, chart)
    return { ...source, categories: hierarchy.map(node => node.name), series: [{ field: valueField, name: source.fields[valueField]!, values: hierarchy.map(node => node.value) }], hierarchy }
  }
  if (shape === 'gauge') {
    const mapping = chart.options?.mapping?.gauge ?? { valueIndex: chart.valueFields[0] ?? (source.fields.length > 1 ? 1 : 0), labelIndex: source.fields.length > 1 ? chart.categoryField : undefined }
    chartValueField(source, mapping.valueIndex, '仪表数值字段')
    if (mapping.labelIndex !== undefined) chartField(source, mapping.labelIndex, '仪表标签字段')
    return { ...source, categories: source.records.map((row, index) => String(mapping.labelIndex === undefined ? source.fields[mapping.valueIndex] : row[mapping.labelIndex] ?? index + 1)), series: [{ field: mapping.valueIndex, name: source.fields[mapping.valueIndex]!, values: source.records.map(row => chartNumber(row[mapping.valueIndex])) }] }
  }
  if (shape === 'distribution') {
    const numericFields = source.fields.map((_field, index) => index).filter(field => source.records.some(record => chartNumber(record[field]) !== null))
    const fields = chart.valueFields.length ? chart.valueFields : chart.type === 'histogram' ? numericFields.slice(0, 1) : numericFields
    fields.forEach(field => chartValueField(source, field, '样本字段'))
    return { ...source, categories: source.records.map((_row, index) => String(index + 1)), series: fields.map(field => ({ field, name: source.fields[field]!, values: source.records.map(row => chartNumber(row[field])) })) }
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
  const mapping = chartEdgeMapping(source, chart)
  const weights = new Map<string, number>(), links = new Map<string, ChartEdge>()
  const undirected = chart.type === 'chord'
  for (const row of source.records) {
    const first = String(row[mapping.sourceIndex] ?? '').trim(), second = String(row[mapping.targetIndex] ?? '').trim()
    if (!first || !second) continue
    const value = mapping.valueIndex === undefined ? 1 : chartNumber(row[mapping.valueIndex])
    if (value === null || value === 0) continue
    if (value < 0) throw new RangeError('关系图的权重不能为负数。')
    const [sourceName, targetName] = undirected && first > second ? [second, first] : [first, second]
    const key = JSON.stringify([sourceName, targetName])
    const link = links.get(key) ?? { source: sourceName, target: targetName, value: 0 }
    link.value = finiteChartValue(link.value + value, '关系边权重合计'); links.set(key, link)
    weights.set(first, finiteChartValue((weights.get(first) ?? 0) + value, '关系节点权重合计'))
    weights.set(second, finiteChartValue((weights.get(second) ?? 0) + value, '关系节点权重合计'))
  }
  const nodes = [...weights].map(([name, value]) => ({ name, value }))
  const edges = [...links.values()]
  const degrees = new Map(nodes.map(node => [node.name, 0])), outgoing = new Map<string, string[]>()
  for (const link of edges) {
    degrees.set(link.target, (degrees.get(link.target) ?? 0) + 1)
    const list = outgoing.get(link.source) ?? []; list.push(link.target); outgoing.set(link.source, list)
  }
  const queue = nodes.filter(node => degrees.get(node.name) === 0).map(node => node.name)
  let visited = 0
  for (let index = 0; index < queue.length; index++) {
    visited++
    for (const target of outgoing.get(queue[index]!) ?? []) { const degree = degrees.get(target)! - 1; degrees.set(target, degree); if (degree === 0) queue.push(target) }
  }
  return { nodes, links: edges, cyclic: visited !== nodes.length }
}

export function chartHierarchyValueField(source: ChartRecords, chart: ChartDataBinding): number {
  return chart.valueFields[0] ?? source.fields.map((_field, index) => index).find(index => index !== chart.categoryField && source.records.some(row => chartNumber(row[index]) !== null)) ?? 1
}

export function chartHierarchyData(source: ChartRecords, chart: ChartDataBinding): ChartHierarchyNode[] {
  const valueField = chartHierarchyValueField(source, chart)
  const fields = chart.options?.mapping?.hierarchyIndexes ?? [chart.categoryField]
  if (!fields.length || new Set(fields).size !== fields.length) throw new RangeError('层级字段须按从根到叶顺序选择且不能重复。')
  for (const field of [...fields, valueField]) chartField(source, field, '层级或权重字段')
  chartValueField(source, valueField, '层级权重字段')
  if (fields.includes(valueField)) throw new RangeError('层级名称字段不能同时作为权重字段。')
  const roots: ChartHierarchyNode[] = []
  for (const row of source.records) {
    const value = chartNumber(row[valueField])
    if (value === null || value === 0) continue
    if (value < 0) throw new RangeError('层级图权重不能为负数。')
    const names = fields.map(field => String(row[field] ?? '').trim())
    if (names.some(name => !name)) continue
    let siblings = roots
    names.forEach((name, depth) => {
      let node = siblings.find(candidate => candidate.name === name)
      if (!node) { node = { name, value: 0 }; siblings.push(node) }
      node.value = finiteChartValue(node.value + value, '层级权重合计')
      if (depth < names.length - 1) { node.children ??= []; siblings = node.children }
    })
  }
  return roots
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
  const mapping = chartCandlestickMapping(source, chart)
  return source.records.map((row, categoryIndex) => {
    const high = chartNumber(row[mapping.highIndex]), low = chartNumber(row[mapping.lowIndex]), close = chartNumber(row[mapping.closeIndex])
    const open = mapping.openIndex === undefined ? undefined : chartNumber(row[mapping.openIndex])
    if (high === null || low === null || close === null || open === null) return null
    if (high < low || close < low || close > high || (open !== undefined && (open < low || open > high))) throw new RangeError(`K 线第 ${categoryIndex + 1} 条数据的最高、最低、开盘或收盘关系无效。`)
    return { ...(open === undefined ? {} : { open }), high, low, close, categoryIndex }
  })
}

