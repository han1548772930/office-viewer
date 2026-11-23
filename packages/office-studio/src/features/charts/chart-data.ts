import type { CellScalar, ChartModel, ChartSemanticOptions } from '../../domain/model'
import type { ChartData } from './chart-engine'
import { resolveChartType } from './chart-catalog'

export interface ChartRecords { fields: string[]; records: CellScalar[][] }
export interface ChartEdge { source: string; target: string; value: number }
export interface ChartGraphData { nodes: Array<{ name: string; value: number }>; links: ChartEdge[]; cyclic: boolean }
export interface ChartHierarchyNode { name: string; value: number; children?: ChartHierarchyNode[] }
export interface ChartCandle { open?: number; high: number; low: number; close: number; categoryIndex: number }

export function chartNumber(value: CellScalar | undefined): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'string' && value.trim() && Number.isFinite(Number(value))) return Number(value)
  return null
}

export function chartField(recordSet: ChartRecords, field: number, label: string): void {
  if (!Number.isSafeInteger(field) || field < 0 || field >= recordSet.fields.length) throw new RangeError(`${label}必须位于图表数据区域内。`)
}

export function readSpecialChartData(source: ChartRecords, chart: ChartModel): ChartData | undefined {
  const shape = resolveChartType(chart.type).dataShape
  if (shape === 'bubble') {
    const mapping = bubbleMapping(source, chart)
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
    return { ...source, categories, series: [...groups.values()] }
  }
  if (shape === 'edge') {
    const graph = chartGraphData(source, chart)
    return { ...source, categories: graph.nodes.map(node => node.name), series: [{ field: chart.options?.mapping?.edge?.valueIndex ?? 2, name: source.fields[chart.options?.mapping?.edge?.valueIndex ?? 2] ?? '关系权重', values: graph.nodes.map(node => node.value) }], graph }
  }
  if (shape === 'ohlc') {
    const candles = chartCandlestickData(source, chart)
    return { ...source, categories: source.records.map(row => String(row[chart.categoryField] ?? '')), series: [{ field: candlestickMapping(source, chart).closeIndex, name: '最高 / 最低 / 收盘', values: candles.map(candle => candle?.close ?? null) }], candles }
  }
  if (shape === 'hierarchy') {
    const hierarchy = chartHierarchyData(source, chart)
    const valueField = hierarchyValueField(source, chart)
    return { ...source, categories: hierarchy.map(node => node.name), series: [{ field: valueField, name: source.fields[valueField]!, values: hierarchy.map(node => node.value) }], hierarchy }
  }
  if (shape === 'gauge') {
    const mapping = chart.options?.mapping?.gauge ?? { valueIndex: chart.valueFields[0] ?? (source.fields.length > 1 ? 1 : 0), labelIndex: source.fields.length > 1 ? chart.categoryField : undefined }
    chartField(source, mapping.valueIndex, '仪表数值字段')
    if (mapping.labelIndex !== undefined) chartField(source, mapping.labelIndex, '仪表标签字段')
    return { ...source, categories: source.records.map((row, index) => String(mapping.labelIndex === undefined ? source.fields[mapping.valueIndex] : row[mapping.labelIndex] ?? index + 1)), series: [{ field: mapping.valueIndex, name: source.fields[mapping.valueIndex]!, values: source.records.map(row => chartNumber(row[mapping.valueIndex])) }] }
  }
  if (shape === 'distribution' && source.fields.length === 1) {
    return { ...source, categories: source.records.map((_row, index) => String(index + 1)), series: [{ field: 0, name: source.fields[0]!, values: source.records.map(row => chartNumber(row[0])) }] }
  }
  return undefined
}

function bubbleMapping(source: ChartRecords, chart: ChartModel): NonNullable<NonNullable<ChartSemanticOptions['mapping']>['bubble']> {
  const values = chart.valueFields.length ? chart.valueFields : source.fields.map((_field, index) => index).filter(index => index !== chart.categoryField)
  const mapping = chart.options?.mapping?.bubble ?? { xIndex: chart.categoryField, yIndex: values[0] ?? 1, sizeIndex: values[1] }
  for (const [name, field] of Object.entries(mapping)) if (field !== undefined) chartField(source, field, `气泡${name}`)
  return mapping
}

export function chartGraphData(source: ChartRecords, chart: ChartModel): ChartGraphData {
  const mapping = chart.options?.mapping?.edge ?? { sourceIndex: chart.categoryField, targetIndex: chart.valueFields[0] ?? 1, valueIndex: chart.valueFields[1] ?? (source.fields.length > 2 ? 2 : undefined) }
  chartField(source, mapping.sourceIndex, '来源字段'); chartField(source, mapping.targetIndex, '目标字段')
  if (mapping.sourceIndex === mapping.targetIndex) throw new RangeError('来源和目标须使用不同字段。')
  if (mapping.valueIndex !== undefined) chartField(source, mapping.valueIndex, '权重字段')
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
    link.value += value; links.set(key, link)
    weights.set(first, (weights.get(first) ?? 0) + value)
    weights.set(second, (weights.get(second) ?? 0) + value)
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

function hierarchyValueField(source: ChartRecords, chart: ChartModel): number {
  return chart.valueFields[0] ?? source.fields.map((_field, index) => index).find(index => index !== chart.categoryField && source.records.some(row => chartNumber(row[index]) !== null)) ?? 1
}

export function chartHierarchyData(source: ChartRecords, chart: ChartModel): ChartHierarchyNode[] {
  const valueField = hierarchyValueField(source, chart)
  const fields = chart.options?.mapping?.hierarchyIndexes ?? [chart.categoryField]
  if (!fields.length || new Set(fields).size !== fields.length) throw new RangeError('层级字段须按从根到叶顺序选择且不能重复。')
  for (const field of [...fields, valueField]) chartField(source, field, '层级或权重字段')
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
      node.value += value
      if (depth < names.length - 1) { node.children ??= []; siblings = node.children }
    })
  }
  return roots
}

function candlestickMapping(source: ChartRecords, chart: ChartModel): NonNullable<NonNullable<ChartSemanticOptions['mapping']>['candlestick']> {
  const fields = chart.valueFields.length ? chart.valueFields : source.fields.map((_name, index) => index).filter(index => index !== chart.categoryField)
  const mapping = chart.options?.mapping?.candlestick ?? (fields.length >= 4
    ? { openIndex: fields[0], highIndex: fields[1]!, lowIndex: fields[2]!, closeIndex: fields[3]! }
    : { highIndex: fields[0]!, lowIndex: fields[1]!, closeIndex: fields[2]! })
  for (const [name, field] of Object.entries(mapping)) if (field !== undefined) chartField(source, field, name)
  const roles = Object.values(mapping).filter((field): field is number => field !== undefined)
  if (roles.length < 3 || new Set(roles).size !== roles.length) throw new RangeError('K 线须选择不同的最高、最低、收盘字段，开盘字段可选。')
  return mapping
}

export function chartCandlestickData(source: ChartRecords, chart: ChartModel): Array<ChartCandle | null> {
  const mapping = candlestickMapping(source, chart)
  return source.records.map((row, categoryIndex) => {
    const high = chartNumber(row[mapping.highIndex]), low = chartNumber(row[mapping.lowIndex]), close = chartNumber(row[mapping.closeIndex])
    const open = mapping.openIndex === undefined ? undefined : chartNumber(row[mapping.openIndex])
    if (high === null || low === null || close === null || open === null) return null
    if (high < low || close < low || close > high || (open !== undefined && (open < low || open > high))) throw new RangeError(`K 线第 ${categoryIndex + 1} 条数据的最高、最低、开盘或收盘关系无效。`)
    return { ...(open === undefined ? {} : { open }), high, low, close, categoryIndex }
  })
}

/** Percent stacks retain the sign of each source value and use the absolute category total. */
export function normalizedChartValues(data: ChartData, chart: ChartModel): Array<Array<number | null>> {
  const percent = resolveChartType(chart.type).percent
  const totals = data.categories.map((_category, index) => data.series.reduce((sum, series) => sum + Math.abs(series.values[index] ?? 0), 0))
  return data.series.map(series => series.values.map((value, index) => value === null
    ? chart.options?.invalidValueStrategy === 'zero' ? 0 : null
    : percent ? totals[index] ? 100 * value / totals[index]! : 0 : value))
}
