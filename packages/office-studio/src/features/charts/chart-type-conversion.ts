import type { ChartType } from '../../domain/model'
import type { ChartConfig } from './chart-engine'
import { resolveChartType } from './chart-catalog'
import { chartNumber, type ChartRecords } from './chart-data'
import { cloneOptionValue } from './chart-options'

/** Switching a chart converts field roles while preserving its source and authored appearance. */
export function convertChartType(source: ChartConfig, type: ChartType, records?: ChartRecords): ChartConfig {
  const previous = resolveChartType(source.type), descriptor = resolveChartType(type)
  const next = cloneOptionValue(source) as ChartConfig
  next.type = type
  const mapping = (next.options ??= {}).mapping ??= {}
  const valid = (field: number | undefined): field is number => field !== undefined && Number.isSafeInteger(field) && field >= 0
    && (!records || field < records.fields.length)
  const numeric = records ? records.sourceValueFields ?? records.fields.map((_name, index) => index).filter(field => records.records.some(record => chartNumber(record[field]) !== null)) : []
  const choose = (preferred: Array<number | undefined>, excluded: number[] = []): number | undefined => [...preferred, ...numeric].find(field => valid(field)
    && !excluded.includes(field) && (!records || numeric.includes(field)))
  const labels = records?.fields.map((_name, index) => index).filter(index => !numeric.includes(index)) ?? []
  const oldCategory = source.categoryField ?? 0
  const authoredValues = source.valueFields ?? []
  const oldValues = previous.dataShape === 'bubble' ? [source.options?.mapping?.bubble?.yIndex ?? authoredValues[0]]
    : previous.dataShape === 'edge' ? [source.options?.mapping?.edge?.valueIndex ?? authoredValues[1]]
      : previous.dataShape === 'gauge' ? [source.options?.mapping?.gauge?.valueIndex ?? authoredValues[0]]
        : previous.dataShape === 'ohlc' ? [source.options?.mapping?.candlestick?.closeIndex ?? authoredValues.at(-1)] : authoredValues
  const selected = oldValues.filter(valid).filter(field => !records || numeric.includes(field))
  const inferred = selected.length ? selected : numeric.filter(field => field !== oldCategory)

  if (previous.dataShape !== descriptor.dataShape) {
    if (previous.dataShape === 'category' || previous.dataShape === 'heatmap') mapping.categoryIndexes ??= [oldCategory]
    const oldLabel = previous.dataShape === 'bubble' ? source.options?.mapping?.bubble?.labelIndex
      : previous.dataShape === 'edge' ? source.options?.mapping?.edge?.sourceIndex ?? oldCategory
        : previous.dataShape === 'hierarchy' ? source.options?.mapping?.hierarchyIndexes?.[0] ?? oldCategory
          : previous.dataShape === 'gauge' ? source.options?.mapping?.gauge?.labelIndex : oldCategory
    const category = [...(mapping.categoryIndexes ?? []), oldLabel, ...labels, oldCategory].find(valid) ?? 0
    if (descriptor.dataShape === 'xy') {
      const x = choose([source.options?.mapping?.bubble?.xIndex, oldCategory, ...numeric])
      next.categoryField = x ?? oldCategory
      next.valueFields = [...new Set([...inferred, ...numeric].filter(field => field !== next.categoryField))]
      next.seriesToCategory = {}
    } else if (descriptor.dataShape === 'bubble') {
      const existing = mapping.bubble
      const x = choose([existing?.xIndex, oldCategory, ...numeric])
      const y = choose([existing?.yIndex, ...inferred], x === undefined ? [] : [x])
      const size = choose([existing?.sizeIndex, ...authoredValues], [x, y].filter(valid))
      next.categoryField = x ?? oldCategory
      next.valueFields = [y, size].filter(valid)
      mapping.bubble = x !== undefined && y !== undefined ? { ...existing, xIndex: x, yIndex: y, sizeIndex: size,
        labelIndex: existing && valid(existing.labelIndex) ? existing.labelIndex : labels.find(field => field !== x && field !== y),
      } : undefined
      next.seriesToCategory = {}
    } else if (descriptor.dataShape === 'edge') {
      const existing = mapping.edge
      const first = existing && valid(existing.sourceIndex) ? existing.sourceIndex : category
      const second = existing && valid(existing.targetIndex) && existing.targetIndex !== first ? existing.targetIndex
        : [...labels, ...(records?.fields.map((_name, index) => index) ?? authoredValues)].find(field => field !== first)
      const weight = choose([existing?.valueIndex, ...inferred], [first, second].filter(valid))
      next.categoryField = first
      next.valueFields = [second, weight].filter(valid)
      mapping.edge = second === undefined ? undefined : { sourceIndex: first, targetIndex: second, valueIndex: weight }
      next.seriesToCategory = {}
    } else if (descriptor.dataShape === 'hierarchy') {
      const weight = choose(inferred)
      const layers = [...new Set((mapping.hierarchyIndexes ?? mapping.categoryIndexes ?? records?.sourceCategoryFields ?? [category]).filter(valid).filter(field => field !== weight))]
      if (!layers.length) { const layer = [...labels, category].find(field => field !== weight); if (layer !== undefined) layers.push(layer) }
      next.categoryField = layers[0] ?? category
      next.valueFields = weight === undefined ? [] : [weight]
      mapping.hierarchyIndexes = layers.length ? layers : undefined
      next.seriesToCategory = {}
    } else if (descriptor.dataShape === 'ohlc') {
      const values = [...new Set([...authoredValues, ...numeric].filter(valid).filter(field => (!records || numeric.includes(field)) && field !== category))]
      next.categoryField = category
      next.valueFields = values.slice(0, values.length >= 4 ? 4 : 3)
      if (!mapping.candlestick || Object.values(mapping.candlestick).some(field => field !== undefined && !valid(field))) {
        mapping.candlestick = values.length >= 4 ? { openIndex: values[0], highIndex: values[1]!, lowIndex: values[2]!, closeIndex: values[3]! }
          : values.length >= 3 ? { highIndex: values[0]!, lowIndex: values[1]!, closeIndex: values[2]! } : undefined
      }
      next.seriesToCategory = {}
    } else if (descriptor.dataShape === 'gauge') {
      const previousGauge = mapping.gauge
      const value = choose([mapping.gauge?.valueIndex, ...inferred])
      next.categoryField = category
      next.valueFields = value === undefined ? [] : [value]
      mapping.gauge = value === undefined ? undefined : { valueIndex: value,
        labelIndex: previousGauge && valid(previousGauge.labelIndex) ? previousGauge.labelIndex : category === value ? undefined : category }
      next.seriesToCategory = {}
    } else if (descriptor.dataShape === 'distribution') {
      next.valueFields = [...new Set([...inferred, ...numeric])]
      next.seriesToCategory = {}
    } else {
      next.categoryField = category
      next.valueFields = [...new Set(inferred.filter(field => !(mapping.categoryIndexes ?? [category]).includes(field)))]
      next.seriesToCategory = {}
    }
  }
  if (['pie', 'funnel', 'wordCloud', 'pareto', 'histogram'].includes(descriptor.family)) {
    next.valueFields = (next.valueFields?.length ? next.valueFields : inferred).slice(0, 1)
  }
  if (descriptor.stacked || descriptor.family === 'radar') {
    next.yAxis = { ...next.yAxis, scaleType: 'linear' }
    if (next.options?.allSeriesStyle?.rightAxis) next.options.allSeriesStyle.rightAxis = false
    for (const style of Object.values(next.options?.series ?? {})) if (style.rightAxis) style.rightAxis = false
    for (const entry of next.options?.namedSeries ?? []) if (entry.style.rightAxis) entry.style.rightAxis = false
  }
  if (Boolean(previous.percent) !== Boolean(descriptor.percent)) {
    next.yAxis = { ...next.yAxis, min: undefined, max: undefined }
    if (next.options?.rightYAxis) { next.options.rightYAxis.min = undefined; next.options.rightYAxis.max = undefined }
  }
  return next
}
