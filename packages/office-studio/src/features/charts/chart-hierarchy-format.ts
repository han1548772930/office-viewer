import type { ChartRecords } from './chart-data'
import { CHART_PREPARATION_CHUNK } from './chart-compute-lifecycle'

/** Fixed za/xae infers one source formatter from the first positive formatted
 * numeric observation. Raw values and display strings remain separate. */
export function* chartHierarchySourceFormatterSteps(source: ChartRecords, field: number): Generator<void, (value: number) => string, void> {
  for (let row = 0; row < source.records.length; row++) {
    const value = source.records[row]?.[field], label = source.formattedRecords?.[row]?.[field]
    if (typeof value === 'number' && Number.isFinite(value) && value > 0 && typeof label === 'string' && label) {
      const formatter = chartHierarchySourceFormatter(label)
      if (formatter) return formatter
    }
    if ((row + 1) % CHART_PREPARATION_CHUNK === 0) yield
  }
  return String
}

export function chartHierarchySourceFormatter(label: string): ((value: number) => string) | undefined {
  const numeric = /[\d][\d.,\s]*/.exec(label)
  if (!numeric) return undefined
  const body = numeric[0].trimEnd(), prefix = label.slice(0, numeric.index), suffix = label.slice(numeric.index + body.length)
  if (/^[eE][+-]?\d/.test(suffix)) return undefined
  const separators = [...body].filter(character => character === '.' || character === ','), last = separators.at(-1)
  const places = last ? body.length - body.lastIndexOf(last) - 1 : 0
  const decimal = last && (separators.includes(last === '.' ? ',' : '.') || places !== 3) ? last : undefined
  const grouping = /\s/.exec(body)?.[0] ?? separators.find(character => character !== decimal) ?? (decimal === undefined ? separators[0] : undefined)
  const precision = decimal ? places : 0, factor = suffix.includes('%') ? 100 : 1
  return value => {
    try {
      const [integer, fraction] = (value * factor).toFixed(precision).split('.')
      const grouped = grouping ? integer!.replace(/\B(?=(\d{3})+(?!\d))/g, grouping) : integer!
      return `${prefix}${fraction === undefined ? grouped : `${grouped}${decimal}${fraction}`}${suffix}`
    } catch { return String(value) }
  }
}
