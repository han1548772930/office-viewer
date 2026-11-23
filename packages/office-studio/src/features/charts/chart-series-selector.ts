import type { ChartSemanticOptions, ChartSeriesOptions, ChartSeriesSelector } from '../../domain/model'

export function validateChartSeriesSelector(selector: ChartSeriesSelector): void {
  if (typeof selector === 'number') {
    if (!Number.isSafeInteger(selector) || selector < 0 || selector > 1_048_575) throw new RangeError('图表系列字段须为 0 到 1048575 的整数。')
    return
  }
  if (!selector || typeof selector !== 'object' || Array.isArray(selector) || Object.getPrototypeOf(selector) !== Object.prototype
    || Object.keys(selector).length !== 1 || Object.getOwnPropertySymbols(selector).length || typeof selector.name !== 'string') {
    throw new TypeError('图表系列选择器须为数值字段或仅包含 name 字符串的对象。')
  }
}

export function chartNamedSeriesStyle(options: ChartSemanticOptions | undefined, sourceName: string | undefined): ChartSeriesOptions | undefined {
  return sourceName === undefined ? undefined : options?.namedSeries?.find(entry => entry.name === sourceName)?.style
}

export function chartSeriesOverride(options: ChartSemanticOptions | undefined, selector: ChartSeriesSelector): ChartSeriesOptions | undefined {
  return typeof selector === 'number' ? options?.series?.[selector] : chartNamedSeriesStyle(options, selector.name)
}

/** Inactive selectors survive filtering and source changes. An active name must identify one source series. */
export function validateChartNamedSeriesBindings(options: ChartSemanticOptions | undefined, series: readonly { name: string; sourceName?: string }[]): void {
  if (!options?.namedSeries?.length) return
  const names = new Set(options.namedSeries.map(entry => entry.name)), counts = new Map<string, number>()
  for (const source of series) {
    const name = source.sourceName ?? source.name
    if (names.has(name)) counts.set(name, (counts.get(name) ?? 0) + 1)
  }
  for (const [name, count] of counts) if (count > 1) throw new RangeError(`名称 ${JSON.stringify(name)} 匹配多个图表系列；请使用源字段选择器。`)
}
