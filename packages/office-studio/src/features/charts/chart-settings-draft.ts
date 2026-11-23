import type { CellRange, ChartModel, ChartSource } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import type { ChartConfig } from './chart-engine'
import { cloneChartOptions, cloneOptionValue } from './chart-options'
import { resolveChartType } from './chart-catalog'

export interface ChartSettingsDraft { config: ChartConfig; sourceText: string }
export interface ChartSettingsField { value: number; label: string }
export interface ChartSettingsSourceSheet { id: string; name: string; cells: Record<string, { value?: string | number | boolean | null; rawValue?: string | number | boolean | null; cachedValue?: string }>; rowCount?: number; columnCount?: number }

/** Internal UI drafts are already typed. Preserve explicit undefined resets while fields are being edited. */
export function cloneChartDraftConfig(config: ChartConfig): ChartConfig { return cloneOptionValue(config) as ChartConfig }

export function createChartSettingsDraft(source: ChartConfig | ChartModel): ChartSettingsDraft {
  const config: ChartConfig = {
    type: source.type, source: cloneModel(source.source), position: { ...source.position }, size: { ...source.size },
    title: source.title, subtitle: source.subtitle, palette: source.palette ? [...source.palette] : undefined,
    categoryField: source.categoryField ?? 0, valueFields: [...(source.valueFields ?? [])],
    seriesToCategory: source.seriesToCategory ? cloneModel(source.seriesToCategory) : {}, pivotTableId: source.pivotTableId,
    legend: { visible: true, position: 'bottom', ...cloneModel(source.legend ?? {}) },
    xAxis: { scaleType: 'linear', ...cloneModel(source.xAxis ?? {}) }, yAxis: { scaleType: 'linear', ...cloneModel(source.yAxis ?? {}) },
    options: source.options ? cloneModel(source.options) : {},
  }
  return { config, sourceText: chartRangeNotation(source.source.range) }
}

export function chartConfigFromDraft(draft: ChartSettingsDraft): ChartConfig {
  const range = parseChartSourceRange(draft.sourceText)
  if (!range) throw new RangeError('请输入有效的数据范围，例如 A1:C4。')
  return {
    ...draft.config, source: { ...draft.config.source, range },
    title: draft.config.title?.trim() || undefined, subtitle: draft.config.subtitle?.trim() || undefined,
    valueFields: [...(draft.config.valueFields ?? [])], palette: draft.config.palette ? [...draft.config.palette] : undefined,
    options: cloneChartOptions(draft.config.options),
  }
}

export function chartSettingsPatch(draft: ChartSettingsDraft): Partial<Omit<ChartModel, 'id' | 'sheetId' | 'revision'>> {
  const config = chartConfigFromDraft(draft)
  return { ...config, categoryField: config.categoryField ?? 0, valueFields: [...(config.valueFields ?? [])],
    legend: { visible: true, position: 'bottom', ...config.legend },
    xAxis: { scaleType: 'linear', ...config.xAxis }, yAxis: { scaleType: 'linear', ...config.yAxis } }
}

export function chartDraftFields(draft: ChartSettingsDraft, sheet: ChartSettingsSourceSheet | undefined): ChartSettingsField[] {
  const range = parseChartSourceRange(draft.sourceText)
  if (!range || !sheet) return []
  const count = draft.config.source.orientation === 'columns' ? range.endCol - range.startCol + 1 : range.endRow - range.startRow + 1
  // Column-oriented data is bounded by Excel's field count. Row-oriented
  // sources may be much larger; callers offer search instead of expanding an
  // unbounded list of empty fields into DOM options.
  if (count > 16_384) return []
  return Array.from({ length: count }, (_unused, value) => {
    const row = range.startRow + (draft.config.source.orientation === 'rows' ? value : 0)
    const column = range.startCol + (draft.config.source.orientation === 'columns' ? value : 0)
    const address = `${chartColumnName(column)}${row + 1}`, label = sheet.cells[address]?.value
    return { value, label: label == null || label === '' ? address : `${address} · ${label}` }
  })
}

export function chartDraftFieldCount(draft: ChartSettingsDraft): number {
  const range = parseChartSourceRange(draft.sourceText)
  return !range ? 0 : draft.config.source.orientation === 'columns' ? range.endCol - range.startCol + 1 : range.endRow - range.startRow + 1
}

export function chartDraftNumericFields(draft: ChartSettingsDraft, sheet: ChartSettingsSourceSheet | undefined): ChartSettingsField[] {
  const range = parseChartSourceRange(draft.sourceText), fields = chartDraftFields(draft, sheet)
  if (!range || !sheet) return []
  const numeric = new Set<number>()
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const address = chartSourceAddress(key), value = cell.rawValue ?? cell.cachedValue ?? cell.value
    if (!address || address.row < range.startRow || address.row > range.endRow || address.column < range.startCol || address.column > range.endCol
      || typeof value === 'boolean' || value == null || typeof value === 'string' && value.trim() === '' || !Number.isFinite(Number(value))) continue
    if (draft.config.source.orientation === 'columns') { if (address.row > range.startRow) numeric.add(address.column - range.startCol) }
    else if (address.column > range.startCol) numeric.add(address.row - range.startRow)
  }
  const descriptor = resolveChartType(draft.config.type), categories = descriptor.dataShape === 'distribution' ? [] : draft.config.options?.mapping?.categoryIndexes ?? [draft.config.categoryField ?? 0]
  return fields.filter(field => numeric.has(field.value) && !categories.includes(field.value))
}

export function chartDraftCategories(draft: ChartSettingsDraft, sheet: ChartSettingsSourceSheet | undefined): string[] {
  const range = parseChartSourceRange(draft.sourceText)
  if (!range || !sheet) return []
  const count = Math.min(16_384, draft.config.source.orientation === 'columns' ? range.endRow - range.startRow : range.endCol - range.startCol)
  const fields = draft.config.options?.mapping?.categoryIndexes ?? [draft.config.categoryField ?? 0]
  return Array.from({ length: count }, (_unused, index) => fields.map(field => {
    const row = range.startRow + (draft.config.source.orientation === 'columns' ? index + 1 : field)
    const column = range.startCol + (draft.config.source.orientation === 'columns' ? field : index + 1)
    const cell = sheet.cells[`${chartColumnName(column)}${row + 1}`]
    return String(cell?.rawValue ?? cell?.cachedValue ?? cell?.value ?? '')
  }).join(' / '))
}

export function parseChartSourceRange(value: string): CellRange | undefined {
  const parts = value.trim().split(':')
  if (!parts[0] || parts.length > 2) return undefined
  const first = chartSourceAddress(parts[0]), last = chartSourceAddress(parts[1] ?? parts[0])
  if (!first || !last) return undefined
  return { startRow: Math.min(first.row, last.row), endRow: Math.max(first.row, last.row), startCol: Math.min(first.column, last.column), endCol: Math.max(first.column, last.column) }
}

export function chartRangeNotation(range: ChartSource['range']): string {
  const first = `${chartColumnName(range.startCol)}${range.startRow + 1}`, last = `${chartColumnName(range.endCol)}${range.endRow + 1}`
  return first === last ? first : `${first}:${last}`
}

function chartSourceAddress(value: string): { row: number; column: number } | undefined {
  const match = /^\$?([A-Z]+)\$?([1-9]\d*)$/i.exec(value.trim())
  if (!match) return undefined
  let column = 0
  for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2])
  return Number.isSafeInteger(row) && row <= 1_048_576 && column <= 16_384 ? { row: row - 1, column: column - 1 } : undefined
}

export function chartColumnName(index: number): string {
  let result = ''
  for (let column = index; column >= 0; column = Math.floor(column / 26) - 1) result = String.fromCharCode(column % 26 + 65) + result
  return result
}
