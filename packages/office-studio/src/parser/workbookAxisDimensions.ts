import { canonicalSheetAxisSpans, type SheetAxisSpan } from '../domain/sheet-axis.ts'
import { array, record } from './validation.ts'

export interface NormalizedWorkbookAxisDimensions {
  axisUnit: 'cssPixel'
  rowHeights: Record<number, number>
  columnWidths: Record<number, number>
  rowHeightSpans: SheetAxisSpan<number>[]
  columnWidthSpans: SheetAxisSpan<number>[]
  rowAutoHeightSpans: SheetAxisSpan<boolean>[]
  rowAutoHeights: Record<number, number>
  defaultRowHeight?: number
  defaultColumnWidth?: number
  defaultRowAutoHeight?: boolean
  rowCount?: number
  columnCount?: number
  hiddenRows: number[]
  hiddenColumns: number[]
  rowHiddenSpans: SheetAxisSpan<boolean>[]
  columnHiddenSpans: SheetAxisSpan<boolean>[]
  rowExtent: number
  columnExtent: number
  compatibility: 'cssPixel' | 'legacyUnspecified'
}

function axisIndex(value: unknown, maximum: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value > maximum) throw new RangeError('解析的行列尺寸坐标无效。')
  return value
}
function axisSize(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) throw new RangeError('解析的行列尺寸必须是有限非负数。')
  return value
}
function axisPoints(source: unknown, maximum: number): Record<number, number> {
  const result: Record<number, number> = {}
  if (source === undefined || source === null) return result
  if (typeof source !== 'object') throw new TypeError('解析的行列尺寸点必须是数组或坐标映射。')
  const entries = Array.isArray(source) ? source.map(item => {
    const value = record(item)
    return [value.index, value.size] as const
  }) : Object.entries(record(source)).map(([key, value]) => {
    if (!/^(?:0|[1-9]\d*)$/.test(key)) throw new RangeError('解析的行列尺寸坐标无效。')
    return [Number(key), value] as const
  })
  for (const [index, size] of entries) {
    const position = axisIndex(index, maximum)
    if (Object.hasOwn(result, position)) throw new RangeError('解析的行列尺寸点重复。')
    result[position] = axisSize(size)
  }
  return result
}
function axisCount(source: unknown, maximum: number): number | undefined {
  if (source === undefined || source === null) return undefined
  if (typeof source !== 'number' || !Number.isSafeInteger(source) || source <= 0 || source > maximum + 1) throw new RangeError('解析的工作表行列总数无效。')
  return source
}
function hiddenIndices(source: unknown, maximum: number): number[] {
  if (source === undefined || source === null) return []
  if (!Array.isArray(source)) throw new TypeError('解析的隐藏行列坐标必须是数组。')
  return [...new Set(source.map(value => axisIndex(value, maximum)))].sort((a, b) => a - b)
}
function sizeSpans(source: unknown, maximum: number): SheetAxisSpan<number>[] {
  if (source !== undefined && source !== null && !Array.isArray(source)) throw new TypeError('解析的行列尺寸区间必须是数组。')
  return canonicalSheetAxisSpans(array(source).map(item => {
    const span = record(item)
    return { start: axisIndex(span.start, maximum), end: axisIndex(span.end, maximum), value: axisSize(span.value) }
  }))
}
function modeSpans(source: unknown, maximum = 1_048_575): SheetAxisSpan<boolean>[] {
  if (source !== undefined && source !== null && !Array.isArray(source)) throw new TypeError('解析的自动行高区间必须是数组。')
  return canonicalSheetAxisSpans(array(source).map(item => {
    const span = record(item)
    if (typeof span.value !== 'boolean') throw new TypeError('解析的自动行高模式无效。')
    return { start: axisIndex(span.start, maximum), end: axisIndex(span.end, maximum), value: span.value }
  }))
}
function extent(points: Record<number, number>, spans: readonly SheetAxisSpan<unknown>[]): number {
  let size = 0
  for (const key of Object.keys(points)) size = Math.max(size, Number(key) + 1)
  for (const span of spans) size = Math.max(size, span.end + 1)
  return size
}

/** Legacy unmarked DTOs retain the numbers the existing view used; no heuristic reinterprets small sizes. */
export function normalizeWorkbookAxisDimensions(raw: Record<string, unknown>): NormalizedWorkbookAxisDimensions {
  if (raw.axisUnit !== undefined && raw.axisUnit !== null && raw.axisUnit !== 'cssPixel') throw new TypeError('不支持的工作表尺寸单位。')
  if (raw.defaultRowAutoHeight !== undefined && raw.defaultRowAutoHeight !== null && typeof raw.defaultRowAutoHeight !== 'boolean') throw new TypeError('默认自动行高模式无效。')
  const compatibility = raw.axisUnit === 'cssPixel' ? 'cssPixel' : 'legacyUnspecified'
  const restored = compatibility === 'cssPixel' && record(record(raw.sourceAxisDimensions).projectAxisExtension).restored === true
  const rowHeights = axisPoints(raw.rowHeights, 1_048_575), columnWidths = axisPoints(raw.columnWidths, 16_383)
  const rowAutoHeights = axisPoints(raw.rowAutoHeights, 1_048_575)
  const rowHeightSpans = sizeSpans(raw.rowHeightSpans, 1_048_575), columnWidthSpans = sizeSpans(raw.columnWidthSpans, 16_383)
  const rowAutoHeightSpans = modeSpans(raw.rowAutoHeightSpans)
  const rowHiddenSpans = modeSpans(raw.rowHiddenSpans).filter(span => span.value), columnHiddenSpans = modeSpans(raw.columnHiddenSpans, 16_383).filter(span => span.value)
  const hiddenRows = hiddenIndices(raw.hiddenRows, 1_048_575), hiddenColumns = hiddenIndices(raw.hiddenColumns, 16_383)
  return {
    axisUnit: 'cssPixel', rowHeights, columnWidths, rowAutoHeights, rowHeightSpans, columnWidthSpans, rowAutoHeightSpans, rowHiddenSpans, columnHiddenSpans,
    defaultRowHeight: raw.defaultRowHeight == null ? restored ? undefined : compatibility === 'cssPixel' ? 26 : 24 : axisSize(raw.defaultRowHeight),
    defaultColumnWidth: raw.defaultColumnWidth == null ? restored ? undefined : compatibility === 'cssPixel' ? 112 : 96 : axisSize(raw.defaultColumnWidth),
    defaultRowAutoHeight: typeof raw.defaultRowAutoHeight === 'boolean' ? raw.defaultRowAutoHeight : undefined,
    rowCount: axisCount(raw.rowCount, 1_048_575), columnCount: axisCount(raw.columnCount, 16_383),
    hiddenRows, hiddenColumns,
    rowExtent: Math.max(extent(rowHeights, [...rowHeightSpans, ...rowAutoHeightSpans, ...rowHiddenSpans]), extent(rowAutoHeights, []), (hiddenRows.at(-1) ?? -1) + 1),
    columnExtent: Math.max(extent(columnWidths, [...columnWidthSpans, ...columnHiddenSpans]), (hiddenColumns.at(-1) ?? -1) + 1), compatibility,
  }
}
