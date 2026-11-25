import type { SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canonicalSheetAxisSpans, type SheetAxisSpan } from '../../domain/sheet-axis.ts'
import { mapSheetAxisIndex, sheetAxisMaximum, type SheetAxisEdit } from '../../domain/sheet-axis-structure.ts'
import { normalizeSheetStyleInput, type SheetStyleInput } from './sheet-style-inheritance.ts'

export interface SheetPublicAxisData {
  sources: Record<number, Record<string, unknown>>
  sizes: Record<number, number>
  styles: NonNullable<SpreadsheetSheet['rowStyles']>
  autoHeights: Record<number, number>
  autoHeightSpans: SheetAxisSpan<boolean>[]
  hiddenSpans: SheetAxisSpan<boolean>[]
}
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
function axisIndex(key: string, maximum: number): number {
  const index = Number(key)
  if (!/^(0|[1-9]\d*)$/.test(key) || !Number.isSafeInteger(index) || index >= maximum) throw new RangeError('公开行列数据索引越界。')
  return index
}
function size(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) throw new RangeError('行列尺寸必须是有限非负数。')
  return value
}
function flag(value: unknown): boolean {
  if (value !== 0 && value !== 1 && typeof value !== 'boolean') throw new TypeError('行列状态必须是 0/1 或布尔值。')
  return value === 1 || value === true
}

/** Public h/ah/ia/w/hd are CSS-pixel geometry and author states, not passive metadata. */
export function normalizeSheetPublicAxisData(input: unknown, axis: 'row' | 'column'): SheetPublicAxisData {
  const result: SheetPublicAxisData = { sources: {}, sizes: {}, styles: {}, autoHeights: {}, autoHeightSpans: [], hiddenSpans: [] }
  if (input === undefined) return result
  if (!object(input)) throw new TypeError('公开行列数据必须是索引对象。')
  for (const [key, raw] of Object.entries(input)) {
    const index = axisIndex(key, sheetAxisMaximum(axis))
    if (!object(raw)) throw new TypeError('公开行列属性必须是对象。')
    result.sources[index] = cloneModel(raw)
    const dimension = axis === 'row' ? raw.h : raw.w
    if (dimension !== undefined) result.sizes[index] = size(dimension)
    if (Object.hasOwn(raw, 's') && raw.s !== undefined) result.styles[index] = normalizeSheetStyleInput(raw.s as SheetStyleInput)
    if (raw.hd !== undefined) result.hiddenSpans.push({ start: index, end: index, value: flag(raw.hd) })
    if (axis !== 'row') continue
    const automatic = raw.ia === undefined ? undefined : flag(raw.ia)
    const measured = raw.ah === undefined ? undefined : size(raw.ah)
    if (measured !== undefined && measured > 0) result.autoHeights[index] = measured
    // The fixed getter uses valid ah when automatic, otherwise h. Preserve an explicit auto row's h fallback.
    const mode = automatic ?? (measured !== undefined && measured > 0 ? true : dimension !== undefined ? false : undefined)
    if (mode !== undefined) result.autoHeightSpans.push({ start: index, end: index, value: mode })
    if (mode && result.autoHeights[index] === undefined && dimension !== undefined) result.autoHeights[index] = result.sizes[index]!
  }
  result.autoHeightSpans = canonicalSheetAxisSpans(result.autoHeightSpans)
  result.hiddenSpans = canonicalSheetAxisSpans(result.hiddenSpans)
  return result
}

type SheetAxisInput = Partial<SpreadsheetSheet> & { rowData?: unknown; columnData?: unknown }
function sizePoints(input: unknown, maximum: number): Record<number, number> {
  if (input === undefined) return {}
  if (!object(input)) throw new TypeError('行列尺寸必须是索引对象。')
  return Object.fromEntries(Object.entries(input).map(([key, value]) => [axisIndex(key, maximum), size(value)]))
}
function spans<T>(input: readonly SheetAxisSpan<T>[] | undefined, maximum: number, validate: (value: T) => unknown): SheetAxisSpan<T>[] {
  if (input === undefined) return []
  if (!Array.isArray(input)) throw new TypeError('行列属性范围必须是数组。')
  const result = canonicalSheetAxisSpans(input.map(span => {
    if (!span || span.end >= maximum) throw new RangeError('行列属性范围越界。')
    validate(span.value); return cloneModel(span)
  }))
  return result
}
function hiddenPoints(input: readonly number[] | undefined, maximum: number): number[] | undefined {
  if (input === undefined) return undefined
  if (!Array.isArray(input) || input.some(index => !Number.isSafeInteger(index) || index < 0 || index >= maximum)) throw new RangeError('隐藏行列索引越界。')
  return [...new Set(input)].sort((left, right) => left - right)
}

/** Validate all public axes and local overrides before applying the initial axis data atomically. */
export function applySheetPublicAxisInput(sheet: SpreadsheetSheet, input: SheetAxisInput): void {
  const row = normalizeSheetPublicAxisData(input.rowData, 'row'), column = normalizeSheetPublicAxisData(input.columnData, 'column')
  const rows = sheetAxisMaximum('row'), columns = sheetAxisMaximum('column')
  const bool = (value: boolean) => { if (typeof value !== 'boolean') throw new TypeError('行列范围状态必须是布尔值。') }
  const candidate: Partial<SpreadsheetSheet> = {
    rowHeights: { ...row.sizes, ...sizePoints(input.rowHeights, rows) },
    columnWidths: { ...column.sizes, ...sizePoints(input.columnWidths, columns) },
    rowAutoHeights: { ...row.autoHeights, ...sizePoints(input.rowAutoHeights, rows) },
    rowHeightSpans: spans(input.rowHeightSpans, rows, size),
    columnWidthSpans: spans(input.columnWidthSpans, columns, size),
    rowAutoHeightSpans: canonicalSheetAxisSpans([...row.autoHeightSpans, ...spans(input.rowAutoHeightSpans, rows, bool)]),
    rowHiddenSpans: canonicalSheetAxisSpans([...row.hiddenSpans, ...spans(input.rowHiddenSpans, rows, bool)]),
    columnHiddenSpans: canonicalSheetAxisSpans([...column.hiddenSpans, ...spans(input.columnHiddenSpans, columns, bool)]),
    hiddenRows: hiddenPoints(input.hiddenRows, rows),
    hiddenColumns: hiddenPoints(input.hiddenColumns, columns),
    rowStyles: Object.keys(row.styles).length || input.rowStyles ? { ...row.styles, ...cloneModel(input.rowStyles ?? {}) } : undefined,
    columnStyles: Object.keys(column.styles).length || input.columnStyles ? { ...column.styles, ...cloneModel(input.columnStyles ?? {}) } : undefined,
    axisUnit: 'cssPixel',
  }
  for (const key of ['defaultRowHeight', 'defaultColumnWidth'] as const) if (input[key] !== undefined) {
    const value = size(input[key]); if (value === 0) throw new RangeError('默认行列尺寸必须为正数。')
    candidate[key] = value
  }
  if (input.defaultRowAutoHeight !== undefined) {
    bool(input.defaultRowAutoHeight); candidate.defaultRowAutoHeight = input.defaultRowAutoHeight
  }
  if (input.axisUnit !== undefined && input.axisUnit !== 'cssPixel') throw new TypeError('行列尺寸单位必须是 cssPixel。')
  if (input.rowData !== undefined || input.columnData !== undefined) candidate.metadata = { ...sheet.metadata, publicAxisData: {
    ...(input.rowData !== undefined ? { rowData: row.sources } : {}),
    ...(input.columnData !== undefined ? { columnData: column.sources } : {}),
  } }
  Object.assign(sheet, candidate)
  if (!Object.keys(candidate.rowAutoHeights!).length) delete sheet.rowAutoHeights
}

/** Unknown public axis fields keep their source coordinates through structural edits without becoming authoritative caches. */
export function transformSheetPublicAxisSources(sheet: SpreadsheetSheet, edit: SheetAxisEdit): void {
  const archive = sheet.metadata?.publicAxisData
  if (!object(archive)) return
  const key = edit.axis === 'row' ? 'rowData' : 'columnData', input = archive[key]
  if (!object(input)) return
  const result: Record<number, unknown> = {}, maximum = sheetAxisMaximum(edit.axis)
  for (const [key, value] of Object.entries(input)) {
    const index = axisIndex(key, maximum), moved = mapSheetAxisIndex(index, edit)
    if (moved !== undefined && moved < maximum) result[moved] = cloneModel(value)
    if (edit.kind === 'copy' && index >= edit.start && index < edit.start + edit.count && edit.destination + index - edit.start < maximum) {
      result[edit.destination + index - edit.start] = cloneModel(value)
    }
  }
  sheet.metadata = { ...sheet.metadata, publicAxisData: { ...archive, [key]: result } }
}
