import type { OutlineAxis, SheetAxisSpan, SheetOutlineGroup, SheetOutlineState, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { getSheetOutlineHiddenSpans, validateSheetOutlineState } from './sheet-outline.ts'
import type { SheetOutlineExtent, SheetOutlineHiddenRange } from './sheet-outline-types.ts'

const MAX_ROWS = 1_048_576, MAX_COLUMNS = 16_384
const MAX_SOURCE_BYTES = 32 * 1024 * 1024
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)

export interface SheetOutlineImportResult extends SheetOutlineExtent {
  outline?: SheetOutlineState
  officeSource?: Record<string, unknown>
  nativeSource: boolean
  hiddenReasonsExact: boolean
  manualVisibilityNormalized: boolean
  diagnostics: string[]
}

function optionalBoolean(value: unknown, fallback: boolean, label: string): boolean {
  if (value === undefined) return fallback
  if (typeof value !== 'boolean') throw new TypeError(`大纲 ${label} 必须是布尔值。`)
  return value
}
function axisIndex(value: unknown, maximum: number, label: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0 || value >= maximum) throw new TypeError(`大纲 ${label} 超出 Office 轴范围。`)
  return value
}
function extent(value: number, maximum: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) throw new TypeError(`大纲 ${label} 工作表大小无效。`)
  return value
}

/** Parser DTO -> domain state. Invalid groups reject the candidate instead of
 * disappearing from the workbook; unknown group/state/source JSON is copied. */
export function normalizeSheetOutlineImport(rawValue: unknown, initialExtent: SheetOutlineExtent): SheetOutlineImportResult {
  if (!object(rawValue)) throw new TypeError('大纲导入缺少工作表 DTO。')
  const raw = rawValue
  let rowCount = extent(initialExtent.rowCount, MAX_ROWS, '行'), columnCount = extent(initialExtent.columnCount, MAX_COLUMNS, '列')
  const sourceValue = raw.sourceOutline
  if (sourceValue !== undefined && sourceValue !== null && !object(sourceValue)) throw new TypeError('大纲来源必须是对象。')
  let officeSource: Record<string, unknown> | undefined
  if (object(sourceValue)) {
    const serialized = JSON.stringify(sourceValue)
    if (new TextEncoder().encode(serialized).byteLength > MAX_SOURCE_BYTES) throw new TypeError('大纲 Office 来源超过 32 MiB 预算。')
    officeSource = cloneModel(sourceValue)
  }
  const native = officeSource && object(officeSource.native) ? officeSource.native : undefined
  const guard = officeSource && object(officeSource.projectOutlineExtension) ? officeSource.projectOutlineExtension : undefined
  const axisSource = object(raw.sourceAxisDimensions) ? raw.sourceAxisDimensions : undefined
  const axisGuard = axisSource && object(axisSource.projectAxisExtension) ? axisSource.projectAxisExtension : undefined
  const nativeSource = native !== undefined
  const hiddenReasonsExact = !nativeSource || guard?.hiddenReasonsExact === true || axisGuard?.restored === true
  const manualVisibilityNormalized = officeSource?.manualVisibilityNormalized === true
  const diagnostics = native && Array.isArray(native.diagnostics) ? native.diagnostics.filter((value): value is string => typeof value === 'string') : []
  let outline: SheetOutlineState | undefined
  if (raw.outline !== undefined && raw.outline !== null) {
    if (!object(raw.outline) || !Array.isArray(raw.outline.groups)) throw new TypeError('大纲导入缺少 groups。')
    const state = raw.outline
    const groupsInput = raw.outline.groups
    const groups: SheetOutlineGroup[] = groupsInput.map((value: unknown, index: number) => {
      if (!object(value) || value.axis !== 'row' && value.axis !== 'column') throw new TypeError('大纲成员轴无效。')
      const maximum = value.axis === 'row' ? MAX_ROWS : MAX_COLUMNS
      const start = axisIndex(value.start, maximum, '起点'), end = axisIndex(value.end, maximum, '终点')
      if (end < start) throw new TypeError('大纲成员范围倒置。')
      const id = value.id === undefined ? `xlsx-outline-legacy-${index}` : value.id
      if (typeof id !== 'string' || !id.trim()) throw new TypeError('大纲成员 ID 无效。')
      const preserved = cloneModel(value)
      delete preserved.unitId; delete preserved.subUnitId
      if (value.axis === 'row') rowCount = Math.max(rowCount, end + 1)
      else columnCount = Math.max(columnCount, end + 1)
      return { ...preserved, id, axis: value.axis, start, end, collapsed: optionalBoolean(value.collapsed, false, '折叠状态') }
    })
    outline = { ...cloneModel(state), groups, summaryAbove: optionalBoolean(state.summaryAbove, false, '行汇总位置'),
      summaryLeft: optionalBoolean(state.summaryLeft, false, '列汇总位置') }
  }
  // Empty summary rows/columns may lie outside the original used-cell rectangle.
  if (native) for (const [key, maximum] of [['rowRuns', MAX_ROWS], ['columnRuns', MAX_COLUMNS]] as const) {
    const runsValue = native[key]
    if (runsValue !== undefined && !Array.isArray(runsValue)) throw new TypeError('大纲原生轴来源不是数组。')
    const runs = Array.isArray(runsValue) ? runsValue : []
    for (const run of runs) {
      if (!object(run)) throw new TypeError('大纲原生轴成员无效。')
      if (typeof run.level !== 'number' || !Number.isSafeInteger(run.level) || run.level < 0 || run.level > 8) throw new TypeError('大纲原生层级无效。')
      const end = axisIndex(run.end, maximum, '原生终点'), start = axisIndex(run.start, maximum, '原生起点')
      if (start > end || typeof run.collapsed !== 'boolean') throw new TypeError('大纲原生范围或折叠状态无效。')
      if (run.level === 0 && !run.collapsed) continue
      if (key === 'rowRuns') rowCount = Math.max(rowCount, end + 1)
      else columnCount = Math.max(columnCount, end + 1)
    }
  }
  if (outline) validateSheetOutlineState(outline, { rowCount, columnCount })
  return { outline, officeSource, rowCount, columnCount, nativeSource, hiddenReasonsExact, manualVisibilityNormalized, diagnostics }
}

/** Main normalizer hook; parsing and native inference stay in Rust/Office IO. */
export function applySheetOutlineImport(sheet: SpreadsheetSheet, raw: unknown): SheetOutlineImportResult {
  const result = normalizeSheetOutlineImport(raw, { rowCount: sheet.rowCount ?? 200, columnCount: sheet.columnCount ?? 26 })
  if (result.outline) sheet.outline = result.outline
  sheet.rowCount = result.rowCount; sheet.columnCount = result.columnCount
  if (result.officeSource) sheet.metadata = { ...sheet.metadata, outlineOfficeSource: result.officeSource }
  reconcileSheetOutlineImportedVisibility(sheet, result)
  return result
}

/** Native `hidden` combines causes. Project-exact manual overlaps are retained;
 * inferred outlines/filters are removed only from ambiguous native manual sets.
 * Rust already normalized its DTO, so that path does not subtract twice. */
export function reconcileSheetOutlineImportedVisibility(sheet: SpreadsheetSheet, result: SheetOutlineImportResult, filteredRows: readonly number[] = []): void {
  if (!result.nativeSource || result.hiddenReasonsExact) return
  if (result.outline && !result.manualVisibilityNormalized) for (const axis of ['row', 'column'] as const) {
    const count = axis === 'row' ? result.rowCount : result.columnCount
    excludeHiddenRanges(sheet, axis, getSheetOutlineHiddenSpans(result.outline, axis, count))
  }
  if (filteredRows.length) {
    const ranges = mergeRanges(filteredRows.map(row => ({ start: axisIndex(row, MAX_ROWS, '筛选行'), end: row })))
    excludeHiddenRanges(sheet, 'row', ranges)
  }
}

function mergeRanges(input: readonly SheetOutlineHiddenRange[]): SheetOutlineHiddenRange[] {
  const result: SheetOutlineHiddenRange[] = []
  for (const range of [...input].sort((left, right) => left.start - right.start || left.end - right.end)) {
    const previous = result.at(-1)
    if (previous && range.start <= previous.end + 1) previous.end = Math.max(previous.end, range.end)
    else result.push({ start: range.start, end: range.end })
  }
  return result
}
function within(ranges: readonly SheetOutlineHiddenRange[], index: number): boolean {
  let low = 0, high = ranges.length
  while (low < high) { const mid = (low + high) >>> 1; if (ranges[mid]!.start <= index) low = mid + 1; else high = mid }
  return low > 0 && index <= ranges[low - 1]!.end
}
function subtract(spans: readonly SheetAxisSpan<boolean>[] | undefined, cuts: readonly SheetOutlineHiddenRange[]): SheetAxisSpan<boolean>[] | undefined {
  if (!spans) return undefined
  const result: SheetAxisSpan<boolean>[] = []
  let cutIndex = 0
  for (const span of spans) {
    if (!span.value) { result.push(cloneModel(span)); continue }
    while (cutIndex < cuts.length && cuts[cutIndex]!.end < span.start) cutIndex += 1
    let cursor = span.start
    for (let index = cutIndex; index < cuts.length && cuts[index]!.start <= span.end; index += 1) {
      const cut = cuts[index]!
      if (cut.start > cursor) result.push({ ...span, start: cursor, end: cut.start - 1 })
      cursor = Math.max(cursor, cut.end + 1)
      if (cursor > span.end) break
    }
    if (cursor <= span.end) result.push({ ...span, start: cursor, end: span.end })
  }
  return result
}
function excludeHiddenRanges(sheet: SpreadsheetSheet, axis: OutlineAxis, input: readonly SheetOutlineHiddenRange[]): void {
  const ranges = mergeRanges(input)
  if (axis === 'row') {
    sheet.hiddenRows = sheet.hiddenRows?.filter(index => !within(ranges, index))
    sheet.rowHiddenSpans = subtract(sheet.rowHiddenSpans, ranges)
  } else {
    sheet.hiddenColumns = sheet.hiddenColumns?.filter(index => !within(ranges, index))
    sheet.columnHiddenSpans = subtract(sheet.columnHiddenSpans, ranges)
  }
}
