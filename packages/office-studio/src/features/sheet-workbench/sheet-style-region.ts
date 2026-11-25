import type { CellStyle, SheetCellStylePatch, SheetStyleSource } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import { canonicalSheetAxisSpans, type SheetAxisSpan } from '../../domain/sheet-axis.ts'
import { composeSheetCellStyles, normalizeSheetStyleInput, resolveSheetStyleSource, sheetInheritedCellStyle,
  type SheetStyleContext, type SheetStyleSheet } from './sheet-style-inheritance.ts'
import { captureSheetRangeThemeRegion, sheetRangeThemeRegionStyle, type SheetRangeThemeRegionSnapshot } from './sheet-range-theme-resolver.ts'
import { assertSheetRangeThemeRegionSnapshot } from './sheet-range-theme-model.ts'

/** Relative, resolved formatting survives source closure without expanding sparse blank cells. */
export interface SheetStyleRegionSnapshot {
  defaultStyle: CellStyle
  rowStylePriority: boolean
  rowStyles?: Record<number, SheetCellStylePatch | null>
  columnStyles?: Record<number, SheetCellStylePatch | null>
  rowStyleSpans?: SheetAxisSpan<SheetCellStylePatch>[]
  columnStyleSpans?: SheetAxisSpan<SheetCellStylePatch>[]
  rangeThemeSnapshot?: SheetRangeThemeRegionSnapshot
}
export type SheetStyleRegionAxis = readonly number[] | { start: number; end: number }

/** Empty border/extension objects explicitly replace destination inherited formatting. */
export function explicitSheetCellStyle(style: CellStyle): CellStyle {
  return { ...composeSheetCellStyles(style), border: cloneModel(style.border ?? {}), metadata: cloneModel(style.metadata ?? {}) }
}

export function isDefaultSheetCellStyle(style: CellStyle): boolean {
  return JSON.stringify(explicitSheetCellStyle(style)) === JSON.stringify(explicitSheetCellStyle(composeSheetCellStyles()))
}

function axisProjection(axis: SheetStyleRegionAxis, limit: number) {
  const contiguous = 'start' in axis
  const length = contiguous ? axis.end - axis.start + 1 : axis.length
  const at = (index: number) => contiguous ? axis.start + index : axis[index]!
  if (!Number.isSafeInteger(length) || length < 1 || length > limit) throw new RangeError('样式采样行列范围无效。')
  if (contiguous) {
    if (!Number.isSafeInteger(axis.start) || !Number.isSafeInteger(axis.end) || axis.start < 0 || axis.end >= limit) throw new RangeError('样式采样行列范围无效。')
  } else for (let index = 0; index < length; index++) {
    if (!Number.isSafeInteger(at(index)) || at(index) < 0 || at(index) >= limit || index > 0 && at(index) <= at(index - 1)) throw new RangeError('样式采样坐标必须递增且位于工作表内。')
  }
  const lowerBound = (position: number) => {
    if (contiguous) return Math.max(0, Math.min(length, position - axis.start))
    let low = 0, high = length
    while (low < high) {
      const middle = (low + high) >>> 1
      if (at(middle) < position) low = middle + 1
      else high = middle
    }
    return low
  }
  return { length, at, lowerBound }
}

function captureAxis(points: Record<number, SheetStyleSource> | undefined, spans: SheetAxisSpan<SheetStyleSource>[] | undefined,
  axis: SheetStyleRegionAxis, limit: number, context: SheetStyleContext) {
  const projection = axisProjection(axis, limit), resolvedPoints: Record<number, SheetCellStylePatch | null> = {}
  const resolvedSpans: SheetAxisSpan<SheetCellStylePatch>[] = []
  for (const [index, source] of Object.entries(points ?? {})) {
    const position = Number(index), relative = projection.lowerBound(position)
    if (relative >= projection.length || projection.at(relative) !== position) continue
    const style = resolveSheetStyleSource(source, context)
    resolvedPoints[relative] = style ? cloneModel(style) : null
  }
  for (const span of canonicalSheetAxisSpans(spans ?? [])) {
    const start = projection.lowerBound(span.start), end = projection.lowerBound(span.end + 1) - 1
    if (end < start) continue
    const style = resolveSheetStyleSource(span.value, context)
    if (style) resolvedSpans.push({ start, end, value: cloneModel(style) })
  }
  return { points: Object.keys(resolvedPoints).length ? resolvedPoints : undefined, spans: resolvedSpans.length ? resolvedSpans : undefined }
}

export function captureSheetStyleRegion(sheet: SheetStyleSheet, rows: SheetStyleRegionAxis, columns: SheetStyleRegionAxis,
  context: SheetStyleContext = {}): SheetStyleRegionSnapshot {
  const row = captureAxis(sheet.rowStyles, sheet.rowStyleSpans, rows, 1_048_576, context)
  const column = captureAxis(sheet.columnStyles, sheet.columnStyleSpans, columns, 16_384, context)
  return {
    defaultStyle: explicitSheetCellStyle(composeSheetCellStyles(resolveSheetStyleSource(context.defaultStyle, context), resolveSheetStyleSource(sheet.defaultStyle, context))),
    rowStylePriority: sheet.rowStylePriority !== false,
    rowStyles: row.points, columnStyles: column.points, rowStyleSpans: row.spans, columnStyleSpans: column.spans,
    ...(sheet.rangeThemeRules?.length ? { rangeThemeSnapshot: captureSheetRangeThemeRegion(sheet, rows, columns, context) } : {}),
  }
}

export function sheetStyleRegionCellStyle(source: SheetStyleRegionSnapshot, row: number, col: number): CellStyle {
  return explicitSheetCellStyle(composeSheetCellStyles(sheetInheritedCellStyle(source, row, col),
    source.rangeThemeSnapshot ? sheetRangeThemeRegionStyle(source.rangeThemeSnapshot, row, col) : undefined))
}

export function assertSheetStyleRegion(source: SheetStyleRegionSnapshot, rows: number, columns: number): void {
  if (!source || typeof source !== 'object' || Array.isArray(source) || typeof source.rowStylePriority !== 'boolean') throw new TypeError('区域默认样式快照无效。')
  const style = (input: unknown) => {
    if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('区域样式必须是已解析的样式对象。')
    normalizeSheetStyleInput(input as SheetCellStylePatch)
  }
  style(source.defaultStyle)
  if (source.rangeThemeSnapshot !== undefined) assertSheetRangeThemeRegionSnapshot(source.rangeThemeSnapshot, rows, columns)
  for (const [points, spans, limit] of [[source.rowStyles, source.rowStyleSpans, rows], [source.columnStyles, source.columnStyleSpans, columns]] as const) {
    if (points !== undefined && (!points || typeof points !== 'object' || Array.isArray(points))) throw new TypeError('区域行列样式索引无效。')
    for (const [key, input] of Object.entries(points ?? {})) {
      const index = Number(key)
      if (!/^(0|[1-9]\d*)$/.test(key) || !Number.isSafeInteger(index) || index >= limit) throw new RangeError('区域行列样式索引越界。')
      if (input !== null) style(input)
    }
    if (spans !== undefined && !Array.isArray(spans)) throw new TypeError('区域行列样式范围无效。')
    let previousEnd = -1
    for (const span of spans ?? []) {
      if (!span || !Number.isSafeInteger(span.start) || !Number.isSafeInteger(span.end) || span.start <= previousEnd
        || span.end < span.start || span.end >= limit) throw new RangeError('区域行列样式范围必须有序、不重叠且位于采样区域内。')
      style(span.value); previousEnd = span.end
    }
  }
}
