import { assignSheetAxisSpan, canonicalSheetAxisSpans, type SheetAxisDimensions, type SheetAxisSpan } from './sheet-axis.ts'
import { assertSheetAxisEdit, mapSheetAxisIndex, sheetAxisMaximum, transformSheetAxisInterval, type SheetAxisEdit } from './sheet-axis-structure.ts'

export type SheetAxisStructureEdit = SheetAxisEdit & {
  /** Newly inserted axes inherit the adjacent source size, as in the fixed-version insertion command. */
  inheritedSize?: number
}
export interface SheetAxisDimensionChanges {
  rowHeights?: Record<number, number>
  columnWidths?: Record<number, number>
  rowHeightSpans?: SheetAxisSpan<number>[]
  columnWidthSpans?: SheetAxisSpan<number>[]
  rowAutoHeightSpans?: SheetAxisSpan<boolean>[]
  rowAutoHeights?: Record<number, number>
  rowHiddenSpans?: SheetAxisSpan<boolean>[]
  columnHiddenSpans?: SheetAxisSpan<boolean>[]
}

function transformPoints(values: Readonly<Record<number, number>> | undefined, edit: SheetAxisStructureEdit): Record<number, number> | undefined {
  if (!values) return undefined
  const result: Record<number, number> = {}
  for (const [key, value] of Object.entries(values)) {
    const index = Number(key)
    const moved = mapSheetAxisIndex(index, edit)
    if (moved !== undefined) result[moved] = value
    if (edit.kind === 'copy' && index >= edit.start && index < edit.start + edit.count) {
      const copied = edit.destination + index - edit.start
      if (copied < sheetAxisMaximum(edit.axis)) result[copied] = value
    }
  }
  return result
}

export function transformSheetAxisSpans<T>(spans: readonly SheetAxisSpan<T>[] | undefined, edit: SheetAxisStructureEdit): SheetAxisSpan<T>[] | undefined {
  if (!spans) return undefined
  const result: SheetAxisSpan<T>[] = []
  for (const span of spans) {
    for (const interval of transformSheetAxisInterval(span, edit, false)) result.push({ ...span, ...interval })
    if (edit.kind === 'copy') {
      const start = Math.max(span.start, edit.start), end = Math.min(span.end, edit.start + edit.count - 1)
      if (start <= end && edit.destination + start - edit.start < sheetAxisMaximum(edit.axis)) result.push({ ...span,
        start: edit.destination + start - edit.start, end: Math.min(sheetAxisMaximum(edit.axis) - 1, edit.destination + end - edit.start) })
    }
  }
  return canonicalSheetAxisSpans(result)
}

/** Transform metadata with the axes, retaining zero sizes and automatic-height caches without expanding spans. */
export function transformSheetAxisDimensions(sheet: SheetAxisDimensions, edit: SheetAxisStructureEdit): SheetAxisDimensionChanges {
  assertSheetAxisEdit(edit)
  if (edit.inheritedSize !== undefined && (!Number.isFinite(edit.inheritedSize) || edit.inheritedSize < 0)) throw new RangeError('继承的行列尺寸无效。')
  if (edit.axis === 'column') {
    let spans = transformSheetAxisSpans(sheet.columnWidthSpans, edit)
    if (edit.kind === 'insert' && edit.inheritedSize !== undefined && edit.inheritedSize !== (sheet.defaultColumnWidth ?? 112)) {
      if (edit.start < sheetAxisMaximum(edit.axis)) spans = assignSheetAxisSpan(spans, edit.start, Math.min(sheetAxisMaximum(edit.axis) - 1, edit.start + edit.count - 1), edit.inheritedSize)
    }
    return { columnWidths: transformPoints(sheet.columnWidths, edit), columnWidthSpans: spans,
      columnHiddenSpans: transformSheetAxisSpans(sheet.columnHiddenSpans, edit) }
  }
  let spans = transformSheetAxisSpans(sheet.rowHeightSpans, edit), modes = transformSheetAxisSpans(sheet.rowAutoHeightSpans, edit)
  if (edit.kind === 'insert' && edit.inheritedSize !== undefined && edit.inheritedSize !== (sheet.defaultRowHeight ?? 26)) {
    if (edit.start < sheetAxisMaximum(edit.axis)) {
      const end = Math.min(sheetAxisMaximum(edit.axis) - 1, edit.start + edit.count - 1)
      spans = assignSheetAxisSpan(spans, edit.start, end, edit.inheritedSize)
      modes = assignSheetAxisSpan(modes, edit.start, end, false)
    }
  }
  return { rowHeights: transformPoints(sheet.rowHeights, edit), rowHeightSpans: spans,
    rowAutoHeightSpans: modes, rowAutoHeights: transformPoints(sheet.rowAutoHeights, edit), rowHiddenSpans: transformSheetAxisSpans(sheet.rowHiddenSpans, edit) }
}
