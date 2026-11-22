import type { CellRange } from './model.ts'

export type SheetAxis = 'row' | 'column'
export interface SheetAxisInterval { start: number; end: number }
export type SheetAxisEdit =
  | { axis: SheetAxis; kind: 'insert' | 'delete'; start: number; count: number }
  | { axis: SheetAxis; kind: 'move' | 'copy'; start: number; count: number; destination: number }

export const sheetAxisMaximum = (axis: SheetAxis): number => axis === 'row' ? 1_048_576 : 16_384
export const sheetRangeAxisInterval = (range: CellRange, axis: SheetAxis): SheetAxisInterval => axis === 'row'
  ? { start: range.startRow, end: range.endRow } : { start: range.startCol, end: range.endCol }
export const sheetRangeWithAxisInterval = (range: CellRange, axis: SheetAxis, interval: SheetAxisInterval): CellRange => axis === 'row'
  ? { ...range, startRow: interval.start, endRow: interval.end } : { ...range, startCol: interval.start, endCol: interval.end }

export function assertSheetAxisEdit(edit: SheetAxisEdit): void {
  const maximum = sheetAxisMaximum(edit.axis)
  if (edit.axis !== 'row' && edit.axis !== 'column' || !['insert', 'delete', 'move', 'copy'].includes(edit.kind)
    || !Number.isSafeInteger(edit.start) || edit.start < 0 || edit.start > maximum
    || !Number.isSafeInteger(edit.count) || edit.count <= 0 || edit.count > maximum
    || edit.kind !== 'insert' && edit.start + edit.count > maximum) throw new RangeError('行列结构参数无效。')
  if ((edit.kind === 'move' || edit.kind === 'copy') && (!Number.isSafeInteger(edit.destination)
    || edit.destination < 0 || edit.destination > maximum)) throw new RangeError('行列目标位置无效。')
  if (edit.kind === 'move' && edit.destination > edit.start && edit.destination < edit.start + edit.count) {
    throw new RangeError('移动目标不能位于源行列内部。')
  }
}

export function sheetAxisInsertionEdit(edit: SheetAxisEdit): Extract<SheetAxisEdit, { kind: 'insert' | 'delete' }> | undefined {
  return edit.kind === 'insert' || edit.kind === 'delete' ? edit : edit.kind === 'copy'
    ? { axis: edit.axis, kind: 'insert', start: edit.destination, count: edit.count } : undefined
}

interface AxisSegment extends SheetAxisInterval { offset: number }
/** Piecewise permutation follows the fixed-version ref-range resource contract. */
export function sheetAxisEditSegments(edit: SheetAxisEdit): AxisSegment[] {
  assertSheetAxisEdit(edit)
  const end = sheetAxisMaximum(edit.axis) - 1, insertion = sheetAxisInsertionEdit(edit)
  let segments: AxisSegment[]
  if (insertion) segments = insertion.kind === 'insert'
    ? [{ start: 0, end: insertion.start - 1, offset: 0 }, { start: insertion.start, end, offset: insertion.count }]
    : [{ start: 0, end: insertion.start - 1, offset: 0 }, { start: insertion.start + insertion.count, end, offset: -insertion.count }]
  else if (edit.kind === 'move' && edit.destination < edit.start) segments = [
    { start: 0, end: edit.destination - 1, offset: 0 },
    { start: edit.destination, end: edit.start - 1, offset: edit.count },
    { start: edit.start, end: edit.start + edit.count - 1, offset: edit.destination - edit.start },
    { start: edit.start + edit.count, end, offset: 0 },
  ]
  else if (edit.kind === 'move' && edit.destination > edit.start + edit.count) segments = [
    { start: 0, end: edit.start - 1, offset: 0 },
    { start: edit.start, end: edit.start + edit.count - 1, offset: edit.destination - edit.start - edit.count },
    { start: edit.start + edit.count, end: edit.destination - 1, offset: -edit.count },
    { start: edit.destination, end, offset: 0 },
  ]
  else segments = [{ start: 0, end, offset: 0 }]
  return segments.filter(segment => segment.end >= segment.start)
}

export function mapSheetAxisIndex(index: number, edit: SheetAxisEdit): number | undefined {
  if (!Number.isSafeInteger(index) || index < 0 || index >= sheetAxisMaximum(edit.axis)) throw new RangeError('行列坐标无效。')
  const segment = sheetAxisEditSegments(edit).find(segment => index >= segment.start && index <= segment.end)
  const result = segment ? index + segment.offset : undefined
  return result !== undefined && result < sheetAxisMaximum(edit.axis) ? result : undefined
}

export function mergeSheetAxisIntervals(intervals: readonly SheetAxisInterval[]): SheetAxisInterval[] {
  const result: SheetAxisInterval[] = []
  for (const interval of [...intervals].sort((left, right) => left.start - right.start || left.end - right.end)) {
    const previous = result.at(-1)
    if (previous && interval.start <= previous.end + 1) previous.end = Math.max(previous.end, interval.end)
    else result.push({ ...interval })
  }
  return result
}

/** Resource membership can split on a move; inserted axes extend an interior interval. */
export function transformSheetAxisInterval(interval: SheetAxisInterval, edit: SheetAxisEdit, expandInsertion = true): SheetAxisInterval[] {
  const maximum = sheetAxisMaximum(edit.axis), result: SheetAxisInterval[] = []
  if (!Number.isSafeInteger(interval.start) || !Number.isSafeInteger(interval.end)
    || interval.start < 0 || interval.end < interval.start || interval.end >= maximum) throw new RangeError('行列资源范围无效。')
  for (const segment of sheetAxisEditSegments(edit)) {
    const start = Math.max(interval.start, segment.start), end = Math.min(interval.end, segment.end)
    if (start <= end && start + segment.offset < maximum) result.push({ start: start + segment.offset, end: Math.min(maximum - 1, end + segment.offset) })
  }
  const insertion = sheetAxisInsertionEdit(edit)
  if (expandInsertion && insertion?.kind === 'insert' && interval.start < insertion.start && interval.end >= insertion.start) {
    result.push({ start: insertion.start, end: Math.min(maximum - 1, insertion.start + insertion.count - 1) })
  }
  return mergeSheetAxisIntervals(result)
}

export function transformSheetAxisRange(range: CellRange, edit: SheetAxisEdit, expandInsertion = true): CellRange[] {
  return transformSheetAxisInterval(sheetRangeAxisInterval(range, edit.axis), edit, expandInsertion)
    .map(interval => sheetRangeWithAxisInterval(range, edit.axis, interval))
}

/**
 * Formula ranges retain one interval. This rewrites the observable algorithm in
 * Univer v1.6.0 sheets/services/ref-range/util.ts::handleBaseMoveRowsCols.
 * Resource membership uses transformSheetAxisInterval instead.
 */
export function transformSheetAxisReferenceInterval(interval: SheetAxisInterval, edit: SheetAxisEdit): SheetAxisInterval | undefined {
  assertSheetAxisEdit(edit)
  if (!Number.isSafeInteger(interval.start) || !Number.isSafeInteger(interval.end) || interval.start < 0
    || interval.end < interval.start || interval.end >= sheetAxisMaximum(edit.axis)) throw new RangeError('公式行列范围无效。')
  const insertion = sheetAxisInsertionEdit(edit)
  if (insertion) {
    const transformed = transformSheetAxisInterval(interval, insertion)
    if (!transformed.length) return undefined
    if (insertion.kind === 'insert' && interval.end + insertion.count >= sheetAxisMaximum(edit.axis)
      && interval.end >= insertion.start && !(interval.start === 0 && interval.end === sheetAxisMaximum(edit.axis) - 1)) return undefined
    return { start: transformed[0]!.start, end: transformed.at(-1)!.end }
  }
  if (edit.kind !== 'move' || edit.destination === edit.start || edit.destination === edit.start + edit.count) return { ...interval }
  const fromEnd = edit.start + edit.count - 1, forward = edit.destination > edit.start
  const targetStart = edit.destination - (forward ? edit.count : 0), targetEnd = targetStart + edit.count - 1
  const intersectionStart = Math.max(edit.start, interval.start), intersectionEnd = Math.min(fromEnd, interval.end)
  const intersectionCount = Math.max(0, intersectionEnd - intersectionStart + 1)
  const insideSource = intersectionCount === interval.end - interval.start + 1
  let start = interval.start, end = interval.end
  if (fromEnd < start) { start -= edit.count; end -= edit.count }
  else if (intersectionCount) {
    if (insideSource) { start = targetStart + interval.start - edit.start; end = start + interval.end - interval.start }
    else if (intersectionStart > edit.start) {
      if (forward) { start -= edit.count; end -= intersectionCount + edit.count }
      else end -= intersectionCount
    } else if (forward) end -= intersectionCount
    else if (start > edit.start && end > fromEnd) { start -= edit.count; end -= edit.count + intersectionCount }
    else end -= intersectionCount
  }
  if (!insideSource) {
    if (targetStart <= start) { start += edit.count; end += edit.count }
    else if (targetStart <= end && targetEnd >= start) {
      if (!forward) {
        if (start < targetStart && end > targetStart) end += edit.count
        else if (start >= targetEnd || start >= targetStart && start <= targetEnd) { start += edit.count; end += edit.count }
      } else if (targetEnd <= start || targetStart <= start && targetEnd >= start) { start += edit.count; end += edit.count }
      else if (targetStart >= start && targetStart <= end) end += edit.count
    }
  }
  return start >= 0 && end >= start && end < sheetAxisMaximum(edit.axis) ? { start, end } : undefined
}
