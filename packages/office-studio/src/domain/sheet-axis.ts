/** Inclusive, sorted, disjoint axis spans keep whole-axis edits proportional to exceptions. */
export interface SheetAxisSpan<T> { start: number; end: number; value: T }
export interface SheetAxisDimensions {
  rowHeights?: Readonly<Record<number, number>>
  columnWidths?: Readonly<Record<number, number>>
  rowHeightSpans?: readonly SheetAxisSpan<number>[]
  columnWidthSpans?: readonly SheetAxisSpan<number>[]
  rowAutoHeightSpans?: readonly SheetAxisSpan<boolean>[]
  rowAutoHeights?: Readonly<Record<number, number>>
  /** Manual hides are independent from zero sizes, filter results and collapsed outlines. */
  rowHiddenSpans?: readonly SheetAxisSpan<boolean>[]
  columnHiddenSpans?: readonly SheetAxisSpan<boolean>[]
  defaultRowHeight?: number
  defaultRowAutoHeight?: boolean
  defaultColumnWidth?: number
}

function appendSpan<T>(output: SheetAxisSpan<T>[], span: SheetAxisSpan<T>): void {
  const previous = output[output.length - 1]
  if (previous && previous.end + 1 === span.start && Object.is(previous.value, span.value)) previous.end = span.end
  else output.push({ ...span })
}

/** Later declarations win. A sweep and priority heap avoid expanding or repeatedly splitting input spans. */
export function canonicalSheetAxisSpans<T>(input: readonly SheetAxisSpan<T>[]): SheetAxisSpan<T>[] {
  const starts = new Map<number, number[]>(), boundaries = new Set<number>()
  input.forEach((span, index) => {
    if (!Number.isSafeInteger(span.start) || !Number.isSafeInteger(span.end) || !Number.isSafeInteger(span.end + 1) || span.start < 0 || span.end < span.start) throw new RangeError('行列尺寸范围无效。')
    const entries = starts.get(span.start) ?? []
    entries.push(index); starts.set(span.start, entries)
    boundaries.add(span.start); boundaries.add(span.end + 1)
  })
  const positions = [...boundaries].sort((a, b) => a - b), heap: number[] = [], output: SheetAxisSpan<T>[] = []
  const push = (value: number) => {
    let index = heap.length; heap.push(value)
    while (index > 0) {
      const parent = (index - 1) >>> 1
      if (heap[parent]! >= value) break
      heap[index] = heap[parent]!; index = parent
    }
    heap[index] = value
  }
  const pop = () => {
    const last = heap.pop()
    if (!heap.length || last === undefined) return
    let index = 0
    while (index * 2 + 1 < heap.length) {
      const left = index * 2 + 1, right = left + 1
      const child = right < heap.length && heap[right]! > heap[left]! ? right : left
      if (heap[child]! <= last) break
      heap[index] = heap[child]!; index = child
    }
    heap[index] = last
  }
  positions.forEach((position, index) => {
    for (const entry of starts.get(position) ?? []) push(entry)
    while (heap.length && input[heap[0]!]!.end < position) pop()
    const next = positions[index + 1], span = heap.length ? input[heap[0]!] : undefined
    if (span && next !== undefined && next > position) appendSpan(output, { start: position, end: next - 1, value: span.value })
  })
  return output
}

export function sheetAxisSpanAt<T>(spans: readonly SheetAxisSpan<T>[] | undefined, index: number): SheetAxisSpan<T> | undefined {
  let low = 0, high = spans?.length ?? 0
  while (low < high) {
    const middle = (low + high) >>> 1
    if (spans![middle]!.start <= index) low = middle + 1
    else high = middle
  }
  const span = spans?.[low - 1]
  return span && index <= span.end ? span : undefined
}

/** Assign or remove one interval without enumerating rows or columns. */
export function assignSheetAxisSpan<T>(spans: readonly SheetAxisSpan<T>[] | undefined, start: number, end: number, value?: T): SheetAxisSpan<T>[] {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || !Number.isSafeInteger(end + 1) || start < 0 || end < start) throw new RangeError('行列尺寸范围无效。')
  const output: SheetAxisSpan<T>[] = []
  for (const span of spans ?? []) {
    if (span.end < start || span.start > end) output.push({ ...span })
    else {
      if (span.start < start) output.push({ ...span, end: start - 1 })
      if (span.end > end) output.push({ ...span, start: end + 1 })
    }
  }
  if (value !== undefined) output.push({ start, end, value })
  output.sort((a, b) => a.start - b.start)
  const merged: SheetAxisSpan<T>[] = []
  for (const span of output) appendSpan(merged, span)
  return merged
}

export function sheetManualRowHeight(sheet: SheetAxisDimensions, row: number): number | undefined {
  return sheet.rowHeights?.[row] ?? sheetAxisSpanAt(sheet.rowHeightSpans, row)?.value
}
export function sheetRowIsAutoHeight(sheet: SheetAxisDimensions, row: number): boolean {
  return sheetAxisSpanAt(sheet.rowAutoHeightSpans, row)?.value ?? (sheet.defaultRowAutoHeight !== false && sheetManualRowHeight(sheet, row) === undefined)
}
export function sheetRowHeight(sheet: SheetAxisDimensions, row: number): number {
  return (sheetRowIsAutoHeight(sheet, row) ? sheet.rowAutoHeights?.[row] : sheetManualRowHeight(sheet, row)) ?? sheet.defaultRowHeight ?? 26
}
export function sheetColumnWidth(sheet: SheetAxisDimensions, column: number): number {
  return sheet.columnWidths?.[column] ?? sheetAxisSpanAt(sheet.columnWidthSpans, column)?.value ?? sheet.defaultColumnWidth ?? (column === 0 ? 152 : 112)
}

export function effectiveSheetRowSizes(sheet: SheetAxisDimensions): { sizes: Record<number, number>; spans: SheetAxisSpan<number>[] } {
  let spans = (sheet.rowHeightSpans ?? []).map(span => ({ ...span }))
  for (const mode of sheet.rowAutoHeightSpans ?? []) if (mode.value) spans = assignSheetAxisSpan(spans, mode.start, mode.end)
  const sizes: Record<number, number> = {}
  for (const key of Object.keys(sheet.rowHeights ?? {})) sizes[Number(key)] = sheetRowHeight(sheet, Number(key))
  for (const key of Object.keys(sheet.rowAutoHeights ?? {})) if (sheetRowIsAutoHeight(sheet, Number(key))) sizes[Number(key)] = sheetRowHeight(sheet, Number(key))
  return { sizes, spans }
}

/** Sparse points and constant spans share a prefix index; a million identical rows need one span. */
export class SparseAxisIndex {
  readonly count: number
  readonly defaultSize: number
  private readonly indices: number[]
  private readonly adjustments: number[]
  private readonly values: Map<number, number>
  private readonly spans: SheetAxisSpan<number>[]
  private readonly spanAdjustments: number[]
  constructor(count: number, defaultSize: number, sizes: Readonly<Record<number, number>> = {}, hidden: readonly number[] = [], spans: readonly SheetAxisSpan<number>[] = [], hiddenSpans: readonly SheetAxisSpan<boolean>[] = []) {
    this.count = Number.isFinite(count) ? Math.max(0, Math.floor(count)) : 0
    this.defaultSize = Number.isFinite(defaultSize) ? Math.max(0, defaultSize) : 0
    const manualHidden = canonicalSheetAxisSpans(hiddenSpans.filter(span => span.end >= 0 && span.start < this.count)
      .map(span => ({ ...span, start: Math.max(0, span.start), end: Math.min(this.count - 1, span.end) }))).filter(span => span.value)
    this.spans = canonicalSheetAxisSpans([
      ...spans.filter(span => Number.isFinite(span.value) && span.end >= 0 && span.start < this.count)
        .map(span => ({ start: Math.max(0, span.start), end: Math.min(this.count - 1, span.end), value: Math.max(0, span.value) })),
      ...manualHidden.map(span => ({ ...span, value: 0 })),
    ])
    this.spanAdjustments = [0]
    for (const span of this.spans) this.spanAdjustments.push(this.spanAdjustments[this.spanAdjustments.length - 1]! + (span.end - span.start + 1) * (span.value - this.defaultSize))
    this.values = new Map(Object.entries(sizes).filter(([key, value]) => Number.isInteger(Number(key)) && Number(key) >= 0 && Number(key) < this.count && Number.isFinite(value)).map(([key, value]) => [Number(key), Math.max(0, value)]))
    for (const index of this.values.keys()) if (sheetAxisSpanAt(manualHidden, index)) this.values.set(index, 0)
    for (const index of hidden) if (Number.isInteger(index) && index >= 0 && index < this.count) this.values.set(index, 0)
    this.indices = [...this.values.keys()].sort((a, b) => a - b)
    this.adjustments = [0]
    for (const index of this.indices) this.adjustments.push(this.adjustments[this.adjustments.length - 1]! + this.values.get(index)! - (sheetAxisSpanAt(this.spans, index)?.value ?? this.defaultSize))
  }
  size(index: number): number { return this.values.get(index) ?? sheetAxisSpanAt(this.spans, index)?.value ?? this.defaultSize }
  offset(index: number): number {
    let low = 0, high = this.indices.length
    while (low < high) { const middle = (low + high) >>> 1; if (this.indices[middle]! < index) low = middle + 1; else high = middle }
    const points = this.adjustments[low]!
    low = 0; high = this.spans.length
    while (low < high) { const middle = (low + high) >>> 1; if (this.spans[middle]!.start < index) low = middle + 1; else high = middle }
    const span = this.spans[low - 1]
    const ranges = span ? this.spanAdjustments[low - 1]! + Math.min(index - span.start, span.end - span.start + 1) * (span.value - this.defaultSize) : 0
    return index * this.defaultSize + ranges + points
  }
  locate(coordinate: number): number {
    let low = 0, high = this.count
    while (low < high) { const middle = (low + high) >>> 1; if (this.offset(middle + 1) <= coordinate) low = middle + 1; else high = middle }
    return Math.min(this.count - 1, Math.max(0, low))
  }
  nextVisible(start: number, direction: -1 | 1): number | undefined {
    if (!this.count || start < 0 && direction < 0 || start >= this.count && direction > 0) return undefined
    const index = Math.max(0, Math.min(this.count - 1, start))
    if (this.size(index) > 0) return index
    const offset = this.offset(index)
    if (direction < 0 && offset <= 0) return undefined
    const candidate = this.locate(direction > 0 ? offset : offset - Math.max(Number.EPSILON, Math.abs(offset) * Number.EPSILON))
    return this.size(candidate) > 0 ? candidate : undefined
  }
  visible(start: number, end: number, minimum = 0, maximum = this.count): number[] {
    const result: number[] = []
    let index = Math.max(minimum, this.locate(Math.max(0, start)))
    while (index < Math.min(maximum, this.count) && this.offset(index) < end) {
      if (this.size(index) > 0) { result.push(index); index += 1 }
      else {
        const next = this.nextVisible(index, 1)
        if (next === undefined || next <= index) break
        index = next
      }
    }
    return result
  }
}
