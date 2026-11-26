import { cloneModel } from '../../domain/serialization.ts'

export interface SparklineMatrixRange { startRow: number; endRow: number; startColumn: number; endColumn: number }
export type SparklineMatrixData<T> = Record<number, Record<number, T>>
type MatrixCallback = boolean | null | undefined | void
const MAX_DENSE_MATRIX_CELLS = 1_048_576

/** Detached sparse matrix for the observable v1.6.0 sparkline range payload. */
export class SparklineRangeMatrix<T> {
  private matrix: SparklineMatrixData<T>
  constructor(matrix: SparklineMatrixData<T> = {}) {
    if (!matrix || typeof matrix !== 'object' || Array.isArray(matrix)) throw new TypeError('迷你图矩阵必须为稀疏对象。')
    for (const [row, values] of Object.entries(matrix)) { this.assertKey(row); this.assertRow(values) }
    this.matrix = cloneModel(matrix)
  }
  static MakeObjectMatrixSize<T>(size: number): SparklineRangeMatrix<T> {
    if (!Number.isSafeInteger(size) || size < 1) throw new RangeError('迷你图矩阵大小必须为正整数。')
    return new SparklineRangeMatrix<T>({ [size - 1]: {} })
  }
  getMatrix(): SparklineMatrixData<T> { return this.matrix }
  getRow(row: number): Record<number, T> | undefined { return this.matrix[row] }
  getRowOrCreate(row: number): Record<number, T> { this.assertIndex(row); return this.matrix[row] ??= {} }
  getValue(row: number, column: number): T | undefined { return this.matrix[row]?.[column] }
  setValue(row: number, column: number, value: T): void { this.assertIndex(column); this.getRowOrCreate(row)[column] = value }
  /** Rows supplied to the detached matrix retain the upstream aliasing contract. */
  setRow(row: number, values: Record<number, T>): void { this.assertIndex(row); this.assertRow(values); this.matrix[row] = values }
  swapRow(source: number, target: number): void {
    this.assertIndex(source); this.assertIndex(target)
    const first = this.matrix[source], second = this.matrix[target]
    if (second === undefined) delete this.matrix[source]; else this.matrix[source] = second
    if (first === undefined) delete this.matrix[target]; else this.matrix[target] = first
  }
  moveRows(start: number, count: number, target: number): void { this.moveAxis(this.matrix, start, count, target) }
  moveColumns(start: number, count: number, target: number): void {
    this.assertMove(start, count, target)
    this.forEach((_row, values) => { this.moveAxis(values, start, count, target) })
  }
  insertRows(start: number, count: number): void { this.insertAxis(this.matrix, start, count) }
  insertColumns(start: number, count: number): void {
    this.assertSpan(start, count)
    this.forEach((_row, values) => { this.assertShift(values, start, count) })
    this.forEach((_row, values) => { this.insertAxis(values, start, count) })
  }
  removeRows(start: number, count: number): void { this.removeAxis(this.matrix, start, count) }
  removeColumns(start: number, count: number): void {
    this.assertSpan(start, count)
    this.forEach((_row, values) => { this.removeAxis(values, start, count) })
  }
  realDeleteValue(row: number, column: number): void {
    const values = this.matrix[row]
    if (!values) return
    delete values[column]
    if (!Object.keys(values).length) delete this.matrix[row]
  }
  reset(): void { this.matrix = {} }
  hasValue(): boolean { return Object.values(this.matrix).some(row => Object.keys(row).length > 0) }
  forEach(callback: (row: number, values: Record<number, T>) => MatrixCallback): this {
    for (const row of this.rows()) if (callback(row, this.matrix[row]!) === false) break
    return this
  }
  forRow(callback: (row: number, columns: number[]) => MatrixCallback): this {
    return this.forEach((row, values) => callback(row, this.keys(values)))
  }
  forValue(callback: (row: number, column: number, value: T) => MatrixCallback): this {
    return this.forRow((row, columns) => {
      for (const column of columns) if (callback(row, column, this.matrix[row]![column]!) === false) return false
    })
  }
  getSizeOf(): number { return this.rows().length }
  getLength(): number { return (this.rows().at(-1) ?? 0) + 1 }
  getRange(): SparklineMatrixRange {
    const range = this.getStartEndScope()
    return { startRow: 0, endRow: this.getLength() - 1, startColumn: 0, endColumn: Math.max(0, range.endColumn) }
  }
  getStartEndScope(): SparklineMatrixRange {
    const rows = this.rows()
    let startColumn = Infinity, endColumn = -1
    this.forRow((_row, columns) => { if (columns.length) { startColumn = Math.min(startColumn, columns[0]!); endColumn = Math.max(endColumn, columns.at(-1)!) } })
    return { startRow: rows[0] ?? -1, endRow: rows.at(-1) ?? -1, startColumn: Number.isFinite(startColumn) ? startColumn : -1, endColumn }
  }
  getRealRange(): SparklineMatrixRange {
    const range = this.getStartEndScope()
    return { startRow: Math.max(0, range.startRow), endRow: Math.max(0, range.endRow), startColumn: Math.max(0, range.startColumn), endColumn: Math.max(0, range.endColumn) }
  }
  getDataRange(): SparklineMatrixRange {
    const range = this.getStartEndScope()
    return { startRow: Math.max(0, range.startRow), endRow: range.endRow, startColumn: Math.max(0, range.startColumn), endColumn: Math.max(0, range.endColumn) }
  }
  /** Fragment is dense and rebased, including explicitly present undefined cells. */
  getFragment(startRow: number, endRow: number, startColumn: number, endColumn: number): SparklineRangeMatrix<T> {
    this.assertBounds(startRow, endRow, startColumn, endColumn)
    this.assertDenseSize(endRow - startRow + 1, endColumn - startColumn + 1)
    const result = new SparklineRangeMatrix<T>()
    for (let row = startRow; row <= endRow; row++) {
      const values: Record<number, T> = {}
      for (let column = startColumn; column <= endColumn; column++) values[column - startColumn] = this.getValue(row, column) as T
      result.setRow(row - startRow, values)
    }
    return result
  }
  /** The fixed source includes truthy values only and deep-copies each selected value. */
  getSlice(startRow: number, endRow: number, startColumn: number, endColumn: number): SparklineRangeMatrix<T> {
    this.assertBounds(startRow, endRow, startColumn, endColumn)
    const result = new SparklineRangeMatrix<T>()
    this.forValue((row, column, value) => {
      if (value && row >= startRow && row <= endRow && column >= startColumn && column <= endColumn) result.setValue(row, column, cloneModel(value))
    })
    return result
  }
  getSliceDataAndCellCountByRows(startRow: number, endRow: number): { sliceData: SparklineRangeMatrix<T>; cellCount: number } {
    this.assertIndex(startRow); this.assertIndex(endRow)
    if (endRow < startRow) throw new RangeError('迷你图矩阵切片终点不能位于起点之前。')
    const sliceData = new SparklineRangeMatrix<T>()
    let cellCount = 0
    this.forEach((row, values) => {
      if (row >= startRow && row <= endRow) { sliceData.setRow(row, values); cellCount += Object.keys(values).length }
    })
    return { sliceData, cellCount }
  }
  getRealRowRange(): { startRow: number; endRow: number } { const rows = this.rows(); return { startRow: rows[0] ?? 0, endRow: rows.at(-1) ?? 0 } }
  toNativeArray(): T[] { const result: T[] = []; this.forValue((_row, _col, value) => { result.push(value) }); return result }
  toArray(): T[][] {
    const result: T[][] = []
    this.forEach((row, values) => {
      this.assertArrayIndex(row); const columns: T[] = []; result[row] = columns
      for (const column of this.keys(values)) { this.assertArrayIndex(column); columns[column] = values[column]! }
    })
    return result
  }
  toFullArray(): T[][] {
    const { endRow, endColumn } = this.getRange()
    this.assertDenseSize(endRow + 1, endColumn + 1)
    const result: T[][] = Array.from({ length: endRow + 1 }, () => new Array<T>(endColumn + 1).fill(undefined as T))
    this.forValue((row, column, value) => { result[row]![column] = value })
    return result
  }
  getArrayData(): SparklineMatrixData<T> {
    const range = this.getRealRange(), result = new SparklineRangeMatrix<T>()
    this.forValue((row, column, value) => { result.setValue(row - range.startRow, column - range.startColumn, value) })
    return result.clone()
  }
  getDiscreteRanges(): SparklineMatrixRange[] {
    const ranges: SparklineMatrixRange[] = []
    this.forValue((row, column) => {
      const range = ranges.find(range => row >= range.startRow && row <= range.endRow + 1 && column >= range.startColumn && column <= range.endColumn + 1)
      if (range) { range.endRow = Math.max(row, range.endRow); range.endColumn = Math.max(column, range.endColumn) }
      else ranges.push({ startRow: row, endRow: row, startColumn: column, endColumn: column })
    })
    return ranges
  }
  /** Merge only replaces existing cells with a non-null value, as ObjectMatrix does. */
  merge(other: SparklineRangeMatrix<T | null | undefined>): void {
    this.forValue((row, column) => { const value = other.getValue(row, column); if (value != null) this.setValue(row, column, value) })
  }
  concatRows(other: SparklineRangeMatrix<T>): void { other.forEach((row, values) => { this.setRow(row, values) }) }
  clone(): SparklineMatrixData<T> {
    const result: SparklineMatrixData<T> = {}
    this.forEach((row, values) => {
      const columns: Record<number, T> = {}; result[row] = columns
      for (const column of this.keys(values)) if (values[column] !== undefined) columns[column] = cloneModel(values[column]!)
    })
    return result
  }
  toJSON(): SparklineMatrixData<T> { return this.clone() }
  private rows(): number[] { return this.keys(this.matrix) }
  private keys(values: object): number[] { return Object.keys(values).map(key => { this.assertKey(key); return Number(key) }).sort((a, b) => a - b) }
  private assertKey(key: string): void { const index = Number(key); this.assertIndex(index); if (String(index) !== key) throw new RangeError('迷你图矩阵索引必须为规范整数。') }
  private assertRow(values: Record<number, T>): void {
    if (!values || typeof values !== 'object' || Array.isArray(values)) throw new TypeError('迷你图矩阵行必须为稀疏对象。')
    for (const key of Object.keys(values)) this.assertKey(key)
  }
  private assertIndex(index: number): void { if (!Number.isSafeInteger(index) || index < 0 || index >= Number.MAX_SAFE_INTEGER) throw new RangeError('迷你图矩阵索引无效。') }
  private assertArrayIndex(index: number): void { if (index > 4_294_967_294) throw new RangeError('迷你图矩阵索引超出 JavaScript 数组范围。') }
  private assertSpan(start: number, count: number): void { this.assertIndex(start); this.assertIndex(count); this.assertIndex(start + count) }
  private assertBounds(startRow: number, endRow: number, startColumn: number, endColumn: number): void {
    for (const index of [startRow, endRow, startColumn, endColumn]) this.assertIndex(index)
    if (endRow < startRow || endColumn < startColumn) throw new RangeError('迷你图矩阵切片终点不能位于起点之前。')
  }
  private assertDenseSize(rows: number, columns: number): void {
    if (!Number.isSafeInteger(rows * columns) || rows * columns > MAX_DENSE_MATRIX_CELLS) throw new RangeError('迷你图矩阵稠密输出超出 1048576 个单元格的内存预算，请使用稀疏切片。')
  }
  private assertShift<V>(values: Record<number, V>, start: number, count: number): void {
    for (const index of this.keys(values)) if (index >= start) this.assertIndex(index + count)
  }
  private insertAxis<V>(values: Record<number, V>, start: number, count: number): void {
    this.assertSpan(start, count); this.assertShift(values, start, count)
    if (!count) return
    for (const index of this.keys(values).reverse()) if (index >= start) { values[index + count] = values[index]!; delete values[index] }
  }
  private removeAxis<V>(values: Record<number, V>, start: number, count: number): void {
    this.assertSpan(start, count)
    if (!count) return
    const end = start + count
    for (const index of this.keys(values)) {
      if (index < start) continue
      const value = values[index]; delete values[index]
      if (index >= end && value !== undefined) values[index - count] = value
    }
  }
  private assertMove(start: number, count: number, target: number): void {
    this.assertSpan(start, count); this.assertIndex(target); this.assertIndex(target + count)
    if (start <= target && start + count > target) throw new RangeError('迷你图矩阵移动的目标不能位于源区间内部。')
  }
  private moveAxis<V>(values: Record<number, V>, start: number, count: number, target: number): void {
    this.assertMove(start, count, target)
    if (!count) return
    const end = start + count, moved: Array<[number, V]> = []
    // Remap only stored indices; no synchronous scan across empty worksheet coordinates.
    for (const index of this.keys(values)) {
      const destination = index >= start && index < end ? (target < start ? target : target - count) + index - start
        : target < start && index >= target && index < start ? index + count
          : target >= end && index >= end && index < target ? index - count : index
      if (destination === index) continue
      const value = values[index]; delete values[index]
      if (value !== undefined) moved.push([destination, value])
    }
    for (const [index, value] of moved) values[index] = value
  }
}
