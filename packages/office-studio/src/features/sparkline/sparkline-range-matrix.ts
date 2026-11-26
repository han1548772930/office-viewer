import { cloneModel } from '../../domain/serialization.ts'

/** Detached sparse matrix for the observable v1.6.0 sparkline range payload. */
export class SparklineRangeMatrix<T> {
  private matrix: Record<number, Record<number, T>>
  constructor(matrix: Record<number, Record<number, T>> = {}) { this.matrix = cloneModel(matrix) }
  getMatrix(): Record<number, Record<number, T>> { return this.matrix }
  getRow(row: number): Record<number, T> | undefined { return this.matrix[row] }
  getRowOrCreate(row: number): Record<number, T> { this.assertIndex(row); return this.matrix[row] ??= {} }
  getValue(row: number, column: number): T | undefined { return this.matrix[row]?.[column] }
  setValue(row: number, column: number, value: T): void { this.assertIndex(column); this.getRowOrCreate(row)[column] = value }
  setRow(row: number, values: Record<number, T>): void { this.assertIndex(row); this.matrix[row] = cloneModel(values) }
  realDeleteValue(row: number, column: number): void {
    const values = this.matrix[row]
    if (!values) return
    delete values[column]
    if (!Object.keys(values).length) delete this.matrix[row]
  }
  reset(): void { this.matrix = {} }
  hasValue(): boolean { return Object.values(this.matrix).some(row => Object.keys(row).length > 0) }
  forEach(callback: (row: number, values: Record<number, T>) => boolean | null | undefined | void): this {
    for (const row of this.rows()) if (callback(row, this.matrix[row]!) === false) break
    return this
  }
  forRow(callback: (row: number, columns: number[]) => boolean | null | undefined | void): this {
    return this.forEach((row, values) => callback(row, Object.keys(values).map(Number).sort((a, b) => a - b)))
  }
  forValue(callback: (row: number, column: number, value: T) => boolean | null | undefined | void): this {
    return this.forRow((row, columns) => {
      for (const column of columns) if (callback(row, column, this.matrix[row]![column]!) === false) return false
    })
  }
  getSizeOf(): number { return this.rows().length }
  getLength(): number { return (this.rows().at(-1) ?? -1) + 1 }
  getRange(): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
    const range = this.getStartEndScope()
    return { startRow: 0, endRow: range.endRow, startColumn: 0, endColumn: Math.max(0, range.endColumn) }
  }
  getStartEndScope(): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
    const rows = this.rows()
    let startColumn = Infinity, endColumn = -1
    this.forRow((_row, columns) => { if (columns.length) { startColumn = Math.min(startColumn, columns[0]!); endColumn = Math.max(endColumn, columns.at(-1)!) } })
    return { startRow: rows[0] ?? -1, endRow: rows.at(-1) ?? -1, startColumn: Number.isFinite(startColumn) ? startColumn : -1, endColumn }
  }
  getRealRange(): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
    const range = this.getStartEndScope()
    return { startRow: Math.max(0, range.startRow), endRow: Math.max(0, range.endRow), startColumn: Math.max(0, range.startColumn), endColumn: Math.max(0, range.endColumn) }
  }
  getRealRowRange(): { startRow: number; endRow: number } { const rows = this.rows(); return { startRow: rows[0] ?? 0, endRow: rows.at(-1) ?? 0 } }
  toNativeArray(): T[] { const result: T[] = []; this.forValue((_row, _col, value) => { result.push(value) }); return result }
  toArray(): T[][] { const result: T[][] = []; this.forValue((row, col, value) => { (result[row] ??= [])[col] = value }); return result }
  clone(): Record<number, Record<number, T>> { return cloneModel(this.matrix) }
  toJSON(): Record<number, Record<number, T>> { return this.clone() }
  private rows(): number[] { return Object.keys(this.matrix).map(Number).sort((a, b) => a - b) }
  private assertIndex(index: number): void { if (!Number.isSafeInteger(index) || index < 0) throw new RangeError('迷你图矩阵索引无效。') }
}
