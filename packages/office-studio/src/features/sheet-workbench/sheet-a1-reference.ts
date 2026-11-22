import { parseSheetAddress } from './sheet-workbench.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'

export interface SheetA1Reference { reference: string; sheetName?: string }
export type SheetA1Range = GridRange & { rangeType?: 1 | 2 }

/** Quoted names may contain ! and escaped apostrophes. External/3D references use a different contract. */
export function splitSheetA1Reference(value: string): SheetA1Reference {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError('A1 范围不能为空。')
  const text = value.trim(), separator = text.lastIndexOf('!')
  if (separator < 0) return { reference: text }
  const name = text.slice(0, separator), reference = text.slice(separator + 1).trim()
  if (!reference) throw new TypeError('A1 范围缺少坐标。')
  let sheetName: string
  if (name.startsWith("'")) {
    if (!/^'(?:[^']|'')+'$/.test(name)) throw new TypeError('A1 工作表名称引号无效。')
    sheetName = name.slice(1, -1).replaceAll("''", "'")
  } else {
    if (!name || /['!\s]/.test(name)) throw new TypeError('A1 工作表名称需要有效的引号。')
    sheetName = name
  }
  if (/[\[\]:]/.test(sheetName)) throw new TypeError('此范围入口不支持外部工作簿或三维引用。')
  return { reference, sheetName }
}

/** Whole-axis bounds are captured from the target sheet at creation, matching the fixed-version Facade. */
export function parseSheetA1Range(value: string, bounds: { rowCount?: number; columnCount?: number } = {}): SheetA1Range {
  const { reference } = splitSheetA1Reference(value)
  const rowCount = bounds.rowCount ?? 1_048_576, columnCount = bounds.columnCount ?? 16_384
  if (!Number.isSafeInteger(rowCount) || rowCount <= 0 || rowCount > 1_048_576
    || !Number.isSafeInteger(columnCount) || columnCount <= 0 || columnCount > 16_384) throw new RangeError('A1 范围的工作表行列总数无效。')
  const parsed = parseSheetAddress(reference, { rowCount, columnCount })
  if (!parsed) throw new RangeError(`A1 范围无效或超出工作表：${value}`)
  const rangeType = /^\$?[1-9]\d*:\$?[1-9]\d*$/.test(reference) ? 1
    : /^\$?[A-Z]+:\$?[A-Z]+$/i.test(reference) ? 2 : undefined
  return { ...parsed.range, ...(rangeType === undefined ? {} : { rangeType }) }
}
