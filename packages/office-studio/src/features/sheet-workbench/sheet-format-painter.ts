import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import type { GridRange } from '../../rendering/spreadsheet/layout.ts'

const FORMAT_KEYS = [
  'bold', 'italic', 'underline', 'strike', 'align', 'vertical', 'background', 'color',
  'fontSize', 'fontFamily', 'numberFormat', 'rawNumberFormat', 'wrap', 'border', 'rotation',
  'rotationVertical', 'indent', 'shrinkToFit',
] as const satisfies readonly (keyof ExcelCellView)[]

export interface FormatPainterSource {
  range: GridRange
  cells: readonly (readonly (ExcelCellView | undefined)[])[]
}

export interface FormatPainterPatch {
  row: number
  column: number
  cell: ExcelCellView
}

/** Tiles the captured source formatting over each target range without replacing cell content. */
export function createFormatPainterPatches(
  source: FormatPainterSource,
  targets: readonly GridRange[],
  targetCells: Readonly<Record<string, ExcelCellView | undefined>>,
  address: (row: number, column: number) => string,
): FormatPainterPatch[] {
  const sourceHeight = source.range.endRow - source.range.startRow + 1
  const sourceWidth = source.range.endColumn - source.range.startColumn + 1
  if (sourceHeight < 1 || sourceWidth < 1) return []

  const patches: FormatPainterPatch[] = []
  for (const target of targets) {
    for (let row = target.startRow; row <= target.endRow; row += 1) {
      for (let column = target.startColumn; column <= target.endColumn; column += 1) {
        const sourceRow = (row - target.startRow) % sourceHeight
        const sourceColumn = (column - target.startColumn) % sourceWidth
        const sourceCell = source.cells[sourceRow]?.[sourceColumn]
        const current = targetCells[address(row, column)] ?? { value: '' }
        const next: ExcelCellView = { ...current }
        for (const key of FORMAT_KEYS) {
          const value = sourceCell?.[key]
          if (value !== undefined) Object.assign(next, { [key]: cloneFormatValue(value) })
          else delete (next as unknown as Record<string, unknown>)[key]
        }
        patches.push({ row, column, cell: next })
      }
    }
  }
  return patches
}

function cloneFormatValue<T>(value: T): T {
  if (Array.isArray(value)) return value.map(cloneFormatValue) as T
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(Object.entries(value).map(([key, nested]) => [key, cloneFormatValue(nested)])) as T
  }
  return value
}
