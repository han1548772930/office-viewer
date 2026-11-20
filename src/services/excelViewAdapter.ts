import type { WorkbookModel, SpreadsheetCell } from '../domain/model'

export interface ExcelCellView {
  value: string
  formula?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  align?: 'left' | 'center' | 'right'
  background?: string
  color?: string
  fontSize?: number
  numberFormat?: 'general' | 'number' | 'percent'
  cachedValue?: string
  fontFamily?: string
}

export interface ExcelSheetView {
  id: string
  name: string
  cells: Record<string, ExcelCellView>
  rowCount: number
  columnCount: number
  merges?: string[]
  freeze?: { rows: number; columns: number }
}

export interface ExcelWorkbookView {
  name: string
  sheets: ExcelSheetView[]
}

const cellKey = (row: number, col: number): string => {
  let result = ''
  let index = col
  while (index >= 0) {
    result = String.fromCharCode((index % 26) + 65) + result
    index = Math.floor(index / 26) - 1
  }
  return `${result}${row + 1}`
}

const parseCellKey = (key: string): { row: number; col: number } | undefined => {
  const match = /^([A-Z]+)(\d+)$/.exec(key.toUpperCase())
  if (!match) return undefined
  let col = 0
  for (const character of match[1]) col = col * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, col: col - 1 }
}

const toViewCell = (cell: SpreadsheetCell): ExcelCellView => ({
  value: cell.formula ? '' : cell.display,
  formula: cell.formula ?? undefined,
  bold: cell.style.font.bold,
  italic: cell.style.font.italic,
  underline: cell.style.font.underline,
  align: cell.style.horizontal === 'justify' ? 'left' : cell.style.horizontal,
  background: cell.style.fill,
  color: cell.style.font.color,
  fontSize: cell.style.font.size,
  fontFamily: cell.style.font.family,
  numberFormat: cell.style.numberFormat === 'General' ? 'general' : cell.style.numberFormat.includes('%') ? 'percent' : 'number',
  cachedValue: cell.formula ? cell.display : undefined,
})

export function toExcelWorkbookView(workbook: WorkbookModel): ExcelWorkbookView {
  return {
    name: workbook.name,
    sheets: workbook.sheets.map((sheet) => {
      const cells: Record<string, ExcelCellView> = {}
      for (const cell of Object.values(sheet.cells)) cells[cellKey(cell.row, cell.col)] = toViewCell(cell)
      const maxRow = Math.max(100, ...Object.values(sheet.cells).map((cell) => cell.row + 1))
      const maxCol = Math.max(12, ...Object.values(sheet.cells).map((cell) => cell.col + 1))
      return {
        id: sheet.id,
        name: sheet.name,
        cells,
        rowCount: maxRow,
        columnCount: maxCol,
        merges: sheet.mergedRanges.map((range) => `${cellKey(range.startRow, range.startCol)}:${cellKey(range.endRow, range.endCol)}`),
        freeze: { rows: sheet.frozenRows, columns: sheet.frozenCols },
      }
    }),
  }
}

export function applyExcelWorkbookView(workbook: WorkbookModel, view: ExcelWorkbookView): WorkbookModel {
  return {
    ...workbook,
    name: view.name,
    sheets: view.sheets.map((source, sheetIndex) => {
      const original = workbook.sheets.find((sheet) => sheet.id === source.id) ?? workbook.sheets[sheetIndex]
      const cells: Record<string, SpreadsheetCell> = {}
      for (const [key, value] of Object.entries(source.cells)) {
        const parsed = parseCellKey(key)
        if (!parsed) continue
        const originalCell = original?.cells[`${parsed.row}:${parsed.col}`]
        cells[`${parsed.row}:${parsed.col}`] = {
          id: `${parsed.row}:${parsed.col}`,
          row: parsed.row,
          col: parsed.col,
          value: value.formula ? (value.cachedValue ?? null) : value.value,
          formula: value.formula ?? null,
          display: value.formula ? (value.cachedValue ?? value.formula) : value.value,
          style: {
            font: {
              family: value.fontFamily ?? originalCell?.style.font.family ?? 'Aptos',
              size: value.fontSize ?? originalCell?.style.font.size ?? 11,
              bold: value.bold ?? originalCell?.style.font.bold ?? false,
              italic: value.italic ?? originalCell?.style.font.italic ?? false,
              underline: value.underline ?? originalCell?.style.font.underline ?? false,
              strike: originalCell?.style.font.strike ?? false,
              color: value.color ?? originalCell?.style.font.color ?? '#1f2937',
            },
            fill: value.background ?? originalCell?.style.fill ?? '#ffffff',
            horizontal: value.align ?? originalCell?.style.horizontal ?? 'left',
            vertical: originalCell?.style.vertical ?? 'middle',
            wrap: originalCell?.style.wrap ?? false,
            numberFormat: value.numberFormat === 'percent' ? '0%' : value.numberFormat === 'number' ? '#,##0.00' : 'General',
          },
        }
      }
      return {
        ...(original ?? {
          id: source.id,
          visible: true,
          mergedRanges: [],
          columnWidths: {},
          rowHeights: {},
          frozenRows: 0,
          frozenCols: 0,
          showGridlines: true,
        }),
        id: source.id,
        name: source.name,
        cells,
        mergedRanges: source.merges?.map((merge) => {
          const [start, end] = merge.split(':')
          const first = parseCellKey(start)
          const last = parseCellKey(end ?? start)
          return { startRow: first?.row ?? 0, startCol: first?.col ?? 0, endRow: last?.row ?? first?.row ?? 0, endCol: last?.col ?? first?.col ?? 0 }
        }) ?? [],
        frozenRows: source.freeze?.rows ?? original?.frozenRows ?? 0,
        frozenCols: source.freeze?.columns ?? original?.frozenCols ?? 0,
      }
    }),
  }
}
