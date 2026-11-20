import { defaultCellStyle, type CellScalar, type SpreadsheetCell, type SpreadsheetSheet, type WorkbookModel } from '../domain/model.ts'
import type { ParseEnvelope } from './contracts.ts'
import { array, finite, positive, record, text } from './validation.ts'

function normalizeCell(value: unknown): SpreadsheetCell {
  const raw = record(value)
  const row = Math.max(0, Math.floor(finite(raw.row)))
  const col = Math.max(0, Math.floor(finite(raw.col)))
  const style = record(raw.style)
  const font = record(style.font)
  const base = defaultCellStyle()
  const scalar: CellScalar = raw.value === null || typeof raw.value === 'string' ||
    typeof raw.value === 'boolean' || (typeof raw.value === 'number' && Number.isFinite(raw.value)) ? raw.value : null
  const formula = typeof raw.formula === 'string' && raw.formula.length ? (raw.formula.startsWith('=') ? raw.formula : '='+raw.formula) : null
  return {
    id: row+':'+col, row, col, value: scalar, formula,
    display: text(raw.display, scalar === null ? '' : String(scalar)),
    dataType: text(raw.dataType),
    cachedDataType: typeof raw.cachedDataType === 'string' ? raw.cachedDataType : null,
    richText: array(raw.richText).map(record),
    hyperlink: typeof raw.hyperlink === 'string' ? raw.hyperlink : null,
    style: {
      ...base,
      font: {
        family: text(font.name, base.font.family), size: positive(font.sizePt, base.font.size),
        bold: font.bold === true, italic: font.italic === true, underline: font.underline === true,
        strike: font.strikethrough === true, color: text(font.color, base.font.color),
      },
      fill: text(style.fill, base.fill),
      numberFormat: text(style.numberFormat, 'General'),
      horizontal: style.horizontal === 'center' || style.horizontal === 'right' || style.horizontal === 'justify' ? style.horizontal : 'left',
      vertical: style.vertical === 'top' || style.vertical === 'bottom' ? style.vertical : 'middle',
      wrap: style.wrap === true, rotation: finite(style.rotation), indent: finite(style.indent),
      shrinkToFit: style.shrinkToFit === true, border: record(style.border), metadata: style,
    },
    metadata: { source: raw },
  }
}

function normalizeSheet(value: unknown, index: number, id: string): SpreadsheetSheet {
  const raw = record(value)
  const cells: Record<string, SpreadsheetCell> = {}
  let rowCount = 100
  let columnCount = 26
  for (const source of array(raw.cells)) {
    const cell = normalizeCell(source)
    cells[cell.id] = cell
    rowCount = Math.max(rowCount, cell.row + 1)
    columnCount = Math.max(columnCount, cell.col + 1)
  }
  const dimensions = record(raw.dimensions)
  const frozen = record(raw.freezePanes)
  const sizes = (source: unknown): Record<number, number> => {
    const result: Record<number, number> = {}
    for (const item of array(source)) {
      const axis = record(item)
      if (finite(axis.index, -1) >= 0) result[Math.floor(finite(axis.index))] = positive(axis.size, 20)
    }
    return result
  }
  return {
    id: id+'-sheet-'+index, name: text(raw.name, 'Sheet'+(index+1)),
    visible: raw.visible !== 'hidden' && raw.visible !== 'veryHidden' && raw.visible !== false,
    cells,
    mergedRanges: array(raw.mergedRanges).map((value) => {
      const range = record(value)
      return { startRow: finite(range.firstRow), startCol: finite(range.firstCol), endRow: finite(range.lastRow), endCol: finite(range.lastCol) }
    }),
    columnWidths: sizes(raw.columnWidths), rowHeights: sizes(raw.rowHeights),
    frozenRows: Math.max(0, Math.floor(finite(frozen.rows))),
    frozenCols: Math.max(0, Math.floor(finite(frozen.columns))),
    showGridlines: raw.gridlines !== false,
    hiddenRows: array(raw.hiddenRows).filter((n): n is number => typeof n === 'number' && Number.isFinite(n)),
    hiddenColumns: array(raw.hiddenColumns).filter((n): n is number => typeof n === 'number' && Number.isFinite(n)),
    defaultColumnWidth: positive(raw.defaultColumnWidth, 96), defaultRowHeight: positive(raw.defaultRowHeight, 24),
    rowCount: Math.max(rowCount, finite(dimensions.lastRow, -1) + 1),
    columnCount: Math.max(columnCount, finite(dimensions.lastCol, -1) + 1),
    zoom: positive(raw.zoom, 100) / 100,
    metadata: { source: raw },
  }
}

export function normalizeWorkbook(envelope: ParseEnvelope): WorkbookModel {
  const data = envelope.payload.data
  const id = crypto.randomUUID()
  const sheets = array(data.sheets).map((sheet, index) => normalizeSheet(sheet, index, id))
  if (!sheets.length) throw new Error('解析结果不包含工作表，未用空白工作簿替代原文件')
  return {
    id, kind: 'workbook', name: envelope.fileName, sheets,
    activeSheetId: sheets[finite(data.activeSheet)]?.id ?? sheets[0]!.id,
    date1904: data.date1904 === true, modified: false,
    metadata: {
      format: envelope.format, parserVersion: envelope.parserVersion,
      definedNames: array(data.definedNames), localDefinedNames: array(data.localDefinedNames),
      properties: record(data.properties), structureProtected: data.structureProtected, textTruncated: data.textTruncated,
    },
  }
}
