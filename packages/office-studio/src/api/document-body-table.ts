import type { DocumentBlock, DocumentDefaultStyle } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { createParagraph, documentId, runText, type DocumentCell, type DocumentTable, type DocumentTableBorder, type DocumentTableCellMargin, type DocumentTableCellStyle, type DocumentTableColumnTypeConfig } from '../features/document/model/document-model'
import { documentTableGrid } from '../features/document/model/document-table-grid'
import { assertDocumentTableColumnType, documentTableRowProperties } from '../features/document/model/document-table-style'
import { bodyBoolean, bodyColor, bodyLength, bodyNumber, bodyRecord, type UniverBodyRecord } from './document-body-style'
import { documentBodyArray, documentBodyIdentity } from './document-body-types'
import { documentBodyOfficeSource, documentBodySourceMetadata } from './document-body-source'

const px = (v: number) => ({ v, u: 3 })
const size = (v: number) => ({ type: 1, width: px(v) })
function integer(value: unknown, fallback: number, label: string, max: number): number {
  const number = bodyNumber(value) ?? fallback
  if (!Number.isSafeInteger(number) || number < 0 || number > max) throw new RangeError(`${label} 超出表格范围。`)
  return number
}
function width(value: unknown, defaults: DocumentDefaultStyle): number | undefined {
  if (value === undefined || value === null) return undefined
  const raw = bodyRecord(value)
  return bodyLength(raw.width ?? value, defaults.text.fontSize)
}
function margins(value: unknown, defaults: DocumentDefaultStyle): DocumentTableCellMargin | undefined {
  if (value === undefined || value === null) return undefined
  const raw = bodyRecord(value), length = (key: string) => bodyLength(raw[key], defaults.text.fontSize) ?? 0
  const result = { top: length('top'), right: length('end'), bottom: length('bottom'), left: length('start') }
  if (Object.values(result).some(number => number < 0)) throw new RangeError('表格边距不能为负。')
  return result
}
function border(value: unknown, defaults: DocumentDefaultStyle): DocumentTableBorder | undefined {
  if (value === undefined || value === null) return undefined
  const raw = bodyRecord(value), lineWidth = bodyLength(raw.width, defaults.text.fontSize) ?? 0
  if (lineWidth < 0) throw new RangeError('表格边框宽度不能为负。')
  const dash = bodyNumber(raw.dashStyle) ?? 1
  return { width: lineWidth, color: bodyColor(raw.color) ?? 'auto', style: dash === 2 ? 'dot' : dash === 3 ? 'dash' : 'solid' }
}
function cellStyle(raw: UniverBodyRecord, defaults: DocumentDefaultStyle): DocumentTableCellStyle {
  const fill = bodyColor(raw.backgroundColor), margin = margins(raw.margin, defaults), align = bodyNumber(raw.vAlign)
  const borders: NonNullable<DocumentTableCellStyle['borders']> = {}
  for (const [from, to] of [['borderTop', 'top'], ['borderRight', 'right'], ['borderBottom', 'bottom'], ['borderLeft', 'left']] as const) {
    const line = border(raw[from], defaults); if (line) borders[to] = line
  }
  const extension = documentBodyOfficeSource(raw, '表格单元格')
  return { ...documentBodySourceMetadata(extension.style, '表格单元格', ['univerBodyCellSource', 'nativeContinuationContent', 'fill', 'background', 'margin', 'verticalAlign', 'borders']),
    univerBodyCellSource: cloneModel(raw), ...(fill === undefined ? {} : { fill }), ...(margin ? { margin } : {}),
    ...(align === undefined ? {} : { verticalAlign: align === 3 ? 'center' : align === 4 ? 'bottom' : 'top' }),
    ...(Object.keys(borders).length ? { borders } : {}) }
}

function columnType(value: unknown): DocumentTableColumnTypeConfig {
  const raw = documentBodySourceMetadata(value, '表格列类型'), type = raw.type
  if (type !== 'text' && type !== 'number' && type !== 'date' && type !== 'checkbox' && type !== 'dropdown') throw new TypeError('表格列类型无效。')
  const config: DocumentTableColumnTypeConfig = { type }
  if (raw.dropdownOptions !== undefined) config.dropdownOptions = documentBodyArray(raw.dropdownOptions, '表格下拉选项').map(option => {
    if (typeof option.label !== 'string' || option.color !== undefined && typeof option.color !== 'string') throw new TypeError('表格下拉选项标签或颜色无效。')
    return { id: documentBodyIdentity(option.id, '表格下拉选项'), label: option.label, ...(typeof option.color === 'string' ? { color: option.color } : {}) }
  })
  assertDocumentTableColumnType(config)
  return config
}

/** Public tables store every covered cell; the editor stores only real merge owners. */
export function documentTableFromBody(raw: UniverBodyRecord, contents: DocumentBlock[][][], defaults: DocumentDefaultStyle): DocumentTable {
  const rows = documentBodyArray(raw.tableRows, 'tableRows'), columns = documentBodyArray(raw.tableColumns, 'tableColumns')
  const extension = documentBodyOfficeSource(raw, '表格'), fill = bodyColor(raw.backgroundColor)
  if (!rows.length || rows.length !== contents.length || rows.length > 10_000) throw new RangeError('表格行定义与内容流不匹配。')
  const table: DocumentTable = { rows: [], columnWidths: columns.map(column => width(column.size, defaults) ?? 80),
    style: { ...documentBodySourceMetadata(extension.style, '表格', ['univerBodyTableSource', 'fill', 'background', 'align', 'indent', 'layout', 'width', 'cellMargin', 'borders']),
      univerBodyTableSource: cloneModel({ ...raw, tableRows: undefined }), ...(fill === undefined ? {} : { fill }),
      align: raw.align === 1 ? 'center' : raw.align === 2 ? 'right' : 'left',
      indent: bodyLength(raw.indent, defaults.text.fontSize) ?? 0, layout: raw.layout === 1 ? 'fixed' : 'auto',
      ...(width(raw.size, defaults) === undefined ? {} : { width: width(raw.size, defaults) }),
      ...(raw.cellMargin === undefined ? {} : { cellMargin: margins(raw.cellMargin, defaults) }), borders: {} }, headerRowCount: 0 }
  const defaultsBorders = bodyRecord(raw.borderDefaults)
  for (const side of ['top', 'right', 'bottom', 'left', 'insideHorizontal', 'insideVertical'] as const) {
    const line = border(defaultsBorders[side], defaults); if (line) table.style!.borders![side] = line
  }
  if (extension.titleRow !== undefined) { if (typeof extension.titleRow !== 'boolean') throw new TypeError('表格标题行状态无效。'); table.titleRow = extension.titleRow }
  if (extension.columnTypes !== undefined) {
    table.columnTypes = {}
    for (const [key, value] of Object.entries(documentBodySourceMetadata(extension.columnTypes, '表格列类型'))) {
      if (!/^(0|[1-9]\d*)$/.test(key) || !Number.isSafeInteger(Number(key)) || Number(key) >= 512) throw new RangeError('表格列类型索引无效。')
      table.columnTypes[Number(key)] = columnType(value)
    }
  }
  if (extension.rowStyles !== undefined) table.rowStyles = documentBodyArray(extension.rowStyles, '表格来源行格式').map(cloneModel)
  const owners = new Map<string, { cell: DocumentCell; row: number; column: number }>()
  rows.forEach((rowSource, rowIndex) => {
    const cells = documentBodyArray(rowSource.tableCells, 'tableCells'), rowContents = contents[rowIndex]!
    if (cells.length !== rowContents.length) throw new RangeError('表格单元格定义与内容流不匹配。')
    const before = integer(rowSource.gridBefore, 0, '表格前置网格', 512), after = integer(rowSource.gridAfter, 0, '表格后置网格', 512)
    const dense = cells.some(cell => cell.rowSpan === 0 || cell.columnSpan === 0) || columns.length === before + cells.length + after
    const rowExtension = documentBodyOfficeSource(rowSource, '表格行')
    const heightSource = bodyRecord(rowSource.trHeight), height = bodyLength(heightSource.val, defaults.text.fontSize)
    if (height !== undefined && (height < 0 || height === 0 && heightSource.hRule !== undefined && heightSource.hRule !== 0)) throw new RangeError('表格固定或最小行高必须为正。')
    const row = { id: rowExtension.rowId === undefined ? documentId('row') : documentBodyIdentity(rowExtension.rowId, '表格行'), cells: [] as DocumentCell[],
      height: height === 0 ? undefined : height, heightRule: heightSource.hRule === 2 ? 'exact' as const : heightSource.hRule === 1 ? 'atLeast' as const : 'auto' as const,
      allowSplit: !(bodyBoolean(rowSource.cantSplit) ?? false), style: { ...documentBodySourceMetadata(rowExtension.style, '表格行', ['univerBodyRowSource']),
        univerBodyRowSource: cloneModel({ ...rowSource, tableCells: undefined }) } }
    table.rows.push(row)
    if (bodyBoolean(rowSource.repeatHeaderRow) && table.headerRowCount === rowIndex) table.headerRowCount = rowIndex + 1
    const omitted = (column: number, side: 'before' | 'after') => row.cells.push({ id: documentId('cell'), blocks: [createParagraph('', defaults.text)],
      style: { univerBodyOmittedCell: side }, columnIndex: column, rowSpan: 1, colSpan: 1 })
    for (let column = 0; column < before; column++) omitted(column, 'before')
    let column = before
    cells.forEach((source, cellIndex) => {
      const rowSpan = integer(source.rowSpan, 1, '跨行数', 10_000), colSpan = integer(source.columnSpan, 1, '跨列数', 512), blocks = rowContents[cellIndex]!
      if (rowSpan === 0 || colSpan === 0) {
        const owner = owners.get(`${rowIndex}:${column}`)
        if (!owner) throw new RangeError('表格覆盖单元格缺少合并起点。')
        const sourceOffset = rowIndex - owner.row
        if (sourceOffset > 0 && column === owner.column) {
          const map = owner.cell.style.nativeContinuationContent ??= {}
          map[sourceOffset] = { version: 1, ownerCellId: owner.cell.id, rowId: row.id, colSpan: owner.cell.colSpan,
            blocks, source: { ...documentBodySourceMetadata(documentBodyOfficeSource(source, '续格').continuationSource, '续格', ['univerBodyCellSource']),
              univerBodyCellSource: cloneModel(source) } }
        } else if (blocks.some(block => runText(block.runs).length || block.type !== 'paragraph')) {
          owner.cell.blocks.push(...blocks)
        }
        column++; return
      }
      if (rowIndex + rowSpan > rows.length || column + colSpan > 512) throw new RangeError('表格合并跨度越过真实网格。')
      const cellId = documentBodyOfficeSource(source, '表格单元格').cellId
      const cell: DocumentCell = { id: cellId === undefined ? documentId('cell') : documentBodyIdentity(cellId, '表格单元格'), blocks: blocks.length ? blocks : [createParagraph('', defaults.text)],
        style: cellStyle(source, defaults), columnIndex: column, rowSpan, colSpan }
      row.cells.push(cell)
      for (let coveredRow = rowIndex; coveredRow < rowIndex + rowSpan; coveredRow++) for (let coveredColumn = column; coveredColumn < column + colSpan; coveredColumn++) {
        const key = `${coveredRow}:${coveredColumn}`
        if (owners.has(key)) throw new RangeError('表格合并区域重叠。')
        owners.set(key, { cell, row: rowIndex, column })
      }
      column += dense ? 1 : colSpan
    })
    for (let index = 0; index < after; index++) omitted(column + index, 'after')
  })
  const grid = documentTableGrid(table)
  for (const key of Object.keys(table.columnTypes ?? {})) if (Number(key) >= grid.columnCount) throw new RangeError('表格列类型超出真实网格。')
  for (let index = 0; index < table.rows.length; index++) documentTableRowProperties(table, index)
  return table
}

function marginToBody(value: DocumentTableCellMargin): UniverBodyRecord { return { top: px(value.top), end: px(value.right), bottom: px(value.bottom), start: px(value.left) } }
function borderToBody(value: DocumentTableBorder, source?: unknown): UniverBodyRecord {
  return { ...cloneModel(bodyRecord(source)), color: value.color === 'auto' ? { auto: 1 } : { rgb: value.color }, width: px(value.width), dashStyle: value.style === 'dot' ? 2 : value.style === 'dash' ? 3 : 1 }
}
export function documentTableCellToBody(cell: DocumentCell, source?: unknown): UniverBodyRecord {
  const result = cloneModel(bodyRecord(source ?? cell.style.univerBodyCellSource)), fill = cell.style.fill ?? cell.style.background
  result.rowSpan = cell.rowSpan; result.columnSpan = cell.colSpan
  result.officeStudio = { version: 1, cellId: cell.id, style: documentBodySourceMetadata(cell.style, '表格单元格',
    ['univerBodyCellSource', 'nativeContinuationContent', 'fill', 'background', 'margin', 'verticalAlign', 'borders']) }
  if (fill !== undefined) result.backgroundColor = { rgb: fill }; else delete result.backgroundColor
  if (cell.style.margin) result.margin = marginToBody(cell.style.margin); else delete result.margin
  if (cell.style.verticalAlign) result.vAlign = cell.style.verticalAlign === 'center' ? 3 : cell.style.verticalAlign === 'bottom' ? 4 : 2
  for (const [side, key] of [['top', 'borderTop'], ['right', 'borderRight'], ['bottom', 'borderBottom'], ['left', 'borderLeft']] as const) {
    const line = cell.style.borders?.[side]
    if (line) result[key] = borderToBody(line, result[key]); else delete result[key]
  }
  return result
}
export function documentTableToBody(table: DocumentTable, id: string, rows: UniverBodyRecord[]): UniverBodyRecord {
  const result = cloneModel(bodyRecord(table.style?.univerBodyTableSource)), style = table.style ?? {}
  result.tableId = id; result.tableRows = rows
  result.officeStudio = { version: 1, style: documentBodySourceMetadata(style, '表格',
    ['univerBodyTableSource', 'fill', 'background', 'align', 'indent', 'layout', 'width', 'cellMargin', 'borders']),
    ...(table.titleRow === undefined ? {} : { titleRow: table.titleRow }), ...(table.columnTypes === undefined ? {} : { columnTypes: cloneModel(table.columnTypes) }),
    ...(table.rowStyles === undefined ? {} : { rowStyles: cloneModel(table.rowStyles) }) }
  if (style.fill ?? style.background) result.backgroundColor = { rgb: style.fill ?? style.background }
  else delete result.backgroundColor
  result.tableColumns = table.columnWidths?.map(value => ({ size: size(value) })) ?? []
  result.align = style.align === 'center' ? 1 : style.align === 'right' ? 2 : 0
  result.indent = px(style.indent ?? 0); result.layout = style.layout === 'fixed' ? 1 : 0
  if (style.width !== undefined) result.size = size(style.width); else delete result.size
  if (style.cellMargin) result.cellMargin = marginToBody(style.cellMargin); else delete result.cellMargin
  const borders: UniverBodyRecord = {}
  for (const side of ['top', 'right', 'bottom', 'left', 'insideHorizontal', 'insideVertical'] as const) {
    const line = style.borders?.[side]; if (line) borders[side] = borderToBody(line, bodyRecord(result.borderDefaults)[side])
  }
  if (Object.keys(borders).length) result.borderDefaults = borders; else delete result.borderDefaults
  return result
}
export function documentTableRowToBody(table: DocumentTable, index: number, cells: UniverBodyRecord[]): UniverBodyRecord {
  const row = table.rows[index]!, result = cloneModel(bodyRecord(row.style?.univerBodyRowSource)), properties = documentTableRowProperties(table, index)
  result.officeStudio = { version: 1, rowId: row.id, style: documentBodySourceMetadata(row.style, '表格行', ['univerBodyRowSource']) }
  result.tableCells = cells
  if (properties.height === undefined) delete result.trHeight
  else result.trHeight = { ...bodyRecord(result.trHeight), val: px(properties.height), hRule: properties.heightRule === 'exact' ? 2 : properties.heightRule === 'atLeast' ? 1 : 0 }
  result.cantSplit = Number(properties.allowSplit === false); result.repeatHeaderRow = Number(index < (table.headerRowCount ?? 0))
  return result
}
