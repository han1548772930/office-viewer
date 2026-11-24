import type { DocumentBlock } from '../domain/model'
import type {
  DocumentCell, DocumentTable, DocumentTableBorder, DocumentTableBorders, DocumentTableCellMargin,
  DocumentTableCellStyle, DocumentTableColumnTypeConfig, DocumentTableRow, DocumentTableStyle,
} from '../features/document/model/document-model'
import { documentTableGrid } from '../features/document/model/document-table-grid'
import { assertDocumentTableContinuationRows } from '../features/document/model/document-table-continuations'
import {
  assertDocumentTableBorder, assertDocumentTableColor, assertDocumentTableColumnType, assertDocumentTableMargin,
  DOCUMENT_TABLE_DEFAULT_BORDER, documentTableCellMargin, documentTableColumnWidths, documentTableRowProperties,
} from '../features/document/model/document-table-style'
import { array, record, text } from './validation'
import { normalizeParsedDocumentBlockGroup } from './documentAuxiliaryContent'

type NormalizeBlock = (raw: unknown, id: string) => DocumentBlock
const sides = ['top', 'right', 'bottom', 'left'] as const
const borderSides = ['top', 'right', 'bottom', 'left', 'insideHorizontal', 'insideVertical'] as const
const twips = (px: number): number => Math.round(px * 15)
const nativeStyleKeys = new Set([
  'fill', 'background', 'backgroundColor', 'verticalAlign', 'margin', 'marginTwips', 'borders',
  'align', 'alignment', 'sourceAlignment', 'sourceInheritedAlignment', 'indent', 'indentTwips', 'offsetY', 'floating',
  'width', 'widthTwips', 'preferredWidth', 'sourceWidth', 'layout', 'cellMargin', 'cellMarginTwips', 'columnWidthsTwips',
  'gridSpan', 'rowSpan', 'vMergeContinuation', 'textDirection', 'sourceTextDirection', 'noWrap', 'sourceNoWrap',
  'styleId', 'sourceStyleId', 'sourceLook', 'bidiVisual', 'sourceBidiVisual', 'sourceInheritedBidiVisual',
  'cellSpacingTwips', 'collapsedGridColumns', 'nativeProperties', 'nativeContinuations', 'nativeContinuationContent',
  'height', 'heightPx', 'heightTwips', 'heightRule', 'allowSplit', 'cantSplit', 'isHeader', 'sourceInheritedHeader',
  'gridBefore', 'gridAfter', 'widthBefore', 'widthAfter', 'sourceRowGeometry',
])

function retainedStyle(saved: unknown, current: Record<string, unknown>): Record<string, unknown> {
  return { ...Object.fromEntries(Object.entries(record(saved)).filter(([key]) => !nativeStyleKeys.has(key))), ...current }
}

function integer(value: unknown, label: string, fallback?: number): number {
  if (value === undefined || value === null) {
    if (fallback !== undefined) return fallback
    throw new TypeError(`文档表格 ${label} 缺失`)
  }
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) throw new RangeError(`文档表格 ${label} 必须为非负安全整数`)
  return value
}

function optionalNumber(value: unknown, label: string, positive = false): number | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'number' || !Number.isFinite(value) || positive && value <= 0) throw new RangeError(`文档表格 ${label} 数值无效`)
  return value
}

function positiveNumber(value: unknown, label: string): number {
  const number = optionalNumber(value, label, true)
  if (number === undefined) throw new TypeError(`文档表格 ${label} 缺失`)
  return number
}

function id(value: unknown, fallback: string): string {
  if (value === undefined || value === null) return fallback
  if (typeof value !== 'string' || !value.trim() || value.length > 4096) throw new TypeError('文档表格标识无效')
  return value
}

function nativeColor(value: unknown, fallback: string): string {
  if (value === undefined || value === null || value === 'auto') return fallback
  if (typeof value !== 'string') throw new TypeError('文档表格原生颜色无效')
  return /^[\da-f]{6}(?:[\da-f]{2})?$/i.test(value) ? `#${value}` : value
}

function margin(value: unknown, native: boolean, fallback?: DocumentTableCellMargin): DocumentTableCellMargin | undefined {
  if (value === undefined || value === null) return fallback ? { ...fallback } : undefined
  const raw = record(value), result = { top: 0, right: 0, bottom: 0, left: 0 }
  for (const side of sides) {
    const value = optionalNumber(raw[native ? `${side}Twips` : side], `边距 ${side}`)
    result[side] = value === undefined ? fallback?.[side] ?? 0 : native ? value / 15 : value
  }
  assertDocumentTableMargin(result)
  return result
}

function border(value: unknown): DocumentTableBorder {
  const raw = record(value)
  const type = text(raw.type)
  const width = optionalNumber(raw.width, '边框宽度') ?? (type === 'nil' || type === 'none' ? 0 : (optionalNumber(raw.sizeEighthPt, '原生边框宽度') ?? 4) / 6)
  const style = raw.style === 'dot' || type === 'dotted' ? 'dot' : raw.style === 'dash' || type === 'dashed' || type === 'dashSmallGap' ? 'dash' : 'solid'
  const result: DocumentTableBorder = { ...raw, width, style, color: nativeColor(raw.color, '#000000') }
  assertDocumentTableBorder(result)
  return result
}

function borders(value: unknown, table: boolean): DocumentTableBorders {
  const raw = record(value), result: DocumentTableBorders = {}
  for (const side of table ? borderSides : sides) {
    const source = raw[side] ?? raw[side === 'insideHorizontal' ? 'insideH' : side === 'insideVertical' ? 'insideV' : side]
    if (source !== undefined && source !== null) result[side] = border(source)
    else if (table) result[side] = { width: 0, color: '#000000', style: 'solid' }
  }
  return result
}

function tableStyle(value: unknown, background: unknown): DocumentTableStyle {
  const raw = record(value)
  const width = optionalNumber(raw.width, '宽度', true) ?? (optionalNumber(raw.widthTwips, '原生宽度', true) === undefined ? undefined : Number(raw.widthTwips) / 15)
  const indent = optionalNumber(raw.indent, '缩进') ?? (optionalNumber(raw.indentTwips, '原生缩进') ?? 0) / 15
  const floating = record(raw.floating)
  const offsetY = optionalNumber(raw.offsetY, '垂直偏移') ?? (optionalNumber(floating.yTwips, '原生垂直偏移') === undefined ? undefined : Number(floating.yTwips) / 15)
  const align = raw.align ?? raw.alignment ?? 'left'
  if (align !== 'left' && align !== 'center' && align !== 'right') throw new TypeError('文档表格对齐方式无效')
  const layout = raw.layout === 'autofit' || raw.layout === 'auto' ? 'auto' : raw.layout === undefined || raw.layout === 'fixed' ? 'fixed' : undefined
  if (!layout) throw new TypeError('文档表格布局方式无效')
  const result: DocumentTableStyle = { ...raw, align, indent, layout,
    ...(width === undefined ? {} : { width }), ...(offsetY === undefined ? {} : { offsetY }),
    cellMargin: margin(raw.cellMargin ?? raw.cellMarginTwips, raw.cellMargin === undefined, { top: 0, right: 7.2, bottom: 0, left: 7.2 }),
    borders: borders(raw.borders, true),
  }
  const fill = raw.fill ?? raw.background ?? background
  if (fill !== undefined && fill !== null) result.fill = nativeColor(fill, '#ffffff')
  return result
}

function cellStyle(value: unknown): DocumentTableCellStyle {
  const raw = record(value), result: DocumentTableCellStyle = { ...raw }
  const fill = raw.fill ?? raw.background ?? raw.backgroundColor
  if (fill !== undefined && fill !== null) result.fill = nativeColor(fill, '#ffffff')
  if (raw.verticalAlign !== undefined && !['top', 'center', 'bottom'].includes(String(raw.verticalAlign))) throw new TypeError('文档单元格垂直对齐无效')
  result.verticalAlign = raw.verticalAlign === 'center' || raw.verticalAlign === 'bottom' ? raw.verticalAlign : 'top'
  const cellMargin = margin(raw.margin ?? raw.marginTwips, raw.margin === undefined)
  if (cellMargin) result.margin = cellMargin
  result.borders = borders(raw.borders, false)
  return result
}

function validateStyle(style: DocumentTableStyle | DocumentTableCellStyle, table: boolean): void {
  for (const key of ['fill', 'background'] as const) {
    const value = style[key]
    if (value !== undefined) { if (typeof value !== 'string') throw new TypeError('文档表格背景颜色无效'); assertDocumentTableColor(value) }
  }
  const edge = table ? (style as DocumentTableStyle).cellMargin : (style as DocumentTableCellStyle).margin
  if (edge !== undefined) assertDocumentTableMargin(edge)
  if (style.borders !== undefined) {
    if (!style.borders || typeof style.borders !== 'object' || Array.isArray(style.borders)) throw new TypeError('文档表格边框集合无效')
    const allowed: readonly string[] = table ? borderSides : sides
    for (const [side, value] of Object.entries(style.borders)) {
      if (!allowed.includes(side)) continue
      if (!value || typeof value.width !== 'number' || typeof value.color !== 'string') throw new TypeError('文档表格边框属性无效')
      assertDocumentTableBorder(value)
    }
  }
  if (table) {
    const value = style as DocumentTableStyle
    if (value.align !== undefined && !['left', 'center', 'right'].includes(value.align)) throw new TypeError('文档整表对齐无效')
    if (value.layout !== undefined && !['auto', 'fixed'].includes(value.layout)) throw new TypeError('文档整表布局无效')
    optionalNumber(value.width, '宽度', true); optionalNumber(value.indent, '缩进'); optionalNumber(value.offsetY, '垂直偏移')
  } else if (style.verticalAlign !== undefined && !['top', 'center', 'bottom'].includes(String(style.verticalAlign))) throw new TypeError('文档单元格垂直对齐无效')
}

function officeSource(details: Record<string, unknown>): Record<string, unknown> | undefined {
  if (details.officeTable === undefined) return undefined
  const source = record(details.officeTable)
  if (source.version !== 1 || typeof source.nativeMatches !== 'boolean' || typeof source.topologyMatches !== 'boolean'
    || !source.data || typeof source.data !== 'object' || Array.isArray(source.data)
    || typeof source.tableId !== 'string' || !source.tableId.trim() || source.tableId.length > 4096
    || source.nativeMatches && !source.topologyMatches) throw new TypeError('文档表格来源协议无效')
  return source
}

export function normalizedDocumentTableId(value: unknown, fallback: string): string {
  const raw = record(value), source = officeSource(record(raw.details))
  return source ? id(source.tableId, fallback) : id(raw.id, fallback)
}

function propertySource(value: unknown, baseline: Record<string, unknown>): Record<string, unknown> | undefined {
  if (value === undefined || value === null) return undefined
  const source = record(value), namespaces = record(source.namespaces)
  if (source.version !== 1 || typeof source.propertiesXml !== 'string' || source.propertiesXml.length > 16 * 1024 * 1024
    || !source.namespaces || typeof source.namespaces !== 'object' || Array.isArray(source.namespaces)
    || Object.values(namespaces).some(value => typeof value !== 'string')) throw new TypeError('文档表格原生属性来源无效')
  return { ...source, namespaces: { ...namespaces }, baseline }
}

function tableBaseline(table: DocumentTable, widths: number[]): Record<string, unknown> {
  const style = table.style ?? {}, nativeWidths = widths.map(twips)
  const effectiveBorders = Object.fromEntries(borderSides.map(side => [side, borderBaseline(style.borders?.[side] ?? DOCUMENT_TABLE_DEFAULT_BORDER)]))
  return { tblW: [nativeWidths.reduce((sum, width) => sum + width, 0), nativeWidths], jc: style.align ?? 'left', tblInd: style.indent ?? 0,
    tblLayout: style.layout ?? 'fixed', tblCellMar: sides.map(side => documentTableCellMargin(table)[side]), tblBorders: effectiveBorders,
    shd: style.fill ?? style.background ?? null, tblpPr: [style.offsetY ?? null, style.indent ?? 0] }
}

function borderBaseline(border: DocumentTableBorder): Record<string, unknown> {
  return { width: border.width, color: border.color, style: border.style }
}

function cellBaseline(table: DocumentTable, cell: DocumentCell, nativeWidth: number): Record<string, unknown> {
  const style = cell.style
  const borderValues = Object.fromEntries(sides.flatMap(side => {
    const border = style.borders?.[side]
    return border === undefined ? [] : [[side, borderBaseline(border)]]
  }))
  return { tcW: nativeWidth, tcBorders: borderValues, shd: style.fill ?? style.background ?? null,
    tcMar: style.margin ? sides.map(side => documentTableCellMargin(table, cell)[side]) : null, vAlign: style.verticalAlign ?? 'top' }
}

function authoredTypes(value: unknown, count: number): Record<number, DocumentTableColumnTypeConfig> | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'object' || Array.isArray(value)) throw new TypeError('文档表格列类型集合无效')
  const result: Record<number, DocumentTableColumnTypeConfig> = {}
  for (const [column, value] of Object.entries(value)) {
    const index = Number(column), raw = record(value)
    if (!Number.isSafeInteger(index) || index < 0 || index >= count || String(index) !== column) throw new RangeError('文档表格列类型超出逻辑网格')
    const type = raw.type
    if (type !== 'text' && type !== 'number' && type !== 'date' && type !== 'checkbox' && type !== 'dropdown') throw new TypeError('文档表格列类型无效')
    const options = raw.dropdownOptions === undefined ? undefined : array(raw.dropdownOptions).map(value => {
      const raw = record(value)
      if (typeof raw.id !== 'string' || typeof raw.label !== 'string' || raw.color !== undefined && typeof raw.color !== 'string') throw new TypeError('文档表格下拉选项无效')
      return { id: raw.id, label: raw.label, ...(typeof raw.color === 'string' ? { color: raw.color } : {}) }
    })
    if (raw.dropdownOptions !== undefined && !Array.isArray(raw.dropdownOptions)) throw new TypeError('文档表格下拉选项必须为数组')
    const config: DocumentTableColumnTypeConfig = { type, ...(options ? { dropdownOptions: options } : {}) }
    assertDocumentTableColumnType(config); result[index] = config
  }
  return result
}

function cellBlocks(source: Record<string, unknown>, cellId: string, property: Record<string, unknown> | undefined,
  normalizeBlock: NormalizeBlock): DocumentBlock[] {
  let blocks = normalizeParsedDocumentBlockGroup(array(source.children), `${cellId}-block`, normalizeBlock)
  const lastSource = record(array(source.children).at(-1)), lastBlock = blocks.at(-1)
  if (property?.syntheticTail === true && lastBlock?.type === 'paragraph' && lastBlock.runs.every(run => !run.text)
    && !array(lastSource.children).length && array(record(lastSource.details).inlines).every(inline => record(inline).type === 'text' && !text(record(inline).text))) blocks = blocks.slice(0, -1)
  if (!blocks.length && source.children === undefined && text(source.content)) blocks = [normalizeBlock({ type: 'text', content: source.content }, `${cellId}-text`)]
  return blocks
}

function horizontalMergeProvenance(cell: DocumentCell, rowId: string, property: Record<string, unknown> | undefined): void {
  if (property?.sourceHorizontalMerge === undefined) return
  const source = record(property.sourceHorizontalMerge)
  if (source.version !== 1 || !Array.isArray(source.members) || source.members.length < 2) throw new TypeError('文档表格横向合并来源协议无效')
  const previous = record(cell.style.nativeHorizontalMergeSource)
  const rows = array(previous.rows).map(record).filter(row => row.rowId !== rowId)
  // Provenance never assigns grid ownership or revives old text. Actual member
  // bodies have already been joined by Rust, with their paragraph IDs intact.
  cell.style.nativeHorizontalMergeSource = { version: 1, rows: [...rows, { rowId, members: source.members }] }
}

/** No placeholder cells, span clamps or compacting of fully covered rows. */
export function normalizeDocumentTable(value: unknown, tableId: string, normalizeBlock: NormalizeBlock): DocumentTable {
  const raw = record(value), details = record(raw.details), native = record(details.nativeTableProperties)
  if (details.nativeTableProperties !== undefined && (native.version !== 1 || !Array.isArray(native.rows))) throw new TypeError('文档表格原生属性协议无效')
  const source = officeSource(details), authored = record(source?.data)
  const sourceCells: Record<string, unknown>[] = []
  for (const child of array(raw.children)) {
    const item = record(child)
    if (item.type === 'table_head') sourceCells.push(...array(item.children).map(record))
    else if (item.type === 'table_cell') sourceCells.push(item)
  }
  const indexed = sourceCells.map((cell, index) => {
    const details = record(cell.details)
    const row = integer(details.rowIndex, 'rowIndex', 0)
    const rowSpan = integer(details.rowSpan, 'rowSpan', 1), colSpan = integer(details.columnSpan, 'columnSpan', 1)
    if (rowSpan < 1 || colSpan < 1 || colSpan > 512) throw new RangeError('文档表格跨度无效或超过逻辑列预算')
    return { cell, index, details, row, rowSpan, colSpan, column: details.columnIndex === undefined ? undefined : integer(details.columnIndex, 'columnIndex'),
      continuation: record(details.style).vMergeContinuation === true }
  })
  const inferredRows = indexed.reduce((count, cell) => Math.max(count, cell.row + (cell.continuation ? 1 : cell.rowSpan)), 0)
  const rowCount = integer(details.rowCount, 'rowCount', inferredRows)
  if (!rowCount || rowCount > 10_000 || indexed.some(cell => cell.row >= rowCount || !cell.continuation && cell.row + cell.rowSpan > rowCount)) throw new RangeError('文档表格实际行数或跨行范围无效')
  if (native.rows !== undefined && array(native.rows).length !== rowCount) throw new RangeError('文档表格原生行数与实际行数不一致')
  const declaredColumnCount = details.columnCount === undefined ? undefined : integer(details.columnCount, 'columnCount')
  const explicitColumnCount = indexed.reduce((count, source) => Math.max(count, (source.column ?? 0) + source.colSpan), 0)
  if (explicitColumnCount > 512 || declaredColumnCount !== undefined && (declaredColumnCount < 1 || declaredColumnCount > 512)
    || rowCount * Math.max(explicitColumnCount, declaredColumnCount ?? 0) > 1_000_000) throw new RangeError('文档表格逻辑网格超过资源预算')
  const rawRows = array(details.rowStyles)
  const rows: DocumentTableRow[] = Array.from({ length: rowCount }, (_, index): DocumentTableRow => {
    const style = record(rawRows[index]), height = optionalNumber(style.height ?? style.heightPx, '行高', true)
      ?? (optionalNumber(style.heightTwips, '原生行高', true) === undefined ? undefined : Number(style.heightTwips) / 15)
    const rawRule = style.heightRule
    if (rawRule !== undefined && rawRule !== 'auto' && rawRule !== 'atLeast' && rawRule !== 'exact') throw new TypeError('文档表格行高规则无效')
    if (style.allowSplit !== undefined && typeof style.allowSplit !== 'boolean' || style.cantSplit !== undefined && typeof style.cantSplit !== 'boolean') throw new TypeError('文档表格跨页拆行设置无效')
    return { id: `${tableId}-row-${index}`, cells: [], ...(height === undefined ? {} : { height }),
      heightRule: rawRule ?? (height === undefined ? 'auto' : 'atLeast'), allowSplit: typeof style.allowSplit === 'boolean' ? style.allowSplit : style.cantSplit !== true, style: { ...style } }
  })
  const table: DocumentTable = { rows, style: tableStyle(details.style, native.backgroundColor), rowStyles: rawRows.map(value => ({ ...record(value) })) }
  const widths = record(details.style).columnWidthsTwips ?? details.columnWidthsTwips
  if (widths !== undefined && widths !== null) {
    if (!Array.isArray(widths)) throw new TypeError('文档表格原生列宽必须为数组')
    table.columnWidths = widths.map(value => positiveNumber(value, '原生列宽') / 15)
  } else if (details.columnWidths !== undefined) {
    if (!Array.isArray(details.columnWidths)) throw new TypeError('文档表格列宽必须为数组')
    table.columnWidths = details.columnWidths.map(value => positiveNumber(value, '列宽'))
  }
  let headerCount = 0
  while (headerCount < rowCount && record(rawRows[headerCount]).isHeader === true) headerCount++
  table.headerRowCount = integer(details.headerRowCount, 'headerRowCount', headerCount)
  const nativeCellSources = new Map<string, Record<string, unknown>>()
  array(native.rows).forEach((row, rowIndex) => array(record(row).cells).forEach(cell => {
    const source = record(cell)
    const key = `${rowIndex}:${integer(source.columnIndex, '原生列索引')}`
    if (nativeCellSources.has(key)) throw new RangeError('文档表格原生单元格来源重复')
    nativeCellSources.set(key, source)
  }))
  for (const source of indexed.filter(source => !source.continuation)) {
    const cellId = id(source.cell.id, `${tableId}-cell-${source.index}`)
    const property = source.column === undefined ? undefined : nativeCellSources.get(`${source.row}:${source.column}`)
    const blocks = cellBlocks(source.cell, cellId, property, normalizeBlock)
    rows[source.row]!.cells.push({ id: cellId, blocks, rowSpan: source.rowSpan, colSpan: source.colSpan,
      ...(source.column === undefined ? {} : { columnIndex: source.column }), style: cellStyle(source.details.style) })
  }
  let grid = documentTableGrid(table)
  for (const owner of grid.anchors) owner.cell.columnIndex ??= owner.column
  const declaredColumns = declaredColumnCount ?? grid.columnCount
  if (declaredColumns !== grid.columnCount) throw new RangeError('文档表格声明列数与实际网格不一致')
  if (source?.topologyMatches === true) {
    const authoredRows = array(authored.rows)
    if (authoredRows.length !== rowCount) throw new RangeError('文档表格来源的实际行数不一致')
    for (const [index, row] of rows.entries()) {
      const saved = record(authoredRows[index]), cells = array(saved.cells)
      if (cells.length !== row.cells.length) throw new RangeError('文档表格来源的锚点数量不一致')
      row.id = id(saved.id, row.id)
      Object.assign(row, Object.fromEntries(Object.entries(saved).filter(([key]) => !['id', 'cells', 'height', 'heightRule', 'allowSplit', 'style'].includes(key))))
      row.style = retainedStyle(saved.style, row.style ?? {})
      row.cells.forEach(cell => {
        const saved = cells.map(record).find(saved => integer(saved.columnIndex, '来源列索引') === cell.columnIndex)
        if (!saved || integer(saved.rowSpan, '来源跨行数') !== cell.rowSpan || integer(saved.colSpan, '来源跨列数') !== cell.colSpan) throw new RangeError('文档表格来源的锚点跨度不一致')
        cell.id = id(saved.id, cell.id)
        Object.assign(cell, Object.fromEntries(Object.entries(saved).filter(([key]) => !['id', 'blocks', 'columnIndex', 'rowSpan', 'colSpan', 'style'].includes(key))))
        if (source.nativeMatches === true) cell.style = { ...record(saved.style) }
        else cell.style = retainedStyle(saved.style, cell.style)
      })
      if (source.nativeMatches === true) {
        row.style = { ...record(saved.style) }
        delete row.height; delete row.heightRule; delete row.allowSplit
        if (saved.height !== undefined) row.height = optionalNumber(saved.height, '来源行高', true)
        if (saved.heightRule !== undefined) {
          if (saved.heightRule !== 'auto' && saved.heightRule !== 'atLeast' && saved.heightRule !== 'exact') throw new TypeError('文档表格来源行高规则无效')
          row.heightRule = saved.heightRule
        }
        if (saved.allowSplit !== undefined) { if (typeof saved.allowSplit !== 'boolean') throw new TypeError('文档表格来源拆行设置无效'); row.allowSplit = saved.allowSplit }
      }
    }
    table.columnTypes = authoredTypes(authored.columnTypes, grid.columnCount)
    Object.assign(table, Object.fromEntries(Object.entries(authored).filter(([key]) => !['rows', 'style', 'columnWidths', 'rowStyles', 'headerRowCount', 'titleRow', 'columnTypes'].includes(key))))
    if (authored.titleRow !== undefined) { if (typeof authored.titleRow !== 'boolean') throw new TypeError('文档表格标题行设置无效'); table.titleRow = authored.titleRow }
    if (source.nativeMatches === true) {
      table.style = { ...record(authored.style) }
      if (authored.columnWidths !== undefined) {
        if (!Array.isArray(authored.columnWidths)) throw new TypeError('文档表格来源列宽必须为数组')
        table.columnWidths = authored.columnWidths.map(value => positiveNumber(value, '来源列宽'))
      } else delete table.columnWidths
      table.headerRowCount = integer(authored.headerRowCount, '来源表头行数', 0)
      if (authored.rowStyles !== undefined) { if (!Array.isArray(authored.rowStyles)) throw new TypeError('文档表格来源行样式必须为数组'); table.rowStyles = authored.rowStyles.map(value => ({ ...record(value) })) }
    } else table.style = retainedStyle(authored.style, table.style ?? {})
    grid = documentTableGrid(table)
  }
  validateStyle(table.style ?? {}, true)
  for (const row of rows) for (const cell of row.cells) validateStyle(cell.style, false)
  const nativeWidth = Array.isArray(widths) ? widths.reduce((sum: number, width: unknown) => sum + positiveNumber(width, '原生列宽') / 15, 0) : undefined
  const effectiveWidths = documentTableColumnWidths(table, grid.columnCount, table.style?.width ?? (nativeWidth && nativeWidth > 0 ? nativeWidth : 600))
  const tableSource = propertySource(native.nativeProperties, tableBaseline(table, effectiveWidths))
  if (tableSource) table.style = { ...table.style, nativeProperties: tableSource }
  rows.forEach((row, index) => {
    const values = documentTableRowProperties(table, index)
    const nativeSource = propertySource(record(array(native.rows)[index]).nativeProperties,
      { trHeight: [values.height ?? null, values.heightRule], cantSplit: values.allowSplit, tblHeader: index < (table.headerRowCount ?? 0) })
    if (nativeSource) row.style = { ...row.style, nativeProperties: nativeSource }
  })
  for (const owner of grid.anchors) {
    // Incoming native text is authoritative even when native geometry still
    // matches. Never resurrect an old continuation from the authored sidecar.
    delete owner.cell.style.nativeContinuations
    delete owner.cell.style.nativeContinuationContent
    const width = effectiveWidths.slice(owner.column, owner.column + owner.colSpan).reduce((sum, width) => sum + twips(width), 0)
    const nativeSource = propertySource(nativeCellSources.get(`${owner.row}:${owner.column}`)?.nativeProperties, cellBaseline(table, owner.cell, width))
    if (nativeSource) owner.cell.style.nativeProperties = nativeSource
    horizontalMergeProvenance(owner.cell, rows[owner.row]!.id, nativeSource)
  }
  for (const source of indexed.filter(source => source.continuation)) {
    if (source.column === undefined) throw new RangeError('原生跨行续格缺少逻辑列')
    const owner = grid.cellAt(source.row, source.column)
    if (!owner || owner.row >= source.row || owner.column !== source.column || owner.colSpan !== source.colSpan) throw new RangeError('原生跨行续格没有匹配锚点')
    const width = effectiveWidths.slice(owner.column, owner.column + owner.colSpan).reduce((sum, width) => sum + twips(width), 0)
    const nativeSource = propertySource(nativeCellSources.get(`${source.row}:${source.column}`)?.nativeProperties, cellBaseline(table, owner.cell, width))
    horizontalMergeProvenance(owner.cell, rows[source.row]!.id, nativeSource)
    if (nativeSource) owner.cell.style.nativeContinuations = { ...record(owner.cell.style.nativeContinuations), [source.row - owner.row]: {
      ...nativeSource, ownerCellId: owner.cell.id, rowId: rows[source.row]!.id, colSpan: owner.colSpan,
    } }
    const children = array(source.cell.children)
    const independentContent = text(source.cell.content).length > 0 || children.length > 1 || children.some(child => {
      const raw = record(child), details = record(raw.details)
      return !['text', 'paragraph'].includes(text(raw.type)) || array(raw.children).length > 0
        || details.documentTextFields !== undefined || details.nativeFieldSource !== undefined || Object.keys(record(details.numbering)).length > 0
    })
    if (independentContent) {
      const property = nativeCellSources.get(`${source.row}:${source.column}`)
      const continuationId = id(source.cell.id, `${owner.cell.id}-continued-${source.row - owner.row}`)
      owner.cell.style.nativeContinuationContent = { ...owner.cell.style.nativeContinuationContent, [source.row - owner.row]: {
        version: 1, ownerCellId: owner.cell.id, rowId: rows[source.row]!.id, colSpan: owner.colSpan,
        blocks: cellBlocks(source.cell, continuationId, property, normalizeBlock), source: source.cell,
      } }
    }
  }
  for (const owner of grid.anchors) assertDocumentTableContinuationRows(table, owner.cell, owner.row)
  return table
}
