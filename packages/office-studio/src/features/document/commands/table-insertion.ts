import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import {
  createParagraph, createTable, documentBlockAncestry, findDocumentBlock, runText, tableForBlock, type DocumentPoint, type DocumentTable,
} from '../model/document-model'
import { documentTableGrid } from '../model/document-table-grid'
import { assertDocumentTableColumnType, assertDocumentTableMargin } from '../model/document-table-style'
import { documentEffectiveSectionSetup } from '../services/document-settings'
import { sectionIdForBlock } from '../services/section-header-footer'
import { applyDocumentTableStylePatch, assertDocumentTableStylePatch } from './table-mutations'
import type { DocumentTableCreateCommand, DocumentTableInsertCommand } from './table-command-types'

export const DOCUMENT_TABLE_INSERT_CELL_BUDGET = 250000

function positive(value: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`表格${label}必须为正整数`)
  return value
}
function size(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) throw new RangeError(`表格${label}必须为正 CSS px`)
  return value
}

export function assertDocumentTableInsertionSize(rows: number, columns: number) {
  positive(rows, '行数')
  positive(columns, '列数')
  if (rows > 10000 || columns > 512) throw new RangeError('表格单次插入最多支持 10000 行、512 列')
  if (rows * columns > DOCUMENT_TABLE_INSERT_CELL_BUDGET) throw new RangeError('表格单次插入超出 250000 个单元格的内存预算')
  return { rows, columns }
}

export function assertDocumentTableInsertion(model: DocumentModel, options: DocumentTableInsertCommand | DocumentTableCreateCommand): { rows: number; columns: number } {
  if (options.values !== undefined && (!Array.isArray(options.values) || options.values.some(row => !Array.isArray(row) || row.some(value => typeof value !== 'string')))) throw new TypeError('表格初始数据必须是文本矩阵')
  const { rows, columns } = assertDocumentTableInsertionSize(
    options.id === 'doc.command.create-table' ? options.rowCount : options.rows ?? options.values?.length ?? 1,
    options.id === 'doc.command.create-table' ? options.colCount : options.cols ?? (options.values ?? []).reduce((maximum, row) => Math.max(maximum, row.length), 1),
  )
  if (options.values && (options.values.length > rows || options.values.some(row => row.length > columns))) throw new RangeError('表格尺寸不能丢弃提供的初始数据')
  if (options.tableId !== undefined && (typeof options.tableId !== 'string' || !options.tableId.trim() || findDocumentBlock(model, options.tableId))) throw new RangeError('表格标识无效或已存在')
  if (options.position !== undefined && !['selection', 'documentEnd', 'offset'].includes(options.position)) throw new TypeError('表格插入位置无效')
  if (options.offset !== undefined && !Number.isSafeInteger(options.offset)) throw new RangeError('表格插入偏移必须为安全整数')
  if (options.segmentId !== undefined && typeof options.segmentId !== 'string') throw new TypeError('表格故事标识无效')
  if (options.segmentId) throw new RangeError('页眉页脚表格需要故事块模型；当前候选实现尚未接入')
  if (options.width !== undefined) size(options.width, '宽度')
  for (const [values, maximum, label] of [[options.columnWidths, columns, '列宽'], [options.rowHeights, rows, '行高']] as const) if (values !== undefined) {
    if (!Array.isArray(values) || values.length > maximum) throw new RangeError(`表格${label}数量超出实际网格`)
    for (const value of values) if (value !== undefined) size(value, label)
  }
  if (options.cellMargin !== undefined) assertDocumentTableMargin(options.cellMargin)
  if (options.titleRow !== undefined && typeof options.titleRow !== 'boolean') throw new TypeError('表格标题行参数无效')
  if (options.headerRowCount !== undefined && (!Number.isSafeInteger(options.headerRowCount) || options.headerRowCount < 0 || options.headerRowCount > rows)) throw new RangeError('表格表头行数无效')
  if (options.columnTypes !== undefined) {
    if (!options.columnTypes || typeof options.columnTypes !== 'object' || Array.isArray(options.columnTypes)) throw new TypeError('表格列类型集合无效')
    for (const [index, config] of Object.entries(options.columnTypes)) {
      if (!/^\d+$/.test(index) || Number(index) >= columns) throw new RangeError('表格列类型超出实际网格')
      assertDocumentTableColumnType(config)
    }
  }
  if (options.style !== undefined) {
    assertDocumentTableStylePatch(options.style)
    if (options.style.align !== undefined && (model.metadata?.documentFlavor === 'modern' || model.metadata?.documentFlavor === 2 || model.metadata?.mode === 'modern')) throw new RangeError('整表对齐适用于传统分页文档')
  }
  if (options.metadata !== undefined && (!options.metadata || typeof options.metadata !== 'object' || Array.isArray(options.metadata))) throw new TypeError('表格扩展元数据无效')
  return { rows, columns }
}

/** Values and metadata are built before insertion; no per-cell follow-up commands. */
export function createDocumentTableBlock(
  model: DocumentModel,
  request: DocumentTableInsertCommand | DocumentTableCreateCommand,
  point?: DocumentPoint,
  sourceModel: DocumentModel = model,
): DocumentBlock {
  const { rows, columns } = assertDocumentTableInsertion(model, request)
  const options = { ...request }
  if (options.id === 'doc.command.create-table' && options.width === undefined) {
    const ancestry = point ? documentBlockAncestry(model, point.cellBlockId ?? point.blockId) : undefined
    const hostSectionId = model.metadata?.documentEditingSectionId
    const sectionId = typeof hostSectionId === 'string' ? hostSectionId : sectionIdForBlock(model, ancestry?.[0]?.id ?? point?.blockId ?? '')
    const page = documentEffectiveSectionSetup(sourceModel, sectionId).page
    options.width = page.width - page.marginLeft - page.marginRight
  }
  const table: DocumentTable = createTable(rows, columns)
  table.columnWidths = Array.from({ length: columns }, (_, column) => options.columnWidths?.[column] ?? (options.width === undefined ? 100 : options.width / columns))
  table.style = { width: table.columnWidths.reduce((sum, width) => sum + width, 0), layout: 'fixed',
    ...(options.cellMargin ? { cellMargin: { ...options.cellMargin } } : {}) }
  table.headerRowCount = options.headerRowCount ?? (options.titleRow ? 1 : 0)
  table.titleRow = options.titleRow ?? false
  if (options.columnTypes) table.columnTypes = cloneModel(options.columnTypes)
  for (const [row, source] of table.rows.entries()) {
    source.height = options.rowHeights?.[row] ?? 30
    source.heightRule = 'atLeast'
    for (const [column, cell] of source.cells.entries()) cell.blocks = (options.values?.[row]?.[column] ?? '').replace(/\r\n?/g, '\n').split('\n').map(text => createParagraph(text))
  }
  if (table.titleRow) {
    const title = table.rows[0]!.cells[0]!, content = table.rows[0]!.cells.flatMap(cell => cell.blocks).filter(paragraph => runText(paragraph.runs).length)
    title.blocks = content.length ? content : [title.blocks[0]!]
    for (const paragraph of title.blocks) paragraph.style.align = 'center'
    title.colSpan = columns; table.rows[0]!.cells = [title]
  }
  if (options.style) applyDocumentTableStylePatch(table, documentTableGrid(table), options.style, table.style.width!)
  documentTableGrid(table)
  const block: DocumentBlock = { ...createParagraph(), type: 'table', runs: [], metadata: { ...cloneModel(options.metadata ?? {}), table } }
  if (options.tableId !== undefined) block.id = options.tableId
  // Retain extension metadata; the typed table always owns its geometry and values.
  if (!tableForBlock(block)) throw new TypeError('表格模型构建失败')
  return block
}
