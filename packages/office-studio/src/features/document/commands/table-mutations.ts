import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import {
  clampPoint, collapsedSelection, createParagraph, createRun, documentBlockAncestry, documentId,
  findDocumentBlock, findDocumentBlockContainer, runText, sliceRuns, tableForBlock, textTargets,
  type DocumentCell, type DocumentSelection, type DocumentTable, type DocumentTableBorder, type DocumentTableBorderSide,
  type DocumentTableRow, type DocumentRun,
} from '../model/document-model'
import {
  assertDocumentTableRange, documentTableCellsInRange, documentTableGrid, documentTableWholeRange, expandDocumentTableRange,
  documentTableCellTarget, type DocumentTableCellRange, type DocumentTableGrid, type DocumentTableGridCell,
} from '../model/document-table-grid'
import {
  assertDocumentTableBorder, assertDocumentTableColor, assertDocumentTableColumnType, assertDocumentTableMargin, DOCUMENT_TABLE_BORDER_PRESETS,
  DOCUMENT_TABLE_BORDER_SIDES, DOCUMENT_TABLE_DEFAULT_BORDER, DOCUMENT_TABLE_MIN_ROW_HEIGHT,
  documentTableCellBorder, documentTableCellBorders, documentTableCellFill, documentTableCellMargin, documentTableColumnWidths, documentTableRowProperties,
  applyDocumentTableRowPropertiesPatch, assertDocumentTableRowPropertiesPatch,
  assertDocumentTableBorderPatch, assertDocumentTableMarginPatch, documentTableBorderWithPatch,
} from '../model/document-table-style'
import { documentEffectiveSectionSetup } from '../services/document-settings'
import { sectionIdForBlock } from '../services/section-header-footer'
import { planDocumentTableFieldAnchors, reconcileDocumentTableAnchors } from '../services/document-table-anchors'
import { documentPointAtClampedBodyOffset, insertDocumentBlocksAtPoint } from '../services/document-block-insertion'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { assertTextRangeEditable, transformContentControlsForBlockEdit, transformContentControlsForBodyRangeEdit, validateDocumentContentControlRanges } from '../services/content-controls'
import { documentFieldLocked, documentTextFieldParagraphs, documentTextFieldRanges, transformDocumentTextFieldsForBlockRangeEdit } from '../model/document-text-fields'
import { transformDocumentCommentAnchorsForBodyRangeEdit } from '../comments/document-comment-anchors'
import { reconcileDocumentNotesAfterStructure } from '../notes/document-notes'
import { assertDocumentPermission, permissionTargetsForBlock, permissionTargetsForChangedBlocks, type DocumentPermissionContext } from '../permissions/document-permissions'
import type { DocumentTableCommand, DocumentTableStylePatch } from './table-command-types'
import { createDocumentTableViewSelection, rebaseDocumentTableViewSelection, resolveDocumentTableViewSelection } from '../model/document-table-selection'
import { copyDocumentBlocksForPaste } from '../services/document-block-paste'
import { applyDocumentPasteDestinationStyle, type DocumentPasteDestinationStyle } from '../services/document-clipboard'
import { validateDocumentRunPatch } from '../model/document-run-range'
import { documentTableCellBlocks, documentTableCellContentContainers, documentTableContinuationEntries, rebindDocumentTableContinuations } from '../model/document-table-continuations'

export interface DocumentTableMutationTarget { blockId: string; range: DocumentTableCellRange }
export interface DocumentTableMutationPlan { model: DocumentModel; selection: DocumentSelection; changed: boolean }
interface Placement { cell: DocumentCell; row: number; column: number; rowSpan: number; colSpan: number }

function index(value: number, size: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 0 || value >= size) throw new RangeError(`表格${label}超出范围`)
  return value
}
function count(value = 1): number {
  if (!Number.isSafeInteger(value) || value < 1) throw new RangeError('表格行列数量必须为正整数')
  return value
}
function dimension(value: number): number {
  if (!Number.isFinite(value) || value <= 0) throw new RangeError('表格尺寸必须为正 CSS px')
  return value
}
function placements(grid: DocumentTableGrid): Placement[] {
  return grid.anchors.map(owner => ({ cell: owner.cell, row: owner.row, column: owner.column, rowSpan: owner.rowSpan, colSpan: owner.colSpan }))
}
function rebuild(table: DocumentTable, cells: Placement[], rows = table.rows): void {
  table.rows = rows
  for (const row of rows) row.cells = []
  cells.sort((left, right) => left.row - right.row || left.column - right.column)
  for (const owner of cells) {
    owner.cell.columnIndex = owner.column
    owner.cell.rowSpan = owner.rowSpan
    owner.cell.colSpan = owner.colSpan
    rows[owner.row]!.cells.push(owner.cell)
    rebindDocumentTableContinuations(table, owner.cell, owner.row)
  }
}
function blankCell(source: DocumentCell, column: number, colSpan = 1): DocumentCell {
  const paragraph = createParagraph()
  const template = documentTableCellBlocks(source).find(block => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type))
  if (template) paragraph.style = { ...template.style }
  const style = cloneModel(source.style)
  delete style.nativeContinuations
  delete style.nativeContinuationContent
  delete style.selectedDropdownOptionId
  return { ...source, id: documentId('cell'), columnIndex: column, rowSpan: 1, colSpan,
    style, blocks: [paragraph] }
}

export function documentTableReferenceWidth(model: DocumentModel, tableId: string): number {
  const ancestry = documentBlockAncestry(model, tableId)
  const section = documentEffectiveSectionSetup(model, sectionIdForBlock(model, ancestry?.[0]?.id ?? tableId))
  let available = (section.page.width - section.page.marginLeft - section.page.marginRight - section.columnGap * (section.columnCount - 1)) / section.columnCount
  for (let depth = 0; depth < (ancestry?.length ?? 0) - 1; depth++) {
    const parent = ancestry![depth]!, child = ancestry![depth + 1]!, table = tableForBlock(parent)
    if (!table) continue
    const grid = documentTableGrid(table), owner = grid.anchors.find(cell => documentTableCellBlocks(cell.cell).some(block => block.id === child.id))
    if (!owner) continue
    const widths = documentTableColumnWidths(table, grid.columnCount, available), margin = documentTableCellMargin(table, owner.cell)
    available = Math.max(1, widths.slice(owner.column, owner.column + owner.colSpan).reduce((sum, width) => sum + width, 0) - margin.left - margin.right)
  }
  return available
}

export function documentTableSelectionForCell(model: DocumentModel, blockId: string, row: number, column: number, offset = 0): DocumentSelection {
  const block = findDocumentBlock(model, blockId), table = block && tableForBlock(block)
  if (!table) throw new RangeError('表格已不存在')
  const owner = documentTableGrid(table).cellAt(row, column)
  if (!owner) throw new RangeError('表格单元格已不存在')
  const descendants = new Set(allBlocks(documentTableCellBlocks(owner.cell)).map(block => block.id))
  const target = textTargets(model).find(candidate => descendants.has(candidate.cellBlockId ?? candidate.blockId))
  if (target) return collapsedSelection(clampPoint(model, { ...target, offset }))
  const paragraph = createParagraph()
  owner.cell.blocks.push(paragraph)
  return collapsedSelection({ ...documentTableCellTarget(blockId, owner, paragraph.id), offset: 0 })
}

function allBlocks(blocks: readonly DocumentBlock[]): DocumentBlock[] {
  return blocks.flatMap(block => [block, ...((tableForBlock(block)?.rows.flatMap(row => row.cells.flatMap(cell => allBlocks(documentTableCellBlocks(cell))))) ?? [])])
}

/** Permissions cover unchanged paragraphs whose cell membership or geometry is changed too. */
function selectedCellsOnly(command: DocumentTableCommand): boolean {
  return ['doc.command.table-merge-cells', 'doc.command.table-unmerge-cells', 'doc.command.table-set-background',
    'doc.command.table-set-border', 'doc.command.table-set-border-color', 'doc.command.table-set-border-width', 'doc.command.table-set-border-style',
    'doc.command.table-set-vertical-align', 'doc.command.table-set-text-align', 'doc.command.table-set-cell-margin', 'doc.command.table-set-cell-text', 'doc.command.table-set-range-text'].includes(command.id)
    && !('scope' in command && command.scope === 'table')
}

export function documentTablePermissionTargets(model: DocumentModel, grid: DocumentTableGrid, target: DocumentTableMutationTarget, command: DocumentTableCommand) {
  const owners = selectedCellsOnly(command) ? documentTableCellsInRange(grid, expandDocumentTableRange(grid, target.range)) : grid.anchors
  const affected = new Set(owners)
  if (['doc.command.table-set-border', 'doc.command.table-set-border-color', 'doc.command.table-set-border-width', 'doc.command.table-set-border-style'].includes(command.id)) {
    const range = expandDocumentTableRange(grid, target.range)
    for (const owner of owners) for (const side of DOCUMENT_TABLE_BORDER_SIDES) {
      if (command.id === 'doc.command.table-set-border' && !borderPresetApplies(owner, side, range, command.preset)) continue
      for (const neighbour of sharedBorderNeighbours(grid, owner, side)) affected.add(neighbour)
    }
  }
  const offsetTarget = command.id === 'doc.command.table-move' && command.targetOffset !== undefined ? documentPointAtClampedBodyOffset(model, command.targetOffset) : undefined
  const ids = [target.blockId, ...[...affected].flatMap(owner => allBlocks(documentTableCellBlocks(owner.cell)).map(block => block.id)),
    ...(offsetTarget ? [offsetTarget.cellBlockId ?? offsetTarget.blockId] : []),
    ...('targetBlockId' in command && command.targetBlockId ? [command.targetBlockId] : [])]
  return ids.flatMap(id => permissionTargetsForBlock(model, id))
}

function insertRows(table: DocumentTable, grid: DocumentTableGrid, at: number, amount: number, templateRow: number): void {
  const cells = placements(grid), source = table.rows[templateRow]!
  const rows = [...table.rows]
  const { cells: _cells, ...properties } = source
  rows.splice(at, 0, ...Array.from({ length: amount }, (): DocumentTableRow => ({ ...cloneModel(properties), id: documentId('row'), cells: [] })))
  for (const owner of cells) {
    if (owner.row >= at) owner.row += amount
    else if (owner.row + owner.rowSpan > at) owner.rowSpan += amount
  }
  for (let row = at; row < at + amount; row++) {
    for (let column = 0; column < grid.columnCount;) {
      const spanning = cells.find(owner => row >= owner.row && row < owner.row + owner.rowSpan && column >= owner.column && column < owner.column + owner.colSpan)
      if (spanning) { column = spanning.column + spanning.colSpan; continue }
      const template = grid.cellAt(templateRow, column)!
      let colSpan = Math.min(template.column + template.colSpan - column, grid.columnCount - column)
      const nextSpanning = cells.filter(owner => row >= owner.row && row < owner.row + owner.rowSpan && owner.column > column).sort((a, b) => a.column - b.column)[0]
      if (nextSpanning) colSpan = Math.min(colSpan, nextSpanning.column - column)
      cells.push({ cell: blankCell(template.cell, column, colSpan), row, column, rowSpan: 1, colSpan })
      column += colSpan
    }
  }
  if (at < (table.headerRowCount ?? 0)) table.headerRowCount = (table.headerRowCount ?? 0) + amount
  if (table.rowStyles) table.rowStyles.splice(at, 0, ...Array.from({ length: amount }, () => cloneModel(table.rowStyles![templateRow] ?? {})))
  rebuild(table, cells, rows)
}

function insertColumns(table: DocumentTable, grid: DocumentTableGrid, at: number, amount: number, templateColumn: number): void {
  const cells = placements(grid)
  for (const owner of cells) {
    if (owner.column >= at) owner.column += amount
    else if (owner.column + owner.colSpan > at) owner.colSpan += amount
  }
  for (let row = 0; row < grid.rowCount; row++) for (let column = at; column < at + amount; column++) {
    if (cells.some(owner => row >= owner.row && row < owner.row + owner.rowSpan && column >= owner.column && column < owner.column + owner.colSpan)) continue
    cells.push({ cell: blankCell(grid.cellAt(row, templateColumn)!.cell, column), row, column, rowSpan: 1, colSpan: 1 })
  }
  if (table.columnWidths?.length) {
    const total = table.columnWidths.reduce((sum, width) => sum + width, 0), inserted = table.columnWidths[templateColumn]!
    const scale = total / (total + inserted * amount)
    const widths = table.columnWidths.map(width => width * scale)
    widths.splice(at, 0, ...Array.from({ length: amount }, () => inserted * scale))
    table.columnWidths = widths
  }
  if (table.columnTypes) table.columnTypes = Object.fromEntries(Object.entries(table.columnTypes).map(([column, config]) => [Number(column) >= at ? Number(column) + amount : Number(column), config]))
  rebuild(table, cells)
}

function deleteRows(table: DocumentTable, grid: DocumentTableGrid, start: number, end: number): void {
  const amount = end - start + 1
  const cells = placements(grid).flatMap(owner => {
    const overlap = Math.max(0, Math.min(owner.row + owner.rowSpan, end + 1) - Math.max(owner.row, start))
    if (overlap === owner.rowSpan) return []
    return [{ ...owner, row: owner.row < start ? owner.row : owner.row > end ? owner.row - amount : start, rowSpan: owner.rowSpan - overlap }]
  })
  const header = table.headerRowCount ?? 0
  table.headerRowCount = header - Math.max(0, Math.min(end + 1, header) - start)
  if (start === 0) table.titleRow = false
  if (table.rowStyles) table.rowStyles.splice(start, amount)
  const rows = [...table.rows]
  rows.splice(start, amount)
  rebuild(table, cells, rows)
}

function deleteColumns(table: DocumentTable, grid: DocumentTableGrid, start: number, end: number): void {
  const amount = end - start + 1
  const cells = placements(grid).flatMap(owner => {
    const overlap = Math.max(0, Math.min(owner.column + owner.colSpan, end + 1) - Math.max(owner.column, start))
    if (overlap === owner.colSpan) return []
    return [{ ...owner, column: owner.column < start ? owner.column : owner.column > end ? owner.column - amount : start, colSpan: owner.colSpan - overlap }]
  })
  if (table.columnWidths?.length) {
    table.columnWidths.splice(start, amount)
    if (table.style?.width !== undefined) table.style.width = table.columnWidths.reduce((sum, width) => sum + width, 0)
  }
  if (table.columnTypes) table.columnTypes = Object.fromEntries(Object.entries(table.columnTypes).flatMap(([column, config]) => Number(column) >= start && Number(column) <= end ? [] : [[Number(column) > end ? Number(column) - amount : Number(column), config]]))
  rebuild(table, cells)
}

function mergeCells(table: DocumentTable, grid: DocumentTableGrid, source: DocumentTableCellRange): void {
  const range = expandDocumentTableRange(grid, source), selected = documentTableCellsInRange(grid, range)
  if (selected.length < 2) throw new RangeError('合并需要至少两个单元格')
  const first = grid.cellAt(range.startRow, range.startColumn)!, lastTop = grid.cellAt(range.startRow, range.endColumn)!, lastBottom = grid.cellAt(range.endRow, range.startColumn)!
  const selectedIds = new Set(selected.map(owner => owner.cell.id))
  const content = selected.flatMap(owner => documentTableCellBlocks(owner.cell))
  for (const owner of selected) { delete owner.cell.style.nativeContinuationContent; delete owner.cell.style.nativeContinuations }
  first.cell.blocks = content
  first.cell.style.borders = { top: documentTableCellBorder(grid, first, 'top'), left: documentTableCellBorder(grid, first, 'left'),
    right: documentTableCellBorder(grid, lastTop, 'right'), bottom: documentTableCellBorder(grid, lastBottom, 'bottom') }
  const merged: Placement = { cell: first.cell, row: range.startRow, column: range.startColumn,
    rowSpan: range.endRow - range.startRow + 1, colSpan: range.endColumn - range.startColumn + 1 }
  rebuild(table, [...placements(grid).filter(owner => !selectedIds.has(owner.cell.id)), merged])
}

function unmergeCells(table: DocumentTable, grid: DocumentTableGrid, range: DocumentTableCellRange): void {
  const selected = documentTableCellsInRange(grid, expandDocumentTableRange(grid, range)).filter(owner => owner.rowSpan > 1 || owner.colSpan > 1)
  if (!selected.length) throw new RangeError('选区内没有合并单元格')
  const ids = new Set(selected.map(owner => owner.cell.id)), cells = placements(grid).filter(owner => !ids.has(owner.cell.id))
  for (const owner of selected) {
    const borders = documentTableCellBorders(grid, owner)
    const continuations = new Map(documentTableContinuationEntries(owner.cell).map(entry => [entry.offset, entry.content]))
    const native = owner.cell.style.nativeContinuations as Record<string, unknown> | undefined
    delete owner.cell.style.nativeContinuationContent
    delete owner.cell.style.nativeContinuations
    for (let row = owner.row; row < owner.row + owner.rowSpan; row++) for (let column = owner.column; column < owner.column + owner.colSpan; column++) {
      const cell = row === owner.row && column === owner.column ? owner.cell : blankCell(owner.cell, column)
      if (row > owner.row && column === owner.column) {
        const continuation = continuations.get(row - owner.row)
        if (continuation) cell.blocks = continuation.blocks
        const properties = native?.[row - owner.row]
        if (properties && typeof properties === 'object') cell.style.nativeProperties = properties
      }
      cell.style = { ...cell.style, borders: {
        top: row === owner.row ? borders.top : { ...DOCUMENT_TABLE_DEFAULT_BORDER },
        bottom: row === owner.row + owner.rowSpan - 1 ? borders.bottom : { ...DOCUMENT_TABLE_DEFAULT_BORDER },
        left: column === owner.column ? borders.left : { ...DOCUMENT_TABLE_DEFAULT_BORDER },
        right: column === owner.column + owner.colSpan - 1 ? borders.right : { ...DOCUMENT_TABLE_DEFAULT_BORDER },
      } }
      cells.push({ cell, row, column, rowSpan: 1, colSpan: 1 })
    }
  }
  rebuild(table, cells)
}

function assertAxisBoundary(grid: DocumentTableGrid, boundary: number, axis: 'row' | 'column'): void {
  if (grid.anchors.some(owner => owner[axis] < boundary && owner[axis] + (axis === 'row' ? owner.rowSpan : owner.colSpan) > boundary)) {
    throw new RangeError('移动或表头边界不能切开合并单元格')
  }
}

function moveAxis(table: DocumentTable, grid: DocumentTableGrid, axis: 'row' | 'column', from: number, amount: number, to: number, position: 'before' | 'after'): void {
  const size = axis === 'row' ? grid.rowCount : grid.columnCount
  index(from, size, axis === 'row' ? '起始行' : '起始列'); index(to, size, axis === 'row' ? '目标行' : '目标列')
  if (from + amount > size) throw new RangeError('移动的表格行列数量超出范围')
  if (to >= from && to < from + amount) return
  const boundary = to + (position === 'after' ? 1 : 0)
  for (const edge of [from, from + amount, boundary]) assertAxisBoundary(grid, edge, axis)
  const order = Array.from({ length: size }, (_, value) => value), moving = order.splice(from, amount)
  order.splice(boundary > from ? boundary - amount : boundary, 0, ...moving)
  const mapped = new Map(order.map((oldIndex, newIndex) => [oldIndex, newIndex]))
  const cells = placements(grid).map(owner => ({ ...owner, [axis]: mapped.get(owner[axis])! }))
  if (axis === 'row') {
    if (table.rowStyles) table.rowStyles = order.map(oldIndex => table.rowStyles![oldIndex] ?? {})
    if (table.titleRow && order[0] !== 0) table.titleRow = false
    rebuild(table, cells, order.map(oldIndex => table.rows[oldIndex]!))
  } else {
    if (table.columnWidths?.length) table.columnWidths = order.map(oldIndex => table.columnWidths![oldIndex]!)
    if (table.columnTypes) table.columnTypes = Object.fromEntries(Object.entries(table.columnTypes).map(([oldIndex, config]) => [mapped.get(Number(oldIndex))!, config]))
    rebuild(table, cells)
  }
}

function sharedBorderNeighbours(grid: DocumentTableGrid, owner: DocumentTableGridCell, side: DocumentTableBorderSide): Set<DocumentTableGridCell> {
  const horizontal = side === 'top' || side === 'bottom'
  const fixed = side === 'top' ? owner.row - 1 : side === 'bottom' ? owner.row + owner.rowSpan : side === 'left' ? owner.column - 1 : owner.column + owner.colSpan
  const adjacent = new Set<DocumentTableGridCell>()
  for (let value = horizontal ? owner.column : owner.row; value < (horizontal ? owner.column + owner.colSpan : owner.row + owner.rowSpan); value++) {
    const neighbour = horizontal ? grid.cellAt(fixed, value) : grid.cellAt(value, fixed)
    if (neighbour && neighbour !== owner) adjacent.add(neighbour)
  }
  return adjacent
}

function setSharedBorder(grid: DocumentTableGrid, owner: DocumentTableGridCell, side: DocumentTableBorderSide, border: DocumentTableBorder): void {
  owner.cell.style.borders = { ...owner.cell.style.borders, [side]: { ...border } }
  const opposite: DocumentTableBorderSide = side === 'top' ? 'bottom' : side === 'bottom' ? 'top' : side === 'left' ? 'right' : 'left'
  for (const neighbour of sharedBorderNeighbours(grid, owner, side)) neighbour.cell.style.borders = { ...neighbour.cell.style.borders, [opposite]: { ...border } }
}

function borderPresetApplies(owner: DocumentTableGridCell, side: DocumentTableBorderSide, range: DocumentTableCellRange, preset: Extract<DocumentTableCommand, { id: 'doc.command.table-set-border' }>['preset']): boolean {
  const outer = side === 'top' ? owner.row === range.startRow : side === 'bottom' ? owner.row + owner.rowSpan - 1 === range.endRow
    : side === 'left' ? owner.column === range.startColumn : owner.column + owner.colSpan - 1 === range.endColumn
  return preset === 'all' || preset === 'none' || preset === 'outer' && outer || preset === 'inner' && !outer
    || preset === side && outer || preset === 'innerHorizontal' && !outer && (side === 'top' || side === 'bottom')
    || preset === 'innerVertical' && !outer && (side === 'left' || side === 'right')
}

function setBorders(grid: DocumentTableGrid, source: DocumentTableCellRange, command: Extract<DocumentTableCommand, { id: 'doc.command.table-set-border' }>): void {
  if (!DOCUMENT_TABLE_BORDER_PRESETS.includes(command.preset)) throw new TypeError('表格边框预设无效')
  if (command.borderPatch !== undefined) {
    if (command.border !== undefined) throw new TypeError('表格边框完整值与局部更改不能同时提供')
    assertDocumentTableBorderPatch(command.borderPatch)
  }
  const range = expandDocumentTableRange(grid, source)
  const changes = documentTableCellsInRange(grid, range).flatMap(owner => DOCUMENT_TABLE_BORDER_SIDES.flatMap(side => {
    if (!borderPresetApplies(owner, side, range, command.preset)) return []
    const base = command.borderPatch ? documentTableCellBorder(grid, owner, side) : command.border ?? DOCUMENT_TABLE_DEFAULT_BORDER
    const border = { ...(command.borderPatch ? documentTableBorderWithPatch(base, command.borderPatch) : base),
      ...(command.preset === 'none' ? { width: 0 } : {}) }
    assertDocumentTableBorder(border)
    return [{ owner, side, border }]
  }))
  for (const { owner, side, border } of changes) setSharedBorder(grid, owner, side, border)
}

function patchBorders(grid: DocumentTableGrid, range: DocumentTableCellRange, patch: Partial<DocumentTableBorder>): void {
  assertDocumentTableBorderPatch(patch)
  const changes = documentTableCellsInRange(grid, expandDocumentTableRange(grid, range)).flatMap(owner => DOCUMENT_TABLE_BORDER_SIDES.map(side => {
    const border = documentTableBorderWithPatch(documentTableCellBorder(grid, owner, side), patch)
    assertDocumentTableBorder(border)
    return { owner, side, border }
  }))
  for (const { owner, side, border } of changes) setSharedBorder(grid, owner, side, border)
}

export function assertDocumentTableStylePatch(patch: DocumentTableStylePatch): void {
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new TypeError('表格样式参数无效')
  if (patch.backgroundColor !== undefined) assertDocumentTableColor(patch.backgroundColor)
  if (patch.align !== undefined && !['left', 'center', 'right'].includes(patch.align)) throw new TypeError('整表对齐无效')
  if (patch.layout !== undefined && !['auto', 'fixed'].includes(patch.layout)) throw new TypeError('表格布局规则无效')
  for (const value of [patch.indent, patch.offsetY]) if (value !== undefined && !Number.isFinite(value)) throw new RangeError('表格位置必须为有限 CSS px')
  if (patch.width !== undefined) dimension(patch.width)
  if (patch.columnWidths !== undefined && (!Array.isArray(patch.columnWidths) || patch.columnWidths.length < 1 || patch.columnWidths.length > 512
    || patch.columnWidths.some(width => !Number.isFinite(width) || width <= 0))) throw new RangeError('表格列宽必须包含有效的正 CSS px')
  if (patch.cellMargin !== undefined && patch.cellMargin !== null) assertDocumentTableMargin(patch.cellMargin)
  if (patch.border !== undefined) {
    if (!patch.border || typeof patch.border !== 'object' || Array.isArray(patch.border)) throw new TypeError('表格边框参数无效')
    if (patch.border.preset !== undefined && !DOCUMENT_TABLE_BORDER_PRESETS.includes(patch.border.preset)) throw new TypeError('表格边框预设无效')
    if (patch.border.patch !== undefined) {
      if ([patch.border.border, patch.border.color, patch.border.width, patch.border.style].some(value => value !== undefined)) throw new TypeError('表格边框完整值与局部更改不能同时提供')
      assertDocumentTableBorderPatch(patch.border.patch)
    }
    if (patch.border.border !== undefined) assertDocumentTableBorder(patch.border.border)
    if (patch.border.color !== undefined) assertDocumentTableColor(patch.border.color)
    assertDocumentTableBorder({ ...DOCUMENT_TABLE_DEFAULT_BORDER, ...patch.border.border,
      ...(patch.border.color !== undefined ? { color: patch.border.color ?? DOCUMENT_TABLE_DEFAULT_BORDER.color } : {}),
      ...(patch.border.width !== undefined ? { width: patch.border.width } : {}),
      ...(patch.border.style !== undefined ? { style: patch.border.style } : {}) })
  }
}

export function applyDocumentTableStylePatch(table: DocumentTable, grid: DocumentTableGrid, patch: DocumentTableStylePatch, referenceWidth: number): void {
  assertDocumentTableStylePatch(patch)
  if (patch.columnWidths && patch.columnWidths.length !== grid.columnCount) throw new RangeError('表格测量列宽与逻辑列数不一致')
  table.style ??= {}
  if (patch.backgroundColor !== undefined) {
    if (patch.backgroundColor === null) { delete table.style.fill; delete table.style.background }
    else table.style.fill = patch.backgroundColor
    for (const owner of grid.anchors) {
      delete owner.cell.style.background
      if (patch.backgroundColor === null) delete owner.cell.style.fill
      else owner.cell.style.fill = patch.backgroundColor
    }
  }
  if (patch.border) {
    const border = { ...DOCUMENT_TABLE_DEFAULT_BORDER, ...patch.border.border,
      ...(patch.border.color !== undefined ? { color: patch.border.color ?? DOCUMENT_TABLE_DEFAULT_BORDER.color } : {}),
      ...(patch.border.width !== undefined ? { width: patch.border.width } : {}),
      ...(patch.border.style !== undefined ? { style: patch.border.style } : {}) }
    setBorders(grid, documentTableWholeRange(grid), { id: 'doc.command.table-set-border', preset: patch.border.preset ?? 'all',
      ...(patch.border.patch ? { borderPatch: patch.border.patch } : { border }) })
  }
  if (patch.columnWidths) {
    table.columnWidths = [...patch.columnWidths]
    table.style.width = patch.columnWidths.reduce((sum, width) => sum + width, 0)
  }
  if (patch.width !== undefined) {
    const widths = documentTableColumnWidths(table, grid.columnCount, referenceWidth), total = widths.reduce((sum, width) => sum + width, 0)
    table.columnWidths = widths.map(width => width * patch.width! / total)
    table.style.width = patch.width
  }
  for (const property of ['align', 'indent', 'offsetY', 'layout'] as const) if (patch[property] !== undefined) {
    Object.assign(table.style, { [property]: patch[property] })
  }
  if (patch.cellMargin !== undefined) {
    if (patch.cellMargin === null) delete table.style.cellMargin
    else table.style.cellMargin = { ...patch.cellMargin }
  }
}

function deleteTable(model: DocumentModel, block: DocumentBlock): DocumentSelection {
  const container = findDocumentBlockContainer(model, block.id)!, at = container.indexOf(block)
  const previousTargets = textTargets(model), ids = new Set(allBlocks([block]).map(child => child.id))
  const firstIndex = previousTargets.findIndex(target => ids.has(target.cellBlockId ?? target.blockId))
  const following = previousTargets.slice(Math.max(0, firstIndex)).find(target => !ids.has(target.cellBlockId ?? target.blockId))
  const preceding = previousTargets.slice(0, Math.max(0, firstIndex)).at(-1)
  container.splice(at, 1)
  if (!documentTextFieldParagraphs({ blocks: container }).length) container.splice(Math.min(at, container.length), 0, createParagraph())
  if (following) return collapsedSelection(clampPoint(model, { ...following, offset: 0 }))
  if (preceding) return collapsedSelection(clampPoint(model, { ...preceding, offset: Number.MAX_SAFE_INTEGER }))
  const target = textTargets(model)[0]
  return target ? collapsedSelection({ ...target, offset: 0 }) : collapsedSelection({ blockId: '', offset: 0 })
}

function isTableCellTextNoop(cell: DocumentCell, paragraphs: readonly { text: string }[], normalized: string): boolean {
  return paragraphs.length > 0 && paragraphs.map(paragraph => paragraph.text).join('\n') === normalized
    && documentTableCellBlocks(cell).every(block => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type))
}

function assertCellTextEditable(model: DocumentModel, owner: DocumentTableGridCell, text: string): void {
  const ids = new Set(allBlocks(documentTableCellBlocks(owner.cell)).map(block => block.id)), inputs = documentTextFieldParagraphs(model).filter(paragraph => ids.has(paragraph.block.id))
  if (isTableCellTextNoop(owner.cell, inputs, text.replace(/\r\n?/g, '\n'))) return
  if (model.contentControls?.some(control => ids.has(control.blockId) && ['controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(control.lock ?? ''))) throw new RangeError('单元格文本不能覆盖锁定内容控件')
  const first = inputs[0], last = inputs.at(-1)
  if (!first || !last) return
  assertTextRangeEditable(model, first.absoluteStart, last.absoluteEnd)
  if (documentTextFieldRanges(model).some(range => documentFieldLocked(range.field) && (range.absoluteStart === range.absoluteEnd
    ? range.absoluteStart >= first.absoluteStart && range.absoluteEnd <= last.absoluteEnd
    : range.absoluteStart < last.absoluteEnd && range.absoluteEnd > first.absoluteStart))) throw new RangeError('单元格文本不能覆盖锁定字段')
}

function tableRangeTextEdits(grid: DocumentTableGrid, range: DocumentTableCellRange,
  command: Extract<DocumentTableCommand, { id: 'doc.command.table-set-range-text' }>): Array<{ owner: DocumentTableGridCell; text: string }> {
  if (!['clear', 'replace', 'matrix'].includes(command.mode)) throw new TypeError('表格内容操作模式无效')
  if (command.marks) validateDocumentRunPatch(command.marks)
  const owners = documentTableCellsInRange(grid, expandDocumentTableRange(grid, range))
  if (command.mode !== 'matrix') {
    if (command.mode === 'replace' && typeof command.text !== 'string') throw new TypeError('表格替换内容必须是文本')
    return owners.map((owner, index) => ({ owner, text: command.mode === 'replace' && index === 0 ? command.text! : '' }))
  }
  if (!Array.isArray(command.values) || !command.values.length || command.values.some(row => !Array.isArray(row) || !row.length || row.some(value => typeof value !== 'string'))) throw new TypeError('表格粘贴数据必须是文本矩阵')
  const columns = command.values.reduce((maximum, row) => Math.max(maximum, row.length), 0)
  if (command.values.length * columns > 250_000) throw new RangeError('表格单次粘贴超出单元格内存预算')
  if (range.startRow + command.values.length - 1 > range.endRow || range.startColumn + columns - 1 > range.endColumn) throw new RangeError('表格粘贴数据超出目标范围')
  const edited = new Map<string, { owner: DocumentTableGridCell; text: string }>()
  for (let row = 0; row < command.values.length; row++) for (let column = 0; column < columns; column++) {
    const logicalRow = range.startRow + row, logicalColumn = range.startColumn + column, owner = grid.cellAt(logicalRow, logicalColumn)!
    const text = command.values[row]![column] ?? ''
    if (owner.row !== logicalRow || owner.column !== logicalColumn) {
      if (text) throw new RangeError('粘贴数据包含合并覆盖位置的独立内容，请先取消合并')
      continue
    }
    edited.set(owner.cell.id, { owner, text })
  }
  return [...edited.values()]
}

function setCellText(model: DocumentModel, blockId: string, grid: DocumentTableGrid, row: number, column: number, text: string, marks?: Extract<DocumentTableCommand, { id: 'doc.command.table-set-range-text' }>['marks']): DocumentSelection {
  const owner = grid.cellAt(row, column)
  if (!owner || typeof text !== 'string') throw new RangeError('表格单元格或文本无效')
  assertCellTextEditable(model, owner, text)
  const normalized = text.replace(/\r\n?/g, '\n'), before = cloneModel(model), source = documentTextFieldParagraphs(before)
  const ids = new Set(allBlocks(documentTableCellBlocks(owner.cell)).map(block => block.id)), inputs = source.filter(paragraph => ids.has(paragraph.block.id))
  const first = inputs[0], last = inputs.at(-1)
  if (!marks && isTableCellTextNoop(owner.cell, inputs, normalized)) {
    return documentTableSelectionForCell(model, blockId, row, column, normalized.length)
  }
  const template = first?.block ?? createParagraph(), parts = normalized.split('\n'), outputs: DocumentBlock[] = parts.map((_line, lineIndex) => ({
    ...(lineIndex === 0 ? cloneModel(template) : createParagraph()), style: { ...template.style },
    runs: [], metadata: lineIndex === 0 ? { ...template.metadata } : {},
  }))
  const { id: _id, text: _text, metadata: _metadata, href: _href, ...appearance } = (template.runs[0] as DocumentRun | undefined) ?? createRun()
  for (const [lineIndex, output] of outputs.entries()) {
    output.runs = [createRun(parts[lineIndex]!, { ...appearance, ...marks })]
    delete output.metadata.html
    delete output.metadata.codeBlock
    delete output.metadata.latexFormulaBlock
  }
  if (first && last) {
    assertTextRangeEditable(model, first.absoluteStart, last.absoluteEnd)
    if (documentTextFieldRanges(before).some(range => documentFieldLocked(range.field) && (range.absoluteStart === range.absoluteEnd
      ? range.absoluteStart >= first.absoluteStart && range.absoluteEnd <= last.absoluteEnd
      : range.absoluteStart < last.absoluteEnd && range.absoluteEnd > first.absoluteStart))) throw new RangeError('单元格文本不能覆盖锁定字段')
    transformDocumentTextFieldsForBlockRangeEdit(inputs.map(paragraph => paragraph.block), 0, last.text.length, normalized, outputs, undefined, model)
  }
  owner.cell.blocks = outputs
  delete owner.cell.style.nativeContinuationContent
  if (first && last) {
    transformContentControlsForBodyRangeEdit(model, source, first.absoluteStart, last.absoluteEnd, normalized.length, false)
    transformDocumentCommentAnchorsForBodyRangeEdit(model, source, documentTextFieldParagraphs(model), first.absoluteStart, last.absoluteEnd, normalized)
  }
  reconcileDocumentNotesAfterStructure(model, new Set(inputs.filter(paragraph => !outputs.some(block => block.id === paragraph.block.id)).map(paragraph => paragraph.block.id)))
  validateDocumentContentControlRanges(model); documentTextFieldRanges(model)
  const lastOutput = outputs.at(-1)!, current = documentTableGrid(tableForBlock(findDocumentBlock(model, blockId)!)!).cellAt(row, column)!
  return collapsedSelection({ ...documentTableCellTarget(blockId, current, lastOutput.id), offset: runText(lastOutput.runs).length })
}

function cellText(cell: DocumentCell): string {
  return allBlocks(documentTableCellBlocks(cell)).filter(block => !tableForBlock(block)).map(block => runText(block.runs)).join('\n')
}
function sortTable(table: DocumentTable, grid: DocumentTableGrid, column: number, direction: 'ascending' | 'descending', locale?: string): void {
  if (direction !== 'ascending' && direction !== 'descending') throw new TypeError('表格排序方向无效')
  index(column, grid.columnCount, '排序列')
  const header = Math.max(table.headerRowCount ?? 0, table.titleRow ? 1 : 0)
  assertAxisBoundary(grid, header, 'row')
  const groups: number[][] = []
  for (let row = header; row < grid.rowCount;) {
    let end = row + 1
    for (let current = row; current < end; current++) for (const owner of grid.rows[current]!) end = Math.max(end, owner.row + owner.rowSpan)
    groups.push(Array.from({ length: end - row }, (_, offset) => row + offset)); row = end
  }
  const collator = new Intl.Collator(locale, { numeric: true, sensitivity: 'base' }), config = table.columnTypes?.[column]
  const key = (row: number): string | number => {
    const text = cellText(grid.cellAt(row, column)!.cell).trim()
    if (config?.type === 'number') { const number = Number(text); return text && Number.isFinite(number) ? number : text }
    if (config?.type === 'date' && /^\d{4}-\d{2}-\d{2}(?:T.*)?$/.test(text)) { const date = Date.parse(text); if (Number.isFinite(date)) return date }
    if (config?.type === 'checkbox') return /^(?:true|1|yes|checked|☑|✓)$/i.test(text) ? 1 : 0
    return text
  }
  const sorted = groups.map((rows, order) => ({ rows, order, key: key(rows[0]!) })).sort((left, right) => {
    const comparison = typeof left.key === 'number' && typeof right.key === 'number' ? left.key - right.key : collator.compare(String(left.key), String(right.key))
    return (direction === 'descending' ? -comparison : comparison) || left.order - right.order
  })
  const order = [...Array.from({ length: header }, (_, row) => row), ...sorted.flatMap(group => group.rows)]
  const mapped = new Map(order.map((oldRow, newRow) => [oldRow, newRow]))
  if (table.rowStyles) table.rowStyles = order.map(oldRow => table.rowStyles![oldRow] ?? {})
  rebuild(table, placements(grid).map(owner => ({ ...owner, row: mapped.get(owner.row)! })), order.map(oldRow => table.rows[oldRow]!))
}

function insertTitleRow(table: DocumentTable, grid: DocumentTableGrid, text: string): void {
  insertRows(table, grid, 0, 1, 0)
  const created = documentTableGrid(table)
  if (created.columnCount > 1) mergeCells(table, created, { startRow: 0, endRow: 0, startColumn: 0, endColumn: created.columnCount - 1 })
  table.titleRow = true; table.headerRowCount = Math.max(1, table.headerRowCount ?? 0)
  const title = table.rows[0]!.cells[0]!
  title.blocks = [createParagraph(text)]
  title.blocks[0]!.style.align = 'center'
}

/** A topology projection shares untouched text and never mutates the source. */
function tableFieldProjection(model: DocumentModel, block: DocumentBlock, command: DocumentTableCommand, target: DocumentTableMutationTarget) {
  const table = tableForBlock(block)!, source = documentTextFieldParagraphs(model)
  const ids = new Set(allBlocks([block]).map(child => child.id)), body = source.map(paragraph => paragraph.block)
  const owned = body.filter(paragraph => ids.has(paragraph.id)), first = body.findIndex(paragraph => ids.has(paragraph.id))
  const transformedFieldIds = new Set<string>()
  if (command.id === 'doc.command.table-move') {
    if (command.targetOffset !== undefined) {
      const point = documentPointAtClampedBodyOffset(model, command.targetOffset), destinationId = point.cellBlockId ?? point.blockId
      const flat = body.map(paragraph => ({ ...paragraph, metadata: { ...paragraph.metadata } })), destination = flat.find(paragraph => paragraph.id === destinationId)!
      const tail = createParagraph(); tail.style = { ...destination.style }; tail.runs = sliceRuns(destination.runs, point.offset)
      const shell: DocumentModel = { ...model, blocks: flat }
      for (const field of documentTextFieldRanges(model)) transformedFieldIds.add(field.field.id)
      transformDocumentTextFieldsForBlockRangeEdit([destination], point.offset, point.offset, '\n', [destination, tail], undefined, shell)
      destination.runs = sliceRuns(destination.runs, 0, point.offset)
      const moving = flat.filter(paragraph => ids.has(paragraph.id)), remaining = flat.filter(paragraph => !ids.has(paragraph.id))
      remaining.splice(remaining.indexOf(destination) + 1, 0, ...moving, tail)
      return { paragraphs: documentTextFieldParagraphs({ blocks: remaining }), transformedFieldIds }
    }
    if (!command.targetBlockId) return undefined
    const destination = findDocumentBlock(model, command.targetBlockId)!, destinationIds = new Set(allBlocks([destination]).map(child => child.id))
    const remaining = body.filter(paragraph => !ids.has(paragraph.id)), indices = remaining.flatMap((paragraph, index) => destinationIds.has(paragraph.id) ? [index] : [])
    // Non-text objects still have a structural position in the source DFS.
    // Reordering across them must validate the same paragraph order as execute.
    const ordered = allBlocks(model.blocks).filter(child => !ids.has(child.id)), rank = new Map(ordered.map((child, index) => [child.id, index]))
    const destinationRank = rank.get(destination.id)!
    const at = indices.length ? command.position === 'after' ? indices.at(-1)! + 1 : indices[0]!
      : remaining.reduce((total, paragraph) => total + ((rank.get(paragraph.id) ?? Number.POSITIVE_INFINITY) < destinationRank ? 1 : 0), 0)
    remaining.splice(at, 0, ...owned)
    return { paragraphs: documentTextFieldParagraphs({ blocks: remaining }), transformedFieldIds }
  }
  if (!owned.length || first < 0) return undefined
  const copy: DocumentTable = { ...table,
    rows: table.rows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell, style: { ...cell.style,
      ...(cell.style.nativeContinuationContent ? { nativeContinuationContent: Object.fromEntries(Object.entries(cell.style.nativeContinuationContent).map(([offset, content]) => [offset, { ...content, blocks: [...content.blocks] }])) } : {}) }, blocks: [...cell.blocks] })) })),
    columnWidths: table.columnWidths && [...table.columnWidths], columnTypes: table.columnTypes && { ...table.columnTypes },
    rowStyles: table.rowStyles && [...table.rowStyles], style: table.style && { ...table.style },
  }
  const grid = documentTableGrid(copy), range = target.range
  let remove = false
  switch (command.id) {
    case 'doc.command.table-insert-row': {
      const row = command.row ?? (command.position === 'before' ? range.startRow : range.endRow)
      insertRows(copy, grid, row + (command.position === 'after' ? 1 : 0), count(command.count), row); break
    }
    case 'doc.command.table-insert-column': {
      const column = command.column ?? (command.position === 'before' ? range.startColumn : range.endColumn)
      insertColumns(copy, grid, column + (command.position === 'after' ? 1 : 0), count(command.count), column); break
    }
    case 'doc.command.table-delete-rows':
      if (range.startRow === 0 && range.endRow === grid.rowCount - 1) remove = true
      else deleteRows(copy, grid, range.startRow, range.endRow)
      break
    case 'doc.command.table-delete-columns':
      if (range.startColumn === 0 && range.endColumn === grid.columnCount - 1) remove = true
      else deleteColumns(copy, grid, range.startColumn, range.endColumn)
      break
    case 'doc.command.table-delete': remove = true; break
    case 'doc.command.table-merge-cells': mergeCells(copy, grid, range); break
    case 'doc.command.table-unmerge-cells': unmergeCells(copy, grid, range); break
    case 'doc.command.table-move-rows': moveAxis(copy, grid, 'row', command.fromRow ?? range.startRow, count(command.count ?? range.endRow - range.startRow + 1), command.toRow, command.position); break
    case 'doc.command.table-move-columns': moveAxis(copy, grid, 'column', command.fromColumn ?? range.startColumn, count(command.count ?? range.endColumn - range.startColumn + 1), command.toColumn, command.position); break
    case 'doc.command.table-insert-title-row': insertTitleRow(copy, grid, command.text ?? ''); break
    case 'doc.command.table-sort': sortTable(copy, grid, command.column ?? range.startColumn, command.direction, command.locale); break
    default: return undefined
  }
  if (!remove) documentTableGrid(copy)
  const paragraphs = remove ? [] : documentTextFieldParagraphs({ blocks: [{ ...block, metadata: { ...block.metadata, table: copy } }] }).map(paragraph => paragraph.block)
  body.splice(first, owned.length, ...paragraphs)
  return { paragraphs: documentTextFieldParagraphs({ blocks: body }), transformedFieldIds }
}

/** Shared preflight for Ribbon, Canvas, Facade and the detached mutation planner. */
export function assertDocumentTableMutation(model: DocumentModel, _selection: DocumentSelection, command: DocumentTableCommand, target: DocumentTableMutationTarget, context: DocumentPermissionContext): void {
  const block = findDocumentBlock(model, target.blockId), table = block && tableForBlock(block)
  if (!block || !table) throw new RangeError('表格已不存在')
  const grid = documentTableGrid(table), range = target.range
  assertDocumentTableRange(grid, range)
  if ('position' in command && command.position !== undefined && command.position !== 'before' && command.position !== 'after') throw new TypeError('表格插入或移动位置无效')
  if ('scope' in command && command.scope !== undefined && command.scope !== 'table' && command.scope !== 'selection') throw new TypeError('表格样式作用域无效')
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', documentTablePermissionTargets(model, grid, target, command))
  const selected = documentTableCellsInRange(grid, expandDocumentTableRange(grid, range))
  const modern = model.metadata?.documentFlavor === 'modern' || model.metadata?.documentFlavor === 2 || model.metadata?.mode === 'modern'
  const validateMove = (axis: 'row' | 'column', from: number, to: number, amount: number, position: 'before' | 'after') => {
    const size = axis === 'row' ? grid.rowCount : grid.columnCount
    index(from, size, '移动源'); index(to, size, '移动目标'); count(amount)
    if (from + amount > size) throw new RangeError('移动的表格行列数量超出范围')
    if (to < from || to >= from + amount) for (const boundary of [from, from + amount, to + (position === 'after' ? 1 : 0)]) assertAxisBoundary(grid, boundary, axis)
  }
  switch (command.id) {
    case 'doc.command.table-insert-row':
      index(command.row ?? (command.position === 'before' ? range.startRow : range.endRow), grid.rowCount, '插入行'); count(command.count)
      if (grid.rowCount + count(command.count) > 10_000 || (grid.rowCount + count(command.count)) * grid.columnCount > 1_000_000 || count(command.count) * grid.columnCount > 250000) throw new RangeError('表格插行超出解析、写出或单次单元格内存预算')
      break
    case 'doc.command.table-insert-column':
      index(command.column ?? (command.position === 'before' ? range.startColumn : range.endColumn), grid.columnCount, '插入列'); count(command.count)
      if (grid.columnCount + count(command.count) > 512 || (grid.columnCount + count(command.count)) * grid.rowCount > 1_000_000 || count(command.count) * grid.rowCount > 250000) throw new RangeError('表格插列超出解析、写出或单次单元格内存预算')
      break
    case 'doc.command.table-merge-cells': if (selected.length < 2) throw new RangeError('合并需要至少两个单元格'); break
    case 'doc.command.table-unmerge-cells': if (!selected.some(owner => owner.rowSpan > 1 || owner.colSpan > 1)) throw new RangeError('选区内没有合并单元格'); break
    case 'doc.command.table-move-rows': validateMove('row', command.fromRow ?? range.startRow, command.toRow, command.count ?? range.endRow - range.startRow + 1, command.position); break
    case 'doc.command.table-move-columns': validateMove('column', command.fromColumn ?? range.startColumn, command.toColumn, command.count ?? range.endColumn - range.startColumn + 1, command.position); break
    case 'doc.command.table-move': {
      const destination = command.targetOffset !== undefined ? documentPointAtClampedBodyOffset(model, command.targetOffset) : undefined
      const destinationId = destination?.cellBlockId ?? destination?.blockId ?? command.targetBlockId
      if (destinationId && allBlocks([block]).some(child => child.id === destinationId)) throw new RangeError('表格不能移动到自身或后代内部')
      if (destinationId && !findDocumentBlock(model, destinationId)) throw new RangeError('表格移动目标已不存在')
      if (destination) {
        assertDocumentTextSelectionEditable(model, collapsedSelection(destination))
        const tail = createParagraph(), controls: DocumentModel = { ...model, contentControls: getDetachedControls(model) }
        transformContentControlsForBlockEdit(controls, destinationId!, destination.offset, destination.offset, 0, { blockId: tail.id, firstBreakOffset: destination.offset, lastParagraphOffset: destination.offset })
      }
      for (const value of [command.deltaX, command.deltaY]) if (value !== undefined && !Number.isFinite(value)) throw new RangeError('表格移动距离必须是有限 CSS px')
      break
    }
    case 'doc.command.table-resize-row': {
      const start = command.row === undefined ? range.startRow : index(command.row, grid.rowCount, '尺寸行'), end = command.row === undefined ? range.endRow : start
      assertDocumentTableRowPropertiesPatch(table, Array.from({ length: end - start + 1 }, (_, offset) => start + offset), command)
      break
    }
    case 'doc.command.table-resize-column':
      dimension(command.width); if (command.column !== undefined) index(command.column, grid.columnCount, '尺寸列')
      if (command.widths !== undefined && (!Array.isArray(command.widths) || command.widths.length !== grid.columnCount || command.widths.some(width => !Number.isFinite(width) || width <= 0))) throw new RangeError('表格测量列宽与逻辑列数不一致')
      break
    case 'doc.command.table-distribute-rows': if (command.heights && (command.heights.length !== grid.rowCount || command.heights.some(height => !Number.isFinite(height) || height <= 0))) throw new RangeError('表格测量行高与逻辑行数不一致'); break
    case 'doc.command.table-set-background': assertDocumentTableColor(command.color); break
    case 'doc.command.table-set-border':
      if (!DOCUMENT_TABLE_BORDER_PRESETS.includes(command.preset)) throw new TypeError('表格边框预设无效')
      if (command.borderPatch !== undefined) {
        if (command.border !== undefined) throw new TypeError('表格边框完整值与局部更改不能同时提供')
        assertDocumentTableBorderPatch(command.borderPatch)
      } else assertDocumentTableBorder(command.border ?? DOCUMENT_TABLE_DEFAULT_BORDER)
      break
    case 'doc.command.table-set-border-color': assertDocumentTableColor(command.color); break
    case 'doc.command.table-set-border-width': if (!Number.isFinite(command.width) || command.width < 0) throw new RangeError('表格边框宽度无效'); break
    case 'doc.command.table-set-border-style': if (!['solid', 'dot', 'dash'].includes(command.style)) throw new TypeError('表格边框样式无效'); break
    case 'doc.command.table-set-vertical-align': if (!['top', 'center', 'bottom'].includes(command.value)) throw new TypeError('表格单元格垂直对齐无效'); break
    case 'doc.command.table-set-text-align': if (!['left', 'center', 'right', 'justify'].includes(command.value)) throw new TypeError('表格文字对齐无效'); break
    case 'doc.command.table-set-align':
      if (modern) throw new RangeError('整表对齐适用于传统分页文档')
      if (!['left', 'center', 'right'].includes(command.value)) throw new TypeError('整表对齐无效')
      break
    case 'doc.command.table-set-cell-margin':
      if (command.marginPatch !== undefined) {
        if (command.margin !== undefined) throw new TypeError('表格边距完整值与局部更改不能同时提供')
        assertDocumentTableMarginPatch(command.marginPatch)
      } else if (command.margin === undefined) throw new TypeError('表格边距值缺失')
      else if (command.margin !== null) assertDocumentTableMargin(command.margin)
      break
    case 'doc.command.table-set-cell-text': {
      const owner = grid.cellAt(command.row, command.column)
      if (!owner || typeof command.text !== 'string') throw new RangeError('表格单元格或文本无效')
      if (command.optionId !== undefined) {
        const config = table.columnTypes?.[owner.column], option = config?.dropdownOptions?.find(option => option.id === command.optionId)
        if (config?.type !== 'dropdown' || !option || option.label !== command.text) throw new RangeError('表格下拉选项与当前列类型不一致')
      }
      assertCellTextEditable(model, owner, command.text)
      break
    }
    case 'doc.command.table-set-range-text':
      for (const edit of tableRangeTextEdits(grid, range, command)) assertCellTextEditable(model, edit.owner, edit.text)
      break
    case 'doc.command.table-set-header-row-count':
      if (!Number.isSafeInteger(command.count) || command.count < 0 || command.count > grid.rowCount) throw new RangeError('表格表头行数无效')
      assertAxisBoundary(grid, command.count, 'row'); break
    case 'doc.command.table-insert-title-row':
      if (table.titleRow) throw new RangeError('表格已包含标题行')
      if (grid.rowCount >= 10_000 || (grid.rowCount + 1) * grid.columnCount > 1_000_000 || grid.columnCount > 250_000) throw new RangeError('表格标题行超出解析与写出预算')
      if (command.text !== undefined && typeof command.text !== 'string') throw new TypeError('表格标题必须是文本')
      break
    case 'doc.command.table-set-column-type': index(command.column ?? range.startColumn, grid.columnCount, '类型列'); assertDocumentTableColumnType(command.config); break
    case 'doc.command.table-sort':
      index(command.column ?? range.startColumn, grid.columnCount, '排序列')
      if (command.direction !== 'ascending' && command.direction !== 'descending') throw new TypeError('表格排序方向无效')
      if (command.locale !== undefined && typeof command.locale !== 'string') throw new TypeError('表格排序语言无效')
      assertAxisBoundary(grid, Math.max(table.headerRowCount ?? 0, table.titleRow ? 1 : 0), 'row'); break
    case 'doc.command.table-set-style':
      assertDocumentTableStylePatch(command.patch)
      if (command.patch.columnWidths && command.patch.columnWidths.length !== grid.columnCount) throw new RangeError('表格测量列宽与逻辑列数不一致')
      if (modern && command.patch.align !== undefined) throw new RangeError('整表对齐适用于传统分页文档')
      break
  }
  if (['doc.command.table-delete', 'doc.command.table-delete-rows', 'doc.command.table-delete-columns'].includes(command.id)) {
    const deleting = command.id === 'doc.command.table-delete' ? grid.anchors : grid.anchors.filter(owner => command.id === 'doc.command.table-delete-rows'
      ? owner.row >= range.startRow && owner.row + owner.rowSpan - 1 <= range.endRow
      : owner.column >= range.startColumn && owner.column + owner.colSpan - 1 <= range.endColumn)
    const ids = new Set(deleting.flatMap(owner => allBlocks(documentTableCellBlocks(owner.cell)).map(child => child.id)))
    if (model.contentControls?.some(control => ids.has(control.blockId) && ['controlLocked', 'sdtLocked', 'sdtContentLocked'].includes(control.lock ?? ''))) throw new RangeError('表格选区包含锁定内容控件')
  }
  const projection = tableFieldProjection(model, block, command, target)
  if (projection) planDocumentTableFieldAnchors(model, projection.paragraphs, projection.transformedFieldIds)
}

function getDetachedControls(model: DocumentModel) { return model.contentControls?.map(control => cloneModel(control)) }

function applyMutation(model: DocumentModel, selection: DocumentSelection, command: DocumentTableCommand, target: DocumentTableMutationTarget): { selection: DocumentSelection; structural: boolean; transformedFieldIds?: ReadonlySet<string> } {
  const block = findDocumentBlock(model, target.blockId)!, table = tableForBlock(block)!, grid = documentTableGrid(table), range = target.range
  assertDocumentTableRange(grid, range)
  const contentBefore = JSON.stringify(model)
  let next = selection, structural = false, transformedFieldIds: ReadonlySet<string> | undefined
  const originalAncestors = documentBlockAncestry(model, block.id) ?? []
  const selected = () => documentTableCellsInRange(grid, expandDocumentTableRange(grid, range))
  switch (command.id) {
    case 'doc.command.table-insert-row': {
      const sourceRow = index(command.row ?? (command.position === 'before' ? range.startRow : range.endRow), grid.rowCount, '插入行')
      const at = sourceRow + (command.position === 'after' ? 1 : 0)
      insertRows(table, grid, at, count(command.count), sourceRow)
      next = documentTableSelectionForCell(model, block.id, at, range.startColumn); structural = true; break
    }
    case 'doc.command.table-insert-column': {
      const sourceColumn = index(command.column ?? (command.position === 'before' ? range.startColumn : range.endColumn), grid.columnCount, '插入列')
      const at = sourceColumn + (command.position === 'after' ? 1 : 0)
      insertColumns(table, grid, at, count(command.count), sourceColumn)
      next = documentTableSelectionForCell(model, block.id, range.startRow, at); structural = true; break
    }
    case 'doc.command.table-delete-rows': {
      if (range.startRow === 0 && range.endRow === grid.rowCount - 1) next = deleteTable(model, block)
      else { deleteRows(table, grid, range.startRow, range.endRow); next = documentTableSelectionForCell(model, block.id, Math.min(range.startRow, table.rows.length - 1), range.startColumn) }
      structural = true; break
    }
    case 'doc.command.table-delete-columns': {
      if (range.startColumn === 0 && range.endColumn === grid.columnCount - 1) next = deleteTable(model, block)
      else { deleteColumns(table, grid, range.startColumn, range.endColumn); next = documentTableSelectionForCell(model, block.id, range.startRow, Math.min(range.startColumn, grid.columnCount - (range.endColumn - range.startColumn + 1) - 1)) }
      structural = true; break
    }
    case 'doc.command.table-delete': next = deleteTable(model, block); structural = true; break
    case 'doc.command.table-merge-cells': mergeCells(table, grid, range); next = documentTableSelectionForCell(model, block.id, range.startRow, range.startColumn); structural = true; break
    case 'doc.command.table-unmerge-cells': unmergeCells(table, grid, range); structural = true; break
    case 'doc.command.table-move-rows': moveAxis(table, grid, 'row', command.fromRow ?? range.startRow, count(command.count ?? range.endRow - range.startRow + 1), command.toRow, command.position); structural = true; break
    case 'doc.command.table-move-columns': moveAxis(table, grid, 'column', command.fromColumn ?? range.startColumn, count(command.count ?? range.endColumn - range.startColumn + 1), command.toColumn, command.position); structural = true; break
    case 'doc.command.table-move': {
      if (command.targetOffset !== undefined) {
        const insertion = insertDocumentBlocksAtPoint(model, documentPointAtClampedBodyOffset(model, command.targetOffset), [block], block)
        transformedFieldIds = insertion.transformedFieldIds
      } else if (command.targetBlockId) {
        if (allBlocks([block]).some(child => child.id === command.targetBlockId)) throw new RangeError('表格不能移动到自身或后代内部')
        const destination = findDocumentBlockContainer(model, command.targetBlockId), source = findDocumentBlockContainer(model, block.id)!
        if (!destination) throw new RangeError('表格移动目标已不存在')
        source.splice(source.indexOf(block), 1)
        destination.splice(destination.findIndex(child => child.id === command.targetBlockId) + (command.position === 'after' ? 1 : 0), 0, block)
        if (!documentTextFieldParagraphs({ blocks: source }).length) source.push(createParagraph())
      }
      for (const [parameter, property] of [['deltaX', 'indent'], ['deltaY', 'offsetY']] as const) if (command[parameter] !== undefined) {
        if (!Number.isFinite(command[parameter])) throw new RangeError('表格移动距离必须是有限 CSS px')
        if (command[parameter] !== 0) table.style = { ...table.style, [property]: (table.style?.[property] ?? 0) + command[parameter]! }
      }
      structural = true; break
    }
    case 'doc.command.table-resize-row': {
      const start = command.row === undefined ? range.startRow : index(command.row, grid.rowCount, '尺寸行'), end = command.row === undefined ? range.endRow : start
      applyDocumentTableRowPropertiesPatch(table, Array.from({ length: end - start + 1 }, (_, offset) => start + offset), command)
      break
    }
    case 'doc.command.table-resize-column': {
      const widths = command.widths ? [...command.widths] : documentTableColumnWidths(table, grid.columnCount, documentTableReferenceWidth(model, block.id))
      const start = command.column === undefined ? range.startColumn : index(command.column, grid.columnCount, '尺寸列'), end = command.column === undefined ? range.endColumn : start
      for (let column = start; column <= end; column++) widths[column] = dimension(command.width)
      table.columnWidths = widths; table.style = { ...table.style, width: widths.reduce((sum, width) => sum + width, 0), layout: 'fixed' }; break
    }
    case 'doc.command.table-distribute-rows': {
      if (command.heights && (command.heights.length !== grid.rowCount || command.heights.some(height => !Number.isFinite(height) || height <= 0))) throw new RangeError('表格测量行高与逻辑行数不一致')
      const heights = table.rows.map((_row, row) => command.heights?.[row] ?? documentTableRowProperties(table, row).height ?? DOCUMENT_TABLE_MIN_ROW_HEIGHT)
      const height = heights.slice(range.startRow, range.endRow + 1).reduce((sum, value) => sum + value, 0) / (range.endRow - range.startRow + 1)
      for (let row = range.startRow; row <= range.endRow; row++) { table.rows[row]!.height = height; table.rows[row]!.heightRule = 'atLeast' }
      break
    }
    case 'doc.command.table-distribute-columns': {
      const widths = documentTableColumnWidths(table, grid.columnCount, documentTableReferenceWidth(model, block.id))
      const width = widths.slice(range.startColumn, range.endColumn + 1).reduce((sum, value) => sum + value, 0) / (range.endColumn - range.startColumn + 1)
      for (let column = range.startColumn; column <= range.endColumn; column++) widths[column] = width
      table.columnWidths = widths; table.style = { ...table.style, width: widths.reduce((sum, value) => sum + value, 0), layout: 'fixed' }; break
    }
    case 'doc.command.table-set-background': {
      assertDocumentTableColor(command.color)
      const owners = command.scope === 'table' ? grid.anchors : selected()
      if (command.scope === 'table') { table.style ??= {}; if (command.color === null) delete table.style.fill; else table.style.fill = command.color }
      for (const owner of owners) { if (command.color === null) { delete owner.cell.style.fill; delete owner.cell.style.background } else owner.cell.style.fill = command.color }
      break
    }
    case 'doc.command.table-set-border': setBorders(grid, range, command); break
    case 'doc.command.table-set-border-color': assertDocumentTableColor(command.color); patchBorders(grid, range, { color: command.color ?? DOCUMENT_TABLE_DEFAULT_BORDER.color }); break
    case 'doc.command.table-set-border-width': if (!Number.isFinite(command.width) || command.width < 0) throw new RangeError('表格边框宽度无效'); patchBorders(grid, range, { width: command.width }); break
    case 'doc.command.table-set-border-style': if (!['solid', 'dot', 'dash'].includes(command.style)) throw new TypeError('表格边框样式无效'); patchBorders(grid, range, { style: command.style }); break
    case 'doc.command.table-set-vertical-align': {
      if (!['top', 'center', 'bottom'].includes(command.value)) throw new TypeError('表格单元格垂直对齐无效')
      for (const owner of selected()) owner.cell.style.verticalAlign = command.value
      break
    }
    case 'doc.command.table-set-text-align': {
      if (!['left', 'center', 'right', 'justify'].includes(command.value)) throw new TypeError('表格文字对齐无效')
      for (const owner of selected()) for (const paragraph of allBlocks(documentTableCellBlocks(owner.cell)).filter(child => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(child.type))) paragraph.style.align = command.value
      break
    }
    case 'doc.command.table-set-align': {
      if (model.metadata?.documentFlavor === 'modern' || model.metadata?.documentFlavor === 2 || model.metadata?.mode === 'modern') throw new RangeError('整表对齐适用于传统分页文档')
      if (!['left', 'center', 'right'].includes(command.value)) throw new TypeError('整表对齐无效')
      table.style = { ...table.style, align: command.value }; break
    }
    case 'doc.command.table-set-cell-margin': {
      if (command.marginPatch !== undefined) {
        assertDocumentTableMarginPatch(command.marginPatch)
        const patch = Object.fromEntries(Object.entries(command.marginPatch).filter(([, value]) => value !== undefined))
        if (command.scope === 'table') { table.style ??= {}; table.style.cellMargin = { ...documentTableCellMargin(table), ...patch } }
        else for (const owner of selected()) owner.cell.style.margin = { ...documentTableCellMargin(table, owner.cell), ...patch }
      } else {
        if (command.margin === undefined) throw new TypeError('表格边距值缺失')
        if (command.margin !== null) assertDocumentTableMargin(command.margin)
        if (command.scope === 'table') { table.style ??= {}; if (command.margin === null) delete table.style.cellMargin; else table.style.cellMargin = { ...command.margin } }
        else for (const owner of selected()) { if (command.margin === null) delete owner.cell.style.margin; else owner.cell.style.margin = { ...command.margin } }
      }
      break
    }
    case 'doc.command.table-set-cell-text':
      transformedFieldIds = new Set(documentTextFieldRanges(model).map(range => range.field.id))
      next = setCellText(model, block.id, grid, command.row, command.column, command.text)
      if (command.optionId === undefined) delete grid.cellAt(command.row, command.column)!.cell.style.selectedDropdownOptionId
      else grid.cellAt(command.row, command.column)!.cell.style.selectedDropdownOptionId = command.optionId
      structural = true; break
    case 'doc.command.table-set-range-text': {
      const edits = tableRangeTextEdits(grid, range, command)
      transformedFieldIds = new Set(documentTextFieldRanges(model).map(range => range.field.id))
      for (const edit of [...edits].reverse()) {
        const result = setCellText(model, block.id, grid, edit.owner.row, edit.owner.column, edit.text, command.marks)
        delete edit.owner.cell.style.selectedDropdownOptionId
        if (edit.owner === edits[0]?.owner && command.mode === 'replace') next = result
      }
      structural = true; break
    }
    case 'doc.command.table-set-header-row-count': {
      if (!Number.isSafeInteger(command.count) || command.count < 0 || command.count > grid.rowCount) throw new RangeError('表格表头行数无效')
      assertAxisBoundary(grid, command.count, 'row'); table.headerRowCount = command.count; break
    }
    case 'doc.command.table-insert-title-row': {
      insertTitleRow(table, grid, command.text ?? '')
      next = documentTableSelectionForCell(model, block.id, 0, 0); structural = true; break
    }
    case 'doc.command.table-set-column-type': {
      const column = index(command.column ?? range.startColumn, grid.columnCount, '类型列')
      assertDocumentTableColumnType(command.config); table.columnTypes = { ...table.columnTypes, [column]: cloneModel(command.config) }; break
    }
    case 'doc.command.table-sort': sortTable(table, grid, command.column ?? range.startColumn, command.direction, command.locale); structural = true; break
    case 'doc.command.table-set-style': applyDocumentTableStylePatch(table, grid, command.patch, documentTableReferenceWidth(model, block.id)); break
  }
  if (findDocumentBlock(model, block.id)) documentTableGrid(table)
  if (JSON.stringify(model) !== contentBefore) for (const ancestor of [...originalAncestors, ...(documentBlockAncestry(model, block.id) ?? [])]) delete ancestor.metadata.html
  return { selection: { anchor: clampPoint(model, next.anchor), focus: clampPoint(model, next.focus) }, structural, transformedFieldIds }
}

/** Every mutation, anchor plan and permission check completes before the caller publishes the model. */
export function prepareDocumentTableMutation(before: DocumentModel, selection: DocumentSelection, command: DocumentTableCommand, target: DocumentTableMutationTarget, context: DocumentPermissionContext): DocumentTableMutationPlan {
  assertDocumentTableMutation(before, selection, command, target, context)
  const sourceBlock = findDocumentBlock(before, target.blockId)!, sourceGrid = documentTableGrid(tableForBlock(sourceBlock)!)
  assertDocumentTableRange(sourceGrid, target.range)
  assertDocumentPermission(before.permissionRules ?? [], context, 'edit', documentTablePermissionTargets(before, sourceGrid, target, command))
  const model = cloneModel(before), result = applyMutation(model, selection, command, target)
  if (result.structural) reconcileDocumentTableAnchors(before, model, result.selection, result.transformedFieldIds)
  const afterBlock = findDocumentBlock(model, target.blockId), afterGrid = afterBlock && tableForBlock(afterBlock) && documentTableGrid(tableForBlock(afterBlock)!)
  const afterTargets = afterGrid ? documentTablePermissionTargets(model, afterGrid, { blockId: target.blockId,
    range: selectedCellsOnly(command) ? target.range : documentTableWholeRange(afterGrid) }, command) : []
  const previousCells = new Map(sourceGrid.anchors.map(owner => [owner.cell.id, owner.cell]))
  const changedCellTargets = afterGrid?.anchors.flatMap(owner => JSON.stringify(previousCells.get(owner.cell.id)?.style) !== JSON.stringify(owner.cell.style)
    ? allBlocks(documentTableCellBlocks(owner.cell)).flatMap(block => [...permissionTargetsForBlock(before, block.id), ...permissionTargetsForBlock(model, block.id)]) : []) ?? []
  assertDocumentPermission(before.permissionRules ?? [], context, 'edit', [...permissionTargetsForChangedBlocks(before, model), ...afterTargets, ...changedCellTargets])
  const tableSelection = command.id === 'doc.command.table-set-cell-text' || command.id === 'doc.command.table-set-range-text' && command.mode === 'replace' ? undefined
    : rebaseDocumentTableViewSelection(before, model, selection.tableSelection, result.structural)
  return { model, selection: { anchor: clampPoint(model, result.selection.anchor), focus: clampPoint(model, result.selection.focus),
    ...(tableSelection ? { tableSelection } : {}) }, changed: JSON.stringify(before) !== JSON.stringify(model) }
}

function assertCellContentReplaceEditable(model: DocumentModel, owner: DocumentTableGridCell): void {
  const ids = new Set(allBlocks(documentTableCellBlocks(owner.cell)).map(block => block.id))
  if (model.contentControls?.some(control => ids.has(control.blockId)
    && ['controlLocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'].includes(control.lock ?? ''))) {
    throw new RangeError('单元格粘贴不能覆盖锁定内容控件')
  }
  const paragraphs = documentTextFieldParagraphs(model), first = paragraphs.find(paragraph => ids.has(paragraph.block.id))
  const last = paragraphs.filter(paragraph => ids.has(paragraph.block.id)).at(-1)
  if (!first || !last) return
  assertTextRangeEditable(model, first.absoluteStart, last.absoluteEnd)
  if (documentTextFieldRanges(model).some(range => documentFieldLocked(range.field) && (range.absoluteStart === range.absoluteEnd
    ? range.absoluteStart >= first.absoluteStart && range.absoluteEnd <= last.absoluteEnd
    : range.absoluteStart < last.absoluteEnd && range.absoluteEnd > first.absoluteStart))) {
    throw new RangeError('单元格粘贴不能覆盖锁定字段')
  }
}

function pasteCellStyle(grid: DocumentTableGrid, owner: DocumentTableGridCell): DocumentCell['style'] {
  return { ...cloneModel(owner.cell.style), fill: documentTableCellFill(grid.table, owner.cell),
    margin: documentTableCellMargin(grid.table, owner.cell), borders: documentTableCellBorders(grid, owner) }
}

function pasteDestinationStyle(cell: DocumentCell): DocumentPasteDestinationStyle {
  const paragraph = allBlocks(documentTableCellBlocks(cell)).find(block => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type))
    ?? createParagraph()
  const { id: _id, text: _text, ...textStyle } = (paragraph.runs[0] as DocumentRun | undefined) ?? createRun()
  return { paragraphStyle: { ...paragraph.style }, textStyle }
}

/**
 * A rich table occupies its source rectangle at the logical selection origin.
 * Ordinary blocks replace the first selected cell's complete content container.
 * The existing table dimensions and column types remain the destination's state.
 */
export function prepareDocumentTablePasteMutation(before: DocumentModel, selection: DocumentSelection,
  sourceBlocks: readonly DocumentBlock[], context: DocumentPermissionContext, mode: 'source' | 'destination' = 'source'): DocumentTableMutationPlan {
  if (!selection.tableSelection) throw new RangeError('表格粘贴需要结构选区')
  if (mode !== 'source' && mode !== 'destination') throw new TypeError('表格粘贴格式模式无效')
  if (!Array.isArray(sourceBlocks) || sourceBlocks.some(block => !block || !block.metadata || !block.style || !Array.isArray(block.runs))) throw new TypeError('表格富文本粘贴必须包含规范化的文档块')
  const view = resolveDocumentTableViewSelection(before, selection.tableSelection)
  const block = findDocumentBlock(before, view.tableId)!, table = tableForBlock(block)!, grid = documentTableGrid(table)
  const sourceTable = sourceBlocks.length === 1 ? tableForBlock(sourceBlocks[0]!) : undefined
  const sourceGrid = sourceTable && documentTableGrid(sourceTable)
  if (!sourceBlocks.length) return { model: before, selection, changed: false }
  const first = grid.cellAt(view.range.startRow, view.range.startColumn)!
  const singleCell = !sourceGrid || sourceGrid.rowCount === 1 && sourceGrid.columnCount === 1
  const range: DocumentTableCellRange = singleCell
    ? { startRow: first.row, endRow: first.row + first.rowSpan - 1, startColumn: first.column, endColumn: first.column + first.colSpan - 1 }
    : { startRow: view.range.startRow, endRow: view.range.startRow + sourceGrid.rowCount - 1,
      startColumn: view.range.startColumn, endColumn: view.range.startColumn + sourceGrid.columnCount - 1 }
  if (range.endRow >= grid.rowCount || range.endColumn >= grid.columnCount) throw new RangeError('粘贴数据超过现有表格，请先插入足够的行列')
  const expanded = expandDocumentTableRange(grid, range)
  if (expanded.startRow !== range.startRow || expanded.endRow !== range.endRow
    || expanded.startColumn !== range.startColumn || expanded.endColumn !== range.endColumn) {
    throw new RangeError('富文本粘贴范围不能切开已有合并单元格，请先取消合并')
  }
  const targets = documentTableCellsInRange(grid, range)
  const permissionTargets = [view.tableId, ...targets.flatMap(owner => allBlocks(documentTableCellBlocks(owner.cell)).map(child => child.id))]
    .flatMap(id => permissionTargetsForBlock(before, id))
  assertDocumentPermission(before.permissionRules ?? [], context, 'edit', permissionTargets)
  for (const owner of targets) assertCellContentReplaceEditable(before, owner)

  const copies = copyDocumentBlocksForPaste(sourceBlocks), incomingTable = sourceGrid ? tableForBlock(copies[0]!)! : undefined
  const incomingGrid = incomingTable && documentTableGrid(incomingTable)
  const model = cloneModel(before), destinationBlock = findDocumentBlock(model, view.tableId)!
  const destinationTable = tableForBlock(destinationBlock)!, destinationGrid = documentTableGrid(destinationTable)
  if (singleCell) {
    const owner = destinationGrid.cellAt(first.row, first.column)!
    owner.cell.blocks = incomingGrid ? documentTableCellBlocks(incomingGrid.anchors[0]!.cell) : copies
    delete owner.cell.style.nativeContinuationContent
    if (mode === 'destination') applyDocumentPasteDestinationStyle(owner.cell.blocks, pasteDestinationStyle(first.cell))
    if (incomingGrid && mode === 'source') {
      const style = pasteCellStyle(incomingGrid, incomingGrid.anchors[0]!)
      // The destination's vertical merge still owns its physical continuation properties.
      delete style.nativeContinuations
      delete style.nativeContinuationContent
      owner.cell.style = { ...owner.cell.style, ...style }
    }
    delete owner.cell.style.selectedDropdownOptionId
  } else {
    const replacedIds = new Set(documentTableCellsInRange(destinationGrid, range).map(owner => owner.cell.id))
    const rowIds = new Map(incomingTable!.rows.map((row, index) => [row.id, destinationTable.rows[range.startRow + index]!.id]))
    const incoming = incomingGrid!.anchors.map(owner => {
      const row = range.startRow + owner.row, column = range.startColumn + owner.column
      const previous = destinationGrid.cellAt(row, column)!, previousOwnerId = owner.cell.id
      owner.cell.style = pasteCellStyle(incomingGrid!, owner)
      if (mode === 'destination') {
        const right = destinationGrid.cellAt(row, column + owner.colSpan - 1)!, bottom = destinationGrid.cellAt(row + owner.rowSpan - 1, column)!
        owner.cell.style = { ...owner.cell.style, fill: documentTableCellFill(destinationTable, previous.cell),
          margin: documentTableCellMargin(destinationTable, previous.cell), verticalAlign: previous.cell.style.verticalAlign ?? 'top',
          borders: { top: documentTableCellBorder(destinationGrid, previous, 'top'), left: documentTableCellBorder(destinationGrid, previous, 'left'),
            right: documentTableCellBorder(destinationGrid, right, 'right'), bottom: documentTableCellBorder(destinationGrid, bottom, 'bottom') } }
        const destinationStyle = pasteDestinationStyle(previous.cell)
        for (const content of documentTableCellContentContainers(owner.cell)) applyDocumentPasteDestinationStyle(content, destinationStyle)
      }
      // Preserve an existing anchor identity; split children get their new paste IDs.
      if (previous.row === row && previous.column === column) owner.cell.id = previous.cell.id
      owner.cell.columnIndex = column
      delete owner.cell.style.selectedDropdownOptionId
      rebindDocumentTableContinuations(destinationTable, owner.cell, row, previousOwnerId, rowIds)
      return { cell: owner.cell, row, column, rowSpan: owner.rowSpan, colSpan: owner.colSpan }
    })
    rebuild(destinationTable, [...placements(destinationGrid).filter(owner => !replacedIds.has(owner.cell.id)), ...incoming])
  }
  const afterGrid = documentTableGrid(destinationTable)
  if (incomingGrid) {
    const borders = documentTableCellsInRange(afterGrid, range).flatMap(owner => DOCUMENT_TABLE_BORDER_SIDES.map(side => ({
      owner, side, border: documentTableCellBorder(afterGrid, owner, side),
    })))
    for (const { owner, side, border } of borders) setSharedBorder(afterGrid, owner, side, border)
  }
  if (!textTargets(model).length) afterGrid.cellAt(range.startRow, range.startColumn)!.cell.blocks.push(createParagraph())
  const next = createDocumentTableViewSelection(model, view.tableId, range, singleCell ? 'cell' : 'range', undefined, selection)
  reconcileDocumentTableAnchors(before, model, collapsedSelection(next.anchor))
  for (const ancestor of documentBlockAncestry(model, view.tableId) ?? []) delete ancestor.metadata.html
  const previousCells = new Map(grid.anchors.map(owner => [owner.cell.id, owner.cell]))
  const changedCellTargets = afterGrid.anchors.flatMap(owner => JSON.stringify(previousCells.get(owner.cell.id)?.style) !== JSON.stringify(owner.cell.style)
    ? allBlocks(documentTableCellBlocks(owner.cell)).flatMap(child => [...permissionTargetsForBlock(before, child.id), ...permissionTargetsForBlock(model, child.id)]) : [])
  assertDocumentPermission(before.permissionRules ?? [], context, 'edit', [
    ...permissionTargetsForChangedBlocks(before, model), ...changedCellTargets,
    ...[view.tableId, ...[...documentTableCellsInRange(afterGrid, range)].flatMap(owner => allBlocks(documentTableCellBlocks(owner.cell)).map(child => child.id))]
      .flatMap(id => permissionTargetsForBlock(model, id)),
  ])
  documentTableGrid(destinationTable)
  validateDocumentContentControlRanges(model)
  documentTextFieldRanges(model)
  return { model, selection: next, changed: JSON.stringify(before) !== JSON.stringify(model) }
}
