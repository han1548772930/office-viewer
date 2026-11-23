import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import {
  clampPoint, collapsedSelection, findDocumentBlock, runText, runsForTarget, tableForBlock, textTargets,
  type DocumentSelection, type DocumentTableViewSelection, type TextTarget,
} from './document-model'
import {
  assertDocumentTableRange, documentTableCellsInRange, documentTableGrid, documentTableWholeRange, expandDocumentTableRange,
  type DocumentTableCellRange,
} from './document-table-grid'
import { documentTableCellBlocks, documentTableCellContentContainers } from './document-table-continuations'

export function cloneDocumentSelection(selection: DocumentSelection): DocumentSelection {
  return { anchor: { ...selection.anchor }, focus: { ...selection.focus },
    ...(selection.tableSelection ? { tableSelection: { ...selection.tableSelection, range: { ...selection.tableSelection.range } } } : {}) }
}

export function sameDocumentTableViewSelection(left?: DocumentTableViewSelection, right?: DocumentTableViewSelection): boolean {
  if (!left || !right) return left === right
  return left.tableId === right.tableId && left.kind === right.kind && left.pageAffinity === right.pageAffinity
    && left.range.startRow === right.range.startRow && left.range.endRow === right.range.endRow
    && left.range.startColumn === right.range.startColumn && left.range.endColumn === right.range.endColumn
}

export function assertDocumentTableViewSelectionShape(view: DocumentTableViewSelection): void {
  if (!view || typeof view !== 'object' || typeof view.tableId !== 'string' || !view.tableId.trim()
    || !['table', 'row', 'column', 'cell', 'range'].includes(view.kind) || !view.range || typeof view.range !== 'object'
    || ![view.range.startRow, view.range.endRow, view.range.startColumn, view.range.endColumn].every(value => Number.isSafeInteger(value) && value >= 0)
    || view.range.endRow < view.range.startRow || view.range.endColumn < view.range.startColumn
    || view.pageAffinity !== undefined && (!Number.isSafeInteger(view.pageAffinity) || view.pageAffinity < 0)) throw new TypeError('表格结构选区无效')
}

export function resolveDocumentTableViewSelection(model: DocumentModel, view: DocumentTableViewSelection): DocumentTableViewSelection {
  assertDocumentTableViewSelectionShape(view)
  const block = findDocumentBlock(model, view.tableId), table = block && tableForBlock(block)
  if (!table) throw new RangeError('表格结构选区已不存在')
  const grid = documentTableGrid(table)
  assertDocumentTableRange(grid, view.range)
  return { ...view, range: expandDocumentTableRange(grid, view.range) }
}

export function documentTableDescendantIds(blocks: readonly DocumentBlock[]): Set<string> {
  const ids = new Set<string>()
  const visit = (items: readonly DocumentBlock[]) => {
    for (const block of items) {
      ids.add(block.id)
      for (const row of tableForBlock(block)?.rows ?? []) for (const cell of row.cells) for (const content of documentTableCellContentContainers(cell)) visit(content)
    }
  }
  visit(blocks)
  return ids
}

/** Selects actual paragraphs in the logical range, including nested cell text. */
export function documentTableSelectedTextTargets(model: DocumentModel, view: DocumentTableViewSelection): TextTarget[] {
  const normalized = resolveDocumentTableViewSelection(model, view), block = findDocumentBlock(model, normalized.tableId)!
  const owners = documentTableCellsInRange(documentTableGrid(tableForBlock(block)!), normalized.range)
  const ids = new Set(owners.flatMap(owner => [...documentTableDescendantIds(documentTableCellBlocks(owner.cell))]))
  return textTargets(model).filter(target => ids.has(target.cellBlockId ?? target.blockId))
}

/** Read/select never manufactures paragraphs in an object-only cell. */
export function createDocumentTableViewSelection(model: DocumentModel, tableId: string, range: DocumentTableCellRange,
  kind: DocumentTableViewSelection['kind'] = 'range', pageAffinity?: number, fallback?: DocumentSelection): DocumentSelection {
  const view = resolveDocumentTableViewSelection(model, { tableId, range, kind, ...(pageAffinity === undefined ? {} : { pageAffinity }) })
  const targets = documentTableSelectedTextTargets(model, view), first = targets[0], last = targets.at(-1)
  let selection: DocumentSelection
  if (first && last) selection = { anchor: { ...first, offset: 0, ...(pageAffinity === undefined ? {} : { pageAffinity }) },
    focus: { ...last, offset: runText(runsForTarget(model, last)).length, ...(pageAffinity === undefined ? {} : { pageAffinity }) } }
  else if (fallback) selection = { anchor: clampPoint(model, fallback.anchor), focus: clampPoint(model, fallback.focus) }
  else {
    const target = textTargets(model)[0]
    if (!target) throw new RangeError('文档没有合法的输入定位点')
    selection = collapsedSelection({ ...target, offset: 0 })
  }
  return { ...selection, tableSelection: view }
}

/** Rebase structural view ranges by surviving cell IDs, never by stale storage indices. */
export function rebaseDocumentTableViewSelection(before: DocumentModel, after: DocumentModel,
  view?: DocumentTableViewSelection, clearPageAffinity = false): DocumentTableViewSelection | undefined {
  if (!view) return undefined
  const source = findDocumentBlock(before, view.tableId), destination = findDocumentBlock(after, view.tableId)
  const first = source && tableForBlock(source), last = destination && tableForBlock(destination)
  if (!first || !last) return undefined
  const sourceGrid = documentTableGrid(first), nextGrid = documentTableGrid(last)
  assertDocumentTableRange(sourceGrid, view.range)
  const selectedIds = new Set(documentTableCellsInRange(sourceGrid, expandDocumentTableRange(sourceGrid, view.range)).map(owner => owner.cell.id))
  const survivors = nextGrid.anchors.filter(owner => selectedIds.has(owner.cell.id))
  if (!survivors.length) return undefined
  const range = view.kind === 'table' ? documentTableWholeRange(nextGrid) : survivors.reduce((range, owner) => ({
    startRow: Math.min(range.startRow, owner.row), endRow: Math.max(range.endRow, owner.row + owner.rowSpan - 1),
    startColumn: Math.min(range.startColumn, owner.column), endColumn: Math.max(range.endColumn, owner.column + owner.colSpan - 1),
  }), { startRow: nextGrid.rowCount - 1, endRow: 0, startColumn: nextGrid.columnCount - 1, endColumn: 0 })
  const { pageAffinity: _affinity, ...identity } = view
  return { ...identity, range: expandDocumentTableRange(nextGrid, range),
    ...(!clearPageAffinity && view.pageAffinity !== undefined ? { pageAffinity: view.pageAffinity } : {}) }
}
