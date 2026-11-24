import type { DocumentBlock, DocumentContentControl, DocumentModel } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { createParagraph, findDocumentBlock, runText, tableForBlock, type DocumentSelection } from './document-model'
import { documentContentControlContains, documentContentControlContainsSelection, resolveDocumentContentControlTableRows } from './content-control-table-rows'
import { resolveDocumentContentControlBlocks } from './content-control-block-range'
import { documentTableContinuationEntries } from './document-table-continuations'
import { documentTextFields, type DocumentTextField } from './document-text-fields'
import { documentContentControlBlocksValueMutationRange } from './content-control-block-mutations'
import { splitDocumentRunRange } from './document-run-range'

export interface DocumentContentControlGroupEditContext { selection?: DocumentSelection; activeControlId?: string }

function maskedFields(fields: readonly DocumentTextField[], start: number, end: number): DocumentTextField[] {
  const delta = end - start
  return fields.flatMap(field => {
    if (!field.endBlockId && field.startOffset >= start && field.endOffset <= end) return []
    const copy = cloneModel(field)
    delete copy.cachedResult
    const position = (offset: number) => offset <= start ? offset : offset >= end ? offset - delta : start
    copy.startOffset = position(copy.startOffset)
    if (!copy.endBlockId) copy.endOffset = position(copy.endOffset)
    return [copy]
  })
}

function groupRuns(runs: DocumentBlock['runs'], start: number, end: number): DocumentBlock['runs'] {
  let offset = 0
  return runs.flatMap(run => {
    const from = offset, to = from + run.text.length
    offset = to
    if (!run.text.length) return from > start && from < end ? [cloneModel(run)] : []
    const first = Math.max(from, start), last = Math.min(to, end)
    return first < last ? [{ ...cloneModel(run), text: run.text.slice(first - from, last - from) }] : []
  })
}

/** Ignore incidental run splitting while retaining text, format, entities and retained provenance. */
function canonicalBlocks(blocks: DocumentBlock[]): unknown[] {
  return blocks.map(block => {
    const runs: Array<Record<string, unknown>> = []
    for (const run of block.runs) {
      const { id: _runId, text, ...properties } = run
      const previous = runs.at(-1)
      if (previous && text && previous.text && JSON.stringify(properties) === JSON.stringify(Object.fromEntries(Object.entries(previous).filter(([key]) => key !== 'text')))) {
        previous.text = String(previous.text) + text
      } else runs.push({ ...properties, text })
    }
    const table = tableForBlock(block)
    const metadata = { ...block.metadata }
    delete metadata.html
    if (metadata.codeBlock && typeof metadata.codeBlock === 'object' && !Array.isArray(metadata.codeBlock)) {
      metadata.codeBlock = { ...metadata.codeBlock, source: runText(block.runs) }
    }
    if (table) metadata.table = { ...table, rows: table.rows.map(row => ({ ...row, cells: row.cells.map(cell => {
      const style = { ...cell.style }, continuations = documentTableContinuationEntries(cell)
      if (style.nativeContinuationContent !== undefined) style.nativeContinuationContent = Object.fromEntries(continuations.map(({ offset, content }) =>
        [offset, { ...content, blocks: canonicalBlocks(content.blocks) }]))
      return { ...cell, blocks: canonicalBlocks(cell.blocks), style }
    }) })) }
    return { ...block, runs, metadata }
  })
}

function maskChild(model: DocumentModel, child: DocumentContentControl): boolean {
  if (child.blockRange) {
    const range = resolveDocumentContentControlBlocks(model, child.blockRange)
    range.container.splice(range.startIndex, range.blockIds.length, { ...createParagraph(), id: '$content-control:' + child.id, runs: [] })
    return true
  }
  if (child.tableRowRange) {
    const rows = resolveDocumentContentControlTableRows(model, child.tableRowRange), table = rows.table
    const headers = table.headerRowCount ?? 0
    table.headerRowCount = headers - Math.max(0, Math.min(headers, rows.endRow + 1) - rows.startRow)
    table.titleRow = table.titleRow === true && rows.startRow !== 0
    table.rows.splice(rows.startRow, rows.rowIds.length, { id: '$content-control:' + child.id, cells: [] })
    if (table.rowStyles) table.rowStyles.splice(rows.startRow, rows.rowIds.length, {})
    return true
  }
  const block = findDocumentBlock(model, child.blockId)
  if (!block || child.startOffset < 0 || child.endOffset < child.startOffset || child.endOffset > runText(block.runs).length) return false
  const fields = maskedFields(documentTextFields(block), child.startOffset, child.endOffset)
  if (block.metadata.documentTextFields !== undefined) block.metadata.documentTextFields = fields
  const { before, after } = splitDocumentRunRange(block.runs, child.startOffset, child.endOffset)
  block.runs = [...before, ...after]
  return true
}

/** Project physical selected rows without cutting an outer merged cell into an invalid shell table. */
function rowGroupOutsideSignature(model: DocumentModel, group: DocumentContentControl, child: DocumentContentControl): string | undefined {
  const source = resolveDocumentContentControlTableRows(model, group.tableRowRange!), block = cloneModel(source.block)
  const masked: DocumentModel = { ...model, blocks: [block] }
  const sameTableRows = child.tableRowRange?.tableId === source.block.id
    ? new Set(resolveDocumentContentControlTableRows(masked, child.tableRowRange).rowIds) : undefined
  if (!sameTableRows && !maskChild(masked, child)) return undefined
  const rows = resolveDocumentContentControlTableRows(masked, group.tableRowRange!)
  let sentinel = false
  const entries = rows.table.rows.slice(rows.startRow, rows.endRow + 1).flatMap((row, index) => {
    if (sameTableRows?.has(row.id)) {
      if (sentinel) return []
      sentinel = true
      return [{ id: '$content-control:' + child.id }]
    }
    const rowIndex = rows.startRow + index
    const content = rows.grid.anchors.flatMap(owner => {
      const { nativeContinuationContent: _continuationContent, nativeContinuations, ...style } = owner.cell.style
      return [
        ...(owner.row === rowIndex ? [{ cellId: owner.cell.id, column: owner.column, colSpan: owner.colSpan,
          rowSpan: owner.rowSpan, style, blocks: canonicalBlocks(owner.cell.blocks) }] : []),
        ...documentTableContinuationEntries(owner.cell).filter(entry => entry.content.rowId === row.id).map(entry => ({
          cellId: owner.cell.id, column: owner.column, colSpan: entry.content.colSpan, source: entry.content.source,
          properties: nativeContinuations && typeof nativeContinuations === 'object' ? Reflect.get(nativeContinuations, String(entry.offset)) : undefined,
          blocks: canonicalBlocks(entry.content.blocks),
        })),
      ]
    })
    return [{ id: row.id, height: row.height, heightRule: row.heightRule, allowSplit: row.allowSplit, style: row.style,
      legacy: rows.table.rowStyles?.[rowIndex], header: rowIndex < (rows.table.headerRowCount ?? 0), content }]
  })
  return JSON.stringify({ tableId: source.block.id, rows: entries })
}

function outsideSignature(model: DocumentModel, group: DocumentContentControl, child: DocumentContentControl): string | undefined {
  let blocks: DocumentBlock[]
  let relative = child
  if (group.blockRange) blocks = cloneModel(resolveDocumentContentControlBlocks(model, group.blockRange).blocks)
  else if (group.tableRowRange) return rowGroupOutsideSignature(model, group, child)
  else {
    if (child.blockRange || child.tableRowRange || child.blockId !== group.blockId) return undefined
    const source = findDocumentBlock(model, group.blockId)
    if (!source) return undefined
    const block = cloneModel(source)
    block.runs = groupRuns(block.runs, group.startOffset, group.endOffset)
    const fields = documentTextFields(source).filter(field => !field.endBlockId && field.startOffset >= group.startOffset && field.endOffset <= group.endOffset)
      .map(field => ({ ...field, startOffset: field.startOffset - group.startOffset, endOffset: field.endOffset - group.startOffset }))
    if (source.metadata.documentTextFields !== undefined) block.metadata.documentTextFields = fields
    blocks = [block]
    relative = { ...child, startOffset: child.startOffset - group.startOffset, endOffset: child.endOffset - group.startOffset }
  }
  const masked: DocumentModel = { ...model, blocks }
  if (!maskChild(masked, relative)) return undefined
  return JSON.stringify(canonicalBlocks(blocks))
}

function isWithinActiveChain(model: DocumentModel, child: DocumentContentControl, activeId: string): boolean {
  const seen = new Set<string>()
  let control = model.contentControls?.find(control => control.id === activeId)
  while (control && !seen.has(control.id)) {
    if (control.id === child.id) return true
    seen.add(control.id)
    control = model.contentControls?.find(candidate => candidate.id === control!.parentId)
  }
  return false
}

/** Preflight only: the completed candidate must still preserve the group's entire outside signature. */
export function documentContentControlGroupOwnsInput(model: DocumentModel, group: DocumentContentControl, activeId: string): boolean {
  return group.type === 'group' && (model.contentControls ?? []).some(child => child.parentId === group.id
    && documentContentControlContains(model, group, child) && isWithinActiveChain(model, child, activeId))
}

/** A group lock allows a child's contents only when the complete outside structure is unchanged. */
export function documentContentControlGroupEditAllowed(before: DocumentModel, after: DocumentModel, group: DocumentContentControl,
  context: DocumentContentControlGroupEditContext = {}): boolean {
  if (group.type !== 'group') return false
  const current = after.contentControls?.find(control => control.id === group.id)
  if (!current || current.type !== group.type) return false
  const children = (before.contentControls ?? []).filter(child => child.parentId === group.id
    && documentContentControlContains(before, group, child)
    && (!context.selection || documentContentControlContainsSelection(before, child, context.selection))
    && (!context.activeControlId || isWithinActiveChain(before, child, context.activeControlId)))
  return children.some(child => {
    try {
      let next = after.contentControls?.find(control => control.id === child.id)
      if (!next && child.temporary && child.blockRange) {
        const replacement = documentContentControlBlocksValueMutationRange(before, after, child.id)
        if (replacement) next = { ...child, blockRange: { blockIds: replacement.blockIds },
          blockId: replacement.selection.anchor.cellBlockId ?? replacement.selection.anchor.blockId, startOffset: 0,
          endOffset: runText(findDocumentBlock(after, replacement.selection.anchor.cellBlockId ?? replacement.selection.anchor.blockId)!.runs).length }
      }
      if (!next && child.temporary && child.tableRowRange) {
        const range = resolveDocumentContentControlTableRows(after, child.tableRowRange)
        if (range.rowIds.length) next = cloneModel(child)
      }
      if (!next && child.temporary && !child.blockRange && !child.tableRowRange) {
        const source = findDocumentBlock(before, child.blockId), destination = findDocumentBlock(after, child.blockId)
        if (source && destination) next = { ...child, endOffset: child.endOffset + runText(destination.runs).length - runText(source.runs).length }
      }
      if (!next || next.parentId !== child.parentId || next.type !== child.type || !documentContentControlContains(after, current, next)) return false
      const left = outsideSignature(before, group, child), right = outsideSignature(after, current, next)
      return left !== undefined && left === right
    } catch (error) {
      if (error instanceof TypeError || error instanceof RangeError) return false
      throw error
    }
  })
}
