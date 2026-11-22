import type { DocumentBlock, DocumentModel, DocumentPageField, DocumentParagraphStyle, DocumentPermissionRule, DocumentSectionBreakType } from '../../../domain/model'
import { getDocumentContentControls, transformContentControlsForBlockEdit } from '../services/content-controls'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'
import { executeDocumentTableCommand, type DocumentTableCommand } from './table-commands'
import { applySectionHeaderFooter, insertSectionPageField, type SectionHeaderFooterPatch } from '../services/section-header-footer'
import { applyDocumentSectionSetupPatch, type DocumentSectionSetupPatch } from '../services/document-settings'
import { deleteDocumentToc, updateDocumentToc, type DocumentTocPageNumbers, type DocumentTocUpdateMode } from '../toc/document-toc'
import { assertDocumentPermission, canManageDocumentPermissions, permissionTargetsForSelection, validateDocumentPermissionRules, type DocumentPermissionContext } from '../permissions/document-permissions'
import { documentTocIdAtBlock } from '../toc/document-toc'
import {
  applyNamedParagraphStyle, applyRunPatch, clampPoint, collapsedSelection, compactRuns, createParagraph, createRun, createTable, documentId,
  marksAtPoint, orderedSelection, runText, runsForTarget, sameTarget, selectionCollapsed,
  setTargetRuns, sliceRuns, tableForBlock, textTargets,
  type DocumentImage, type DocumentNamedParagraphStyle, type DocumentPoint, type DocumentRun, type DocumentSelection, type RunPatch, type TextTarget,
} from '../model/document-model'

export type DocumentCommand =
  | DocumentTableCommand
  | { id: 'doc.command.insert-text'; text: string; marks?: RunPatch }
  | { id: 'doc.command.delete'; direction: -1 | 1 }
  | { id: 'doc.command.format-text'; patch: RunPatch }
  | { id: 'doc.command.format-paragraph'; patch: Partial<DocumentParagraphStyle> }
  | { id: 'doc.command.set-paragraph-named-style'; namedStyle: DocumentNamedParagraphStyle }
  | { id: 'doc.command.set-block-type'; type: DocumentBlock['type']; level?: number }
  | { id: 'doc.command.toggle-checklist-item'; blockId: string }
  | { id: 'doc.command.insert-page-break' }
  | { id: 'doc.command.insert-column-break' }
  | { id: 'doc.command.insert-section-break'; type: DocumentSectionBreakType }
  | { id: 'doc.command.insert-table'; rows: number; cols: number }
  | { id: 'doc.command.insert-image'; image: DocumentImage }
  | { id: 'doc.command.insert-divider' }
  | { id: 'doc.command.paste-blocks'; blocks: DocumentBlock[] }
  | { id: 'doc.command.table-tab'; direction: -1 | 1 }
  | { id: 'doc.command.set-section-header-footer'; sectionId: string; patch: SectionHeaderFooterPatch }
  | { id: 'doc.command.set-section-setup'; sectionId: string; patch: DocumentSectionSetupPatch }
  | ({ id: 'doc.command.insert-page-field'; sectionId: string } & DocumentPageField)
  | { id: 'doc.command.set-document-permission-rules'; rules: DocumentPermissionRule[] }
  | { id: 'doc.command.update-table-of-contents'; tocId: string; mode: DocumentTocUpdateMode; pageNumbers?: DocumentTocPageNumbers }
  | { id: 'doc.command.delete-table-of-contents'; tocId: string }

export function blockAtPoint(model: DocumentModel, point: TextTarget): DocumentBlock | undefined {
  const block = model.blocks.find((candidate) => candidate.id === point.blockId)
  if (point.row == null || !block) return block
  const cell = tableForBlock(block)?.rows[point.row]?.cells[point.col ?? 0]
  return cell?.blocks.find((candidate) => candidate.id === point.cellBlockId) ?? cell?.blocks[0]
}

function blockContainer(model: DocumentModel, point: TextTarget): DocumentBlock[] {
  if (point.row == null) return model.blocks
  const block = model.blocks.find((candidate) => candidate.id === point.blockId)
  return block ? tableForBlock(block)?.rows[point.row]?.cells[point.col ?? 0]?.blocks ?? model.blocks : model.blocks
}

function findBlockById(blocks: readonly DocumentBlock[], id: string): DocumentBlock | undefined {
  for (const block of blocks) {
    if (block.id === id) return block
    const table = tableForBlock(block)
    for (const row of table?.rows ?? []) {
      for (const cell of row.cells) {
        const nested = findBlockById(cell.blocks, id)
        if (nested) return nested
      }
    }
  }
  return undefined
}

function adjacentChecklistId(model: DocumentModel, target: TextTarget | undefined): string | undefined {
  if (!target || target.row != null) return undefined
  const index = model.blocks.findIndex((block) => block.id === target.blockId)
  if (index < 0) return undefined
  for (const block of [model.blocks[index - 1], model.blocks[index + 1]]) {
    if (block?.type === 'checklist' && block.checklist?.listId) return block.checklist.listId
  }
  return undefined
}

function pointForBlock(template: DocumentPoint, block: DocumentBlock, offset: number): DocumentPoint {
  return template.row == null ? { blockId: block.id, offset } : { ...template, cellBlockId: block.id, offset }
}

/** Model mutation only. The application command bus owns the resulting history snapshot. */
export function executeDocumentCommand(
  model: DocumentModel,
  selection: DocumentSelection,
  command: DocumentCommand,
  permissionContext?: DocumentPermissionContext,
): DocumentSelection {
  const context = permissionContext ?? { documentId: model.id ?? '', userId: 'local-user' }
  if (command.id === 'doc.command.set-document-permission-rules') {
    if (!canManageDocumentPermissions(context)) throw new Error('Only the document permission administrator can update permission rules')
  } else {
    assertDocumentPermission(model.permissionRules ?? [], context, 'edit', 'sectionId' in command ? [{ scope: 'section', id: command.sectionId }] : permissionTargetsForSelection(model, selection))
  }
  switch (command.id) {
    case 'doc.command.table-insert-row':
    case 'doc.command.table-insert-column':
    case 'doc.command.table-delete-rows':
    case 'doc.command.table-delete-columns':
    case 'doc.command.table-delete': return executeDocumentTableCommand(model, selection, command)
    case 'doc.command.insert-text': return replaceSelectionText(model, selection, command.text, command.marks)
    case 'doc.command.delete': return deleteContent(model, selection, command.direction)
    case 'doc.command.format-text': formatText(model, selection, command.patch); return selection
    case 'doc.command.format-paragraph': {
      selectedTargets(model, selection).forEach((target) => {
        const block = blockAtPoint(model, target)
        if (block) block.style = { ...block.style, ...command.patch }
      })
      return selection
    }
    case 'doc.command.set-paragraph-named-style': {
      selectedTargets(model, selection).forEach((target) => {
        const block = blockAtPoint(model, target)
        if (block && ['paragraph', 'heading', 'quote'].includes(block.type)) applyNamedParagraphStyle(block, command.namedStyle)
      })
      return selection
    }
    case 'doc.command.toggle-checklist-item': {
      const block = findBlockById(model.blocks, command.blockId)
      if (!block || block.type !== 'checklist') throw new RangeError('The selected block is not a checklist item.')
      block.checklist = { listId: block.checklist?.listId || documentId('checklist'), checked: block.checklist?.checked !== true }
      return selection
    }
    case 'doc.command.set-block-type': {
      const targets = selectedTargets(model, selection)
      const isList = command.type === 'bullet' || command.type === 'numbered' || command.type === 'checklist'
      const toggleOff = isList && targets.length > 0 && targets.every((target) => blockAtPoint(model, target)?.type === command.type)
      const selectedChecklist = targets.map((target) => blockAtPoint(model, target)).find((block) => block?.type === 'checklist')
      const checklistListId = command.type === 'checklist' && !toggleOff
        ? selectedChecklist?.checklist?.listId ?? adjacentChecklistId(model, targets[0]) ?? documentId('checklist')
        : undefined
      targets.forEach((target) => {
        const block = blockAtPoint(model, target)
        if (!block) return
        const wasList = block.type === 'bullet' || block.type === 'numbered' || block.type === 'checklist'
        block.type = toggleOff ? 'paragraph' : command.type
        block.level = block.type === 'paragraph' ? 0 : command.level ?? (isList && wasList ? block.level : 0)
        if (block.type === 'checklist') block.checklist = { listId: checklistListId ?? block.checklist?.listId ?? documentId('checklist'), checked: block.checklist?.checked === true }
        else delete block.checklist
        if (command.type === 'heading') block.runs = block.runs.map((run) => ({ ...run, bold: true, fontSize: command.level === 1 ? 24 : command.level === 2 ? 18 : 14 }))
        if (command.type === 'paragraph') block.runs = block.runs.map((run) => ({ ...run, bold: false, fontSize: 11 }))
      })
      return selection
    }
    case 'doc.command.insert-page-break': {
      const split = replaceSelectionText(model, selection, '\n')
      const container = blockContainer(model, split.focus)
      const index = container.findIndex((block) => block.id === (split.focus.cellBlockId ?? split.focus.blockId))
      if (split.focus.row == null) container.splice(index, 0, { ...createParagraph(), type: 'pageBreak', runs: [] })
      return split
    }
    case 'doc.command.insert-column-break': return insertBreakBlock(model, selection, 'columnBreak')
    case 'doc.command.insert-section-break': return insertBreakBlock(model, selection, 'sectionBreak', command.type)
    case 'doc.command.insert-table': {
      const block = { ...createParagraph(), type: 'table' as const, metadata: { table: createTable(command.rows, command.cols) } }
      insertBlocks(model, selection, [block])
      const paragraph = tableForBlock(block)?.rows[0]?.cells[0]?.blocks[0]
      return paragraph ? collapsedSelection({ blockId: block.id, row: 0, col: 0, cellBlockId: paragraph.id, offset: 0 }) : selection
    }
    case 'doc.command.insert-image': return insertBlocks(model, selection, [{ ...createParagraph(), type: 'image', metadata: { image: command.image } }])
    case 'doc.command.insert-divider': return insertBlocks(model, selection, [{ ...createParagraph(), type: 'divider', runs: [] }])
    case 'doc.command.paste-blocks': {
      const paragraph = command.blocks.length === 1 && command.blocks[0]?.type === 'paragraph' ? command.blocks[0] : undefined
      const text = paragraph && runText(paragraph.runs)
      if (paragraph && text !== undefined && !/[\r\n]/.test(text)) {
        const next = replaceSelectionText(model, selection, text)
        const offset = next.focus.offset
        const runs = runsForTarget(model, next.focus)
        setTargetRuns(model, next.focus, [...sliceRuns(runs, 0, offset - text.length), ...paragraph.runs as DocumentRun[], ...sliceRuns(runs, offset)])
        return next
      }
      return insertBlocks(model, selection, command.blocks)
    }
    case 'doc.command.table-tab': return moveTableCell(model, selection, command.direction)
    case 'doc.command.set-section-header-footer':
      applySectionHeaderFooter(model, command.sectionId, command.patch)
      return selection
    case 'doc.command.set-section-setup':
      applyDocumentSectionSetupPatch(model, command.sectionId, command.patch)
      return selection
    case 'doc.command.insert-page-field':
      insertSectionPageField(model, command.sectionId, command.story, command.field, command)
      return selection
    case 'doc.command.update-table-of-contents':
    case 'doc.command.delete-table-of-contents': {
      const tocBlocks = model.blocks.filter((block) => documentTocIdAtBlock(block) === command.tocId)
      const targets = selectedTargets(model, selection)
      if (!tocBlocks.length || !targets.length || targets.some((target) => documentTocIdAtBlock(blockAtPoint(model, target)) !== command.tocId)) {
        throw new RangeError('Table of contents commands require a selection inside the target table.')
      }
      const tocPermissionTargets = tocBlocks.flatMap((block) => permissionTargetsForSelection(model, collapsedSelection({ blockId: block.id, offset: 0 })))
      assertDocumentPermission(model.permissionRules ?? [], context, 'edit', tocPermissionTargets)
      if (command.id === 'doc.command.update-table-of-contents') updateDocumentToc(model, command.tocId, command.mode, command.pageNumbers)
      else deleteDocumentToc(model, command.tocId)
      return selection
    }
    case 'doc.command.set-document-permission-rules':
      validateDocumentPermissionRules(command.rules)
      model.permissionRules = command.rules.map(rule => ({ ...rule, userIds: [...rule.userIds] }))
      return selection
  }
}

export function selectedTargets(model: DocumentModel, selection: DocumentSelection): TextTarget[] {
  return targetsForSelection(model, selection, false)
}

function targetsForSelection(model: DocumentModel, selection: DocumentSelection, includeEndBoundary: boolean): TextTarget[] {
  const { start, end } = orderedSelection(model, selection)
  const targets = textTargets(model)
  const first = targets.findIndex((target) => sameTarget(target, start))
  const last = targets.findIndex((target) => sameTarget(target, end))
  if (first < 0 || last < first) return []
  // Paragraph formatting uses a half-open range. A caret at offset zero still targets its own paragraph.
  const excludesLast = !includeEndBoundary && last > first && end.offset === 0
  return targets.slice(first, last + (excludesLast ? 0 : 1))
}

export function replaceSelectionText(model: DocumentModel, selection: DocumentSelection, text: string, marks?: RunPatch): DocumentSelection {
  const { start, end } = orderedSelection(model, selection)
  const startTargetId = start.cellBlockId ?? start.blockId
  const endTargetId = end.cellBlockId ?? end.blockId
  if (startTargetId !== endTargetId) {
    // Replacing a paragraph break relocates the following paragraph, even at its offset zero.
    const targets = targetsForSelection(model, selection, true)
    const affectedIds = new Set(targets.map((target) => target.cellBlockId ?? target.blockId))
    const affectedControl = model.contentControls?.find((control) => affectedIds.has(control.blockId))
    if (affectedControl) throw new RangeError(`A text edit cannot span content control paragraphs: ${affectedControl.id}`)
  }
  const before = sliceRuns(runsForTarget(model, start), 0, start.offset)
  const after = sliceRuns(runsForTarget(model, end), end.offset)
  const startContainer = blockContainer(model, start)
  const endContainer = blockContainer(model, end)
  const inherited = { ...marksAtPoint(model, start), ...marks }
  const startBlock = blockAtPoint(model, start)
  if (!startBlock) return selection
  const normalizedText = text.replace(/\r\n?/g, '\n')
  if (normalizedText === '\n' && selectionCollapsed(selection) && startBlock.type === 'checklist' && !runText(runsForTarget(model, start)).length) {
    startBlock.type = 'paragraph'
    startBlock.level = 0
    delete startBlock.checklist
    return collapsedSelection(start)
  }
  const paragraphs = normalizedText.split('\n')
  const additions = paragraphs.slice(1).map((part, index) => {
    const paragraph = createParagraph(part, inherited)
    paragraph.style = { ...startBlock.style }
    if (['bullet', 'numbered', 'checklist', 'quote'].includes(startBlock.type)) {
      paragraph.type = startBlock.type
      paragraph.level = startBlock.level
      if (startBlock.type === 'checklist') paragraph.checklist = { listId: startBlock.checklist?.listId ?? documentId('checklist'), checked: false }
    }
    if (index === paragraphs.length - 2) paragraph.runs = compactRuns([...paragraph.runs as DocumentRun[], ...after])
    return paragraph
  })
  if (startTargetId === endTargetId) {
    transformContentControlsForBlockEdit(model, startTargetId, start.offset, end.offset, normalizedText.length,
      additions.length ? {
        blockId: additions.at(-1)!.id,
        firstBreakOffset: start.offset + normalizedText.indexOf('\n'),
        lastParagraphOffset: start.offset + normalizedText.lastIndexOf('\n') + 1,
      } : undefined)
  }

  if (startContainer !== endContainer) {
    const targets = targetsForSelection(model, selection, true)
    targets.forEach((target, index) => {
      setTargetRuns(model, target, index === 0 ? [...before, createRun(text, inherited)] : index === targets.length - 1 ? after : [])
    })
    return collapsedSelection({ ...start, offset: start.offset + text.length })
  }

  const startIndex = startContainer.findIndex((block) => block.id === startBlock.id)
  const endBlock = blockAtPoint(model, end)
  const endIndex = startContainer.findIndex((block) => block.id === endBlock?.id)
  if (endIndex > startIndex) startContainer.splice(startIndex + 1, endIndex - startIndex)
  if (paragraphs.length === 1) {
    setTargetRuns(model, start, [...before, createRun(text, inherited), ...after])
    if (startTargetId === endTargetId) {
      transformDocumentCommentAnchorsForTextEdit(model, startTargetId, start.offset, end.offset, text)
    }
    return collapsedSelection({ ...start, offset: start.offset + text.length })
  }
  setTargetRuns(model, start, [...before, createRun(paragraphs[0], inherited)])
  startContainer.splice(startIndex + 1, 0, ...additions)
  const last = additions.at(-1)!
  return collapsedSelection(pointForBlock(start, last, paragraphs.at(-1)!.length))
}

export function graphemeOffsets(text: string): number[] {
  const offsets = Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text), (part) => part.index)
  offsets.push(text.length)
  return offsets
}

function deleteContent(model: DocumentModel, selection: DocumentSelection, direction: -1 | 1): DocumentSelection {
  if (!selectionCollapsed(selection)) return replaceSelectionText(model, selection, '')
  const point = clampPoint(model, selection.focus)
  const text = runText(runsForTarget(model, point))
  const offsets = graphemeOffsets(text)
  const next = direction === -1 ? offsets.filter((offset) => offset < point.offset).at(-1) : offsets.find((offset) => offset > point.offset)
  if (next != null) return replaceSelectionText(model, { anchor: point, focus: { ...point, offset: next } }, '')

  const container = blockContainer(model, point)
  const block = blockAtPoint(model, point)
  const index = container.findIndex((candidate) => candidate.id === block?.id)
  const neighbor = container[index + direction]
  if (neighbor && ['image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(neighbor.type)) {
    container.splice(index + direction, 1)
    return collapsedSelection(point)
  }
  if (neighbor && !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(neighbor.type)) {
    const neighborPoint = pointForBlock(point, neighbor, direction === -1 ? runText(neighbor.runs).length : 0)
    return replaceSelectionText(model, { anchor: point, focus: neighborPoint }, '')
  }
  return selection
}

function formatText(model: DocumentModel, selection: DocumentSelection, patch: RunPatch): void {
  const { start, end } = orderedSelection(model, selection)
  // Typing marks at a caret are session state; an empty range must not rewrite persistent run IDs.
  if (sameTarget(start, end) && start.offset === end.offset) return
  selectedTargets(model, selection).forEach((target) => {
    const runs = runsForTarget(model, target)
    const from = sameTarget(target, start) ? start.offset : 0
    const to = sameTarget(target, end) ? end.offset : runText(runs).length
    const selected = sliceRuns(runs, from, to)
    if (!selected.length || selected.every((run) => Object.entries(patch).every(([key, value]) => Object.is(Reflect.get(run, key), value)))) return
    const changed = selected.map((run) => applyRunPatch(run, patch))
    setTargetRuns(model, target, [...sliceRuns(runs, 0, from), ...changed, ...sliceRuns(runs, to)])
  })
}

function insertBlocks(model: DocumentModel, selection: DocumentSelection, blocks: DocumentBlock[]): DocumentSelection {
  if (!blocks.length) return selection
  const { start, end } = orderedSelection(model, selection)
  const tail = createParagraph()
  const controlPlan = { ...model, contentControls: getDocumentContentControls(model) }
  if (sameTarget(start, end)) {
    transformContentControlsForBlockEdit(controlPlan, start.cellBlockId ?? start.blockId, start.offset, end.offset, 0, {
      blockId: tail.id, firstBreakOffset: start.offset, lastParagraphOffset: start.offset,
    })
  }
  const collapsed = replaceSelectionText(model, selection, '')
  const point = collapsed.focus
  const block = blockAtPoint(model, point)
  if (!block) return collapsed
  const container = blockContainer(model, point)
  const index = container.findIndex((candidate) => candidate.id === block.id)
  tail.style = { ...block.style }
  tail.runs = sliceRuns(block.runs as DocumentRun[], point.offset)
  if (sameTarget(start, end) && model.contentControls) model.contentControls = controlPlan.contentControls
  setTargetRuns(model, point, sliceRuns(block.runs as DocumentRun[], 0, point.offset))
  container.splice(index + 1, 0, ...blocks, tail)
  return collapsedSelection(pointForBlock(point, tail, 0))
}

function insertBreakBlock(model: DocumentModel, selection: DocumentSelection, type: 'columnBreak' | 'sectionBreak', sectionType?: DocumentSectionBreakType): DocumentSelection {
  if (selection.anchor.row != null || selection.focus.row != null) throw new RangeError('Breaks can only be inserted in the document body.')
  const split = replaceSelectionText(model, selection, '\n')
  const index = model.blocks.findIndex((block) => block.id === split.focus.blockId)
  if (index < 0) return split
  const marker = createParagraph()
  marker.type = type
  marker.runs = []
  if (type === 'sectionBreak' && sectionType) marker.sectionBreak = { type: sectionType }
  model.blocks.splice(index, 0, marker)
  return split
}

function moveTableCell(model: DocumentModel, selection: DocumentSelection, direction: -1 | 1): DocumentSelection {
  const point = selection.focus
  if (point.row == null) return selection
  const block = model.blocks.find((candidate) => candidate.id === point.blockId)
  const table = block && tableForBlock(block)
  if (!table) return selection
  let row = point.row
  let col = (point.col ?? 0) + direction
  if (col < 0) { row -= 1; col = (table.rows[row]?.cells.length ?? 1) - 1 }
  if (col >= (table.rows[row]?.cells.length ?? 0)) { row += 1; col = 0 }
  if (row < 0) return selection
  if (row >= table.rows.length) {
    if (direction === -1) return selection
    const inserted = executeDocumentTableCommand(model, selection, { id: 'doc.command.table-insert-row', position: 'after' })
    const insertedRow = inserted.focus.row ?? row
    const paragraph = table.rows[insertedRow]?.cells[0]?.blocks[0]
    return paragraph ? collapsedSelection({ blockId: point.blockId, row: insertedRow, col: 0, cellBlockId: paragraph.id, offset: 0 }) : selection
  }
  const paragraph = table.rows[row]?.cells[col]?.blocks[0]
  return paragraph ? collapsedSelection({ blockId: point.blockId, row, col, cellBlockId: paragraph.id, offset: 0 }) : selection
}
