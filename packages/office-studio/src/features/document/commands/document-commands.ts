import type { DocumentBlock, DocumentModel, DocumentPageField, DocumentParagraphStyle, DocumentPermissionRule, DocumentSectionBreakType } from '../../../domain/model'
import { getDocumentContentControls, transformContentControlsForBlockEdit } from '../services/content-controls'
import { transformDocumentCommentAnchorsForTextEdit } from '../comments/thread-comments'
import { executeDocumentTableCommand, type DocumentTableCommand } from './table-commands'
import { applySectionHeaderFooter, insertSectionPageField, type SectionHeaderFooterPatch } from '../services/section-header-footer'
import { applyDocumentParagraphStylePatch, applyDocumentSectionSetupPatch, type DocumentSectionSetupPatch } from '../services/document-settings'
import { applyDocumentTocPageNumberUpdate, deleteDocumentToc, prepareDocumentTocPageNumberUpdate, updateDocumentToc, type DocumentTocPageNumbers, type DocumentTocUpdateMode } from '../toc/document-toc'
import { assertDocumentPermission, canManageDocumentPermissions, permissionTargetsForChangedBlocks, permissionTargetsForChangedStories, permissionTargetsForSelection, validateDocumentPermissionRules, type DocumentPermissionContext } from '../permissions/document-permissions'
import { cloneModel } from '../../../domain/serialization'
import { documentTocIdAtBlock } from '../toc/document-toc'
import { applyPainterParagraph, documentPainterTarget, type DocumentPainterFormat } from '../services/document-format-painter'
import { applyDocumentClearFormatting, applyDocumentStoryClearFormatting, prepareDocumentClearFormatting } from '../services/document-clear-formatting'
import type { DocumentStoryRange } from '../model/document-story'
import type { DocumentStoryInputCommit } from '../model/document-input-session'
import { selectionFontPatch } from '../model/document-font'
import { documentRunStyleAt, formatDocumentRunRange, splitDocumentRunRange, validateDocumentRunPatch } from '../model/document-run-range'
import { applyDocumentStoryTextFormatting } from '../services/document-story-formatting'
import { prepareDocumentStoryFieldRefresh, type DocumentStoryFieldRefreshRequest } from '../services/document-story-fields'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'
import { blockListLevel, blockListPreset, isListPreset, listKind, presetNumberingDefinition, type ListPreset } from '../model/document-lists'
import { blockHasNumbering, blockNumberingDefinition, blockNumberingId, blockNumberingLevelIndex, clearBlockNumbering, cloneNumberingDefinition, consecutiveNumberingGroup, copyBlockNumbering, numberingDefinitionFromParts, numberingRecord } from '../model/document-numbering'
import { officeNumberingPrefixLength } from '../model/document-office-numbering-marker'
import { documentTextFieldRanges, documentTextFields, insertDocumentTextFields, remapDocumentTextFieldIds, transformDocumentTextFieldsForBlockEdit, transformDocumentTextFieldsForBlockRangeEdit } from '../model/document-text-fields'
export { blockAtPoint } from '../model/document-model'
import {
  applyNamedParagraphStyle, blockAtPoint, clampPoint, collapsedSelection, compactRuns, createParagraph, createRun, createTable, documentId, documentRunFormatPatch,
  marksAtPoint, orderedSelection, runText, runsForTarget, sameTarget, selectionCollapsed,
  setTargetRuns, sliceRuns, tableForBlock, textTargets,
  type DocumentImage, type DocumentNamedParagraphStyle, type DocumentPoint, type DocumentRun, type DocumentSelection, type RunPatch, type TextTarget,
} from '../model/document-model'

export type DocumentCommand =
  | DocumentTableCommand
  | { id: 'doc.command.insert-text'; text: string; marks?: RunPatch }
  | { id: 'doc.command.delete'; direction: -1 | 1 }
  | { id: 'doc.command.format-text'; patch: RunPatch; clearFields?: Array<'color' | 'highlight'>; story?: DocumentStoryRange }
  | { id: 'doc.command.apply-format-painter'; format: DocumentPainterFormat }
  | { id: 'doc.command.clear-formatting'; story?: DocumentStoryRange }
  | { id: 'doc.command.format-paragraph'; patch: Partial<DocumentParagraphStyle> }
  | { id: 'doc.command.change-paragraph-indent'; direction: -1 | 1 }
  | { id: 'doc.command.set-paragraph-named-style'; namedStyle: DocumentNamedParagraphStyle }
  | { id: 'doc.command.set-block-type'; type: DocumentBlock['type']; level?: number }
  | { id: 'doc.command.change-list-type'; listType: ListPreset }
  | { id: 'doc.command.change-list-nesting-level'; direction: -1 | 1 }
  | { id: 'doc.command.restart-list-numbering'; start?: number }
  | { id: 'doc.command.continue-list-numbering' }
  | { id: 'doc.command.toggle-checklist-item'; blockId: string }
  | { id: 'doc.command.insert-page-break' }
  | { id: 'doc.command.insert-column-break' }
  | { id: 'doc.command.insert-section-break'; type: DocumentSectionBreakType }
  | { id: 'doc.command.insert-table'; rows: number; cols: number }
  | { id: 'doc.command.insert-image'; image: DocumentImage }
  | { id: 'doc.command.insert-divider' }
  | { id: 'doc.command.paste-blocks'; blocks: DocumentBlock[] }
  | { id: 'doc.command.table-tab'; direction: -1 | 1 }
  | { id: 'doc.command.set-section-header-footer'; sectionId: string; patch: SectionHeaderFooterPatch; input?: DocumentStoryInputCommit }
  | { id: 'doc.command.set-section-setup'; sectionId: string; patch: DocumentSectionSetupPatch }
  | ({ id: 'doc.command.refresh-header-footer-fields' } & Omit<DocumentStoryFieldRefreshRequest, 'stories' | 'bodySelection'>)
  | ({ id: 'doc.command.update-selected-fields' } & DocumentStoryFieldRefreshRequest)
  | ({ id: 'doc.command.insert-page-field'; sectionId: string } & DocumentPageField)
  | { id: 'doc.command.set-document-permission-rules'; rules: DocumentPermissionRule[] }
  | { id: 'doc.command.update-table-of-contents'; tocId: string; mode: DocumentTocUpdateMode; pageNumbers?: DocumentTocPageNumbers; timestamp?: number }
  | { id: 'doc.command.delete-table-of-contents'; tocId: string }

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

export function documentTabCommand(model: DocumentModel, selection: DocumentSelection, shift: boolean): DocumentCommand | undefined {
  const targets = selectedTargets(model, selection), direction = shift ? -1 : 1
  const blocks = targets.map(target => blockAtPoint(model, target))
  const isList = (b?: DocumentBlock) => !!b && blockHasNumbering(b)
  if (blocks.length > 1 && blocks.some(isList)) return { id: 'doc.command.change-list-nesting-level', direction }
  const block = blocks[0], target = targets[0]
  if (block && target && isList(block) && orderedSelection(model, selection).start.offset === 0) {
    const container = blockContainer(model, target), index = container.indexOf(block)
    const id = blockNumberingId(block)
    const hasPrevious = id ? container.slice(0, index).some(b => blockNumberingId(b) === id) : container[index - 1]?.type === block.type
    if (hasPrevious) return { id: 'doc.command.change-list-nesting-level', direction }
  }
  if (selection.focus.row != null) return { id: 'doc.command.table-tab', direction }
}

function numberingStyleKey(block: DocumentBlock): string {
  return blockListPreset(block) ?? JSON.stringify(blockNumberingDefinition(block)?.abstract ?? { type: block.type })
}

function targetForBlock(model: DocumentModel, block: DocumentBlock): TextTarget | undefined {
  return textTargets(model).find(target => blockAtPoint(model, target) === block)
}

function replaceImportedNumberingMarkers(model: DocumentModel, targets: readonly TextTarget[], selection: DocumentSelection, context: DocumentPermissionContext): DocumentSelection {
  const plans = targets.flatMap(target => {
    const block = blockAtPoint(model, target)
    return block?.metadata.officeNumberingMarker ? [{ target, block, length: officeNumberingPrefixLength(block) ?? 0 }] : []
  })
  if (!plans.length) return selection
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', plans.flatMap(({ target, length }) => permissionTargetsForSelection(model, {
    anchor: { ...target, offset: 0 }, focus: { ...target, offset: length },
  })))
  const controlPlan: DocumentModel = { ...model, contentControls: model.contentControls && getDocumentContentControls(model) }
  for (const { target, length } of plans) {
    if (length) transformContentControlsForBlockEdit(controlPlan, target.cellBlockId ?? target.blockId, 0, length, 0)
  }
  model.contentControls = controlPlan.contentControls
  for (const { target, block, length } of plans) {
    if (length) {
      let remaining = length
      const bodyRuns: DocumentRun[] = []
      for (const run of runsForTarget(model, target)) {
        if (!run.text.length || remaining === 0) bodyRuns.push({ ...run })
        else if (remaining >= run.text.length) remaining -= run.text.length
        else { bodyRuns.push(...sliceRuns([run], remaining)); remaining = 0 }
      }
      // Cached labels may precede zero-length footnote/field objects. Keep
      // those identities when the label is explicitly replaced or cleared.
      transformDocumentTextFieldsForBlockEdit(block, 0, length, '', undefined, model)
      block.runs = bodyRuns.length ? bodyRuns : [createRun()]
      delete block.metadata.html
      if (target.row !== undefined) {
        const container = model.blocks.find(candidate => candidate.id === target.blockId)
        if (container) delete container.metadata.html
      }
      transformDocumentCommentAnchorsForTextEdit(model, target.cellBlockId ?? target.blockId, 0, length, '')
    }
    delete block.metadata.officeNumberingMarker
  }
  const point = (current: DocumentPoint): DocumentPoint => {
    const plan = plans.find(plan => sameTarget(plan.target, current))
    return plan?.length ? { ...current, offset: Math.max(0, current.offset - plan.length) } : current
  }
  return { anchor: point(selection.anchor), focus: point(selection.focus) }
}

function legacyNumberingTargets(model: DocumentModel, block: DocumentBlock): TextTarget[] {
  const target = targetForBlock(model, block)
  if (!target) return []
  const container = blockContainer(model, target), index = container.indexOf(block), key = numberingStyleKey(block)
  const compatible = (candidate: DocumentBlock | undefined) => candidate?.type === block.type && !blockNumberingId(candidate) && numberingStyleKey(candidate) === key
  let first = index, last = index
  while (compatible(container[first - 1])) first--
  while (compatible(container[last + 1])) last++
  const ids = new Set(container.slice(first, last + 1).map(candidate => candidate.id))
  return textTargets(model).filter(candidate => ids.has(candidate.cellBlockId ?? candidate.blockId))
}

function orderedNumberingParagraph(block: DocumentBlock | undefined): block is DocumentBlock {
  if (!block || block.type === 'checklist' || !blockHasNumbering(block)) return false
  const { level } = blockListLevel(block)
  if (level) return level.format !== 'bullet' && !level.pictureBullet && level.pictureBulletId === undefined
  const preset = blockListPreset(block)
  return block.type === 'numbered' || !!preset && listKind(preset) === 'numbered'
}

function numberingCommandPlan(model: DocumentModel, selection: DocumentSelection): { block: DocumentBlock; targets: TextTarget[] } | undefined {
  const first = selectedTargets(model, selection)[0], block = first && blockAtPoint(model, first)
  if (!first || !orderedNumberingParagraph(block)) return
  const targets = textTargets(model), index = targets.findIndex(target => sameTarget(target, first)), id = blockNumberingId(block)
  const affected = id ? targets.slice(index).filter(target => {
    const candidate = blockAtPoint(model, target)
    return candidate && candidate.type !== 'checklist' && blockHasNumbering(candidate) && blockNumberingId(candidate) === id
  }) : legacyNumberingTargets(model, block).filter(target => targets.findIndex(candidate => sameTarget(candidate, target)) >= index)
  return { block, targets: affected }
}

function precedingCompatibleNumbering(model: DocumentModel, block: DocumentBlock): DocumentBlock | undefined {
  const targets = textTargets(model), current = targets.findIndex(target => blockAtPoint(model, target) === block), key = numberingStyleKey(block)
  for (let index = current - 1; index >= 0; index--) {
    const candidate = blockAtPoint(model, targets[index]!)
    if (orderedNumberingParagraph(candidate) && numberingStyleKey(candidate) === key) return candidate
  }
}

function sameNumberingSequence(model: DocumentModel, previous: DocumentBlock, block: DocumentBlock): boolean {
  const id = blockNumberingId(block), previousId = blockNumberingId(previous)
  const definition = blockNumberingDefinition(block), previousDefinition = blockNumberingDefinition(previous)
  if (id && previousId && definition?.abstract.consecutiveNumbering === true && previousDefinition?.abstract.consecutiveNumbering === true) {
    return numberingStyleKey(previous) === numberingStyleKey(block)
      && consecutiveNumberingGroup(definition, id).id === consecutiveNumberingGroup(previousDefinition, previousId).id
  }
  if (id || previousId) return !!id && id === previousId
  return legacyNumberingTargets(model, previous).some(target => blockAtPoint(model, target) === block)
}

function numberingPermissionsAllowed(model: DocumentModel, targets: TextTarget[], permissionContext?: DocumentPermissionContext): boolean {
  try {
    assertDocumentPermission(model.permissionRules ?? [], permissionContext ?? { documentId: model.id ?? '', userId: 'local-user' }, 'edit', targets.flatMap(target => permissionTargetsForSelection(model, collapsedSelection({ ...target, offset: 0 }))))
    return true
  } catch { return false }
}

export function canRestartDocumentNumbering(model: DocumentModel, selection: DocumentSelection, permissionContext?: DocumentPermissionContext): boolean {
  const plan = numberingCommandPlan(model, selection)
  if (!plan) return false
  const preset = blockListPreset(plan.block), source = blockNumberingDefinition(plan.block) ?? (preset ? presetNumberingDefinition(preset) : undefined)
  return !!source?.levels.some(level => level.ilvl === blockNumberingLevelIndex(plan.block)) && numberingPermissionsAllowed(model, plan.targets, permissionContext)
}

export function canContinueDocumentNumbering(model: DocumentModel, selection: DocumentSelection, permissionContext?: DocumentPermissionContext): boolean {
  const plan = numberingCommandPlan(model, selection), previous = plan && precedingCompatibleNumbering(model, plan.block)
  if (!plan || !previous || sameNumberingSequence(model, previous, plan.block)) return false
  const targets = blockNumberingId(previous) ? plan.targets : [...legacyNumberingTargets(model, previous), ...plan.targets]
  return numberingPermissionsAllowed(model, targets, permissionContext)
}

export const DOCUMENT_PARAGRAPH_INDENT_STEP = 20
export const DOCUMENT_PARAGRAPH_INDENT_MAXIMUM = 100

type ParagraphIndentChange = { target: TextTarget; block: DocumentBlock; level?: number; firstLine?: number }
function paragraphIndentPlan(model: DocumentModel, selection: DocumentSelection, direction: -1 | 1): ParagraphIndentChange[] {
  if (direction !== -1 && direction !== 1) throw new RangeError('Paragraph indentation direction must be -1 or 1')
  return selectedTargets(model, selection).flatMap<ParagraphIndentChange>(target => {
    const block = blockAtPoint(model, target)
    if (!block || !['paragraph', 'heading', 'quote', 'bullet', 'numbered', 'checklist'].includes(block.type)) return []
    if (blockHasNumbering(block)) {
      const { index, maximum } = blockListLevel(block)
      const level = Math.max(0, Math.min(maximum, index + direction))
      return level === index ? [] : [{ target, block, level }]
    }
    const current = block.style.indentFirstLine ?? 0
    if (!Number.isFinite(current)) throw new RangeError('Paragraph first-line indentation must be finite')
    if (direction === -1 && current <= 0) return []
    if (direction === 1 && current >= DOCUMENT_PARAGRAPH_INDENT_MAXIMUM) return []
    const firstLine = Math.max(0, Math.min(DOCUMENT_PARAGRAPH_INDENT_MAXIMUM, current + direction * DOCUMENT_PARAGRAPH_INDENT_STEP))
    return firstLine === current ? [] : [{ target, block, firstLine }]
  })
}

export function canChangeDocumentParagraphIndent(model: DocumentModel, selection: DocumentSelection, direction: -1 | 1, permissionContext?: DocumentPermissionContext): boolean {
  try {
    const changes = paragraphIndentPlan(model, selection, direction)
    return changes.length > 0 && numberingPermissionsAllowed(model, changes.map(change => change.target), permissionContext)
  } catch { return false }
}

function applyListNestingLevel(block: DocumentBlock, level: number): void {
  if (blockListLevel(block).index === level) return
  const previous = blockListLevel(block).level?.indent ?? {}
  const nextBlock = ['bullet', 'numbered', 'checklist'].includes(block.type) ? { ...block, level }
    : { ...block, metadata: { ...block.metadata, numberingLevel: level } }
  const next = blockListLevel(nextBlock).level?.indent ?? {}
  const style = { ...block.style }
  for (const [property, field] of [['indentStart', 'leftTwips'], ['indentEnd', 'rightTwips'], ['indentFirstLine', 'firstLineTwips'], ['hanging', 'hangingTwips']] as const) {
    const oldValue = typeof previous[field] === 'number' && Number.isFinite(previous[field]) ? previous[field] / 15 : 0
    const newValue = typeof next[field] === 'number' && Number.isFinite(next[field]) ? next[field] / 15 : 0
    const current = style[property] ?? (property === 'indentStart' ? oldValue + style.indent * 24 : oldValue)
    if (style[property] !== undefined || oldValue !== newValue) style[property] = newValue + current - oldValue
  }
  if (['bullet', 'numbered', 'checklist'].includes(block.type)) block.level = level
  else block.metadata.numberingLevel = level
  block.style = style
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
  } else if (command.id !== 'doc.command.change-paragraph-indent' && command.id !== 'doc.command.refresh-header-footer-fields' && command.id !== 'doc.command.update-selected-fields'
    && !((command.id === 'doc.command.clear-formatting' || command.id === 'doc.command.format-text') && command.story)) {
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
    case 'doc.command.format-text': {
      const patch = documentRunFormatPatch(command.patch, command.clearFields)
      if (command.story) applyDocumentStoryTextFormatting(model, command.story, patch, context)
      else formatText(model, selection, selectionFontPatch(model, selection, patch))
      return selection
    }
    case 'doc.command.clear-formatting': {
      if (command.story) { applyDocumentStoryClearFormatting(model, command.story, context); return selection }
      const plan = prepareDocumentClearFormatting(model, selection, context)
      if (!plan) return selection
      const range = replaceImportedNumberingMarkers(model, plan.targets, plan.range, context)
      applyDocumentClearFormatting(model, range, plan.targets, plan.defaults)
      return range
    }
    case 'doc.command.apply-format-painter': {
      let range = documentPainterTarget(model, selection)
      if (!range) return selection
      assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, range))
      if (command.format.paragraph) range = replaceImportedNumberingMarkers(model, selectedTargets(model, range), range, context)
      formatText(model, range, command.format.text)
      const paragraph = command.format.paragraph
      if (paragraph) for (const target of selectedTargets(model, range)) {
        const block = blockAtPoint(model, target)
        if (block) applyPainterParagraph(block, paragraph)
      }
      return range
    }
    case 'doc.command.format-paragraph': {
      const updates = selectedTargets(model, selection).flatMap((target) => {
        const block = blockAtPoint(model, target)
        if (!block) return []
        const prepared = { ...block, style: { ...block.style } }
        applyDocumentParagraphStylePatch(prepared, command.patch)
        return [{ block, style: prepared.style }]
      })
      for (const update of updates) update.block.style = update.style
      return selection
    }
    case 'doc.command.change-paragraph-indent': {
      const changes = paragraphIndentPlan(model, selection, command.direction)
      assertDocumentPermission(model.permissionRules ?? [], context, 'edit', changes.flatMap(change => permissionTargetsForSelection(model, collapsedSelection({ ...change.target, offset: 0 }))))
      selection = replaceImportedNumberingMarkers(model, changes.filter(change => change.level !== undefined).map(change => change.target), selection, context)
      for (const change of changes) {
        if (change.level !== undefined) applyListNestingLevel(change.block, change.level)
        else if (change.firstLine !== undefined) applyDocumentParagraphStylePatch(change.block, { indentFirstLine: change.firstLine })
      }
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
      const target = targetForBlock(model, block)
      if (target) {
        assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, collapsedSelection({ ...target, offset: 0 })))
        selection = replaceImportedNumberingMarkers(model, [target], selection, context)
      }
      block.checklist = { listId: block.checklist?.listId || documentId('checklist'), checked: block.checklist?.checked !== true }
      return selection
    }
    case 'doc.command.change-list-nesting-level': {
      if (command.direction !== -1 && command.direction !== 1) throw new RangeError('List nesting direction must be -1 or 1')
      const changes = selectedTargets(model, selection).flatMap(target => {
        const block = blockAtPoint(model, target)
        if (!block || !blockHasNumbering(block)) return []
        const { index, maximum } = blockListLevel(block)
        const level = Math.max(0, Math.min(maximum, index + command.direction))
        return level === index ? [] : [{ target, block, level }]
      })
      selection = replaceImportedNumberingMarkers(model, changes.map(change => change.target), selection, context)
      for (const change of changes) applyListNestingLevel(change.block, change.level)
      return selection
    }
    case 'doc.command.restart-list-numbering':
    case 'doc.command.continue-list-numbering': {
      const plan = numberingCommandPlan(model, selection)
      if (!plan) throw new RangeError('Numbering changes require a numbered paragraph')
      if (command.id === 'doc.command.continue-list-numbering') {
        const previous = precedingCompatibleNumbering(model, plan.block)
        if (!previous) throw new RangeError('No compatible preceding numbering instance is available')
        if (sameNumberingSequence(model, previous, plan.block)) return selection
        const previousTargets = blockNumberingId(previous) ? [] : legacyNumberingTargets(model, previous)
        assertDocumentPermission(model.permissionRules ?? [], context, 'edit', [...previousTargets, ...plan.targets].flatMap(target => permissionTargetsForSelection(model, collapsedSelection({ ...target, offset: 0 }))))
        selection = replaceImportedNumberingMarkers(model, [...previousTargets, ...plan.targets], selection, context)
        if (!blockNumberingId(previous)) {
          const preset = blockListPreset(previous)
          const id = documentId('list')
          for (const target of previousTargets) {
            const block = blockAtPoint(model, target)
            if (!block) continue
            if (preset) block.list = { id, preset }
            else block.metadata.numberingListId = id
          }
        }
        for (const target of plan.targets) {
          const block = blockAtPoint(model, target)
          if (block) {
            const level = blockNumberingLevelIndex(block)
            copyBlockNumbering(previous, block)
            if (!['numbered', 'bullet', 'checklist'].includes(block.type)) block.metadata.numberingLevel = level
          }
        }
      } else {
        const start = command.start ?? 1
        if (!Number.isSafeInteger(start) || start < 0 || start > 2_147_483_647) throw new RangeError('The numbering start must be an integer between 0 and 2147483647')
        const preset = blockListPreset(plan.block)
        const source = blockNumberingDefinition(plan.block) ?? (preset ? presetNumberingDefinition(preset) : undefined)
        const level = blockNumberingLevelIndex(plan.block)
        if (!source?.levels.some(candidate => candidate.ilvl === level)) throw new RangeError('The imported numbering definition has no editable level for this paragraph')
        assertDocumentPermission(model.permissionRules ?? [], context, 'edit', plan.targets.flatMap(target => permissionTargetsForSelection(model, collapsedSelection({ ...target, offset: 0 }))))
        selection = replaceImportedNumberingMarkers(model, plan.targets, selection, context)
        const id = documentId('list')
        const instance: Record<string, unknown> = { ...source.instance, levelStartOverrides: { ...numberingRecord(source.instance.levelStartOverrides), [String(level)]: start } }
        if (source.abstract.consecutiveNumbering === true) instance.consecutiveListId = id
        const definition = numberingDefinitionFromParts(source.abstract, instance)
        for (const target of plan.targets) {
          const block = blockAtPoint(model, target)
          if (!block) continue
          const targetLevel = blockNumberingLevelIndex(block)
          clearBlockNumbering(block)
          block.metadata.numberingDefinition = cloneNumberingDefinition(definition)
          block.metadata.numberingListId = id
          if (!['numbered', 'bullet', 'checklist'].includes(block.type)) block.metadata.numberingLevel = targetLevel
          if (preset) block.list = { id, preset }
        }
      }
      return selection
    }
    case 'doc.command.change-list-type':
    case 'doc.command.set-block-type': {
      const targets = selectedTargets(model, selection)
      if (command.id === 'doc.command.change-list-type' && !isListPreset(command.listType)) throw new RangeError('Unknown list preset')
      const explicit = command.id === 'doc.command.change-list-type' ? command.listType : undefined
      const type = command.id === 'doc.command.set-block-type' ? command.type : listKind(command.listType)
      const level = command.id === 'doc.command.set-block-type' ? command.level : undefined
      const isList = type === 'bullet' || type === 'numbered' || type === 'checklist'
      const toggleOff = !explicit && isList && targets.length > 0 && targets.every((target) => blockAtPoint(model, target)?.type === type)
      let list: DocumentBlock['list']
      let adjacent: DocumentBlock | undefined
      if (!toggleOff && (type === 'bullet' || type === 'numbered')) {
        const first = targets[0], last = targets.at(-1)
        if (!explicit && targets.length === 1 && first && last) {
          const before = blockContainer(model, first), after = blockContainer(model, last)
          const left = before.indexOf(blockAtPoint(model, first)!), right = after.indexOf(blockAtPoint(model, last)!)
          adjacent = [before[left - 1], after[right + 1]].find(b => b?.type === type && (b.list || blockNumberingDefinition(b)))
        }
        list = adjacent?.list ?? (adjacent && blockNumberingDefinition(adjacent) ? undefined : { id: documentId('list'), preset: explicit ?? (type === 'bullet' ? 'BULLET_LIST' : 'ORDER_LIST') })
      }
      const selectedChecklist = targets.map((target) => blockAtPoint(model, target)).find((block) => block?.type === 'checklist')
      const checklistListId = type === 'checklist' && !toggleOff
        ? selectedChecklist?.checklist?.listId ?? adjacentChecklistId(model, targets[0]) ?? documentId('checklist')
        : undefined
      selection = replaceImportedNumberingMarkers(model, targets, selection, context)
      targets.forEach((target) => {
        const block = blockAtPoint(model, target)
        if (!block) return
        const wasList = block.type === 'bullet' || block.type === 'numbered' || block.type === 'checklist'
        block.type = toggleOff ? 'paragraph' : type
        delete block.metadata.namedStyleType
        block.level = block.type === 'paragraph' ? 0 : level ?? (isList && wasList ? block.level : 0)
        clearBlockNumbering(block)
        if (adjacent) copyBlockNumbering(adjacent, block)
        else if (list) block.list = { ...list }
        if (block.type === 'bullet' || block.type === 'numbered') block.level = blockListLevel(block).index
        if (block.type === 'checklist') block.checklist = { listId: checklistListId ?? block.checklist?.listId ?? documentId('checklist'), checked: block.checklist?.checked === true }
        else delete block.checklist
        if (type === 'heading') block.runs = block.runs.map((run) => ({ ...run, bold: true, fontSize: level === 1 ? 24 : level === 2 ? 18 : 14 }))
        if (type === 'paragraph') block.runs = block.runs.map((run) => ({ ...run, bold: false, fontSize: 11 }))
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
      const blocks = copyBlocksForPaste(command.blocks)
      const paragraph = blocks.length === 1 && blocks[0]?.type === 'paragraph' ? blocks[0] : undefined
      const text = paragraph && runText(paragraph.runs)
      if (paragraph && text !== undefined && !/[\r\n]/.test(text)) {
        const { start, end } = orderedSelection(model, selection)
        const startRuns = runsForTarget(model, start)
        const before = splitDocumentRunRange(startRuns, start.offset, sameTarget(start, end) ? end.offset : runText(startRuns).length).before
        const after = blockContainer(model, start) === blockContainer(model, end)
          ? splitDocumentRunRange(runsForTarget(model, end), sameTarget(start, end) ? start.offset : 0, end.offset).after : []
        const fields = documentTextFields(paragraph)
        const next = replaceSelectionText(model, selection, text)
        if (text) setEditingRuns(model, next.focus, [...before, ...paragraph.runs as DocumentRun[], ...after], {})
        const destination = blockAtPoint(model, next.focus)
        if (destination) insertDocumentTextFields(destination, fields, start.offset)
        return next
      }
      return insertBlocks(model, selection, blocks)
    }
    case 'doc.command.table-tab': return moveTableCell(model, selection, command.direction)
    case 'doc.command.set-section-header-footer': {
      const candidate = { ...model, sectionHeaderFooters: model.sectionHeaderFooters ? cloneModel(model.sectionHeaderFooters) : undefined }
      applySectionHeaderFooter(candidate, command.sectionId, command.patch)
      assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedStories(model, candidate))
      model.sectionHeaderFooters = candidate.sectionHeaderFooters
      return selection
    }
    case 'doc.command.set-section-setup':
      applyDocumentSectionSetupPatch(model, command.sectionId, command.patch)
      return selection
    case 'doc.command.refresh-header-footer-fields':
    case 'doc.command.update-selected-fields': {
      const plan = prepareDocumentStoryFieldRefresh(model, { locale: command.locale, timestamp: command.timestamp,
        onOpen: command.onOpen,
        ...(command.id === 'doc.command.update-selected-fields' ? { stories: command.stories ?? [],
          bodySelection: command.bodySelection ?? (!command.stories?.length ? selection : undefined) } : {}) }, context, selection)
      if (plan) {
        model.sectionHeaderFooters = plan.sectionHeaderFooters
        if (plan.blocks) { model.blocks = plan.blocks; model.contentControls = plan.contentControls; model.comments = plan.comments }
      }
      return plan?.selection ?? selection
    }
    case 'doc.command.insert-page-field': {
      const candidate = { ...model, sectionHeaderFooters: model.sectionHeaderFooters ? cloneModel(model.sectionHeaderFooters) : undefined }
      insertSectionPageField(candidate, command.sectionId, command.story, command.field, command)
      assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedStories(model, candidate))
      model.sectionHeaderFooters = candidate.sectionHeaderFooters
      return selection
    }
    case 'doc.command.update-table-of-contents':
    case 'doc.command.delete-table-of-contents': {
      const tocBlocks = model.blocks.filter((block) => documentTocIdAtBlock(block) === command.tocId)
      const targets = selectedTargets(model, selection)
      if (!tocBlocks.length || !targets.length || targets.some((target) => documentTocIdAtBlock(blockAtPoint(model, target)) !== command.tocId)) {
        throw new RangeError('Table of contents commands require a selection inside the target table.')
      }
      const tocPermissionTargets = tocBlocks.flatMap((block) => permissionTargetsForSelection(model, collapsedSelection({ blockId: block.id, offset: 0 })))
      assertDocumentPermission(model.permissionRules ?? [], context, 'edit', tocPermissionTargets)
      if (command.id === 'doc.command.update-table-of-contents' && command.mode === 'pageNumbersOnly') {
        const plan = prepareDocumentTocPageNumberUpdate(model, command.tocId, command.pageNumbers, selection, command.timestamp)
        if (!plan?.changed) return selection
        assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForChangedBlocks(model, { ...model, blocks: plan.blocks }))
        applyDocumentTocPageNumberUpdate(model, plan)
        return plan.selection ?? selection
      }
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
  if (marks) validateDocumentRunPatch(marks)
  assertDocumentTextSelectionEditable(model, selection)
  const { start, end } = orderedSelection(model, selection)
  if (!text && sameTarget(start, end) && start.offset === end.offset) return selection
  const startTargetId = start.cellBlockId ?? start.blockId
  const endTargetId = end.cellBlockId ?? end.blockId
  if (startTargetId !== endTargetId) {
    // Replacing a paragraph break relocates the following paragraph, even at its offset zero.
    const targets = targetsForSelection(model, selection, true)
    const affectedIds = new Set(targets.map((target) => target.cellBlockId ?? target.blockId))
    const affectedControl = model.contentControls?.find((control) => affectedIds.has(control.blockId))
    if (affectedControl) throw new RangeError(`A text edit cannot span content control paragraphs: ${affectedControl.id}`)
  }
  const startRuns = runsForTarget(model, start), endRuns = runsForTarget(model, end)
  const split = splitDocumentRunRange(startRuns, start.offset, sameTarget(start, end) ? end.offset : runText(startRuns).length)
  const before = split.before
  const after = sameTarget(start, end) ? split.after : end.offset === 0 ? [...endRuns] : splitDocumentRunRange(endRuns, 0, end.offset).after
  const startContainer = blockContainer(model, start)
  const endContainer = blockContainer(model, end)
  // Replacing a range inherits its first selected character. A caret still
  // inherits its left neighbor; otherwise typing after clear could revive the
  // unselected prefix's appearance at a run boundary.
  const inherited = { ...documentRunStyleAt(runsForTarget(model, start), start.offset, sameTarget(start, end) ? end.offset : start.offset + 1), ...marks }
  const startBlock = blockAtPoint(model, start)
  if (!startBlock) return selection
  const normalizedText = text.replace(/\r\n?/g, '\n')
  if (normalizedText === '\n' && selectionCollapsed(selection) && ['bullet', 'numbered', 'checklist'].includes(startBlock.type) && !runText(runsForTarget(model, start)).length) {
    startBlock.type = 'paragraph'
    startBlock.level = 0
    delete startBlock.checklist
    clearBlockNumbering(startBlock)
    return collapsedSelection(start)
  }
  const paragraphs = normalizedText.split('\n')
  const additions = paragraphs.slice(1).map((part, index) => {
    const paragraph = createParagraph(part, inherited)
    paragraph.style = { ...startBlock.style }
    if (['bullet', 'numbered', 'checklist', 'quote'].includes(startBlock.type)) {
      paragraph.type = startBlock.type
      paragraph.level = startBlock.level
      copyBlockNumbering(startBlock, paragraph)
      if (startBlock.type === 'checklist') paragraph.checklist = { listId: startBlock.checklist?.listId ?? documentId('checklist'), checked: false }
    }
    if (index === paragraphs.length - 2) paragraph.runs = [...paragraph.runs, ...after]
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
      const block = blockAtPoint(model, target)
      if (block) {
        if (index === 0) transformDocumentTextFieldsForBlockEdit(block, start.offset, runText(block.runs).length, normalizedText, undefined, model)
        else if (index === targets.length - 1) transformDocumentTextFieldsForBlockEdit(block, 0, end.offset, '', undefined, model)
        else transformDocumentTextFieldsForBlockEdit(block, 0, runText(block.runs).length, '', undefined, model)
      }
      setEditingRuns(model, target, index === 0 ? [...before, createRun(normalizedText, inherited)] : index === targets.length - 1 ? after : [], inherited)
    })
    return collapsedSelection({ ...start, offset: start.offset + normalizedText.length })
  }

  const startIndex = startContainer.findIndex((block) => block.id === startBlock.id)
  const endBlock = blockAtPoint(model, end)
  const endIndex = startContainer.findIndex((block) => block.id === endBlock?.id)
  transformDocumentTextFieldsForBlockRangeEdit(startContainer.slice(startIndex, endIndex + 1), start.offset, end.offset, normalizedText, [startBlock, ...additions], undefined, model)
  if (endIndex > startIndex) startContainer.splice(startIndex + 1, endIndex - startIndex)
  if (paragraphs.length === 1) {
    setEditingRuns(model, start, [...before, createRun(normalizedText, inherited), ...after], inherited)
    if (startTargetId === endTargetId) {
      transformDocumentCommentAnchorsForTextEdit(model, startTargetId, start.offset, end.offset, normalizedText)
    }
    return collapsedSelection({ ...start, offset: start.offset + normalizedText.length })
  }
  setEditingRuns(model, start, [...before, createRun(paragraphs[0], inherited)], inherited)
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

function setEditingRuns(model: DocumentModel, target: TextTarget, runs: DocumentRun[], fallback: RunPatch): void {
  const block = blockAtPoint(model, target)
  if (!block) return
  block.runs = runs.length ? runs : [createRun('', fallback)]
  delete block.metadata.html
  if (target.row !== undefined) {
    const container = model.blocks.find(candidate => candidate.id === target.blockId)
    if (container) delete container.metadata.html
  }
}

function formatText(model: DocumentModel, selection: DocumentSelection, patch: RunPatch): void {
  validateDocumentRunPatch(patch)
  assertDocumentTextSelectionEditable(model, selection)
  const { start, end } = orderedSelection(model, selection)
  // Typing marks at a caret are session state; an empty range must not rewrite persistent run IDs.
  if (sameTarget(start, end) && start.offset === end.offset) return
  const updates = selectedTargets(model, selection).flatMap((target) => {
    const runs = runsForTarget(model, target)
    const from = sameTarget(target, start) ? start.offset : 0
    const to = sameTarget(target, end) ? end.offset : runText(runs).length
    const changed = formatDocumentRunRange(runs, from, to, patch)
    if (JSON.stringify(runs) === JSON.stringify(changed)) return []
    const block = blockAtPoint(model, target)
    return block ? [{ block, target, runs: changed }] : []
  })
  for (const update of updates) {
    update.block.runs = update.runs
    delete update.block.metadata.html
    if (update.target.row !== undefined) {
      const container = model.blocks.find(block => block.id === update.target.blockId)
      if (container) delete container.metadata.html
    }
  }
}

function copyBlocksForPaste(blocks: readonly DocumentBlock[]): DocumentBlock[] {
  const copies = JSON.parse(JSON.stringify(blocks)) as DocumentBlock[]
  const blockIds = new Map<string, string>(), paragraphs: DocumentBlock[] = []
  const visit = (container: DocumentBlock[]) => {
    for (const block of container) {
      const nextId = documentId(block.type)
      if (blockIds.has(block.id)) throw new TypeError(`复制范围的段落标识重复：${block.id}`)
      blockIds.set(block.id, nextId)
      block.id = nextId
      block.runs = block.runs.map(run => ({ ...run, id: documentId('run') }))
      paragraphs.push(block)
      const table = tableForBlock(block)
      if (!table) continue
      for (const row of table.rows) {
        row.id = documentId('row')
        for (const cell of row.cells) { cell.id = documentId('cell'); visit(cell.blocks) }
      }
    }
  }
  visit(copies)
  for (const block of paragraphs) remapDocumentTextFieldIds(block, blockIds)
  documentTextFieldRanges({ blocks: copies })
  return copies
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
  transformDocumentTextFieldsForBlockRangeEdit([block], point.offset, point.offset, '\n', [block, tail], undefined, model)
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
