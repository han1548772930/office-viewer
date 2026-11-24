import type { DocumentBlock, DocumentModel, DocumentPermissionAction, DocumentPermissionRule, DocumentPermissionScope } from '../../../domain/model'
import type { DocumentSelection, TextTarget } from '../model/document-model'
import { documentBlockAncestry, findDocumentBlock, orderedSelection, tableForBlock, textTargets } from '../model/document-model'
import { documentBlockHostForId, documentBlockHostId, documentEditingBlockHost, documentStoredBlockHosts, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import { documentTableCellsInRange, documentTableGrid } from '../model/document-table-grid'
import { documentTableDescendantIds, resolveDocumentTableViewSelection } from '../model/document-table-selection'
import { documentTableCellBlocks } from '../model/document-table-continuations'
import { flattenDocumentBody } from '../services/document-facade'
import { documentSectionIds, resolveSectionHeaderFooter, sectionIdForBlock } from '../services/section-header-footer'
import { DOCUMENT_STORY_KEYS } from '../model/document-story'
import { sameSectionStory, sectionStoryLinkedToPrevious } from '../model/document-section-stories'

export interface DocumentPermissionTarget {
  scope: Exclude<DocumentPermissionScope, 'document'>
  id: string
}

export interface DocumentPermissionContext {
  documentId: string
  userId: string
  defaultAllow?: boolean
  /** Local workbench administrator identity, not authenticated server authorization. */
  permissionAdminUserId?: string
}

export function canManageDocumentPermissions(context: DocumentPermissionContext): boolean {
  return context.userId === (context.permissionAdminUserId ?? context.userId)
}

export interface DocumentPermissionDecision {
  allowed: boolean
  action: DocumentPermissionAction
  deniedBy?: DocumentPermissionRule
}

export class DocumentPermissionError extends Error {
  readonly code = 'DOCUMENT_PERMISSION_DENIED'
  readonly decision: DocumentPermissionDecision

  constructor(decision: DocumentPermissionDecision) {
    super(`Document ${decision.action} permission denied`)
    this.name = 'DocumentPermissionError'
    this.decision = decision
  }
}

/** Validates permission rules before storing them in a local document snapshot. */
export function validateDocumentPermissionRules(rules: readonly DocumentPermissionRule[]): void {
  const ids = new Set<string>()
  for (const rule of rules) {
    if (!rule.id.trim() || ids.has(rule.id)) throw new Error('Permission rule IDs must be non-empty and unique')
    ids.add(rule.id)
    if (!rule.userIds.every(userId => typeof userId === 'string' && userId.trim().length > 0)) {
      throw new Error(`Permission rule ${rule.id} contains an empty user ID`)
    }
    if (rule.scope === 'document') {
      if (rule.targetId !== undefined) throw new Error(`Document rule ${rule.id} cannot have a target ID`)
    } else if (!rule.targetId?.trim()) {
      throw new Error(`Permission rule ${rule.id} requires a target ID`)
    }
    if (rule.scope !== 'document' && rule.action !== 'view' && rule.action !== 'edit') {
      throw new Error(`Scoped permission rule ${rule.id} only supports edit or view`)
    }
  }
}

/**
 * Resolves local UI/command permission. A document-level denial always wins.
 * For scoped view/edit actions, every target must allow access and any matching deny
 * overrides matching allows at that target.
 */
export function resolveDocumentPermission(
  rules: readonly DocumentPermissionRule[],
  context: DocumentPermissionContext,
  action: DocumentPermissionAction,
  targets: readonly DocumentPermissionTarget[] = [],
): DocumentPermissionDecision {
  validateDocumentPermissionRules(rules)
  const documentRule = matchingRule(rules, context.userId, action, 'document')
  if (documentRule?.effect === 'deny') return { allowed: false, action, deniedBy: documentRule }
  if (action !== 'view' && action !== 'edit' || targets.length === 0) {
    return { allowed: documentRule?.effect === 'allow' || context.defaultAllow !== false, action }
  }
  const documentAllows = documentRule?.effect === 'allow' || context.defaultAllow !== false

  for (const target of targets) {
    const scopedRules = rules.filter(rule =>
      rule.scope === target.scope && rule.targetId === target.id && rule.action === action && appliesToUser(rule, context.userId),
    )
    const denial = scopedRules.find(rule => rule.effect === 'deny')
    if (denial) return { allowed: false, action, deniedBy: denial }
    if (!scopedRules.some(rule => rule.effect === 'allow') && !documentAllows) {
      return { allowed: false, action }
    }
  }
  return { allowed: true, action }
}

export function assertDocumentPermission(
  rules: readonly DocumentPermissionRule[],
  context: DocumentPermissionContext,
  action: DocumentPermissionAction,
  targets: readonly DocumentPermissionTarget[] = [],
): void {
  const decision = resolveDocumentPermission(rules, context, action, targets)
  if (!decision.allowed) throw new DocumentPermissionError(decision)
}

export function permissionTargetsForSelection(model: DocumentModel, selection: DocumentSelection): DocumentPermissionTarget[] {
  if (selection.tableSelection) {
    const view = resolveDocumentTableViewSelection(model, selection.tableSelection), block = findDocumentBlock(model, view.tableId)!
    const owners = documentTableCellsInRange(documentTableGrid(tableForBlock(block)!), view.range), result = new Map<string, DocumentPermissionTarget>()
    addBlockPermissionTargets(model, view.tableId, result)
    for (const owner of owners) for (const id of documentTableDescendantIds(documentTableCellBlocks(owner.cell))) addBlockPermissionTargets(model, id, result)
    return [...result.values()]
  }
  const { start, end } = orderedSelection(model, selection)
  const targets = textTargets(model)
  const firstIndex = targets.findIndex(target => samePoint(target, start))
  const lastIndex = targets.findIndex(target => samePoint(target, end))
  if (firstIndex < 0 || lastIndex < firstIndex) return []
  const endIndex = lastIndex > firstIndex && end.offset === 0 ? lastIndex - 1 : lastIndex
  const result = new Map<string, DocumentPermissionTarget>()
  for (const target of targets.slice(firstIndex, endIndex + 1)) {
    addBlockPermissionTargets(model, target.cellBlockId ?? target.blockId, result)
  }
  return [...result.values()]
}

export function permissionTargetsForRange(model: DocumentModel, startOffset: number, endOffset: number): DocumentPermissionTarget[] {
  const start = Math.max(0, Math.min(startOffset, endOffset))
  const end = Math.max(startOffset, endOffset)
  const segments = flattenDocumentBody(model).segments
  const result = new Map<string, DocumentPermissionTarget>()
  for (const segment of segments) {
    const intersects = start === end
      ? start >= segment.startOffset && start <= segment.endOffset
      : start < segment.endOffset && end > segment.startOffset
    if (intersects) addBlockPermissionTargets(model, segment.cellBlockId ?? segment.blockId, result)
  }
  return [...result.values()]
}

function addBlockPermissionTargets(model: DocumentModel, blockId: string, result: Map<string, DocumentPermissionTarget>): void {
  const host = documentEditingBlockHost(model) ?? documentBlockHostForId(model, blockId)
  const hostBlocks = host && !documentEditingBlockHost(model) ? documentStoredBlockHosts(model).find(entry => documentBlockHostId(entry.host) === documentBlockHostId(host))?.blocks : undefined
  const ancestry = blockPermissionAncestry(hostBlocks ?? model.blocks, blockId)
  const block = ancestry?.at(-1)
  if (!block) return
  const sectionId = host ? permissionSectionIdForHost(model, host) : sectionIdForBlock(model, ancestry![0]!.id)
  result.set(`section:${sectionId}`, { scope: 'section', id: sectionId })
  if (host?.kind === 'noteDecoration') {
    const sourceIds = model.metadata?.documentEditingSectionIds
    const sections = Array.isArray(sourceIds) && sourceIds.every(id => typeof id === 'string') ? sourceIds as string[] : documentSectionIds(model)
    for (const id of sections) result.set(`section:${id}`, { scope: 'section', id })
  }
  if (host) {
    const id = documentBlockHostId(host)
    result.set(`entity:${id}`, { scope: 'entity', id })
    if (host.kind === 'story') {
      const area = host.story.toLowerCase().includes('header') ? 'header' : 'footer'
      result.set(`entity:${host.sectionId}:${area}`, { scope: 'entity', id: `${host.sectionId}:${area}` })
    }
  }
  for (const ancestor of ancestry!) {
    result.set(`paragraph:${ancestor.id}`, { scope: 'paragraph', id: ancestor.id })
    if (isPermissionEntity(ancestor)) result.set(`entity:${ancestor.id}`, { scope: 'entity', id: ancestor.id })
  }
}

function permissionSectionIdForHost(model: DocumentModel, host: DocumentAuxiliaryBlockHost): string {
  if (host.kind === 'story') return host.sectionId
  const scoped = model.metadata?.documentEditingSectionId
  if (typeof scoped === 'string') return scoped
  if (host.kind === 'noteDecoration') return 'first'
  const note = (host.noteKind === 'footnote' ? model.footnotes : model.endnotes)?.[host.noteId]
  const reference = note && documentBlockAncestry(model, note.referenceBlockId)
  return sectionIdForBlock(model, reference?.[0]?.id ?? note?.referenceBlockId ?? '')
}

function allPermissionBlocks(model: DocumentModel): Map<string, DocumentBlock> {
  const result = flattenBlocks(model.blocks)
  for (const entry of documentStoredBlockHosts(model)) for (const [id, block] of flattenBlocks(entry.blocks)) result.set(id, block)
  return result
}

/** Global body operations use the same section/table/paragraph scopes as selected edits. */
export function permissionTargetsForBlock(model: DocumentModel, blockId: string): DocumentPermissionTarget[] {
  const result = new Map<string, DocumentPermissionTarget>()
  addBlockPermissionTargets(model, blockId, result)
  return [...result.values()]
}

function blockPermissionAncestry(blocks: readonly DocumentBlock[], id: string): DocumentBlock[] | undefined {
  for (const block of blocks) {
    if (block.id === id) return [block]
    const table = block.metadata.table as { rows?: Array<{ cells?: Array<{ blocks?: DocumentBlock[] }> }> } | undefined
    for (const row of table?.rows ?? []) for (const cell of row.cells ?? []) {
      const nested = blockPermissionAncestry(documentTableCellBlocks(cell), id)
      if (nested) return [block, ...nested]
    }
  }
  return undefined
}

export function permissionTargetsForChangedBlocks(before: DocumentModel, after: DocumentModel): DocumentPermissionTarget[] {
  const beforeBlocks = allPermissionBlocks(before)
  const afterBlocks = allPermissionBlocks(after)
  const changedIds = new Set<string>()
  for (const [id, block] of beforeBlocks) {
    if (JSON.stringify(block) !== JSON.stringify(afterBlocks.get(id))) changedIds.add(id)
  }
  for (const id of afterBlocks.keys()) if (!beforeBlocks.has(id)) changedIds.add(id)
  const result = new Map<string, DocumentPermissionTarget>()
  for (const id of changedIds) {
    addBlockPermissionTargets(before, id, result)
    addBlockPermissionTargets(after, id, result)
  }
  // Binding properties and backing XML can change without changing a paragraph's
  // displayed text. Publish/undo/redo must protect those same block scopes.
  const controls = new Map((before.contentControls ?? []).map(control => [control.id, control]))
  const nextControls = new Map((after.contentControls ?? []).map(control => [control.id, control]))
  for (const id of new Set([...controls.keys(), ...nextControls.keys()])) {
    const previous = controls.get(id), next = nextControls.get(id)
    if (JSON.stringify(previous) === JSON.stringify(next)) continue
    if (previous) addBlockPermissionTargets(before, previous.blockId, result)
    if (next) addBlockPermissionTargets(after, next.blockId, result)
  }
  const xmlStores = new Map((before.customXmlParts ?? []).filter(part => part.storeItemID).map(part => [part.storeItemID!.toLowerCase(), part]))
  const nextXmlStores = new Map((after.customXmlParts ?? []).filter(part => part.storeItemID).map(part => [part.storeItemID!.toLowerCase(), part]))
  const changedStores = new Set([...new Set([...xmlStores.keys(), ...nextXmlStores.keys()])]
    .filter(store => JSON.stringify(xmlStores.get(store)) !== JSON.stringify(nextXmlStores.get(store))))
  for (const model of [before, after]) for (const control of model.contentControls ?? []) {
    if (control.dataBinding?.storeItemID && changedStores.has(control.dataBinding.storeItemID.toLowerCase())) addBlockPermissionTargets(model, control.blockId, result)
  }
  for (const target of permissionTargetsForChangedStories(before, after)) result.set(`${target.scope}:${target.id}`, target)
  for (const noteKind of ['footnote', 'endnote'] as const) {
    const property = noteKind === 'footnote' ? 'footnotes' : 'endnotes'
    for (const id of new Set([...Object.keys(before[property] ?? {}), ...Object.keys(after[property] ?? {})])) {
      if (JSON.stringify(before[property]?.[id]) === JSON.stringify(after[property]?.[id])) continue
      const host: DocumentAuxiliaryBlockHost = { kind: 'note', noteKind, noteId: id }, entityId = documentBlockHostId(host)
      result.set(`entity:${entityId}`, { scope: 'entity', id: entityId })
      for (const model of [before, after]) {
        if (!model[property]?.[id]) continue
        const sectionId = permissionSectionIdForHost(model, host)
        result.set(`section:${sectionId}`, { scope: 'section', id: sectionId })
      }
    }
  }
  if (JSON.stringify(before.sectionSetup) !== JSON.stringify(after.sectionSetup)) result.set('section:first', { scope: 'section', id: 'first' })
  return [...result.values()]
}

/** Inheritance is resolved before checking actual story changes; canonical empty overrides are not separate edits. */
export function permissionTargetsForChangedStories(before: DocumentModel, after: DocumentModel): DocumentPermissionTarget[] {
  const result = new Map<string, DocumentPermissionTarget>()
  for (const sectionId of new Set([...documentSectionIds(before), ...documentSectionIds(after)])) {
    const previous = resolveSectionHeaderFooter(before, sectionId), next = resolveSectionHeaderFooter(after, sectionId)
    const firstChanged = previous.differentFirstPage !== next.differentFirstPage, evenChanged = previous.differentOddEven !== next.differentOddEven
    for (const story of DOCUMENT_STORY_KEYS) {
      const optionsChanged = firstChanged && (story.startsWith('firstPage') || story === 'header' || story === 'footer')
        || evenChanged && (story.startsWith('evenPage') || story === 'header' || story === 'footer')
      if (sameSectionStory(previous, next, story) && sectionStoryLinkedToPrevious(previous, story) === sectionStoryLinkedToPrevious(next, story) && !optionsChanged) continue
      result.set(`section:${sectionId}`, { scope: 'section', id: sectionId })
      result.set(`entity:${sectionId}:${story}`, { scope: 'entity', id: `${sectionId}:${story}` })
      const area = story.toLowerCase().includes('header') ? 'header' : 'footer'
      result.set(`entity:${sectionId}:${area}`, { scope: 'entity', id: `${sectionId}:${area}` })
    }
  }
  return [...result.values()]
}

function flattenBlocks(blocks: readonly DocumentBlock[]): Map<string, DocumentBlock> {
  const result = new Map<string, DocumentBlock>()
  for (const block of blocks) {
    result.set(block.id, block)
    const table = block.metadata.table as { rows?: Array<{ cells?: Array<{ blocks?: DocumentBlock[] }> }> } | undefined
    for (const row of table?.rows ?? []) for (const cell of row.cells ?? []) {
      for (const [id, nested] of flattenBlocks(documentTableCellBlocks(cell))) result.set(id, nested)
    }
  }
  return result
}

function samePoint(target: TextTarget, point: TextTarget): boolean {
  return target.blockId === point.blockId && target.row === point.row && target.col === point.col
    && (target.cellBlockId ?? target.blockId) === (point.cellBlockId ?? point.blockId)
}

function isPermissionEntity(block: DocumentBlock): boolean {
  return block.type === 'table' || block.type === 'image' || block.type === 'divider'
    || block.type === 'pageBreak' || block.type === 'columnBreak' || block.type === 'sectionBreak'
}

function matchingRule(
  rules: readonly DocumentPermissionRule[],
  userId: string,
  action: DocumentPermissionAction,
  scope: DocumentPermissionScope,
): DocumentPermissionRule | undefined {
  const applicable = rules.filter(rule => rule.scope === scope && rule.action === action && appliesToUser(rule, userId))
  return applicable.find(rule => rule.effect === 'deny') ?? applicable.find(rule => rule.effect === 'allow')
}

function appliesToUser(rule: DocumentPermissionRule, userId: string): boolean {
  return rule.userIds.length === 0 || rule.userIds.includes(userId)
}
