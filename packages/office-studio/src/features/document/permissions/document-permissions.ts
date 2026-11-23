import type { DocumentBlock, DocumentModel, DocumentPermissionAction, DocumentPermissionRule, DocumentPermissionScope } from '../../../domain/model'
import type { DocumentSelection, TextTarget } from '../model/document-model'
import { orderedSelection, textTargets } from '../model/document-model'
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
    if (rule.scope !== 'document' && rule.action !== 'edit') {
      throw new Error(`Scoped permission rule ${rule.id} only supports edit`)
    }
  }
}

/**
 * Resolves local UI/command permission. A document-level denial always wins.
 * For scoped edits, every target must remain editable and any matching deny
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
  if (action !== 'edit' || targets.length === 0) {
    return { allowed: documentRule?.effect === 'allow' || context.defaultAllow !== false, action }
  }
  const documentAllows = documentRule?.effect === 'allow' || context.defaultAllow !== false

  for (const target of targets) {
    const scopedRules = rules.filter(rule =>
      rule.scope === target.scope && rule.targetId === target.id && rule.action === 'edit' && appliesToUser(rule, context.userId),
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
  const block = findBlock(model.blocks, blockId)
  if (!block) return
  const sectionId = sectionIdForBlock(model, block.id)
  result.set(`paragraph:${block.id}`, { scope: 'paragraph', id: block.id })
  result.set(`section:${sectionId}`, { scope: 'section', id: sectionId })
  if (isPermissionEntity(block)) result.set(`entity:${block.id}`, { scope: 'entity', id: block.id })
}

export function permissionTargetsForChangedBlocks(before: DocumentModel, after: DocumentModel): DocumentPermissionTarget[] {
  const beforeBlocks = flattenBlocks(before.blocks)
  const afterBlocks = flattenBlocks(after.blocks)
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
  for (const target of permissionTargetsForChangedStories(before, after)) result.set(`${target.scope}:${target.id}`, target)
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
      for (const [id, nested] of flattenBlocks(cell.blocks ?? [])) result.set(id, nested)
    }
  }
  return result
}

function samePoint(target: TextTarget, point: TextTarget): boolean {
  return target.blockId === point.blockId && target.row === point.row && target.col === point.col
    && (target.cellBlockId ?? target.blockId) === (point.cellBlockId ?? point.blockId)
}

function findBlock(blocks: readonly DocumentBlock[], id: string): DocumentBlock | undefined {
  for (const block of blocks) {
    if (block.id === id) return block
    const table = block.metadata.table as { rows?: Array<{ cells?: Array<{ blocks?: DocumentBlock[] }> }> } | undefined
    for (const row of table?.rows ?? []) for (const cell of row.cells ?? []) {
      const nested = findBlock(cell.blocks ?? [], id)
      if (nested) return nested
    }
  }
  return undefined
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
