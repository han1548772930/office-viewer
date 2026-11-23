import type { DocumentModel, DocumentPermissionRule, DocumentPermissionScope } from '../../../domain/model.ts'
import { validateDocumentPermissionRules } from './document-permissions.ts'
import { findDocumentParagraphs } from '../services/document-facade.ts'
import { documentSectionIds } from '../services/section-header-footer.ts'
import { documentTableCellBlocks } from '../model/document-table-continuations'

export interface DocumentPermissionRuleDraft {
  scope: DocumentPermissionScope
  targetId?: string
  action: DocumentPermissionRule['action']
  effect: DocumentPermissionRule['effect']
  userIds: readonly string[]
}

export interface DocumentPermissionTargetOption {
  scope: Exclude<DocumentPermissionScope, 'document'>
  id: string
  label: string
}

export function documentPermissionTargetOptions(
  model: DocumentModel,
  translate: (key: string, fallback: string) => string = (_key, fallback) => fallback,
): DocumentPermissionTargetOption[] {
  const sections = documentSectionIds(model).map((id, index) => ({
    scope: 'section' as const,
    id,
    label: `${translate('doc.permissions.scope.section', 'Section')} ${index + 1}`,
  }))
  const paragraphs = findDocumentParagraphs(model, '').map((paragraph, index) => ({
    scope: 'paragraph' as const,
    id: paragraph.paragraphId,
    label: `${translate('doc.permissions.scope.paragraph', 'Paragraph')} ${index + 1}${paragraph.text.trim() ? ` · ${paragraph.text.trim().slice(0, 80)}` : ''}`,
  }))
  const entities = permissionEntities(model.blocks).map((block, index) => ({
    scope: 'entity' as const,
    id: block.id,
    label: `${translate('doc.permissions.scope.entity', 'Object')} ${index + 1}`,
  }))
  return [...sections, ...paragraphs, ...entities]
}

export function createDocumentPermissionRule(
  rules: readonly DocumentPermissionRule[],
  draft: DocumentPermissionRuleDraft,
  createId: () => string = () => `permission-${crypto.randomUUID()}`,
): DocumentPermissionRule {
  const id = createId().trim()
  if (!id || rules.some(rule => rule.id === id)) throw new Error('Permission rule IDs must be non-empty and unique')
  const rule = normalizeRule({ ...draft, id })
  validateDocumentPermissionRules([...rules, rule])
  return rule
}

export function upsertDocumentPermissionRule(
  rules: readonly DocumentPermissionRule[],
  rule: DocumentPermissionRule,
): DocumentPermissionRule[] {
  const next = rules.some(candidate => candidate.id === rule.id)
    ? rules.map(candidate => candidate.id === rule.id ? normalizeRule(rule) : { ...candidate, userIds: [...candidate.userIds] })
    : [...rules.map(candidate => ({ ...candidate, userIds: [...candidate.userIds] })), normalizeRule(rule)]
  validateDocumentPermissionRules(next)
  return next
}

export function removeDocumentPermissionRule(
  rules: readonly DocumentPermissionRule[],
  ruleId: string,
): DocumentPermissionRule[] {
  const next = rules.filter(rule => rule.id !== ruleId).map(rule => ({ ...rule, userIds: [...rule.userIds] }))
  validateDocumentPermissionRules(next)
  return next
}

function normalizeRule(rule: DocumentPermissionRuleDraft & { id: string }): DocumentPermissionRule {
  const normalized: DocumentPermissionRule = {
    id: rule.id,
    scope: rule.scope,
    action: rule.action,
    effect: rule.effect,
    userIds: [...new Set(rule.userIds.map(userId => userId.trim()).filter(Boolean))],
  }
  if (rule.scope !== 'document') normalized.targetId = rule.targetId?.trim()
  return normalized
}

function permissionEntities(blocks: DocumentModel['blocks']): DocumentModel['blocks'] {
  const result: DocumentModel['blocks'] = []
  for (const block of blocks) {
    if (['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) result.push(block)
    const table = block.metadata.table as { rows?: Array<{ cells?: Array<{ blocks?: DocumentModel['blocks'] }> }> } | undefined
    for (const row of table?.rows ?? []) for (const cell of row.cells ?? []) result.push(...permissionEntities(documentTableCellBlocks(cell)))
  }
  return result
}
