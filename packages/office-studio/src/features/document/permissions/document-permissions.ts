export type DocumentPermissionAction = 'edit' | 'copy' | 'comment' | 'print' | 'export'
export type DocumentPermissionScope = 'document' | 'section' | 'paragraph' | 'entity'

export interface DocumentPermissionTarget {
  scope: Exclude<DocumentPermissionScope, 'document'>
  id: string
}

export interface DocumentPermissionRule {
  id: string
  scope: DocumentPermissionScope
  targetId?: string
  action: DocumentPermissionAction
  effect: 'allow' | 'deny'
  userIds: string[]
}

export interface DocumentPermissionContext {
  documentId: string
  userId: string
  defaultAllow?: boolean
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

  for (const target of targets) {
    const scopedRules = rules.filter(rule =>
      rule.scope === target.scope && rule.targetId === target.id && rule.action === 'edit' && appliesToUser(rule, context.userId),
    )
    const denial = scopedRules.find(rule => rule.effect === 'deny')
    if (denial) return { allowed: false, action, deniedBy: denial }
    if (!scopedRules.some(rule => rule.effect === 'allow') && context.defaultAllow === false) {
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
