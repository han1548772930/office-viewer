import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  assertDocumentPermission,
  DocumentPermissionError,
  resolveDocumentPermission,
  validateDocumentPermissionRules,
} from '../src/features/document/permissions/document-permissions.ts'

const context = { documentId: 'doc-1', userId: 'writer' }
const rule = (id, scope, action, effect, userIds = [], targetId) => ({ id, scope, action, effect, userIds, ...(targetId ? { targetId } : {}) })

test('document-level actions independently gate edit, copy, comment, print and export', () => {
  const rules = [
    rule('edit', 'document', 'edit', 'deny'),
    rule('copy', 'document', 'copy', 'deny'),
    rule('comment', 'document', 'comment', 'allow'),
    rule('print', 'document', 'print', 'deny', ['guest']),
  ]

  assert.equal(resolveDocumentPermission(rules, context, 'edit').allowed, false)
  assert.equal(resolveDocumentPermission(rules, context, 'copy').allowed, false)
  assert.equal(resolveDocumentPermission(rules, context, 'comment').allowed, true)
  assert.equal(resolveDocumentPermission(rules, context, 'print').allowed, true)
  assert.equal(resolveDocumentPermission(rules, context, 'export').allowed, true)
})

test('paragraph editing checks document and every scoped target, with deny taking precedence', () => {
  const targets = [
    { scope: 'paragraph', id: 'p1' },
    { scope: 'paragraph', id: 'p2' },
  ]
  const rules = [
    rule('p1-allow', 'paragraph', 'edit', 'allow', ['writer'], 'p1'),
    rule('p2-allow', 'paragraph', 'edit', 'allow', ['writer'], 'p2'),
    rule('p2-deny', 'paragraph', 'edit', 'deny', ['writer'], 'p2'),
  ]
  const denied = resolveDocumentPermission(rules, context, 'edit', targets)
  assert.equal(denied.allowed, false)
  assert.equal(denied.deniedBy.id, 'p2-deny')

  const documentDenied = resolveDocumentPermission([...rules, rule('doc-deny', 'document', 'edit', 'deny')], context, 'edit', targets)
  assert.equal(documentDenied.deniedBy.id, 'doc-deny')
})

test('scoped grants and default-deny contexts require a matching user grant', () => {
  const rules = [rule('p1-allow', 'paragraph', 'edit', 'allow', ['writer'], 'p1')]
  const target = [{ scope: 'paragraph', id: 'p1' }]
  assert.equal(resolveDocumentPermission(rules, { ...context, defaultAllow: false }, 'edit', target).allowed, true)
  assert.equal(resolveDocumentPermission(rules, { ...context, userId: 'guest', defaultAllow: false }, 'edit', target).allowed, false)
})

test('assertion exposes a typed denial and rules reject ambiguous definitions', () => {
  const denied = [rule('edit-deny', 'document', 'edit', 'deny')]
  assert.throws(() => assertDocumentPermission(denied, context, 'edit'), DocumentPermissionError)
  assert.throws(() => validateDocumentPermissionRules([rule('bad', 'paragraph', 'copy', 'deny', [], 'p1')]), /only supports edit/)
  assert.throws(() => validateDocumentPermissionRules([rule('bad', 'paragraph', 'edit', 'deny')]), /requires a target ID/)
  assert.throws(() => validateDocumentPermissionRules([rule('same', 'document', 'edit', 'allow'), rule('same', 'document', 'copy', 'deny')]), /unique/)
})
