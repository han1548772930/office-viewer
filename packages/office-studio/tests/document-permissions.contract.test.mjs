import assert from 'node:assert/strict'
import { test } from 'vitest'
import { afterAll, beforeAll } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule
let permissionModule

beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  permissionModule = await server.ssrLoadModule('/packages/office-studio/src/features/document/permissions/document-permissions.ts')
})

afterAll(async () => { await server?.close() })

const context = { documentId: 'doc-1', userId: 'writer' }
const rule = (id, scope, action, effect, userIds = [], targetId) => ({ id, scope, action, effect, userIds, ...(targetId ? { targetId } : {}) })

test('document-level actions independently gate edit, copy, comment, print and export', () => {
  const rules = [
    rule('edit', 'document', 'edit', 'deny'),
    rule('copy', 'document', 'copy', 'deny'),
    rule('comment', 'document', 'comment', 'allow'),
    rule('print', 'document', 'print', 'deny', ['guest']),
  ]

  assert.equal(permissionModule.resolveDocumentPermission(rules, context, 'edit').allowed, false)
  assert.equal(permissionModule.resolveDocumentPermission(rules, context, 'copy').allowed, false)
  assert.equal(permissionModule.resolveDocumentPermission(rules, context, 'comment').allowed, true)
  assert.equal(permissionModule.resolveDocumentPermission(rules, context, 'print').allowed, true)
  assert.equal(permissionModule.resolveDocumentPermission(rules, context, 'export').allowed, true)
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
  const denied = permissionModule.resolveDocumentPermission(rules, context, 'edit', targets)
  assert.equal(denied.allowed, false)
  assert.equal(denied.deniedBy.id, 'p2-deny')

  const documentDenied = permissionModule.resolveDocumentPermission([...rules, rule('doc-deny', 'document', 'edit', 'deny')], context, 'edit', targets)
  assert.equal(documentDenied.deniedBy.id, 'doc-deny')
})

test('scoped grants and default-deny contexts require a matching user grant', () => {
  const rules = [rule('p1-allow', 'paragraph', 'edit', 'allow', ['writer'], 'p1')]
  const target = [{ scope: 'paragraph', id: 'p1' }]
  assert.equal(permissionModule.resolveDocumentPermission(rules, { ...context, defaultAllow: false }, 'edit', target).allowed, true)
  assert.equal(permissionModule.resolveDocumentPermission(rules, { ...context, userId: 'guest', defaultAllow: false }, 'edit', target).allowed, false)
  assert.equal(permissionModule.resolveDocumentPermission([rule('document-allow', 'document', 'edit', 'allow')], { ...context, defaultAllow: false }, 'edit', target).allowed, true)
})

test('assertion exposes a typed denial and rules reject ambiguous definitions', () => {
  const denied = [rule('edit-deny', 'document', 'edit', 'deny')]
  assert.throws(() => permissionModule.assertDocumentPermission(denied, context, 'edit'), permissionModule.DocumentPermissionError)
  assert.throws(() => permissionModule.validateDocumentPermissionRules([rule('bad', 'paragraph', 'copy', 'deny', [], 'p1')]), /only supports edit/)
  assert.throws(() => permissionModule.validateDocumentPermissionRules([rule('bad', 'paragraph', 'edit', 'deny')]), /requires a target ID/)
  assert.throws(() => permissionModule.validateDocumentPermissionRules([rule('same', 'document', 'edit', 'allow'), rule('same', 'document', 'copy', 'deny')]), /unique/)
})

test('Office API persists permission rules and blocks scoped edits through its command bus', () => {
  const studio = new apiModule.OfficeStudio({ currentUserId: 'writer' })
  const office = apiModule.FOffice.newAPI(studio)
  const document = office.createDocument({
    kind: 'document',
    blocks: [
      { id: 'p1', type: 'paragraph', level: 0, runs: [{ id: 'r1', text: 'Alpha', bold: false, italic: false, underline: false, strike: false, color: '#000000', fontFamily: 'Arial', fontSize: 11 }], style: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 }, metadata: {} },
      { id: 'p2', type: 'paragraph', level: 0, runs: [{ id: 'r2', text: 'Omega', bold: false, italic: false, underline: false, strike: false, color: '#000000', fontFamily: 'Arial', fontSize: 11 }], style: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 }, metadata: {} },
    ],
  })
  document.setPermissionRules([rule('deny-p1', 'paragraph', 'edit', 'deny', ['writer'], 'p1')])
  assert.equal(document.getPermissionRules()[0].targetId, 'p1')
  assert.throws(() => document.insertText(1, '!'), permissionModule.DocumentPermissionError)
  assert.equal(document.getText(), 'Alpha\nOmega')
  assert.equal(document.insertText(7, '!'), true)
  assert.equal(document.getText(), 'Alpha\nO!mega')
  studio.dispose()
})
