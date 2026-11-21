import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: 'packages/office-studio/vite.config.ts', server: { middlewareMode: true } })
const rulesApi = await vite.ssrLoadModule('/packages/office-studio/src/features/document/permissions/permission-rules.ts')
const modelApi = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')

test('permission rules can be created, updated and removed without mutating the source', () => {
  const initial = [{ id: 'document-edit', scope: 'document', action: 'edit', effect: 'allow', userIds: ['alice'] }]
  const created = rulesApi.createDocumentPermissionRule(initial, {
    scope: 'paragraph', targetId: 'paragraph-1', action: 'edit', effect: 'deny', userIds: [' bob ', 'bob', ''],
  }, () => 'paragraph-rule')
  assert.deepEqual(created, { id: 'paragraph-rule', scope: 'paragraph', targetId: 'paragraph-1', action: 'edit', effect: 'deny', userIds: ['bob'] })
  const updated = rulesApi.upsertDocumentPermissionRule(initial, { ...created, effect: 'allow' })
  assert.equal(updated.length, 2)
  assert.equal(updated[1].effect, 'allow')
  assert.equal(initial.length, 1)
  assert.deepEqual(rulesApi.removeDocumentPermissionRule(updated, 'paragraph-rule'), initial)
})

test('document rules omit targets and preserve a global empty-user set', () => {
  const rule = rulesApi.createDocumentPermissionRule([], {
    scope: 'document', action: 'export', effect: 'deny', userIds: [], targetId: 'ignored',
  }, () => 'export-rule')
  assert.deepEqual(rule, { id: 'export-rule', scope: 'document', action: 'export', effect: 'deny', userIds: [] })
})

test('invalid scoped actions, missing targets and duplicate IDs are rejected', () => {
  assert.throws(() => rulesApi.createDocumentPermissionRule([], {
    scope: 'section', targetId: 'section-1', action: 'print', effect: 'deny', userIds: [],
  }, () => 'invalid'))
  assert.throws(() => rulesApi.createDocumentPermissionRule([], {
    scope: 'paragraph', action: 'edit', effect: 'deny', userIds: [],
  }, () => 'missing-target'))
  const duplicate = [{ id: 'same', scope: 'document', action: 'edit', effect: 'deny', userIds: [] }]
  assert.throws(() => rulesApi.createDocumentPermissionRule(duplicate, {
    scope: 'document', action: 'edit', effect: 'allow', userIds: [],
  }, () => 'same'))
})

test('permission target labels use localized scope names and omit internal block type names', () => {
  const model = modelApi.createBlankDocument()
  model.blocks[0].runs[0].text = ''
  model.blocks.push({ ...model.blocks[0], id: 'page-break', type: 'pageBreak', runs: [] })
  const targetLabels = rulesApi.documentPermissionTargetOptions(model, (key, fallback) => ({
    'doc.permissions.scope.section': '节',
    'doc.permissions.scope.paragraph': '段落',
    'doc.permissions.scope.entity': '对象',
  }[key] ?? fallback))
  assert.ok(targetLabels.some(target => target.scope === 'paragraph' && target.label === '段落 1'))
  assert.ok(targetLabels.some(target => target.scope === 'entity' && target.label === '对象 1'))
  assert.ok(targetLabels.every(target => !target.label.includes('pageBreak') && !target.label.includes('Paragraph')))
})

await vite.close()
