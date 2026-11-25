import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createServer } from 'vite'

const vite = await createServer({ configFile: 'packages/office-studio/vite.config.ts', server: { middlewareMode: true } })
const rulesApi = await vite.ssrLoadModule('/packages/office-studio/src/features/document/permissions/permission-rules.ts')
const modelApi = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
const localeApi = await vite.ssrLoadModule('/packages/office-studio/src/i18n/docs/locales/permissions.ts')

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

test('scoped View rules can be saved independently from Edit rules', () => {
  const view = rulesApi.createDocumentPermissionRule([], {
    scope: 'paragraph', targetId: 'note-source', action: 'view', effect: 'deny', userIds: ['editor'],
  }, () => 'note-view')
  const edit = rulesApi.createDocumentPermissionRule([view], {
    scope: 'paragraph', targetId: 'note-source', action: 'edit', effect: 'allow', userIds: ['editor'],
  }, () => 'note-edit')
  const saved = rulesApi.upsertDocumentPermissionRule([view], edit)
  assert.deepEqual(saved.map(rule => [rule.action, rule.effect]), [['view', 'deny'], ['edit', 'allow']])
})

test('all permission locales name View and the six header/footer and note decoration targets', () => {
  const keys = ['action.view', 'host.header', 'host.footer', 'host.firstPageHeader', 'host.firstPageFooter',
    'host.evenPageHeader', 'host.evenPageFooter', 'host.footnote', 'host.endnote', 'host.separator',
    'host.continuationSeparator', 'host.continuationNotice', 'host.dormant']
  for (const [locale, messages] of Object.entries(localeApi.docPermissionLocaleMessages)) {
    for (const key of keys) assert.ok(typeof messages['doc.permissions.' + key] === 'string'
      && messages['doc.permissions.' + key].trim().length > 0, locale + ': ' + key)
  }
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

// Auxiliary target inventory draft; no tests or browser acceptance were run for this batch.
test('permission target inventory includes note and active/default/public dormant decoration identities without projecting their text', () => {
  const value = modelApi.createBlankDocument(), block = (id, text) => {
    const owned = JSON.parse(JSON.stringify(value.blocks[0])); owned.id = id; owned.runs[0].id = id + '-run'; owned.runs[0].text = text; return owned
  }, story = block => ({ blocks: [block], runs: block.runs, paragraphs: [{ startOffset: 0, endOffset: block.runs[0].text.length, style: block.style }] })
  value.footnotes = { orphan: { id: 'orphan', text: 'Sensitive note', referenceBlockId: '', referenceRunId: '', content: story(block('note-p', 'Sensitive note')) } }
  const dormant = story(block('dormant-p', 'Sensitive dormant content')); dormant.metadata = {
    univerNoteDecorationSource: { version: 1, kind: 'footnote', type: 'continuationSeparator', noteId: 'public-extra' },
  }
  value.noteSettings = { footnote: { separator: story(block('separator-p', 'Sensitive separator')), metadata: { officeDormantNoteDecorations: { version: 1, contents: [dormant] } } } }
  const before = JSON.parse(JSON.stringify(value)), targets = rulesApi.documentPermissionTargetOptions(value)
  for (const id of ['footnote:orphan', 'noteDecoration:footnote:separator:active', 'noteDecoration:footnote:continuationNotice:active',
    'noteDecoration:footnote:continuationSeparator:0']) assert.ok(targets.some(target => target.scope === 'entity' && target.id === id))
  for (const id of ['note-p', 'separator-p', 'dormant-p']) assert.ok(targets.some(target => target.scope === 'paragraph' && target.id === id))
  assert.equal(new Set(targets.map(target => target.scope + ':' + target.id)).size, targets.length)
  assert.ok(targets.every(target => !target.label.includes('Sensitive'))); assert.deepEqual(value, before)
})

await vite.close()
