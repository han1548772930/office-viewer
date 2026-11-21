import assert from 'node:assert/strict'
import { after, test } from 'node:test'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
after(() => vite.close())

const { normalizeDocument } = await vite.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts')
const { createChecklistParagraph, documentHtml } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { executeDocumentCommand } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')

function envelope(blocks) {
  return {
    schemaVersion: 1, parserVersion: 'contract', format: 'docx', fileName: 'checklist.docx', size: 100,
    model: 'document', payload: { modelType: 'document', data: { documentId: 'checklist-doc', blocks } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  }
}

function parsedChecklist(id, listLabel, content) {
  return { id, type: 'paragraph', content, details: { inlines: [{ type: 'text', text: content }], numbering: { numId: 1, level: 0 }, listLabel } }
}

test('DOCX checklist labels normalize to independent checked paragraph state', () => {
  const model = normalizeDocument(envelope([
    parsedChecklist('unchecked', '☐', 'Prepare files'),
    parsedChecklist('checked', '☑', 'Review changes'),
    parsedChecklist('bullet', '•', 'Keep this a bullet'),
  ]))
  assert.deepEqual(model.blocks.map(block => block.type), ['checklist', 'checklist', 'bullet'])
  assert.deepEqual(model.blocks.slice(0, 2).map(block => block.checklist.checked), [false, true])
  assert.equal(model.blocks[0].checklist.listId, '1')
  assert.equal(model.blocks[1].checklist.listId, '1')
  assert.equal(model.blocks[0].metadata.source.details.listLabel, '☐')
})

test('checklist mutations toggle one task and ordinary paragraphs can join a shared list', () => {
  const first = createChecklistParagraph('One', false, 'tasks')
  const second = createChecklistParagraph('Two', false, 'tasks')
  const model = { ...normalizeDocument(envelope([])), blocks: [first, second] }
  executeDocumentCommand(model, { anchor: { blockId: second.id, offset: 0 }, focus: { blockId: second.id, offset: 0 } }, { id: 'doc.command.toggle-checklist-item', blockId: second.id })
  assert.deepEqual(model.blocks.map(block => block.checklist.checked), [false, true])
  const plain = { ...first, id: 'plain', type: 'paragraph', checklist: undefined }
  model.blocks.push(plain)
  executeDocumentCommand(model, { anchor: { blockId: 'plain', offset: 0 }, focus: { blockId: 'plain', offset: 0 } }, { id: 'doc.command.set-block-type', type: 'checklist' })
  assert.equal(plain.type, 'checklist')
  assert.equal(plain.checklist.listId, 'tasks')
  assert.equal(plain.checklist.checked, false)
})

test('internal HTML snapshot keeps checklist state and list grouping', () => {
  const first = createChecklistParagraph('One', false, 'tasks')
  const second = createChecklistParagraph('Two', true, 'tasks')
  const model = { ...normalizeDocument(envelope([])), blocks: [first, second] }
  const html = documentHtml(model)
  assert.equal((html.match(/<ul /gu) ?? []).length, 1)
  assert.match(html, /data-office-list-type="checklist" data-office-list-id="tasks"/u)
  assert.match(html, /data-office-checked="false"/u)
  assert.match(html, /data-office-checked="true"/u)
  assert.ok(html.indexOf('One') < html.indexOf('Two'))
})
