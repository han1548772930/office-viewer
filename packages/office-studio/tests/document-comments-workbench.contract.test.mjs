import assert from 'node:assert/strict'
import { afterAll, test } from 'vitest'
import { createPinia } from 'pinia'
import { createServer } from 'vite'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
const ribbon = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-ribbon.ts')
const toolbar = await vite.ssrLoadModule('/packages/office-studio/src/components/document/document-toolbar-state.ts')
const modelApi = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const comments = await vite.ssrLoadModule('/packages/office-studio/src/features/document/comments/thread-comments.ts')
const documentFacade = await vite.ssrLoadModule('/packages/office-studio/src/features/document/services/document-facade.ts')
const editor = await vite.ssrLoadModule('/packages/office-studio/src/stores/editorStore.ts')
const domain = await vite.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
afterAll(() => vite.close())

test('comment entries match the two fixed Docs Insert > Others operations and selection state', () => {
  const state = { marks: {}, mixedMarks: [], paragraph: { align: 'left', indent: 0 }, hasSelection: false }
  const tabs = ribbon.createDocumentRibbon({ state, canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false })
  const others = tabs.find(item => item.id === 'insert').groups.find(group => group.id === 'doc-insert-extensions')
  const toggle = others.items.find(item => item.commandId === 'docs.operation.toggle-comment-panel')
  const add = others.items.find(item => item.commandId === 'docs.operation.start-add-comment')

  assert.ok(toggle)
  assert.ok(add)
  assert.equal(toggle.activated, undefined)
  assert.equal(add.disabled, true)
  assert.deepEqual([toggle.id, add.id], ['doc-comment-toggle', 'doc-comment-add'])
  assert.deepEqual([toggle.gridLayout, add.gridLayout], [
    { row: 1, column: 1, showLabel: true },
    { row: 2, column: 1, showLabel: true },
  ])
  assert.equal(ribbon.createDocumentRibbon({ state, canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false, paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: false, commentsPanelOpen: true, canCreateComment: true })
    .find(item => item.id === 'insert').groups.find(group => group.id === 'doc-insert-extensions').items.find(item => item.id === 'doc-comment-toggle').activated, true)
})

test('comment model updates publish through document history and survive undo, redo, and snapshot serialization', () => {
  const pinia = createPinia()
  const session = editor.useEditorStore(pinia)
  const model = domain.createBlankDocument()
  model.id = 'comments-history-doc'
  const paragraph = modelApi.createParagraph('Review this')
  model.blocks = [paragraph]
  session.document.replace(model)
  session.setActiveKind('document')
  const changed = modelApi.cloneDocument(session.document.model)
  const selection = { anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 6 } }
  comments.createDocumentThreadComment(changed, selection, 'Check this wording', { documentId: model.id, userId: 'author', displayName: 'Author' })

  assert.equal(session.replaceDocumentFromView(changed), true)
  assert.equal(Object.keys(session.document.model.comments).length, 1)
  assert.ok(session.undo())
  assert.equal(session.document.model.comments, undefined)
  assert.ok(session.redo())
  assert.equal(Object.values(session.document.model.comments)[0].referenceText, 'Review')
  assert.deepEqual(documentFacade.saveDocumentSnapshot(session.document.model).comments, session.document.model.comments)
})

test('resolved threads remain persisted in the comment model with their resolved state', () => {
  const model = domain.createBlankDocument()
  const paragraph = modelApi.createParagraph('Comment')
  model.blocks = [paragraph]
  const thread = comments.createDocumentThreadComment(model, {
    anchor: { blockId: paragraph.id, offset: 0 }, focus: { blockId: paragraph.id, offset: 7 },
  }, 'Done', { documentId: model.id ?? '', userId: 'author', displayName: 'Author' })
  comments.setDocumentThreadCommentResolved(model, thread.id, true, { documentId: model.id ?? '', userId: 'author', displayName: 'Author' })

  assert.equal(thread.resolved, true)
  assert.equal(Object.values(model.comments).length, 1)
  assert.equal(toolbar.documentToolbarState(model, modelApi.collapsedSelection({ blockId: paragraph.id, offset: 0 }), {}).hasSelection, false)
})
