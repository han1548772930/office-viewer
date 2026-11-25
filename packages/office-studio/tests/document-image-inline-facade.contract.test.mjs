import assert from 'node:assert/strict'
import { File as NodeFile } from 'node:buffer'
import { afterAll, afterEach, beforeAll, beforeEach, test, vi } from 'vitest'
import { createServer } from 'vite'

// Deferred real-Facade/command/history/codec contracts. No Vite server or test has been started.
let server, OfficeStudio, FDocument, domain, modelUtils, atoms, imageModel, commands, codec, sections, storyBlocks
const studios = new Set(), copy = value => structuredClone(value)
const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg=='
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-image-inline-facade', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  ;[domain, modelUtils, atoms, imageModel, commands, codec, sections, storyBlocks] = await Promise.all([
    'domain/model', 'features/document/model/document-model', 'features/document/images/document-image-atom',
    'features/document/images/document-image-model', 'api/document-image-commands', 'api/document-body-codec',
    'features/document/services/section-header-footer', 'features/document/model/document-story-blocks',
  ].map(path => server.ssrLoadModule('/packages/office-studio/src/' + path + '.ts')))
})
beforeEach(() => {
  vi.stubGlobal('File', NodeFile)
  vi.stubGlobal('Image', class {
    onload = null; onerror = null; complete = false; naturalWidth = 0; naturalHeight = 0
    get src() { return this.value ?? '' }
    set src(value) {
      this.value = value
      if (value) queueMicrotask(() => { if (this.value) { this.complete = true; this.naturalWidth = 1; this.naturalHeight = 1; this.onload?.() } })
    }
  })
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear(); vi.unstubAllGlobals() })
afterAll(async () => { await server?.close() })
const paragraph = (id, text = '') => ({ ...modelUtils.createParagraph(text), id })
const atom = (id, patch = {}) => {
  const run = modelUtils.createRun()
  atoms.setDocumentImageAtom(run, id, { src: png, alt: id, width: 80, height: 40, intrinsicWidth: 1, intrinsicHeight: 1, wrappingStyle: 'inline', ...patch })
  return run
}
function open(source = { ...domain.createBlankDocument(), id: 'inline-facade', blocks: [paragraph('p', 'ABCD'), paragraph('tail', 'Tail')] }) {
  const studio = new OfficeStudio({ document: source, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(source.id)
  return { studio, document: new FDocument(studio, source.id) }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0 })

test('public textRange insertion replaces source text with a true paragraph atom and preserves both flip options', async () => {
  const { studio, document } = open(), image = await document.insertImage({ source: png, width: 80, height: 40, flipX: true, flipY: true,
    textRange: { startOffset: 1, endOffset: 3 } })
  assert.ok(image); assert.equal(document.getModel().blocks[0].type, 'paragraph')
  assert.equal(modelUtils.runText(document.getModel().blocks[0].runs), 'A\uFFFCD'); assert.equal(document.getModel().blocks.length, 2)
  const entry = imageModel.documentImageEntries(document.getModel())[0]
  assert.equal(entry.imageId, image.getId()); assert.deepEqual([entry.inline.startOffset, entry.inline.endOffset], [1, 2])
  assert.equal(image.getFlipX(), true); assert.equal(image.getFlipY(), true)
  const snapshot = document.getSnapshot(), custom = snapshot.body.customBlocks.find(block => block.blockId === image.getId())
  assert.equal(custom.startIndex, 1); assert.equal(custom.officeStudio.imageAtom, true)
  assert.equal(snapshot.drawings[image.getId()].docTransform.flipX, true); assert.equal(history(studio, document.id).undo, 1)
})

test('flip transform batches update two atoms once and a later invalid item cannot partly publish', () => {
  const block = paragraph('p'); block.runs = [modelUtils.createRun('L'), atom('a'), modelUtils.createRun('M'), atom('b')]
  const { studio, document } = open({ ...domain.createBlankDocument(), id: 'flip-batch', blocks: [block] })
  const transform = commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.transform
  assert.equal(studio.executeCommandSync(transform, { unitId: document.id, drawings: [
    { drawingId: 'a', key: 'flipX', value: true }, { drawingId: 'b', key: 'flipY', value: true },
  ] }), true)
  assert.equal(document.getImage('a').getFlipX(), true); assert.equal(document.getImage('b').getFlipY(), true)
  assert.equal(history(studio, document.id).undo, 1)
  const before = copy(document.getModel()), oldHistory = history(studio, document.id)
  assert.throws(() => studio.executeCommandSync(transform, { unitId: document.id, drawings: [
    { drawingId: 'a', key: 'flipX', value: false }, { drawingId: 'b', key: 'flipY', value: 'true' },
  ] }))
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), oldHistory)
  assert.equal(document.undo(), true); assert.equal(document.getImage('a').getFlipX(), false)
  assert.equal(document.redo(), true); assert.equal(document.getImage('b').getFlipY(), true)
})

test('atom identity and flips survive save, undo, redo, removal and public snapshot reopening', async () => {
  const { document } = open(), image = await document.insertImage({ source: png, flipX: true, flipY: true,
    textRange: { startOffset: 2, endOffset: 2 } }), imageId = image.getId()
  assert.equal(image.setFlipY(false), true); assert.equal(document.undo(), true); assert.equal(image.getFlipY(), true)
  assert.equal(document.redo(), true); assert.equal(image.getFlipY(), false)
  const saved = document.save(), reopened = open(copy(saved)).document
  assert.equal(reopened.getImage(imageId).getFlipX(), true); assert.equal(reopened.getImage(imageId).getFlipY(), false)
  assert.equal(modelUtils.runText(reopened.getModel().blocks[0].runs), 'AB\uFFFCCD')
  assert.equal(reopened.getImage(imageId).remove(), true); assert.equal(reopened.getImage(imageId), null)
  assert.equal(modelUtils.runText(reopened.getModel().blocks[0].runs), 'ABCD')
  assert.equal(reopened.undo(), true); assert.equal(reopened.getImage(imageId).getFlipX(), true)
  assert.equal(open(reopened.save()).document.getImage(imageId).getFlipX(), true)
})

test('public image queries enforce the drawing entity view permission independently of ordinary paragraph text', () => {
  const block = paragraph('p'); block.runs = [modelUtils.createRun('L'), atom('hidden-image'), modelUtils.createRun('R')]
  const source = { ...domain.createBlankDocument(), id: 'image-view-permission', blocks: [block], permissionRules: [
    { id: 'deny-image', scope: 'entity', targetId: 'hidden-image', action: 'view', effect: 'deny', userIds: ['editor'] },
  ] }
  const { studio, document } = open(source), before = copy(document.getModel())
  assert.throws(() => document.getImage('hidden-image'), /permission denied/i)
  assert.throws(() => document.getImages(), /permission denied/i)
  assert.equal(document.getParagraph('p').getText().includes('L'), true)
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
})

test('header and note atom images keep separate segment commands, source ranges and flips through reopening', () => {
  const body = paragraph('p'); body.runs = [modelUtils.createRun('B'), atom('body-image')]
  const headerParagraph = paragraph('header-p'); headerParagraph.runs = [modelUtils.createRun('H'), atom('header-image')]
  const source = { ...domain.createBlankDocument(), id: 'inline-segments', blocks: [body] }
  const header = storyBlocks.projectDocumentStoryBlocks({ blocks: [headerParagraph], runs: [], paragraphs: [], metadata: {} }).content
  source.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: 'H\uFFFC', storyContent: { header } } }
  const { studio, document } = open(source), segmentId = document.getSnapshot().documentStyle.defaultHeaderId
  const noteParagraph = paragraph('note-p'); noteParagraph.runs = [modelUtils.createRun('N'), atom('note-image')]
  const noteId = document.insertFootnote(0, codec.encodeDocumentBody({ blocks: [noteParagraph], runs: [], paragraphs: [], metadata: {} }))
  assert.ok(segmentId); assert.ok(noteId)
  assert.equal(document.getImage('header-image', segmentId).setFlipX(true), true)
  assert.equal(document.getImage('note-image', noteId).setFlipY(true), true)
  assert.equal(document.getImage('header-image'), null)
  assert.throws(() => studio.executeCommandSync(commands.DOCUMENT_IMAGE_PUBLIC_COMMANDS.transform, { unitId: document.id, segmentId,
    drawings: [{ drawingId: 'note-image', key: 'flipX', value: true }] }))
  const reopened = open(copy(document.save())).document
  assert.equal(reopened.getImage('header-image', segmentId).getFlipX(), true)
  assert.equal(reopened.getImage('note-image', noteId).getFlipY(), true)
  assert.equal(reopened.getImage('body-image').getFlipX(), false)
  assert.equal(imageModel.documentImageEntries(reopened.getSegmentModel(segmentId))[0].inline.startOffset, 1)
})
