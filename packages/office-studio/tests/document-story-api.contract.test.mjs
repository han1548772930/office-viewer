import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Drafts only: no test, type check or build was executed in this turn.
let server, OfficeStudio, FDocument, domain, docModel, storyModel, storyEditor, sections, codeBlocks
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  docModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  storyModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story.ts')
  storyEditor = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/story-editor.ts')
  sections = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/section-header-footer.ts')
  codeBlocks = await server.ssrLoadModule('/packages/office-studio/src/features/document/code-blocks/code-blocks.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const range = (blockId, start = 0, end = start) => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })
const readSection = document => sections.resolveSectionHeaderFooter(document.getModel(), 'first')

function sourceModel(id = 'story-api-doc') {
  const model = domain.createBlankDocument(), paragraph = docModel.createParagraph('body untouched')
  model.id = id; paragraph.id = id + '-body'; model.blocks = [paragraph]
  model.defaultStyle.text = { ...model.defaultStyle.text, fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 15, color: 'auto' }
  model.defaultStyle.paragraph = { ...model.defaultStyle.paragraph, align: 'right', spaceAfter: 12 }
  const text = 'A😀 word\nTail', run = docModel.createRun(text, { bold: true, italic: true, fontSize: 22, color: '#FF0000', href: 'https://example.org/story' })
  run.metadata = { source: { sourceFormat: { version: 1, propertiesXml: '<w:rPr><w:caps/><w:spacing w:val="24"/></w:rPr>', namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } } } }
  const paragraphs = storyModel.documentStoryParagraphs(text, { ...docModel.defaultParagraphStyle(), align: 'center' })
  model.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: text, storyContent: { header: { runs: [run], paragraphs } },
    pageFields: [{ story: 'header', field: 'PAGE', offset: 3, cachedResult: '12', textStyle: { ...storyModel.documentStoryDefaultStyle(model).text, fontSize: 18 } }] } }
  return model
}

function open(model = sourceModel()) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id), body: range(model.blocks[0].id, 2, 5) }
}

test('facade story clear, edit, undo, redo and save use one document history and retain the body selection', () => {
  const { studio, document, body } = open(), before = copy(document.getModel()), events = []
  studio.events.addEvent('CommandExecuted', event => events.push([event.id, event.unitId, event.kind]))
  const request = { sectionId: 'first', story: 'header', start: 9, end: 5 }
  const result = document.clearFormatting(body, request)
  assert.equal(result.changed, true); assert.deepEqual(result.selection, body)
  assert.deepEqual(document.getModel().blocks, before.blocks)
  const afterClear = copy(document.getModel()), cleared = readSection(document)
  assert.deepEqual(cleared.storyContent.header.runs.map(run => [run.text, run.bold, run.fontSize]), [['A😀 ', true, 22], ['word', false, 9], ['\nTail', true, 22]])
  assert.equal(cleared.storyContent.header.runs[1].fontFamily, 'Georgia')
  assert.equal(cleared.storyContent.header.runs[1].href, 'https://example.org/story')
  assert.equal(cleared.storyContent.header.runs[1].metadata.officeStoryTextFormatCleared, true)
  assert.deepEqual(cleared.pageFields, before.sectionHeaderFooters.first.pageFields)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.equal(document.clearFormatting(body, request).changed, false)
  assert.deepEqual(document.getModel(), afterClear)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  const patch = storyEditor.replaceStoryRange(cleared, 'header', 5, 9, 'new', document.getModel())
  const insertion = document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch }, body)
  assert.equal(insertion.changed, true); assert.deepEqual(insertion.selection, body)
  assert.equal(readSection(document).header, 'A😀 new\nTail')
  const inserted = readSection(document).storyContent.header.runs.find(run => run.text === 'new')
  assert.deepEqual([inserted.bold, inserted.fontFamily, inserted.fontSize, inserted.metadata.officeStoryTextFormatCleared], [false, 'Georgia', 9, true])
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), afterClear)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.equal(document.redo(), true)
  const saved = document.save(), reopened = open(copy(saved))
  assert.equal(saved.modified, false); assert.equal(history(studio, document.id).dirty, false)
  assert.deepEqual(reopened.document.getModel().sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.deepEqual(reopened.document.getModel().defaultStyle, saved.defaultStyle)
  assert.deepEqual(history(reopened.studio, saved.id), { undo: 0, redo: 0, dirty: false })
  assert.deepEqual(events.filter(event => event[0] === 'doc.command.clear-formatting'), [
    ['doc.command.clear-formatting', document.id, 'document'], ['doc.command.clear-formatting', document.id, 'document'],
  ])
})

test('story clear cancellation, section permissions, surrogate boundaries and locks reject without a history entry', () => {
  const { studio, document, body } = open(), before = copy(document.getModel())
  const story = { sectionId: 'first', story: 'header', start: 5, end: 9 }
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.clear-formatting') event.cancel = true })
  assert.throws(() => document.clearFormatting(body, story), /canceled/); cancel.dispose()
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  for (const kind of ['entity', 'section', 'surrogate', 'outside-word-lock']) {
    const model = sourceModel(kind), request = { ...story }
    if (kind === 'entity' || kind === 'section') model.permissionRules = [{ id: 'deny', scope: kind, targetId: kind === 'section' ? 'first' : 'first:header', action: 'edit', effect: 'deny', userIds: [] }]
    if (kind === 'surrogate') request.start = 2
    if (kind === 'outside-word-lock') model.sectionHeaderFooters.first.storyContent.header.locks = [{ id: 'emoji', startOffset: 1, endOffset: 3, lock: 'contentLocked' }]
    const blocked = open(model), original = copy(blocked.document.getModel())
    assert.throws(() => blocked.document.clearFormatting(blocked.body, request), undefined, kind)
    assert.deepEqual(blocked.document.getModel(), original, kind)
    assert.deepEqual(history(blocked.studio, blocked.document.id), { undo: 0, redo: 0, dirty: false }, kind)
  }
})

test('direct public story patches enforce content and wrapper locks through reversible edits', () => {
  const source = sourceModel('patch-lock'), content = source.sectionHeaderFooters.first.storyContent.header
  content.locks = [{ id: 'tail', startOffset: 10, endOffset: 14, lock: 'contentLocked' }]
  const { studio, document, body } = open(source), original = copy(document.getModel())
  for (const mutate of [
    patch => { patch.header = patch.header.replace('Tail', 'Fail') },
    patch => { patch.storyContent.header.locks = [] },
    patch => { patch.storyContent.header.runs[0].metadata.source.sourceFormat.propertiesXml = '<w:rPr><w:smallCaps/></w:rPr>' },
  ]) {
    const patch = { header: source.sectionHeaderFooters.first.header, storyContent: { header: copy(readSection(document).storyContent.header) } }
    mutate(patch)
    assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch }, body))
    assert.deepEqual(document.getModel(), original)
  }
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  const patch = storyEditor.replaceStoryRange(readSection(document), 'header', 5, 9, 'longer', document.getModel())
  assert.equal(document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch }, body).changed, true)
  assert.deepEqual(readSection(document).storyContent.header.locks, [{ id: 'tail', startOffset: 12, endOffset: 16, lock: 'contentLocked' }])
  document.undo(); assert.deepEqual(document.getModel(), original)
  document.redo(); assert.equal(readSection(document).header, 'A😀 longer\nTail')
})

test('inactive document story histories and saved buffers remain isolated', () => {
  const first = open(sourceModel('story-first'))
  const secondModel = sourceModel('story-second'); first.studio.createDocument(secondModel)
  const second = new FDocument(first.studio, secondModel.id), secondBody = range(secondModel.blocks[0].id)
  first.studio.activateDocument(second.id)
  first.document.clearFormatting(first.body, { sectionId: 'first', story: 'header', start: 5, end: 9 })
  assert.equal(first.studio.activeDocumentId, second.id)
  assert.deepEqual(second.getModel().sectionHeaderFooters, secondModel.sectionHeaderFooters)
  assert.deepEqual(history(first.studio, first.document.id), { undo: 1, redo: 0, dirty: true })
  assert.deepEqual(history(first.studio, second.id), { undo: 0, redo: 0, dirty: false })
  const patch = storyEditor.replaceStoryRange(readSection(second), 'header', 5, 9, 'other', second.getModel())
  second.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch }, secondBody)
  assert.equal(readSection(first.document).header, 'A😀 word\nTail')
  assert.equal(readSection(second).header, 'A😀 other\nTail')
  first.document.undo(); assert.equal(readSection(second).header, 'A😀 other\nTail')
  second.undo(); assert.equal(readSection(second).header, 'A😀 word\nTail')
})

test('partial code clear and subsequent insertion retain one code source through undo, redo and internal reopen', () => {
  const model = sourceModel('code-story-contract'), block = codeBlocks.createCodeBlock('const value = 1;\nreturn value', 'javascript', false)
  model.blocks = [block]
  const { studio, document } = open(model), before = copy(document.getModel()), selected = range(block.id, 1, 4)
  document.clearFormatting(selected)
  const cleared = copy(document.getModel())
  document.executeCommand({ id: 'doc.command.insert-text', text: 'let' }, selected)
  const edited = copy(document.getModel()), code = edited.blocks[0]
  assert.equal(code.metadata.codeBlock.source, 'clett value = 1;\nreturn value')
  assert.equal(docModel.runText(code.runs), code.metadata.codeBlock.source)
  assert.equal(code.runs.find(run => run.text === 'let').metadata.codeSyntaxHighlight, false)
  assert.equal(code.metadata.codeBlock.lineNumbers, false)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  document.undo(); assert.deepEqual(document.getModel(), cleared)
  document.undo(); assert.deepEqual(document.getModel(), before)
  document.redo(); document.redo(); assert.deepEqual(document.getModel(), edited)
  const reopened = open(copy(document.save()))
  assert.equal(reopened.document.getModel().blocks[0].runs.find(run => run.text === 'let').metadata.codeSyntaxHighlight, false)
  assert.equal(reopened.document.getModel().blocks[0].metadata.codeBlock.source, code.metadata.codeBlock.source)
})
