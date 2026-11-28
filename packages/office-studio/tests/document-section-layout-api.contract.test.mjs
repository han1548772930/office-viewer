import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred public-command/history/layout drafts, 2026-11-22. Not executed.
let server, OfficeStudio, FDocument, domain, docModel, layoutModule, settings, sections, storyModel, storyEditor, codec
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  docModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  layoutModule = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
  settings = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-settings.ts')
  sections = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/section-header-footer.ts')
  storyModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story.ts')
  storyEditor = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/story-editor.ts')
  codec = await server.ssrLoadModule('/packages/office-studio/src/api/document-body-codec.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value))
const measurer = { measure: text => Array.from(text).length * 10 }
const range = blockId => ({ anchor: { blockId, offset: 0 }, focus: { blockId, offset: 0 } })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })
function sourceModel(id = 'section-layout-api') {
  const model = domain.createBlankDocument(), first = docModel.createParagraph('A'.repeat(300)), last = docModel.createParagraph('Second section')
  model.id = id
  const section = { ...docModel.createParagraph(), id: id + '-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  for (const block of [first, last]) Object.assign(block.style, { spaceBefore: 0, spaceAfter: 0, lineHeight: 20, spacingRule: 'exact', wordWrap: true })
  model.page = { width: 320, height: 360, marginLeft: 30, marginRight: 30, marginTop: 30, marginBottom: 30 }
  model.blocks = [first, section, last]
  model.wordCount = first.runs[0].text.length + last.runs[0].text.replace(/\s/g, '').length
  return model
}
function open(model = sourceModel()) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id), body: range(model.blocks[0].id), sectionId: model.blocks[1].id }
}
const layout = document => layoutModule.layoutDocument(document.getModel(), measurer)
const geometry = document => layout(document).pages.map(page => ({ width: page.width, height: page.height, pageNumber: page.pageNumber, sectionId: page.sectionId, columns: page.sections.map(section => section.columns.length), blank: page.isParityBlank }))
const pageCommand = (page, flavor = 1, pageOrient = page.width > page.height ? 1 : 0) => ({ id: 'docs.command.page-setup', pageSize: { width: page.width, height: page.height },
  pageOrient, documentFlavor: flavor, marginTop: page.marginTop, marginBottom: page.marginBottom, marginLeft: page.marginLeft, marginRight: page.marginRight })
function table(id, widths) {
  const data = docModel.createTable(1, widths.length)
  data.columnWidths = [...widths]; data.style = { width: widths.reduce((sum, width) => sum + width, 0) }
  return { ...docModel.createParagraph(), id, type: 'table', runs: [], metadata: { table: data } }
}
function openPublic(snapshot) {
  const studio = new OfficeStudio({ currentUserId: 'editor' })
  studios.add(studio)
  const model = studio.createDocument(copy(snapshot))
  return { studio, document: new FDocument(studio, model.id) }
}

test('section commands change real geometry, share one history and persist through internal save/reopen', () => {
  const { studio, document, body, sectionId } = open(), original = copy(document.getModel()), originalGeometry = geometry(document)
  const first = document.executeCommand({ id: 'doc.command.set-section-setup', sectionId: 'first', patch: { columnCount: 2, columnGap: 20, columnSeparator: 'line' } }, body)
  assert.equal(first.changed, true); assert.deepEqual(first.selection, body)
  assert.deepEqual(layout(document).pages[0].sections[0].columns.map(column => column.width), [120, 120])
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  const firstGeometry = geometry(document)
  document.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch: { width: 620, height: 480, orientation: 'landscape', marginLeft: 50, pageNumberStart: 7 } }, body)
  const edited = copy(document.getModel()), editedGeometry = geometry(document)
  assert.equal(layout(document).pages.at(-1).width, 620)
  assert.equal(layout(document).pages.at(-1).pageNumber, 7)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(geometry(document), firstGeometry)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), original); assert.deepEqual(geometry(document), originalGeometry)
  assert.equal(document.redo(), true); assert.equal(document.redo(), true)
  assert.deepEqual(document.getModel(), edited)
  const saved = copy(document.saveModel()), reopened = open(saved)
  assert.deepEqual(geometry(reopened.document), editedGeometry)
  assert.deepEqual(reopened.document.getModel().sectionSetup, saved.sectionSetup)
  assert.deepEqual(reopened.document.getModel().blocks[1].sectionBreak, saved.blocks[1].sectionBreak)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: false })
  assert.deepEqual(history(reopened.studio, saved.id), { undo: 0, redo: 0, dirty: false })
})

test('invalid effective geometry, permission denial and canceled commands publish neither model nor history', () => {
  const { studio, document, body, sectionId } = open(), before = copy(document.getModel()), beforeGeometry = geometry(document)
  for (const patch of [{ width: 40 }, { marginTop: 180, marginBottom: 180 }, { columnCount: 12, columnGap: 1000 }]) {
    assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch }, body))
    assert.deepEqual(document.getModel(), before)
    assert.deepEqual(geometry(document), beforeGeometry)
    assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  }
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.set-section-setup') event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-setup', sectionId, patch: { columnCount: 2 } }, body))
  cancel.dispose()
  assert.deepEqual(document.getModel(), before)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  const blockedModel = sourceModel('section-layout-denied')
  blockedModel.permissionRules = [{ id: 'deny-section', scope: 'section', targetId: blockedModel.blocks[1].id, action: 'edit', effect: 'deny', userIds: [] }]
  const blocked = open(blockedModel), blockedBefore = copy(blocked.document.getModel())
  assert.throws(() => blocked.document.executeCommand({ id: 'doc.command.set-section-setup', sectionId: blocked.sectionId, patch: { pageNumberStart: 8 } }, blocked.body))
  assert.deepEqual(blocked.document.getModel(), blockedBefore)
  assert.deepEqual(history(blocked.studio, blocked.document.id), { undo: 0, redo: 0, dirty: false })
})

test('layout edits and history remain scoped to the target document when another document is active', () => {
  const first = open(), secondModel = sourceModel('section-layout-other')
  first.studio.createDocument(secondModel); first.studio.activateDocument(secondModel.id)
  const second = new FDocument(first.studio, secondModel.id), secondGeometry = geometry(second)
  first.document.executeCommand({ id: 'doc.command.set-section-setup', sectionId: 'first', patch: { columnCount: 3, columnGap: 10 } }, first.body)
  assert.equal(first.studio.activeDocumentId, second.id)
  assert.deepEqual(geometry(second), secondGeometry)
  assert.deepEqual(history(first.studio, second.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(first.document.undo(), true)
  assert.deepEqual(geometry(second), secondGeometry)
  assert.deepEqual(history(first.studio, first.document.id), { undo: 0, redo: 1, dirty: false })
})

test('page setup uses one canonical transaction for all explicit sections and page-fill tables, repeats as a no-op and preserves internal and public reopening', () => {
  const source = sourceModel('canonical-page-setup')
  source.sectionSetup = { ...source.page, columnCount: 2, columnGap: 20, pageNumberStart: 4 }
  source.blocks[1].sectionBreak.setup = { ...source.page, width: 620, height: 480, orientation: 'landscape', pageNumberStart: 7 }
  source.blocks.push(table('page-fill', [65, 195]), table('fixed-width', [50, 70]))
  const { studio, document, body } = open(source), before = copy(document.getModel()), events = []
  studio.events.addEvent('CommandExecuted', event => { if (event.id === 'docs.command.page-setup') events.push([event.id, event.unitId, event.kind]) })
  document.setInputSelection(body)
  const request = pageCommand({ width: 900, height: 700, marginTop: 40, marginBottom: 50, marginLeft: 60, marginRight: 80 })
  const result = document.executeCommand(request, body), edited = copy(document.getModel())
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, true); assert.deepEqual(result.selection, body)
  assert.ok(layout(document).pages.every(page => page.width === 900 && page.height === 700))
  assert.deepEqual(edited.blocks[3].metadata.table.columnWidths, [190, 570])
  assert.deepEqual(edited.blocks[4], before.blocks[4])
  assert.equal(edited.sectionSetup.columnCount, 2); assert.equal(edited.sectionSetup.pageNumberStart, 4)
  assert.equal(edited.blocks[1].sectionBreak.setup.pageNumberStart, 7)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.equal(document.executeCommand(request, body).changed, false)
  assert.deepEqual(document.getModel(), edited); assert.equal(history(studio, document.id).undo, 1)
  assert.deepEqual(events, Array.from({ length: 2 }, () => ['docs.command.page-setup', document.id, 'document']))
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), edited)
  const savedModel = document.saveModel(), internal = open(copy(savedModel))
  assert.deepEqual(internal.document.getModel().page, savedModel.page)
  assert.deepEqual(internal.document.getModel().sectionSetup, savedModel.sectionSetup)
  const snapshot = document.getSnapshot(), saved = document.save(), publicReopened = openPublic(saved)
  assert.deepEqual(saved, snapshot)
  assert.deepEqual(publicReopened.document.getModel().page, edited.page)
  assert.deepEqual(publicReopened.document.getModel().blocks.filter(block => block.type === 'table').map(block => block.metadata.table.columnWidths), [[190, 570], [50, 70]])
  assert.ok(layout(publicReopened.document).pages.every(page => page.width === 900 && page.height === 700))
})

test('the public flat page DTO selects the target document and shares the facade history, no-op and input marks', () => {
  const { studio, document, body } = open(sourceModel('flat-page-setup'))
  document.setInputSelection(body)
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, fontSize: 18 } }, body)
  const input = document.getInputState(), { id, ...params } = pageCommand({ ...document.getModel().page, width: 500 })
  assert.equal(studio.executeCommandSync(id, params), true)
  assert.deepEqual(document.getInputState().location, input.location)
  assert.deepEqual(document.getInputMarks(), input.marks)
  assert.equal(history(studio, document.id).undo, 1)
  assert.equal(studio.executeCommandSync(id, params), true)
  assert.equal(history(studio, document.id).undo, 1)
  const other = sourceModel('flat-page-other')
  studio.createDocument(other); studio.activateDocument(other.id)
  assert.equal(studio.executeCommandSync(id, { ...params, unitId: document.id, pageSize: { width: 550, height: params.pageSize.height } }), true)
  assert.equal(document.getModel().page.width, 550); assert.equal(studio.getDocumentModel(other.id).page.width, other.page.width)
  assert.equal(history(studio, other.id).undo, 0)
  assert.equal(studio.executeCommandSync(id, null), false)
  assert.equal(studio.executeCommandSync(id, { ...params, unitId: 'missing-page-document' }), false)
})

test('modern public inputs retain section source geometry while page settings and two public reopenings preserve the default page and orientation', () => {
  const source = sourceModel('modern-page-setup')
  source.sectionSetup = { width: 420, height: 620, orientation: 'portrait', marginLeft: 35, marginRight: 45 }
  const modernContent = 960 - settings.DOCUMENT_MODERN_PAGE_MARGIN * 2
  source.blocks.push(table('modern-fill', [modernContent / 4, modernContent * 3 / 4]))
  const original = open(source).document.getSnapshot()
  original.documentStyle = { ...original.documentStyle, documentFlavor: 2, pageSize: { width: 960, height: 1124 }, pageOrient: 0,
    marginTop: 20, marginBottom: 30, marginLeft: 40, marginRight: 50 }
  const { document, studio } = openPublic(original), before = copy(document.getModel()), body = range(before.blocks[0].id)
  assert.equal(before.page.width, 960); assert.equal(before.sectionSetup.width, 420)
  assert.equal(layout(document).pages[0].width, 960)
  const request = pageCommand({ ...before.page, width: 1123, height: 1124 }, 2, 1)
  document.setInputSelection(body)
  assert.equal(document.executeCommand(request, body).changed, true)
  const current = copy(document.getModel()), width = 1123 - settings.DOCUMENT_MODERN_PAGE_MARGIN * 2
  assert.deepEqual(current.sectionSetup, before.sectionSetup)
  assert.deepEqual(current.blocks[1].sectionBreak, before.blocks[1].sectionBreak)
  assert.deepEqual(current.blocks.at(-1).metadata.table.columnWidths, [width / 4, width * 3 / 4])
  assert.equal(current.metadata.documentPageOrientation, 'landscape')
  assert.equal(history(studio, document.id).undo, 1)
  const saved = document.save(), reopened = openPublic(saved)
  assert.deepEqual(saved.documentStyle.pageSize, { width: 1123, height: 1124 }); assert.equal(saved.documentStyle.pageOrient, 1)
  assert.deepEqual(saved.body.sectionBreaks[0].pageSize, { width: 420, height: 620 })
  assert.equal(reopened.document.getModel().page.width, 1123); assert.equal(reopened.document.getModel().sectionSetup.width, 420)
  assert.equal(reopened.document.getModel().metadata.documentPageOrientation, 'landscape')
  const second = openPublic(reopened.document.save())
  assert.equal(layout(second.document).pages[0].width, 1123)
  assert.equal(second.document.getModel().sectionSetup.width, 420)
})

test('invalid, canceled and stale page requests leave the target page and its content history unchanged', () => {
  const { studio, document, body } = open(sourceModel('page-setup-rejection')), before = copy(document.getModel())
  const request = pageCommand({ ...before.page, width: 500 })
  assert.throws(() => document.executeCommand({ ...request, marginLeft: 250, marginRight: 250 }, body))
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
  const cancel = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === request.id) event.cancel = true })
  assert.throws(() => document.executeCommand(request, body), /canceled/)
  cancel.dispose()
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
  const revision = document.getModelRevision(), generation = document.getHistoryGeneration()
  studio.editor.history.clear(document.id)
  const { id, ...params } = request
  assert.throws(() => studio.executeCommandSync(id, { ...params, unitId: document.id, selection: body,
    expectedRevision: revision, expectedHistoryGeneration: generation }))
  assert.deepEqual(document.getModel(), before); assert.equal(history(studio, document.id).undo, 0)
  const observer = studio.events.addEvent('BeforeCommandExecute', event => {
    if (event.id === id) document.mutate('Observer edit', 'page-setup-observer-edit', model => { model.name = 'Changed before page setup' })
  })
  assert.throws(() => document.executeCommand(request, body))
  observer.dispose()
  assert.deepEqual(document.getModel().page, before.page)
  assert.equal(document.getModel().name, 'Changed before page setup'); assert.equal(history(studio, document.id).undo, 1)
})

test('page setup after auxiliary and checked story input preserves the active input location and typing marks in one additional history entry', () => {
  const auxiliary = open(sourceModel('page-setup-note-input')), paragraph = { ...docModel.createParagraph('Note'), id: 'page-note-p' }
  const noteId = auxiliary.document.insertFootnote(0, codec.encodeDocumentBody({ blocks: [paragraph], runs: [], paragraphs: [], metadata: {} }))
  assert.ok(noteId)
  const host = { kind: 'note', noteKind: 'footnote', noteId }, caret = range(paragraph.id)
  auxiliary.document.setInputSelection(caret, undefined, host)
  auxiliary.document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true }, host }, caret)
  const input = auxiliary.document.executeCommand({ id: 'doc.command.insert-text', text: '中', host }, caret)
  const state = auxiliary.document.getInputState(), before = copy(auxiliary.document.getModel()), depth = history(auxiliary.studio, auxiliary.document.id).undo
  const result = auxiliary.document.executeCommand(pageCommand({ ...before.page, width: 500 }), input.selection)
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, true); assert.deepEqual(result.host, host)
  assert.deepEqual(result.selection, input.selection)
  assert.deepEqual(auxiliary.document.getInputState().location, state.location)
  assert.deepEqual(auxiliary.document.getInputMarks(), state.marks)
  assert.deepEqual(auxiliary.document.getModel().footnotes, before.footnotes)
  assert.equal(history(auxiliary.studio, auxiliary.document.id).undo, depth + 1)
  const source = sourceModel('page-setup-story-input'), text = 'abc'
  source.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: text,
    storyContent: { header: { runs: [docModel.createRun(text)], paragraphs: storyModel.documentStoryParagraphs(text, docModel.defaultParagraphStyle()) } } } }
  const story = open(source), originalRange = { sectionId: 'first', story: 'header', start: 1, end: 1 }
  story.document.setInputSelection(story.body, originalRange)
  story.document.executeCommand({ id: 'doc.command.format-text', story: originalRange, patch: { italic: true } }, story.body)
  const session = story.document.getInputState(), declaration = { range: copy(session.location.range), caret: 2,
    modelRevision: story.document.getModelRevision(), generation: session.generation }, model = story.document.getModel()
  const patch = storyEditor.editStoryBuffer(sections.resolveSectionHeaderFooter(model, 'first'), 'header', 'aXbc', 2, model, session.marks, declaration.range)
  story.document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch, input: declaration }, story.body)
  assert.equal(story.document.advanceInputMarks(story.body, { ...originalRange, start: 2, end: 2 }), true)
  const storyState = story.document.getInputState(), storyBefore = copy(story.document.getModel()), storyDepth = history(story.studio, story.document.id).undo
  assert.equal(story.document.executeCommand(pageCommand({ ...storyBefore.page, width: 500 }), story.body).selectionCurrent, true)
  assert.deepEqual(story.document.getInputState().location, storyState.location)
  assert.deepEqual(story.document.getInputMarks(), storyState.marks)
  assert.deepEqual(story.document.getModel().sectionHeaderFooters, storyBefore.sectionHeaderFooters)
  assert.equal(history(story.studio, story.document.id).undo, storyDepth + 1)
})

test('flat explicit body coordinates cannot inherit an active footnote host, while omitted coordinates retain that host', () => {
  for (const explicitBody of [false, true]) {
    const { studio, document, body } = open(sourceModel('flat-page-host-' + explicitBody))
    const paragraph = { ...docModel.createParagraph('Note text'), id: 'flat-page-note-p' }
    const noteId = document.insertFootnote(0, codec.encodeDocumentBody({ blocks: [paragraph], runs: [], paragraphs: [], metadata: {} }))
    const host = { kind: 'note', noteKind: 'footnote', noteId }, caret = range(paragraph.id)
    document.setInputSelection(caret, undefined, host)
    document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true }, host }, caret)
    const state = copy(document.getInputState()), before = copy(document.getModel()), depth = history(studio, document.id).undo
    const { id, ...params } = pageCommand({ ...before.page, width: 500 })
    assert.equal(studio.executeCommandSync(id, { ...params, unitId: document.id, ...(explicitBody ? { selection: body } : {}) }), true)
    assert.equal(document.getModel().page.width, 500)
    assert.deepEqual(document.getModel().footnotes, before.footnotes)
    assert.equal(history(studio, document.id).undo, depth + 1)
    const current = document.getInputState()
    assert.deepEqual(current.location, state.location)
    assert.equal(current.valid, !explicitBody)
    assert.deepEqual(document.getInputMarks(), explicitBody ? {} : state.marks)
    assert.ok(current.location.selection.anchor.blockId); assert.ok(current.location.selection.focus.blockId)
  }
})

test('flat explicit host coordinates remain independent of the current body selection and reject malformed hosts atomically', () => {
  const { studio, document, body } = open(sourceModel('flat-page-explicit-host'))
  const paragraph = { ...docModel.createParagraph('Auxiliary'), id: 'flat-page-host-p' }
  const noteId = document.insertFootnote(0, codec.encodeDocumentBody({ blocks: [paragraph], runs: [], paragraphs: [], metadata: {} }))
  const host = { kind: 'note', noteKind: 'footnote', noteId }, caret = range(paragraph.id)
  document.setInputSelection(body)
  const before = copy(document.getModel()), depth = history(studio, document.id).undo, { id, ...params } = pageCommand({ ...before.page, width: 500 })
  assert.equal(studio.executeCommandSync(id, { ...params, host, selection: caret }), true)
  assert.equal(document.getModel().page.width, 500); assert.deepEqual(document.getModel().footnotes, before.footnotes)
  assert.equal(history(studio, document.id).undo, depth + 1)
  const current = copy(document.getModel()), currentDepth = history(studio, document.id).undo
  for (const invalid of [null, { kind: 'note', noteKind: 'footnote', noteId: '' }, { kind: 'body' }]) {
    assert.throws(() => studio.executeCommandSync(id, { ...params, host: invalid, selection: caret }))
    assert.deepEqual(document.getModel(), current); assert.equal(history(studio, document.id).undo, currentDepth)
  }
  for (const selectionHost of ['body', false, { kind: 'body' }, { kind: 'story', sectionId: 'first', story: 'unknown' }]) {
    assert.throws(() => studio.executeCommandSync(id, { unitId: document.id, command: { ...params, id }, selection: body, selectionHost }))
    assert.deepEqual(document.getModel(), current); assert.equal(history(studio, document.id).undo, currentDepth)
  }
})

test('flat default page requests retain a valid legacy header input, while an explicit body selection cannot inherit that story', () => {
  for (const explicitBody of [false, true]) {
    const source = sourceModel('flat-page-header-' + explicitBody), text = 'Header text'
    source.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: text,
      storyContent: { header: { runs: [docModel.createRun(text)], paragraphs: storyModel.documentStoryParagraphs(text, docModel.defaultParagraphStyle()) } } } }
    const { studio, document, body } = open(source), story = { sectionId: 'first', story: 'header', start: 2, end: 2 }
    document.setInputSelection(body, story)
    document.executeCommand({ id: 'doc.command.format-text', story, patch: { italic: true } }, body)
    const before = copy(document.getModel()), state = copy(document.getInputState()), depth = history(studio, document.id).undo
    const { id, ...params } = pageCommand({ ...before.page, width: 500 })
    assert.equal(studio.executeCommandSync(id, { ...params, ...(explicitBody ? { selection: body } : {}) }), true)
    assert.equal(document.getModel().page.width, 500); assert.deepEqual(document.getModel().sectionHeaderFooters, before.sectionHeaderFooters)
    assert.equal(history(studio, document.id).undo, depth + 1)
    assert.deepEqual(document.getInputState().location, state.location)
    assert.equal(document.getInputState().valid, !explicitBody)
    assert.deepEqual(document.getInputMarks(), explicitBody ? {} : state.marks)
  }
})

test('page setup with an emptied auxiliary block host preserves nonempty stale coordinates and cannot revive its input session', () => {
  const { studio, document } = open(sourceModel('page-empty-auxiliary'))
  const paragraph = { ...docModel.createParagraph('Removed note paragraph'), id: 'page-empty-note-p' }
  const noteId = document.insertFootnote(0, codec.encodeDocumentBody({ blocks: [paragraph], runs: [], paragraphs: [], metadata: {} }))
  const host = { kind: 'note', noteKind: 'footnote', noteId }, caret = range(paragraph.id)
  document.setInputSelection(caret, undefined, host)
  document.mutate('Empty auxiliary content', 'page-test-empty-note', model => {
    model.footnotes[noteId] = { ...model.footnotes[noteId], text: '', content: { blocks: [], runs: [], paragraphs: [], metadata: {} } }
  })
  const state = copy(document.getInputState()), before = copy(document.getModel()), depth = history(studio, document.id).undo
  assert.equal(state.valid, false); assert.deepEqual(state.location.selection, caret)
  const published = []
  const subscription = document.onInputStateChanged(event => published.push(copy(event)))
  const result = document.executeCommand({ ...pageCommand({ ...before.page, width: 500 }), host }, caret)
  subscription.dispose()
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false)
  assert.ok(result.selection.anchor.blockId); assert.ok(result.selection.focus.blockId)
  assert.equal(document.getModel().page.width, 500); assert.deepEqual(document.getModel().footnotes, before.footnotes)
  assert.equal(history(studio, document.id).undo, depth + 1)
  const current = document.getInputState()
  assert.equal(current.valid, false); assert.deepEqual(current.location.selection, caret)
  assert.ok(published.every(event => event.state?.location.kind !== 'body'
    || event.state.location.selection.anchor.blockId && event.state.location.selection.focus.blockId))
})
