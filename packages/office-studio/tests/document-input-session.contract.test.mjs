import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Contract drafts for COMMON-03 / DOC-01 / DOC-02. Execution remains deferred.
let server, OfficeStudio, FDocument, domain, docModel, storyModel, sections, storyEditor, inputApi, indexedInput
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  docModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  storyModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story.ts')
  sections = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/section-header-footer.ts')
  storyEditor = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/story-editor.ts')
  inputApi = await server.ssrLoadModule('/packages/office-studio/src/api/document-input-commands.ts')
  indexedInput = await server.ssrLoadModule('/packages/office-studio/src/api/document-indexed-input.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })

const copy = value => JSON.parse(JSON.stringify(value))
const point = (blockId, offset = 0, coordinates = {}) => ({ blockId, offset, ...coordinates })
const selection = (anchor, focus = anchor) => ({ anchor, focus })
const paragraph = (id, text) => ({ ...docModel.createParagraph(text), id })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0,
  redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })
function openDocument(blocks = [paragraph('body', 'AB')], extra = {}) {
  const model = { ...domain.createBlankDocument(), id: 'input-doc', blocks, ...extra, modified: false }
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' })
  studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}
function openStory(text = 'abc', extra = {}) {
  const section = { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: text,
    storyContent: { header: { runs: [docModel.createRun(text)], paragraphs: storyModel.documentStoryParagraphs(text, docModel.defaultParagraphStyle()) } } }
  const opened = openDocument([paragraph('body', 'Unchanged body')], { sectionHeaderFooters: { first: section }, ...extra })
  return { ...opened, body: selection(point('body', 2)), range: { sectionId: 'first', story: 'header', start: 1, end: 1 } }
}
function inputDeclaration(document, caret) {
  const state = document.getInputState()
  assert.equal(state?.valid, true); assert.equal(state.location.kind, 'story')
  return { range: copy(state.location.range), caret, modelRevision: document.getModelRevision(), generation: state.generation }
}
function storyPatch(document, value, caret, range, marks = document.getInputMarks()) {
  const model = document.getModel(), section = sections.resolveSectionHeaderFooter(model, 'first')
  return storyEditor.editStoryBuffer(section, 'header', value, caret, model, marks, range)
}

test('caret formatting changes only the shared session and repeating it is a no-op', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  const original = document.getModel(), saved = copy(original), revision = document.getModelRevision()
  document.setInputSelection(caret)
  const result = document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, fontSize: 18 } }, caret)
  assert.deepEqual([result.changed, result.inputChanged, result.selectionCurrent, result.modelRevision], [false, true, true, revision])
  assert.equal(document.getModel(), original); assert.deepEqual(document.getModel(), saved)
  assert.deepEqual(document.getInputMarks(caret), { bold: true, fontSize: 18 })
  const generation = document.getInputState().generation
  assert.equal(document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, fontSize: 18 } }, caret).inputChanged, false)
  assert.equal(document.getInputState().generation, generation)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('clear intent and nested metadata survive detached session getters', () => {
  const { document } = openDocument(), caret = selection(point('body', 1)), metadata = { vendor: { value: 'kept' } }
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, color: '#FF0000', highlight: 'yellow', metadata } }, caret)
  metadata.vendor.value = 'changed outside'
  document.executeCommand({ id: 'doc.command.format-text', patch: {}, clearFields: ['color', 'highlight'] }, caret)
  const marks = document.getInputMarks(caret)
  assert.equal(Object.hasOwn(marks, 'color'), true); assert.equal(marks.color, undefined)
  assert.equal(Object.hasOwn(marks, 'highlight'), true); assert.equal(marks.highlight, undefined)
  assert.equal(marks.metadata.vendor.value, 'kept')
  marks.metadata.vendor.value = 'getter mutation'
  const state = document.getInputState(); state.location.selection.focus.offset = 0; state.marks.bold = false
  assert.equal(document.getInputMarks(caret).metadata.vendor.value, 'kept')
  assert.equal(document.getInputMarks(caret).bold, true)
})

test('successive body inputs keep overrides while content history and saved runs stay independent', () => {
  const { studio, document } = openDocument([paragraph('body', 'A'), paragraph('untouched', 'Tail')])
  const originalTail = copy(document.getModel().blocks[1]), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true, fontSize: 17, color: '#C81E1E' } }, caret)
  const first = document.executeCommand({ id: 'doc.command.insert-text', text: '中😀' }, caret)
  assert.equal(first.selectionCurrent, true)
  assert.deepEqual(document.getInputState().location.selection, first.selection)
  const second = document.executeCommand({ id: 'doc.command.insert-text', text: '!' }, first.selection)
  assert.equal(document.getText(), 'A中😀!\nTail')
  assert.deepEqual(document.getModel().blocks[1], originalTail)
  assert.equal(document.getInputMarks(second.selection).bold, true)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  const saved = document.save(), reopened = openDocument(saved.blocks, saved)
  assert.deepEqual(reopened.document.getModel().blocks, saved.blocks)
  assert.deepEqual(reopened.document.getInputMarks(), {}); assert.equal(reopened.document.getInputState(), undefined)
  assert.equal(history(studio, document.id).dirty, false)
})

test('indexed facade input consumes CJK font overrides and explicit automatic color clears', () => {
  const block = paragraph('body', '汉'); block.runs[0].highlight = 'yellow'
  const { studio, document } = openDocument([block]), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { fontFamily: '仿宋' }, clearFields: ['color', 'highlight'] }, caret)
  assert.equal(document.insertText(1, '字'), true)
  const inserted = document.getModel().blocks[0].runs.find(run => run.text === '字')
  assert.deepEqual([inserted.fontFamily, inserted.eastAsiaFontFamily, inserted.color, inserted.highlight], ['仿宋', '仿宋', 'auto', undefined])
  assert.equal(Object.hasOwn(document.getInputMarks(), 'color'), true)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
})

test('indexed offsets address paragraph separators and preserve independent table-cell boundaries', () => {
  const { document } = openDocument([paragraph('body', 'AB'), paragraph('next', 'CD')])
  assert.deepEqual(indexedInput.documentSelectionAtBodyRange(document.getModel(), 2, 3), selection(point('body', 2), point('next', 0)))
  assert.equal(document.replaceRange(2, 3, '!'), true); assert.equal(document.getText(), 'AB!CD')
  const table = paragraph('table', ''); table.type = 'table'; table.runs = []; table.metadata.table = docModel.createTable(1, 2)
  const first = table.metadata.table.rows[0].cells[0].blocks[0], second = table.metadata.table.rows[0].cells[1].blocks[0]
  first.runs = [docModel.createRun('A')]; second.runs = [docModel.createRun('B')]
  const other = openDocument([table], { id: 'table-input' }), original = copy(other.document.getModel())
  assert.throws(() => other.document.replaceRange(0, 3, 'X'), /independent table cells/)
  assert.deepEqual(other.document.getModel(), original)
})

test('indexed replacement refuses to remove an unaddressed non-text block', () => {
  const image = { ...paragraph('image', ''), type: 'image', runs: [], metadata: { image: { src: 'data:image/png;base64,AA==', width: 10, height: 10, alt: 'kept' } } }
  const { studio, document } = openDocument([paragraph('body', 'AB'), image, paragraph('next', 'CD')]), original = copy(document.getModel())
  assert.throws(() => document.replaceRange(1, 4, 'X'), /non-text blocks/)
  assert.deepEqual(document.getModel(), original); assert.equal(history(studio, document.id).undo, 0)
})

test('invalid Unicode boundaries, stale offsets and bad format patches preserve the existing session', () => {
  const { studio, document } = openDocument([paragraph('body', 'A😀B')]), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { italic: true } }, caret)
  const original = copy(document.getModel()), state = document.getInputState()
  for (const invalid of [selection(point('body', 2)), selection(point('body', 99)), selection(point('missing', 0))]) {
    assert.throws(() => document.setInputSelection(invalid))
    assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, invalid))
  }
  assert.throws(() => document.insertText(2, 'X'), RangeError)
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-text', patch: { fontSize: 401 } }, caret), RangeError)
  assert.deepEqual(document.getModel(), original); assert.deepEqual(document.getInputState(), state)
  assert.equal(history(studio, document.id).undo, 0)
})

test('caret formatting and indexed text editing enforce the same paragraph permission', () => {
  const { studio, document } = openDocument([paragraph('body', 'AB')], {
    permissionRules: [{ id: 'deny', scope: 'paragraph', targetId: 'body', action: 'edit', effect: 'deny', userIds: [] }],
  }), caret = selection(point('body', 1)), original = copy(document.getModel())
  document.setInputSelection(caret)
  const state = document.getInputState()
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret), /permission denied/)
  assert.throws(() => document.insertText(1, 'X'), /permission denied/)
  assert.deepEqual(document.getInputState(), state); assert.deepEqual(document.getModel(), original)
  assert.equal(history(studio, document.id).undo, 0)
})

test('locked text prevents both pending formatting and actual insertion inside its range', () => {
  const { studio, document } = openDocument([paragraph('body', 'ABCD')], {
    contentControls: [{ id: 'locked', type: 'text', blockId: 'body', startOffset: 1, endOffset: 3, lock: 'contentLocked' }],
  }), caret = selection(point('body', 2)), original = copy(document.getModel())
  document.setInputSelection(caret)
  const state = document.getInputState()
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret), /locked/)
  assert.throws(() => document.insertText(2, 'X'), /locked/)
  assert.deepEqual(document.getInputState(), state); assert.deepEqual(document.getModel(), original)
  assert.equal(history(studio, document.id).undo, 0)
})

test('explicit navigation clears pending overrides even when it reselects the same caret', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  const generation = document.getInputState().generation
  assert.equal(document.setInputSelection(caret), true)
  assert.deepEqual(document.getInputMarks(), {}); assert.ok(document.getInputState().generation > generation)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('unrelated content changes cannot restore stale marks and a later input establishes a fresh session', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  document.mutate('Rename', 'doc.command.input-contract-rename', model => { model.name = 'Changed' })
  document.mutate('Rename again', 'doc.command.input-contract-rename-again', model => { model.name = 'Changed again' })
  assert.deepEqual(document.getInputMarks(), {}); assert.equal(document.advanceInputMarks(caret), false)
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret)
  assert.equal(result.selectionCurrent, true); assert.equal(document.getInputState().valid, true)
  assert.deepEqual(document.getInputState().location.selection, result.selection)
  assert.deepEqual(document.getInputMarks(), {}); assert.equal(document.getModel().blocks[0].runs.find(run => run.text === 'X').bold, false)
  assert.equal(history(studio, document.id).undo, 3)
})

test('canceled input and formatting commands preserve marks, model revision and content history', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  const original = copy(document.getModel()), state = document.getInputState(), revision = document.getModelRevision()
  const cancellation = studio.events.addEvent('BeforeCommandExecute', event => { event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret), /canceled/)
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-text', patch: { bold: false } }, caret), /canceled/)
  assert.throws(() => document.clearInputMarks(), /canceled/)
  cancellation.dispose()
  assert.deepEqual(document.getInputState(), state); assert.deepEqual(document.getModel(), original)
  assert.equal(document.getModelRevision(), revision); assert.equal(history(studio, document.id).undo, 0)
})

test('successful undo and redo clear transient preferences while canceled undo retains them', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  const inserted = document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret)
  document.executeCommand({ id: 'doc.command.format-text', patch: { italic: true } }, inserted.selection)
  const pending = document.getInputState(), original = copy(document.getModel())
  const cancellation = studio.events.addEvent('BeforeUndo', event => { event.cancel = true })
  assert.throws(() => document.undo(), /canceled/); cancellation.dispose()
  assert.deepEqual(document.getInputState(), pending); assert.deepEqual(document.getModel(), original)
  assert.equal(document.undo(), true); assert.deepEqual(document.getInputMarks(), {})
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  assert.equal(document.redo(), true); assert.deepEqual(document.getInputMarks(), {})
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
})

test('document and product switches clear input preferences without adding content history', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  studio.activateWorkbook(studio.activeWorkbookId); assert.deepEqual(document.getInputMarks(), {})
  studio.activateDocument(document.id)
  document.executeCommand({ id: 'doc.command.format-text', patch: { italic: true } }, caret)
  const other = studio.createDocument({ id: 'other-input-doc', text: 'Other' })
  assert.deepEqual(document.getInputMarks(), {}); assert.deepEqual(new FDocument(studio, other.id).getInputMarks(), {})
  assert.equal(history(studio, document.id).undo, 0); assert.equal(history(studio, other.id).undo, 0)
})

test('table Tab publishes the actual public caret and clears input style without a content mutation', () => {
  const table = paragraph('table', ''); table.type = 'table'; table.runs = []; table.metadata.table = docModel.createTable(1, 2)
  const first = table.metadata.table.rows[0].cells[0].blocks[0], second = table.metadata.table.rows[0].cells[1].blocks[0]
  const { studio, document } = openDocument([table]), caret = selection(point('table', 0, { row: 0, col: 0, cellBlockId: first.id }))
  const original = copy(document.getModel())
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  const result = document.executeCommand({ id: 'doc.command.table-tab', direction: 1 }, caret)
  assert.equal(result.changed, false); assert.equal(result.selectionCurrent, true)
  assert.deepEqual(document.getInputState().location.selection, selection(point('table', 0, { row: 0, col: 1, cellBlockId: second.id })))
  assert.deepEqual(document.getInputMarks(), {}); assert.deepEqual(document.getModel(), original)
  assert.equal(history(studio, document.id).undo, 0)
})

test('scoped observers receive detached state and their failures do not reject accepted formatting', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1)), observed = []
  document.onInputStateChanged(event => { event.state.marks.bold = false; event.state.location.selection.focus.offset = 0; throw new Error('observer failed') })
  const subscription = document.onInputStateChanged(event => observed.push(event))
  const result = document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  assert.equal(result.inputChanged, true); assert.equal(result.selectionCurrent, true)
  assert.equal(observed[0].state.marks.bold, true); assert.deepEqual(observed[0].state.location.selection, caret)
  assert.equal(studio.getEventListenerErrors().at(-1).eventName, 'DocumentInputStateChanged')
  subscription.dispose(); document.clearInputMarks(); assert.equal(observed.length, 1)
})

test('nested navigation stops obsolete input-state publication and keeps the newer caret', () => {
  const { document } = openDocument(), old = selection(point('body', 1)), next = selection(point('body', 0)), observed = []
  let navigated = false
  document.onInputStateChanged(event => { if (!navigated && event.state?.marks.bold) { navigated = true; document.setInputSelection(next) } })
  document.onInputStateChanged(event => observed.push(event.state))
  const result = document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, old)
  assert.equal(result.selectionCurrent, false)
  assert.equal(observed.length, 1); assert.deepEqual(observed[0].location.selection, next); assert.deepEqual(observed[0].marks, {})
  assert.deepEqual(document.getInputState().location.selection, next)
})

test('a subscriber can commit a later input after the first history entry and supersede its returned selection', () => {
  const { studio, document } = openDocument([paragraph('body', 'A')]), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  let nested = false
  document.onInputStateChanged(event => {
    if (!nested && event.state?.valid && event.state.modelRevision === 1) {
      nested = true; document.executeCommand({ id: 'doc.command.insert-text', text: 'Y' }, event.state.location.selection)
    }
  })
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret)
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false); assert.equal(document.getText(), 'AXY')
  assert.equal(document.getInputState().modelRevision, 2); assert.equal(document.getInputState().location.selection.focus.offset, 3)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
})

test('content invalidation is observed after its history entry and a later input has its own undo step', () => {
  const { studio, document } = openDocument([paragraph('body', 'A')]), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  let nested = false
  document.onInputStateChanged(event => {
    if (!nested && event.state?.valid === false) {
      nested = true
      assert.equal(studio.editor.history.isApplying, false)
      assert.equal(history(studio, document.id).undo, 1)
      document.insertText(2, 'Y')
    }
  })
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret)
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false)
  assert.equal(document.getText(), 'AXY'); assert.equal(history(studio, document.id).undo, 2)
  document.undo(); assert.equal(document.getText(), 'AX')
  document.undo(); assert.equal(document.getText(), 'A')
  document.redo(); assert.equal(document.getText(), 'AX')
  document.redo(); assert.equal(document.getText(), 'AXY')
})

test('BeforeUndo and BeforeRedo document switches never replay another document history', () => {
  for (const direction of ['undo', 'redo']) {
    const { studio, document } = openDocument()
    document.insertText(1, 'X')
    if (direction === 'redo') document.undo()
    const otherModel = studio.createDocument({ id: 'history-other', text: 'Other' }, false), other = new FDocument(studio, otherModel.id)
    other.insertText(5, '!')
    const original = copy(document.getModel()), otherBefore = copy(other.getModel())
    const firstHistory = history(studio, document.id), secondHistory = history(studio, other.id)
    studio.events.addEvent(direction === 'undo' ? 'BeforeUndo' : 'BeforeRedo', () => studio.activateDocument(other.id))
    assert.equal(document[direction](), false)
    assert.deepEqual(document.getModel(), original); assert.deepEqual(other.getModel(), otherBefore)
    assert.deepEqual(history(studio, document.id), firstHistory); assert.deepEqual(history(studio, other.id), secondHistory)
    assert.equal(studio.activeDocumentId, other.id)
  }
})

test('an edit in BeforeUndo supersedes the captured history entry without undoing the newer edit', () => {
  const { studio, document } = openDocument()
  document.insertText(1, 'X')
  let nested = false
  studio.events.addEvent('BeforeUndo', () => {
    if (!nested) { nested = true; document.insertText(2, 'Y') }
  })
  assert.equal(document.undo(), false)
  assert.equal(document.getText(), 'AXYB'); assert.equal(history(studio, document.id).undo, 2)
  assert.equal(history(studio, document.id).redo, 0)
})

test('document disposal during a committed input leaves its result accepted and selection obsolete', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  document.onInputStateChanged(event => { if (event.state?.valid && event.state.modelRevision === 1) studio.dispose() })
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret)
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false); assert.equal(document.getInputState(), undefined)
})

test('a BeforeCommandExecute content edit rejects the facade input based on the earlier revision', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  let nested = false
  studio.events.addEvent('BeforeCommandExecute', event => {
    if (!nested && event.id === 'doc.command.insert-text') {
      nested = true; document.mutate('Rename', 'doc.command.before-input-rename', model => { model.name = 'Changed by hook' })
    }
  })
  assert.throws(() => document.insertText(1, 'X'), /内容已更改/)
  assert.equal(document.getText(), 'AB'); assert.equal(document.getName(), 'Changed by hook'); assert.equal(history(studio, document.id).undo, 1)
})

test('navigation in BeforeCommandExecute does not replace the new selection with an older input result', () => {
  const { studio, document } = openDocument(), caret = selection(point('body', 1)), next = selection(point('body', 0))
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true } }, caret)
  let navigated = false
  studio.events.addEvent('BeforeCommandExecute', event => { if (!navigated && event.id === 'doc.command.insert-text') { navigated = true; document.setInputSelection(next) } })
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: 'X' }, caret)
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, false); assert.equal(document.getText(), 'AXB')
  assert.equal(document.getModel().blocks[0].runs.find(run => run.text === 'X').bold, true)
  assert.deepEqual(document.getInputState().location.selection, next); assert.deepEqual(document.getInputMarks(), {})
})

test('a checked story input permits exactly one advance to its declared caret and retains body content', () => {
  const { studio, document, body, range } = openStory(), originalBody = copy(document.getModel().blocks)
  document.setInputSelection(body, range)
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true }, story: range }, body)
  const input = inputDeclaration(document, 2), patch = storyPatch(document, 'aXbc', 2, input.range)
  const result = document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch, input }, body)
  assert.equal(result.changed, true); assert.equal(result.selectionCurrent, true); assert.deepEqual(document.getInputMarks(), {})
  assert.equal(document.advanceInputMarks(body, { ...range, start: 3, end: 3 }), false)
  const next = { ...range, start: 2, end: 2 }
  assert.equal(document.advanceInputMarks(body, next), true); assert.equal(document.advanceInputMarks(body, next), false)
  assert.equal(document.getInputMarks(body, next).bold, true); assert.deepEqual(document.getModel().blocks, originalBody)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
})

test('an untagged story patch and a later content edit both invalidate advance authorization', () => {
  const first = openStory(); first.document.setInputSelection(first.body, first.range)
  first.document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true }, story: first.range }, first.body)
  const patch = storyPatch(first.document, 'aXbc', 2, first.range)
  first.document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch }, first.body)
  assert.equal(first.document.advanceInputMarks(first.body, { ...first.range, start: 2, end: 2 }), false)
  const second = openStory(); second.document.setInputSelection(second.body, second.range)
  const input = inputDeclaration(second.document, 2)
  second.document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch: storyPatch(second.document, 'aXbc', 2, input.range), input }, second.body)
  second.document.mutate('Rename', 'doc.command.after-story-rename', model => { model.name = 'Later' })
  assert.equal(second.document.advanceInputMarks(second.body, { ...second.range, start: 2, end: 2 }), false)
})

test('wrong story source generation, revision, target and caret reject the detached commit atomically', () => {
  const { studio, document, body, range } = openStory(); document.setInputSelection(body, range)
  const input = inputDeclaration(document, 2), patch = storyPatch(document, 'aXbc', 2, input.range)
  const original = copy(document.getModel()), state = document.getInputState()
  for (const invalid of [{ ...input, generation: input.generation + 1 }, { ...input, modelRevision: input.modelRevision + 1 },
    { ...input, range: { ...range, story: 'footer' } }, { ...input, caret: 3 }]) {
    assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch, input: invalid }, body))
    assert.deepEqual(document.getModel(), original); assert.deepEqual(document.getInputState(), state)
  }
  assert.equal(history(studio, document.id).undo, 0)
})

test('the source-range splice supports repeated text, same-text replacement and native Unicode deletion', () => {
  for (const [before, after, start, end, caret, expected] of [
    ['aaa', 'aaaa', 1, 1, 2, { start: 1, end: 1, text: 'a' }],
    ['aaa', 'aaa', 0, 2, 2, { start: 0, end: 2, text: 'aa' }],
    ['A😀B', 'AB', 3, 3, 1, { start: 1, end: 3, text: '' }],
    ['A😀B', 'AB', 1, 1, 1, { start: 1, end: 3, text: '' }],
  ]) {
    assert.deepEqual(storyModel.documentStoryInputEdit(before, after, { start, end }, caret), expected)
    assert.doesNotThrow(() => inputApi.assertDocumentStoryInputSplice(before, after, { start, end }, caret))
  }
  assert.throws(() => inputApi.assertDocumentStoryInputSplice('abc', 'abc', { start: 1, end: 1 }, 1), RangeError)
  assert.throws(() => inputApi.assertDocumentStoryInputSplice('abc', 'abXc', { start: 1, end: 1 }, 3), RangeError)
  assert.throws(() => inputApi.assertDocumentStoryInputSplice('A😀B', 'AXB', { start: 2, end: 3 }, 2), RangeError)
})

test('equivalent field replacement advances its expanded session without dirtying document history', () => {
  const { studio, document, body } = openStory('ab')
  document.mutate('Seed field', 'doc.command.seed-input-field', model => {
    const section = model.sectionHeaderFooters.first
    section.pageFields = [{ story: 'header', field: 'PAGE', offset: 1, cachedResult: '1' }]
    Object.assign(section, storyEditor.replaceStoryRange(section, 'header', 1, 2, copy(section.pageFields[0]), model))
  })
  document.save()
  const range = { sectionId: 'first', story: 'header', start: 1, end: 2 }; document.setInputSelection(body, range)
  const input = inputDeclaration(document, 2), model = document.getModel(), section = sections.resolveSectionHeaderFooter(model, 'first')
  const patch = storyEditor.replaceStoryRange(section, 'header', 1, 2, copy(section.pageFields[0]), model)
  const before = copy(model), revision = document.getModelRevision(), baseline = history(studio, document.id)
  const result = document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch, input }, body)
  assert.equal(result.changed, false); assert.equal(result.modelRevision, revision)
  assert.equal(document.advanceInputMarks(body, { ...range, start: 2, end: 2 }), true)
  assert.deepEqual(document.getModel(), before); assert.deepEqual(history(studio, document.id), baseline)
})

test('story input-state listeners can supersede an advance without reviving an older draft', () => {
  const { document, body, range } = openStory(); document.setInputSelection(body, range)
  document.executeCommand({ id: 'doc.command.format-text', patch: { bold: true }, story: range }, body)
  const input = inputDeclaration(document, 2)
  document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch: storyPatch(document, 'aXbc', 2, input.range), input }, body)
  let navigated = false
  document.onInputStateChanged(event => {
    if (!navigated && event.state?.valid && event.state.modelRevision === 1) { navigated = true; document.setInputSelection(body, { ...range, start: 0, end: 0 }) }
  })
  assert.equal(document.advanceInputMarks(body, { ...range, start: 2, end: 2 }), true)
  assert.deepEqual(document.getInputState().location.range, { ...range, start: 0, end: 0 })
  assert.deepEqual(document.getInputMarks(), {}); assert.equal(document.advanceInputMarks(body, { ...range, start: 2, end: 2 }), false)
})
