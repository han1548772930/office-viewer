import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, runText } from '../src/features/document/model/document-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentRunStyleAt, formatDocumentRunRange, replaceDocumentRunRange, splitDocumentRunRange } from '../src/features/document/model/document-run-range.ts'
import { documentStoryContent, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, documentStoryInputEdit, documentStoryParagraphs, documentStoryTextStyle } from '../src/features/document/model/document-story.ts'
import { documentStoryFontPatch, documentStoryTextFormattingState, prepareDocumentStoryTextFormatting } from '../src/features/document/services/document-story-formatting.ts'
import { documentTextFormattingState } from '../src/features/document/services/document-text-formatting.ts'
import { applySectionHeaderFooter, EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'
import { editStoryBuffer, replaceStoryRange, storyBuffer, storyRunStyleAt } from '../src/features/document/services/story-editor.ts'
import { documentStoryToolbarState, documentToolbarState } from '../src/components/document/document-toolbar-state.ts'

// DOC-02 / COMMON-03 drafts, 2026-11-22 Asia/Shanghai. No execution is claimed.
const copy = value => JSON.parse(JSON.stringify(value))
const context = model => ({ documentId: model.id, userId: 'editor' })
const range = (blockId, start, end = start) => ({ anchor: { blockId, offset: start }, focus: { blockId, offset: end } })
const storyRange = (start, end = start) => ({ sectionId: 'first', story: 'header', start, end })
const source = { sourceFormat: { version: 1, propertiesXml: '<w:rPr><w:caps/></w:rPr>', namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } } }

function bodyModel(id = 'partial-body') {
  const model = createBlankDocument(), first = createParagraph(), second = createParagraph('untouched', { fontFamily: 'Verdana', fontSize: 10 })
  model.id = id; first.id = 'first-body'; second.id = 'second-body'
  first.runs = [
    createRun('ab', { fontFamily: 'Georgia', fontSize: 18, bold: true, italic: true, underline: true, color: '#AA0000', highlight: 'yellow', href: 'https://example.org/first', metadata: { source: copy(source), entityId: 'first-source' } }),
    createRun('', { metadata: { noteReference: { kind: 'footnote', id: 'note-1' }, entityId: 'empty-note' } }),
    createRun('cd', { fontFamily: 'Aptos', fontSize: 12, strike: true, color: '#0000AA', metadata: { source: copy(source), entityId: 'last-source' } }),
  ]
  model.blocks = [first, second]
  return model
}

function storyModel(id = 'partial-story') {
  const model = bodyModel(id), text = 'A😀BC', defaults = documentStoryDefaultStyle(model).text
  const first = createRun('中', { fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 18, color: '#AA0000', metadata: { source: copy(source) } })
  const last = createRun('Z', { fontFamily: 'Aptos', fontSize: 12, bold: true, color: '#0000AA', href: 'https://example.org/cache', metadata: { source: copy(source) } })
  const empty = createRun('', { fontFamily: 'Verdana', fontSize: 28, color: '#00AA00', metadata: { entityId: 'empty-cache' } })
  const textStyle = documentStoryTextStyle(first, defaults)
  model.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: text,
    storyContent: { header: { runs: [createRun('A😀', { fontFamily: 'Arial', fontSize: 24 }), createRun('BC', { italic: true })], paragraphs: documentStoryParagraphs(text, documentStoryDefaultStyle(model).paragraph) } },
    pageFields: [{ story: 'header', field: 'DATE', offset: 3, instruction: 'DATE', cachedResult: '中Z', textStyle,
      metadata: { source: copy(source), officeStoryStyleBaseline: { ...textStyle }, officeStoryFieldCache: { version: 1, runs: [first, last, empty] } } }],
  } }
  return model
}

test('half-open family changes preserve the other marks, links, source records, empty notes and untouched identities', () => {
  const model = bodyModel(), before = copy(model), block = model.blocks[0]
  executeDocumentCommand(model, range(block.id, 1, 3), { id: 'doc.command.format-text', patch: { fontFamily: 'Tahoma' } }, context(model))
  assert.deepEqual(block.runs.map(run => run.text), ['a', 'b', '', 'c', 'd'])
  assert.deepEqual(block.runs.map(run => run.fontFamily), ['Georgia', 'Tahoma', before.blocks[0].runs[1].fontFamily, 'Tahoma', 'Aptos'])
  assert.deepEqual(block.runs[2], before.blocks[0].runs[1])
  assert.equal(block.runs[0].id, before.blocks[0].runs[0].id)
  assert.deepEqual(model.blocks[1], before.blocks[1])
  for (const [index, original] of [[1, before.blocks[0].runs[0]], [3, before.blocks[0].runs[2]]]) {
    const { id: _id, text: _text, fontFamily: _family, ...style } = block.runs[index]
    const { id: _oldId, text: _oldText, fontFamily: _oldFamily, ...expected } = original
    assert.deepEqual(style, expected)
  }
  const uniform = documentToolbarState(model, range(block.id, 1, 3))
  assert.equal(uniform.marks.fontFamily, 'Tahoma'); assert.ok(uniform.mixedMarks.includes('fontSize'))
})

test('a size-only change and no-op formatting retain unrelated attributes and do not regenerate unaffected runs', () => {
  const runs = [createRun('A中', { fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 18, bold: true, color: '#123456', href: 'https://example.org' }), createRun('tail')]
  const changed = formatDocumentRunRange(runs, 0, 2, { fontSize: 24 })
  assert.deepEqual(changed[0], { ...runs[0], fontSize: 24 })
  assert.equal(changed[1], runs[1])
  assert.deepEqual(formatDocumentRunRange(changed, 0, 2, { fontSize: 24 }), changed)
  const model = bodyModel(), before = copy(model)
  executeDocumentCommand(model, range(model.blocks[0].id, 1), { id: 'doc.command.format-text', patch: { fontSize: 26 } }, context(model))
  assert.deepEqual(model, before)
})

test('reversed cross-paragraph formatting excludes the paragraph at the half-open end and preflights every target', () => {
  const model = bodyModel(), first = model.blocks[0], second = model.blocks[1], untouched = copy(second)
  const selected = { anchor: { blockId: second.id, offset: 0 }, focus: { blockId: first.id, offset: 1 } }
  executeDocumentCommand(model, selected, { id: 'doc.command.format-text', patch: { fontSize: 24 } }, context(model))
  assert.deepEqual(second, untouched)
  assert.deepEqual(first.runs.filter(run => run.text).map(run => [run.text, run.fontSize]), [['a', 18], ['b', 24], ['cd', 24]])
  const invalid = bodyModel(); invalid.blocks[1].runs = [createRun('😀')]
  const before = copy(invalid), bad = { anchor: { blockId: invalid.blocks[0].id, offset: 1 }, focus: { blockId: invalid.blocks[1].id, offset: 1 } }
  assert.throws(() => executeDocumentCommand(invalid, bad, { id: 'doc.command.format-text', patch: { fontSize: 20 } }, context(invalid)), /UTF-16/)
  assert.deepEqual(invalid, before)
})

test('malformed character patches reject before any selected run changes', () => {
  for (const patch of [{ fontSize: NaN }, { fontSize: 401 }, { fontSize: 0 }, { fontFamily: '' }, { bold: undefined }, { unknownMark: true }, { color: '#GG0000' }, { highlight: 'invalid-color' }]) {
    const model = bodyModel(), before = copy(model)
    assert.throws(() => executeDocumentCommand(model, range(model.blocks[0].id, 1, 3), { id: 'doc.command.format-text', patch }, context(model)))
    assert.deepEqual(model, before)
  }
})

test('text splitting assigns boundary entities once, preserves outside entities and inherits the first selected character', () => {
  const left = createRun('A', { fontFamily: 'Georgia', metadata: { entityId: 'left' } }), middle = createRun('BC', { fontSize: 24 }), right = createRun('D', { italic: true })
  const atStart = createRun('', { metadata: { entityId: 'start-boundary' } }), inside = createRun('', { metadata: { entityId: 'inside' } }), atEnd = createRun('', { metadata: { entityId: 'end-boundary' } })
  const runs = [left, atStart, createRun('B', { fontSize: 24 }), inside, createRun('C', { fontSize: 24 }), atEnd, right]
  const split = splitDocumentRunRange(runs, 1, 3)
  assert.deepEqual(split.before, [left, atStart]); assert.deepEqual(split.after, [atEnd, right])
  const changed = replaceDocumentRunRange(runs, 1, 3, 'x', createBlankDocument().defaultStyle.text)
  assert.equal(runText(changed), 'AxD'); assert.equal(changed[2].fontSize, 24)
  assert.equal(changed.filter(run => run.id === atStart.id).length, 1); assert.equal(changed.filter(run => run.id === atEnd.id).length, 1)
  assert.equal(changed.some(run => run.id === inside.id), false)
  assert.deepEqual(documentRunStyleAt([left, middle, right], 1, 3).fontSize, 24)
  assert.throws(() => replaceDocumentRunRange(runs, 1, 3, 'x', createBlankDocument().defaultStyle.text, { fontSize: 401 }), /字号/)
})

test('multiline body input preserves the untouched prefix, suffix, empty references and captured formatting', () => {
  const model = bodyModel(), block = model.blocks[0], before = copy(model), entityId = block.runs[1].id
  const caret = executeDocumentCommand(model, range(block.id, 2, 3), { id: 'doc.command.insert-text', text: 'X\nY', marks: { fontSize: 26, color: undefined, highlight: undefined } }, context(model))
  assert.deepEqual(model.blocks.slice(0, 2).map(block => runText(block.runs)), ['abX', 'Yd'])
  assert.equal(model.blocks[0].runs[0].id, before.blocks[0].runs[0].id)
  assert.equal(model.blocks.flatMap(block => block.runs).filter(run => run.id === entityId).length, 1)
  for (const run of model.blocks.slice(0, 2).flatMap(block => block.runs).filter(run => run.text === 'X' || run.text === 'Y')) {
    assert.equal(run.fontSize, 26); assert.equal(run.color, 'auto'); assert.equal(run.highlight, undefined); assert.equal(run.strike, true)
  }
  assert.equal(caret.focus.blockId, model.blocks[1].id); assert.equal(caret.focus.offset, 1)
  assert.deepEqual(model.blocks[2], before.blocks[1])
})

test('content locks and paragraph permissions govern both character formatting and final text insertion', () => {
  const model = bodyModel(), block = model.blocks[0]
  model.contentControls = [{ id: 'text-lock', type: 'richText', blockId: block.id, startOffset: 1, endOffset: 3, lock: 'contentLocked' }]
  const before = copy(model)
  assert.equal(documentTextFormattingState(model, range(block.id, 1, 3), context(model)).available, false)
  for (const command of [{ id: 'doc.command.format-text', patch: { fontSize: 24 } }, { id: 'doc.command.insert-text', text: 'x', marks: { fontSize: 24 } }]) {
    assert.throws(() => executeDocumentCommand(model, range(block.id, 2), command, context(model)))
    assert.deepEqual(model, before)
  }
  const denied = bodyModel(); denied.permissionRules = [{ id: 'deny', scope: 'paragraph', targetId: denied.blocks[0].id, action: 'edit', effect: 'deny', userIds: [] }]
  assert.equal(documentTextFormattingState(denied, range(denied.blocks[0].id, 1, 3), context(denied)).available, false)
})

test('story toolbar reads the selected field caches and its font changes retain mixed sizes, colors, links and empty entities', () => {
  const model = storyModel(), original = copy(model), selected = storyRange(3, 4)
  assert.equal(storyBuffer(model.sectionHeaderFooters.first, 'header').value, 'A😀\uFFFCBC')
  const state = documentStoryToolbarState(model, selected)
  assert.ok(state.mixedMarks.includes('fontFamily')); assert.ok(state.mixedMarks.includes('fontSize')); assert.ok(state.mixedMarks.includes('bold'))
  assert.equal(state.marks.fontFamily, undefined); assert.equal(state.marks.fontSize, undefined)
  executeDocumentCommand(model, range(model.blocks[0].id, 1), { id: 'doc.command.format-text', story: selected, patch: { fontFamily: 'Tahoma' } }, context(model))
  const section = resolveSectionHeaderFooter(model, 'first'), field = section.pageFields[0], cache = documentStoryFieldCache(field).runs
  assert.deepEqual(model.blocks, original.blocks); assert.deepEqual(section.storyContent, original.sectionHeaderFooters.first.storyContent)
  assert.deepEqual(cache.map(run => [run.fontFamily, run.eastAsiaFontFamily]), [['Tahoma', 'Tahoma'], ['Tahoma', 'Tahoma'], ['Tahoma', 'Tahoma']])
  const oldCache = original.sectionHeaderFooters.first.pageFields[0].metadata.officeStoryFieldCache.runs
  assert.deepEqual(cache.map(run => [run.id, run.text, run.fontSize, run.bold, run.color, run.href]), oldCache.map(run => [run.id, run.text, run.fontSize, run.bold, run.color, run.href]))
  assert.equal(cache[2].metadata.entityId, 'empty-cache')
  executeDocumentCommand(model, range(model.blocks[0].id, 1), { id: 'doc.command.format-text', story: selected, patch: { fontSize: 20 } }, context(model))
  const sized = documentStoryFieldCache(resolveSectionHeaderFooter(model, 'first').pageFields[0]).runs
  assert.deepEqual(sized.map(run => [run.fontSize, run.bold, run.color]), [[20, false, '#AA0000'], [20, true, '#0000AA'], [20, false, '#00AA00']])
  assert.equal(prepareDocumentStoryTextFormatting(model, selected, { fontSize: 20 }, context(model)).patch, undefined)
})

test('field replacement inherits its first cache character while the trailing caret inherits the last nonempty character', () => {
  const model = storyModel(), section = model.sectionHeaderFooters.first
  section.pageFields[0].metadata.officeStoryFieldCache.runs.unshift(createRun('', { fontFamily: 'Verdana', fontSize: 32, metadata: { entityId: 'leading-empty-cache' } }))
  const first = storyRunStyleAt(section, 'header', 3, 4, model), last = storyRunStyleAt(section, 'header', 4, 4, model)
  assert.deepEqual([first.fontFamily, first.eastAsiaFontFamily, first.fontSize, first.bold], ['Georgia', '宋体', 18, false])
  assert.deepEqual([last.fontFamily, last.fontSize, last.bold, last.color], ['Aptos', 12, true, '#0000AA'])
  assert.equal(first.href, undefined); assert.equal(last.href, undefined); assert.equal(last.metadata, undefined)
  assert.deepEqual(documentStoryFontPatch(model, storyRange(4), { fontFamily: 'Tahoma' }), { fontFamily: 'Tahoma' })
  applySectionHeaderFooter(model, 'first', replaceStoryRange(section, 'header', 3, 4, '新', model, { fontSize: 26, color: undefined }))
  const changed = resolveSectionHeaderFooter(model, 'first'), inserted = documentStoryContent(changed, 'header', model).runs.find(run => run.text === '新')
  assert.equal(changed.header, 'A😀新BC'); assert.deepEqual(changed.pageFields, [])
  assert.deepEqual([inserted.fontFamily, inserted.eastAsiaFontFamily, inserted.fontSize, inserted.bold, inserted.color, inserted.href], ['Georgia', '宋体', 26, false, 'auto', undefined])
})

test('materialized clear markers preserve later mixed-cache values and optional clears do not revive document defaults', () => {
  const model = storyModel(), section = model.sectionHeaderFooters.first, field = section.pageFields[0]
  field.metadata.officeStoryTextFormatCleared = true
  const original = copy(documentStoryFieldDisplayRuns(field, field.cachedResult, documentStoryDefaultStyle(model).text))
  executeDocumentCommand(model, range(model.blocks[0].id, 0), { id: 'doc.command.format-text', story: storyRange(3, 4), patch: { fontSize: 21 } }, context(model))
  const after = resolveSectionHeaderFooter(model, 'first').pageFields[0], display = documentStoryFieldDisplayRuns(after, after.cachedResult, documentStoryDefaultStyle(model).text)
  assert.equal(after.metadata.officeStoryTextFormatCleared, true)
  assert.deepEqual(display.map(run => [run.bold, run.color, run.href]), original.map(run => [run.bold, run.color, run.href]))
  model.defaultStyle.text = { ...model.defaultStyle.text, eastAsiaFontFamily: '宋体', highlight: 'yellow', verticalAlign: 'superscript' }
  executeDocumentCommand(model, range(model.blocks[0].id, 0), { id: 'doc.command.format-text', story: storyRange(3, 4), patch: { highlight: undefined, verticalAlign: 'baseline' } }, context(model))
  const cleared = resolveSectionHeaderFooter(model, 'first'), inherited = storyRunStyleAt(cleared, 'header', 4, 4, model)
  assert.equal(inherited.highlight, undefined); assert.equal(inherited.verticalAlign, undefined)
  const insertion = replaceStoryRange(cleared, 'header', 4, 4, { field: 'PAGE', cachedResult: '1' }, model)
  const insertedField = insertion.pageFields.find(candidate => candidate.field === 'PAGE')
  assert.equal(insertedField.textStyle.highlight, undefined); assert.equal(insertedField.textStyle.verticalAlign, undefined)
})

test('field input preserves an imported size and deliberately absent optional slots without treating them as new toolbar overrides', () => {
  const model = storyModel(), section = model.sectionHeaderFooters.first, field = section.pageFields[0]
  const cache = field.metadata.officeStoryFieldCache.runs
  for (const run of cache) { run.fontSize = 800; delete run.eastAsiaFontFamily; delete run.highlight; delete run.verticalAlign }
  field.textStyle = { ...field.textStyle, fontSize: 800 }
  delete field.textStyle.eastAsiaFontFamily
  field.metadata.officeStoryStyleBaseline = { ...field.textStyle }
  model.defaultStyle.text = { ...model.defaultStyle.text, eastAsiaFontFamily: '宋体', highlight: 'yellow', verticalAlign: 'superscript' }
  const inherited = storyRunStyleAt(section, 'header', 4, 4, model)
  assert.equal(inherited.fontSize, 800); assert.equal(inherited.eastAsiaFontFamily, undefined)
  const patch = replaceStoryRange(section, 'header', 4, 4, 'X', model)
  const inserted = patch.storyContent.header.runs.find(run => run.text === 'X')
  assert.equal(inserted.fontSize, 800); assert.equal(inserted.eastAsiaFontFamily, undefined)
  assert.equal(inserted.highlight, undefined); assert.equal(inserted.verticalAlign, undefined)
  const page = replaceStoryRange(section, 'header', 4, 4, { field: 'PAGE', cachedResult: '1' }, model).pageFields.find(item => item.field === 'PAGE')
  assert.equal(page.textStyle.fontSize, 800); assert.equal(page.textStyle.eastAsiaFontFamily, undefined)
  assert.throws(() => replaceStoryRange(section, 'header', 4, 4, 'X', model, { fontSize: 401 }), /字号/)
})

test('captured native ranges resolve repeated replacements, same-text formatting and UTF-16 deletion exactly', () => {
  assert.deepEqual(documentStoryInputEdit('aaa', 'aaaa', { start: 0, end: 2 }, 3), { start: 0, end: 2, text: 'aaa' })
  assert.deepEqual(documentStoryInputEdit('aaa', 'aaa', { start: 0, end: 3 }, 3), { start: 0, end: 3, text: 'aaa' })
  assert.deepEqual(documentStoryInputEdit('A😀B', 'AB', { start: 3, end: 3 }, 1), { start: 1, end: 3, text: '' })
  assert.deepEqual(documentStoryInputEdit('A😀B', 'AB', { start: 1, end: 1 }, 1), { start: 1, end: 3, text: '' })
  assert.throws(() => documentStoryInputEdit('A😀B', 'AB', { start: 2, end: 2 }, 1), /光标/)
  assert.throws(() => documentStoryInputEdit('abc', 'Xbc', { start: 2, end: 3 }, 3), /提交前/)
  const model = storyModel(), section = model.sectionHeaderFooters.first
  section.header = 'aaa'; section.pageFields = []; section.storyContent.header = { runs: [createRun('a', { fontFamily: 'Georgia', fontSize: 24, bold: true }), createRun('aa', { fontFamily: 'Aptos', fontSize: 12 })], paragraphs: documentStoryParagraphs('aaa', documentStoryDefaultStyle(model).paragraph) }
  const patch = editStoryBuffer(section, 'header', 'aaa', 3, model, {}, { start: 0, end: 3 })
  assert.deepEqual(patch.storyContent.header.runs.filter(run => run.text).map(run => [run.text, run.fontFamily, run.fontSize, run.bold]), [['aaa', 'Georgia', 24, true]])
})

test('story font formatting rejects linked, locked, unauthorized, unknown and surrogate-split selections atomically', () => {
  for (const reason of ['lock', 'entity', 'surrogate', 'unknown', 'field-metadata', 'field-source-clear', 'field-link-clear']) {
    const model = storyModel(), selected = storyRange(3, 4), patch = { fontSize: 24 }
    if (reason === 'lock') model.sectionHeaderFooters.first.storyContent.header.locks = [{ id: 'field-lock', startOffset: 3, endOffset: 4, lock: 'contentLocked' }]
    if (reason === 'entity') model.permissionRules = [{ id: 'deny', scope: 'entity', targetId: 'first:header', action: 'edit', effect: 'deny', userIds: [] }]
    if (reason === 'surrogate') selected.start = 2
    if (reason === 'unknown') selected.story = 'missing'
    if (reason === 'field-metadata') patch.metadata = { discarded: true }
    if (reason === 'field-source-clear') patch.metadata = undefined
    if (reason === 'field-link-clear') patch.href = undefined
    const before = copy(model)
    if (!reason.startsWith('field-')) assert.equal(documentStoryTextFormattingState(model, selected, context(model)).available, false)
    assert.throws(() => prepareDocumentStoryTextFormatting(model, selected, patch, context(model)))
    assert.deepEqual(model, before)
  }
  const linked = storyModel(); linked.blocks.push({ ...createParagraph(), id: 'linked', type: 'sectionBreak', runs: [], sectionBreak: { type: 'continuous' } })
  linked.sectionHeaderFooters.linked = { ...EMPTY_SECTION_HEADER_FOOTER, linkToPrevious: true }
  assert.equal(documentStoryTextFormattingState(linked, { sectionId: 'linked', story: 'header', start: 0, end: 1 }, context(linked)).available, false)
})

let server, OfficeStudio, FDocument
const studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
function open(model) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(model.id)
  return { studio, document: new FDocument(studio, model.id) }
}
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })

test('real facade body and field formatting retain their models through one history and internal save/reopen', () => {
  const { studio, document } = open(storyModel('partial-history')), before = copy(document.getModel()), selected = range('first-body', 1, 3)
  assert.equal(document.executeCommand({ id: 'doc.command.format-text', patch: { fontFamily: 'Tahoma' } }, selected).changed, true)
  const afterBody = copy(document.getModel())
  assert.equal(document.executeCommand({ id: 'doc.command.format-text', story: storyRange(3, 4), patch: { fontSize: 20 } }, selected).changed, true)
  const afterStory = copy(document.getModel())
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), afterBody)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), afterStory)
  const saved = copy(document.save()), reopened = open(saved)
  assert.deepEqual(reopened.document.getModel().blocks, saved.blocks)
  assert.deepEqual(reopened.document.getModel().sectionHeaderFooters, saved.sectionHeaderFooters)
  assert.equal(reopened.document.getInputState(), undefined)
  assert.deepEqual(history(reopened.studio, saved.id), { undo: 0, redo: 0, dirty: false })
})

test('real facade formatting cancellation leaves content, history and the other document unchanged', () => {
  const { studio, document } = open(storyModel('partial-cancel')), otherModel = bodyModel('other-document')
  studio.createDocument(otherModel)
  const other = new FDocument(studio, otherModel.id), otherBefore = copy(other.getModel()), before = copy(document.getModel())
  const stop = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.format-text') event.cancel = true })
  assert.throws(() => document.executeCommand({ id: 'doc.command.format-text', story: storyRange(3, 4), patch: { fontSize: 20 } }, range('first-body', 1)), /canceled/)
  stop.dispose()
  assert.deepEqual(document.getModel(), before); assert.deepEqual(other.getModel(), otherBefore)
  assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})
