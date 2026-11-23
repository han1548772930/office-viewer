import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, defaultParagraphStyle, runText } from '../src/features/document/model/document-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { documentStoryContent, documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, documentStoryParagraphs, isDocumentTextBoundary } from '../src/features/document/model/document-story.ts'
import { applyDocumentStoryClearFormatting, documentClearFormattingState, prepareDocumentStoryClearFormatting } from '../src/features/document/services/document-clear-formatting.ts'
import { applySectionHeaderFooter, EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'
import { editStoryBuffer, replaceStoryRange, storyBuffer } from '../src/features/document/services/story-editor.ts'
import { documentStoryDisplay } from '../src/features/document/services/story-rendering.ts'
import { documentStoryCaretRect, hitTestDocumentStory, layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { sceneFromDocument } from '../src/features/document/rendering/document-scene.ts'
import { createCodeBlock } from '../src/features/document/code-blocks/code-blocks.ts'
import { normalizeDocument } from '../src/parser/documentNormalizer.ts'

// Drafts for DOC-02; no execution or acceptance result is claimed for this turn.
const copy = value => JSON.parse(JSON.stringify(value))
const context = model => ({ documentId: model.id, userId: 'editor' })
const bodyRange = model => ({ anchor: { blockId: model.blocks[0].id, offset: 0 }, focus: { blockId: model.blocks[0].id, offset: 1 } })
const sourceFormat = (propertiesXml = '<w:rPr><w:b/><w:caps/><w:spacing w:val="24"/></w:rPr>') => ({ version: 1, propertiesXml, namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } })

function storyModel(text = 'A😀 words\nTail') {
  const model = createBlankDocument()
  model.id = 'story-contract'; model.blocks = [createParagraph('body unchanged', { italic: true })]
  model.defaultStyle.text = { ...model.defaultStyle.text, fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 15, color: 'auto' }
  model.defaultStyle.paragraph = { ...model.defaultStyle.paragraph, align: 'right', lineHeight: 1.2, spaceAfter: 12 }
  const paragraphStyle = { ...defaultParagraphStyle(), align: 'center', indentStart: 20 }
  const run = createRun(text, { bold: true, fontSize: 22, color: '#FF0000', href: 'https://example.org', metadata: { source: { sourceFormat: sourceFormat() } } })
  const paragraphs = documentStoryParagraphs(text, paragraphStyle).map(paragraph => ({ ...paragraph, metadata: { source: { sourceFormat: sourceFormat('<w:pPr><w:jc w:val="center"/></w:pPr>') } } }))
  model.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: text, pageFields: [], storyContent: { header: { runs: [run], paragraphs } } } }
  return model
}

test('story clear uses UTF-16 object offsets, retains the body range, and inserts with cleared formatting', () => {
  const model = storyModel(), section = model.sectionHeaderFooters.first
  section.pageFields = [{ story: 'header', field: 'PAGE', offset: 3, cachedResult: '12', textStyle: { ...documentStoryDefaultStyle(model).text, fontSize: 18 } }]
  const originalBody = copy(model.blocks), originalField = copy(section.pageFields[0]), originalNext = copy(section.storyContent.header.paragraphs[1])
  assert.equal(storyBuffer(section, 'header').value, 'A😀\uFFFC words\nTail')
  const request = { sectionId: 'first', story: 'header', start: 10, end: 5 }, range = bodyRange(model)
  const plan = prepareDocumentStoryClearFormatting(model, request, context(model))
  assert.deepEqual(plan.range, { ...request, start: 5, end: 10 })
  assert.deepEqual(executeDocumentCommand(model, range, { id: 'doc.command.clear-formatting', story: request }, context(model)), range)
  const cleared = resolveSectionHeaderFooter(model, 'first')
  assert.deepEqual(model.blocks, originalBody)
  assert.deepEqual(cleared.pageFields[0], originalField)
  assert.deepEqual(cleared.storyContent.header.paragraphs[1], originalNext)
  assert.deepEqual(cleared.storyContent.header.runs.map(run => [run.text, run.bold, run.fontSize]), [['A😀 ', true, 22], ['words', false, 9], ['\nTail', true, 22]])
  const run = cleared.storyContent.header.runs[1]
  assert.equal(run.fontFamily, 'Georgia'); assert.equal(run.href, 'https://example.org')
  assert.equal(run.metadata.officeStoryTextFormatCleared, true)
  assert.ok(run.metadata.source.sourceFormat.propertiesXml.includes('<w:caps/>'))
  assert.equal(cleared.storyContent.header.paragraphs[0].metadata.officeStoryParagraphFormatCleared, true)
  applySectionHeaderFooter(model, 'first', editStoryBuffer(cleared, 'header', 'A😀\uFFFC new\nTail', 8, model))
  const after = resolveSectionHeaderFooter(model, 'first'), inserted = after.storyContent.header.runs.find(run => run.text === 'new')
  assert.equal(after.header, 'A😀 new\nTail'); assert.equal(after.pageFields[0].offset, 3)
  assert.deepEqual([inserted.bold, inserted.fontFamily, inserted.fontSize, inserted.metadata.officeStoryTextFormatCleared], [false, 'Georgia', 9, true])
})

test('caret word expansion and reversed half-open story ranges leave the next paragraph untouched', () => {
  const model = storyModel('word\nnext'), section = model.sectionHeaderFooters.first
  const next = copy(section.storyContent.header.paragraphs[1])
  assert.deepEqual(applyDocumentStoryClearFormatting(model, { sectionId: 'first', story: 'header', start: 2, end: 2 }, context(model)), { sectionId: 'first', story: 'header', start: 0, end: 4 })
  assert.deepEqual(resolveSectionHeaderFooter(model, 'first').storyContent.header.paragraphs[1], next)
  const second = storyModel('word\nnext'), originalNext = copy(second.sectionHeaderFooters.first.storyContent.header.paragraphs[1])
  applyDocumentStoryClearFormatting(second, { sectionId: 'first', story: 'header', start: 5, end: 1 }, context(second))
  assert.deepEqual(resolveSectionHeaderFooter(second, 'first').storyContent.header.paragraphs[1], originalNext)
})

test('an empty lock at the next paragraph boundary does not block the preceding paragraph, while the final paragraph mark stays protected', () => {
  const model = storyModel('word\nnext'), content = model.sectionHeaderFooters.first.storyContent.header
  content.locks = [{ id: 'next-paragraph', startOffset: 5, endOffset: 5, lock: 'contentLocked' }]
  const next = copy(content.paragraphs[1])
  assert.equal(documentClearFormattingState(model, bodyRange(model), context(model), { sectionId: 'first', story: 'header', start: 1, end: 3 }).available, true)
  applyDocumentStoryClearFormatting(model, { sectionId: 'first', story: 'header', start: 1, end: 3 }, context(model))
  assert.deepEqual(resolveSectionHeaderFooter(model, 'first').storyContent.header.paragraphs[1], next)
  assert.deepEqual(resolveSectionHeaderFooter(model, 'first').storyContent.header.locks, content.locks)
  const last = storyModel('word')
  last.sectionHeaderFooters.first.storyContent.header.locks = [{ id: 'final-mark', startOffset: 4, endOffset: 4, lock: 'contentLocked' }]
  const before = copy(last), request = { sectionId: 'first', story: 'header', start: 1, end: 3 }
  assert.equal(documentClearFormattingState(last, bodyRange(last), context(last), request).available, false)
  assert.throws(() => applyDocumentStoryClearFormatting(last, request, context(last)), /已锁定/)
  assert.deepEqual(last, before)
})

test('permissions, linked sections, out-of-range offsets and outside-word content locks reject atomically', () => {
  for (const kind of ['entity', 'section', 'content-lock', 'invalid-surrogate', 'unknown-story', 'linked']) {
    const model = storyModel(), request = { sectionId: 'first', story: 'header', start: 4, end: 9 }
    if (kind === 'entity' || kind === 'section') model.permissionRules = [{ id: kind, scope: kind, targetId: kind === 'entity' ? 'first:header' : 'first', action: 'edit', effect: 'deny', userIds: [] }]
    if (kind === 'content-lock') model.sectionHeaderFooters.first.storyContent.header.locks = [{ id: 'emoji', startOffset: 1, endOffset: 3, lock: 'contentLocked' }]
    if (kind === 'invalid-surrogate') request.start = 2
    if (kind === 'unknown-story') request.story = 'missing'
    if (kind === 'linked') {
      model.blocks.push({ ...createParagraph(), id: 'next-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'continuous' } })
      model.sectionHeaderFooters['next-section'] = { ...EMPTY_SECTION_HEADER_FOOTER, linkToPrevious: true }
      request.sectionId = 'next-section'
    }
    const before = copy(model)
    assert.equal(documentClearFormattingState(model, bodyRange(model), context(model), request).available, false, kind)
    assert.throws(() => applyDocumentStoryClearFormatting(model, request, context(model)), undefined, kind)
    assert.deepEqual(model, before, kind)
  }
})

test('direct patches retain locks under equivalent metadata key order and reject text or source-style changes', () => {
  const model = storyModel('ABCD'), content = model.sectionHeaderFooters.first.storyContent.header
  content.locks = [{ id: 'middle', startOffset: 1, endOffset: 3, lock: 'contentLocked' }]
  const equivalent = copy(content), source = equivalent.runs[0].metadata.source.sourceFormat
  equivalent.runs[0].metadata.source.sourceFormat = { namespaces: source.namespaces, propertiesXml: source.propertiesXml, version: 1 }
  applySectionHeaderFooter(model, 'first', { storyContent: { header: equivalent } })
  for (const change of [
    patch => { patch.header = 'AXCD' },
    patch => { patch.storyContent.header.runs[0].bold = false },
    patch => { patch.storyContent.header.runs[0].metadata.source.sourceFormat.propertiesXml = '<w:rPr><w:b/><w:smallCaps/></w:rPr>' },
    patch => { patch.storyContent.header.locks = [] },
  ]) {
    const before = copy(model), patch = { storyContent: { header: copy(resolveSectionHeaderFooter(model, 'first').storyContent.header) } }
    change(patch)
    assert.throws(() => applySectionHeaderFooter(model, 'first', patch))
    assert.deepEqual(model, before)
  }
})

test('edits before content locks transform anchors while wrapper locks allow replacement and preserve identity', () => {
  const model = storyModel('ABCD'), section = model.sectionHeaderFooters.first
  section.storyContent.header.locks = [{ id: 'tail', startOffset: 2, endOffset: 4, lock: 'contentLocked' }]
  applySectionHeaderFooter(model, 'first', replaceStoryRange(section, 'header', 0, 1, 'long', model))
  assert.equal(resolveSectionHeaderFooter(model, 'first').header, 'longBCD')
  assert.deepEqual(resolveSectionHeaderFooter(model, 'first').storyContent.header.locks, [{ id: 'tail', startOffset: 5, endOffset: 7, lock: 'contentLocked' }])
  const wrapper = storyModel('ABCD'), wrapped = wrapper.sectionHeaderFooters.first
  wrapped.storyContent.header.locks = [{ id: 'wrapper', startOffset: 1, endOffset: 3, lock: 'sdtLocked' }]
  applySectionHeaderFooter(wrapper, 'first', replaceStoryRange(wrapped, 'header', 1, 3, '', wrapper))
  assert.deepEqual(resolveSectionHeaderFooter(wrapper, 'first').storyContent.header.locks, [{ id: 'wrapper', startOffset: 1, endOffset: 1, lock: 'sdtLocked' }])
})

function importedFieldModel() {
  const run = (text, family, size, color, bold = false) => ({ text, style: { fontFamilyAscii: family, fontSizePt: size, color, bold }, sourceFormat: sourceFormat() })
  return normalizeDocument({ schemaVersion: 1, parserVersion: 'draft', format: 'docx', fileName: 'field.docx', size: 1, model: 'document', diagnostics: {}, payload: { modelType: 'document', data: {
    documentId: 'field-model', blocks: [], sectionLayout: [{ headerSections: [{ referenceType: 'default', storyText: {
      text: '', runs: [run('', 'Georgia', 24, '#112233')], paragraphs: [{ startOffset: 0, endOffset: 0, style: {}, sourceFormat: sourceFormat('<w:pPr/>') }],
      fields: [{ field: 'DATE', offset: 0, instruction: 'DATE', cachedResult: 'ABCD', textStyle: { fontFamilyAscii: 'Georgia', fontSizePt: 18, color: '#FF0000', bold: false }, sourceFormat: sourceFormat(),
        cachedRuns: [run('AB', 'Georgia', 18, '#FF0000'), { ...run('CD', 'Aptos', 12, '#0000FF', true), link: 'https://example.org/cache' }] }],
      sourceXml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>',
    } }] }],
  } } })
}

test('normalization preserves mixed field-cache runs, overlays changed properties and retains one input object', () => {
  const model = importedFieldModel(), section = model.sectionHeaderFooters.first, field = section.pageFields[0]
  const cache = documentStoryFieldCache(field)
  assert.deepEqual(cache.runs.map(run => [run.text, run.fontFamily, run.fontSize, run.bold]), [['AB', 'Georgia', 18, false], ['CD', 'Aptos', 12, true]])
  field.textStyle.color = '#008000'
  assert.deepEqual(documentStoryFieldDisplayRuns(field, field.cachedResult, documentStoryDefaultStyle(model).text).map(run => [run.color, run.fontFamily, run.fontSize]), [['#008000', 'Georgia', 18], ['#008000', 'Aptos', 12]])
  assert.deepEqual(cache.runs.map(run => run.color), ['#FF0000', '#0000FF'])
  const display = documentStoryDisplay(section, 'header', 1, 2, model)
  assert.equal(storyBuffer(section, 'header').value, '\uFFFC'); assert.equal(display.text, 'ABCD')
  assert.deepEqual(display.inputRange(1, 3), { start: 0, end: 1 })
  assert.deepEqual(display.inputRange(2, 2), { start: 0, end: 0 })
  assert.deepEqual(display.inputRange(4, 4), { start: 1, end: 1 })
  applyDocumentStoryClearFormatting(model, { sectionId: 'first', story: 'header', start: 0, end: 1 }, context(model))
  const clearedField = resolveSectionHeaderFooter(model, 'first').pageFields[0]
  assert.equal(clearedField.metadata.officeStoryTextFormatCleared, true)
  const clearedCache = documentStoryFieldCache(clearedField).runs
  assert.deepEqual(clearedCache.map(run => [run.text, run.bold, run.fontSize, run.metadata.officeStoryTextFormatCleared]), [['AB', false, 9, true], ['CD', false, 9, true]])
  assert.equal(clearedCache[1].metadata.link, 'https://example.org/cache')
  const displayed = documentStoryFieldDisplayRuns(clearedField, 'ABCD', documentStoryDefaultStyle(model).text)
  assert.deepEqual(displayed.map(run => [run.text, run.bold, run.fontSize]), [['AB', false, 9], ['CD', false, 9]])
  assert.equal(displayed[1].metadata.link, 'https://example.org/cache')
  const invalid = copy(clearedField); invalid.cachedResult = 'different'
  assert.throws(() => documentStoryFieldCache(invalid), /缓存文本/)
})

test('field clear resets every cached run when the primary style already equals the defaults, retaining links and empty entities', () => {
  const model = storyModel(''), section = model.sectionHeaderFooters.first, defaults = documentStoryDefaultStyle(model).text
  const runs = [createRun('AB', defaults), createRun('CD', { ...defaults, bold: true, href: 'https://example.org/cache' }),
    createRun('', { ...defaults, metadata: { entityId: 'empty-cache-entity' } })]
  section.pageFields = [{ story: 'header', field: 'DATE', offset: 0, cachedResult: 'ABCD', instruction: 'DATE', textStyle: { ...defaults },
    metadata: { officeStoryStyleBaseline: { ...defaults }, officeStoryFieldCache: { version: 1, runs } } }]
  const range = { sectionId: 'first', story: 'header', start: 0, end: 1 }
  applyDocumentStoryClearFormatting(model, range, context(model))
  const field = resolveSectionHeaderFooter(model, 'first').pageFields[0], cache = documentStoryFieldCache(field).runs
  assert.deepEqual(cache.map(run => [run.id, run.text, run.bold, run.fontSize]), runs.map(run => [run.id, run.text, false, 9]))
  assert.equal(cache[1].href, 'https://example.org/cache')
  assert.equal(cache[2].metadata.entityId, 'empty-cache-entity')
  assert.deepEqual(documentStoryFieldDisplayRuns(field, 'ABCD', defaults).map(run => run.bold), [false, false, false])
  const before = copy(model)
  applyDocumentStoryClearFormatting(model, range, context(model))
  assert.deepEqual(model, before)
})

test('a wrapped field has only object boundaries and its trailing caret follows the final visual fragment', () => {
  const model = importedFieldModel()
  model.page = { ...model.page, width: 200, marginLeft: 95, marginRight: 95 }
  const layout = layoutDocument(model, { measure: text => Array.from(text).length * 10 }), lines = layout.pageHeaderFooters[0].headerLayoutLines
  assert.ok(lines.length > 1)
  assert.ok(lines.flatMap(line => line.glyphs).every(glyph => glyph.offset === 0 && glyph.end === 1))
  assert.equal(documentStoryCaretRect(lines, 0).y, lines[0].y)
  assert.equal(documentStoryCaretRect(lines, 1).y, lines.at(-1).y)
  for (const line of lines) {
    assert.equal(hitTestDocumentStory(lines, line.glyphs[0].x, line.y + line.height / 2), 0)
    const last = line.glyphs.at(-1)
    assert.equal(hitTestDocumentStory(lines, last.x + last.width - 0.01, line.y + line.height / 2), 1)
  }
})

test('empty story paragraphs inherit the insertion font and collapsed display ranges do not expand', () => {
  const model = storyModel(''), content = documentStoryContent(model.sectionHeaderFooters.first, 'header', model)
  content.runs[0].fontSize = 30; content.paragraphs[0].style.lineHeight = 1
  const display = documentStoryDisplay(model.sectionHeaderFooters.first, 'header', 1, 1, model)
  assert.deepEqual(display.inputRange(0, 0), { start: 0, end: 0 })
  assert.equal(layoutDocument(model, { measure: () => 10 }).pageHeaderFooters[0].headerLayoutLines[0].height, 40)
  applySectionHeaderFooter(model, 'first', editStoryBuffer(model.sectionHeaderFooters.first, 'header', '😀', 2, model))
  assert.equal(resolveSectionHeaderFooter(model, 'first').storyContent.header.runs[0].fontSize, 30)
  assert.equal(isDocumentTextBoundary('😀', 1), false)
})

test('partial code clear survives insertion and Canvas coloring while preserving the code source and line-number setting', () => {
  const model = createBlankDocument(), block = createCodeBlock('const value = 1;\nreturn value', 'javascript', false)
  model.blocks = [block]
  const range = { anchor: { blockId: block.id, offset: 1 }, focus: { blockId: block.id, offset: 4 } }
  executeDocumentCommand(model, range, { id: 'doc.command.clear-formatting' })
  assert.deepEqual(block.runs.map(run => [run.text, run.metadata?.codeSyntaxHighlight]), [['c', undefined], ['ons', false], ['t value = 1;\nreturn value', undefined]])
  const selection = executeDocumentCommand(model, range, { id: 'doc.command.insert-text', text: 'let' })
  assert.equal(block.metadata.codeBlock.source, 'clett value = 1;\nreturn value')
  assert.equal(runText(block.runs), block.metadata.codeBlock.source)
  assert.equal(block.runs.find(run => run.text === 'let').metadata.codeSyntaxHighlight, false)
  const layout = layoutDocument(model, { measure: () => 8 }), scene = sceneFromDocument(model, layout, { width: 1000, height: 1200, scrollLeft: 0, scrollTop: 0, zoom: 100 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  const text = scene.commands.filter(command => command.kind === 'text')
  assert.equal(text.some(command => command.text === '1' && command.align === 'right'), false)
  assert.ok(text.some(command => command.text === 'return' || command.text === 'r'))
  assert.ok(text.filter(command => ['l', 'e', 't'].includes(command.text)).some(command => command.color === '#1f2937'))
  block.metadata.codeBlock.lineNumbers = true
  const numbered = sceneFromDocument(model, layoutDocument(model, { measure: () => 8 }), { width: 1000, height: 1200, scrollLeft: 0, scrollTop: 0, zoom: 100 }, { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
  assert.ok(numbered.commands.some(command => command.kind === 'text' && command.text === '1' && command.align === 'right'))
})
