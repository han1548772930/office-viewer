import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred public-command/history/save drafts, 2026-11-22. Not executed.
let server, OfficeStudio, FDocument, domain, docModel, stories, sections, editor, layoutEngine, fieldService
const studios = new Set(), copy = value => JSON.parse(JSON.stringify(value))
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  docModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  stories = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story.ts')
  sections = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/section-header-footer.ts')
  editor = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/story-editor.ts')
  layoutEngine = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
  fieldService = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-story-fields.ts')
})
afterEach(() => { for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const history = (studio, id) => ({ undo: studio.editor.history.units[id]?.undo.length ?? 0, redo: studio.editor.history.units[id]?.redo.length ?? 0, dirty: studio.editor.history.isDirty(id) })
function sourceModel(id = 'story-links-api') {
  const model = domain.createBlankDocument(), defaults = stories.documentStoryDefaultStyle(model)
  model.id = id
  const first = docModel.createParagraph('Body one'), last = docModel.createParagraph('Body two')
  const boundary = { ...docModel.createParagraph(), id: id + '-section', type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  model.blocks = [first, boundary, last]
  const section = { ...sections.EMPTY_SECTION_HEADER_FOOTER, differentFirstPage: true, differentOddEven: true, storyContent: {}, pageFields: [] }
  for (const story of stories.DOCUMENT_STORY_KEYS) {
    section[story] = story + '😀'
    section.storyContent[story] = { runs: [docModel.createRun(section[story], { ...defaults.text, bold: true, color: '#123456' })],
      paragraphs: stories.documentStoryParagraphs(section[story], defaults.paragraph), metadata: { retained: { values: [1, 2] } } }
    section.pageFields.push({ story, field: 'PAGE', offset: 0, cachedResult: '7', metadata: { officeStoryFieldCache: { version: 1, runs: [
      docModel.createRun('7', { ...defaults.text, underline: true, color: '#654321' }),
    ] } } })
  }
  model.sectionHeaderFooters = { first: section }
  return model
}
function open(model = sourceModel()) {
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(model.id)
  const document = new FDocument(studio, model.id), point = { blockId: model.blocks[0].id, offset: 0 }
  return { studio, document, body: { anchor: point, focus: point }, sectionId: model.blocks[1].id }
}
const resolved = (document, id) => sections.resolveSectionHeaderFooter(document.getModel(), id)
const set = (opened, patch) => opened.document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: opened.sectionId, patch }, opened.body)
const storySnapshot = (section, story) => ({ text: section[story], content: section.storyContent?.[story], fields: section.pageFields?.filter(field => field.story === story) ?? [] })

for (const story of ['header', 'footer', 'firstPageHeader', 'firstPageFooter', 'evenPageHeader', 'evenPageFooter']) test(`public ${story} association shares content history and survives internal save/reopen`, () => {
  const opened = open(), { studio, document, sectionId } = opened, original = copy(document.getModel())
  const inherited = storySnapshot(resolved(document, sectionId), story)
  assert.equal(set(opened, { storyLinkToPrevious: { [story]: false } }).changed, true)
  assert.deepEqual(storySnapshot(resolved(document, sectionId), story), inherited)
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  const unlinked = copy(document.getModel())
  const section = resolved(document, sectionId), patch = editor.replaceStoryRange(section, story, 1, 1, 'Own', document.getModel())
  assert.equal(set(opened, patch).changed, true)
  const edited = copy(document.getModel()), own = storySnapshot(resolved(document, sectionId), story)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), unlinked)
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), original)
  assert.equal(document.redo(), true); assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), edited)
  const saved = copy(document.save()), reopened = open(saved)
  assert.deepEqual(storySnapshot(resolved(reopened.document, sectionId), story), own)
  for (const key of stories.DOCUMENT_STORY_KEYS) assert.equal(sections.sectionStoryLinkedToPrevious(resolved(reopened.document, sectionId), key), key !== story)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: false })
  assert.deepEqual(history(reopened.studio, reopened.document.id), { undo: 0, redo: 0, dirty: false })
})

test('custom DATE/TIME pictures refresh one story through public content history and internal save/reopen', () => {
  const source = sourceModel('custom-date-history'), defaults = stories.documentStoryDefaultStyle(source)
  const originalSection = source.sectionHeaderFooters.first
  originalSection.header = ' / '
  originalSection.storyContent.header = { runs: [docModel.createRun(' / ', defaults.text)], paragraphs: stories.documentStoryParagraphs(' / ', defaults.paragraph) }
  const dateField = { story: 'header', field: 'DATE', offset: 0, instruction: 'DATE \\@ "yyyy/MM/dd" \\*MERGEFORMAT', cachedResult: '2026/11/21', metadata: {
    officeStoryFieldCache: { version: 1, runs: [
      docModel.createRun('2026/', { ...defaults.text, color: '#aa0000', bold: true }),
      docModel.createRun('11/21', { ...defaults.text, color: '#0000aa', italic: true, href: 'https://example.org/date' }),
    ] },
  } }
  const timeField = { story: 'header', field: 'TIME', offset: 3, instruction: 'TIME \\@ "hh:mm:ss AM/PM" \\*Lower', cachedResult: '12:59:58 pm', metadata: {
    officeStoryFieldCache: { version: 1, runs: [
      docModel.createRun('12:', { ...defaults.text, color: '#00aa00', bold: true }),
      docModel.createRun('59:58 pm', { ...defaults.text, color: '#aa00aa', italic: true, href: 'https://example.org/time' }),
    ] },
  } }
  const unknown = { story: 'header', field: 'DATE', offset: 3, instruction: 'DATE \\h', cachedResult: 'retained source' }
  originalSection.pageFields = [...originalSection.pageFields.filter(field => field.story !== 'header'), dateField, timeField, unknown]
  const opened = open(source), { studio, document, sectionId } = opened
  assert.equal(set(opened, { storyLinkToPrevious: { header: false } }).changed, true)
  const before = copy(document.getModel()), previous = resolved(document, sectionId), sourceStory = storySnapshot(resolved(document, 'first'), 'header')
  const input = editor.storyBuffer(previous, 'header').value, footer = storySnapshot(previous, 'footer')
  const date = new Date(2026, 10, 22, 13, 4, 5)
  const refreshed = fieldService.refreshDocumentStoryDateFields(previous, 'header', 'en-US', date, document.getModel())
  assert.equal(set(opened, { pageFields: refreshed }).changed, true)
  const after = copy(document.getModel()), section = resolved(document, sectionId), fields = section.pageFields.filter(field => field.story === 'header')
  assert.equal(fields[0].cachedResult, '2026/11/22'); assert.equal(fields[1].cachedResult, '01:04:05 pm')
  assert.deepEqual(fields[2], unknown)
  assert.equal(editor.storyBuffer(section, 'header').value, input)
  assert.deepEqual(storySnapshot(section, 'footer'), footer)
  assert.deepEqual(storySnapshot(resolved(document, 'first'), 'header'), sourceStory)
  assert.equal(docModel.runText(stories.documentStoryFieldCache(fields[0]).runs), fields[0].cachedResult)
  assert.equal(stories.documentStoryFieldCache(fields[0]).runs[1].href, 'https://example.org/date')
  assert.equal(stories.documentStoryFieldCache(fields[1]).runs.at(-1).href, 'https://example.org/time')
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  const repeat = fieldService.refreshDocumentStoryDateFields(section, 'header', 'en-US', date, document.getModel())
  assert.equal(set(opened, { pageFields: repeat }).changed, false)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), before)
  assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), after)
  const reopened = open(copy(document.save())), reopenedSection = resolved(reopened.document, sectionId)
  assert.deepEqual(storySnapshot(reopenedSection, 'header'), storySnapshot(section, 'header'))
  const render = layoutEngine.layoutDocument(reopened.document.getModel(), { measure: text => Array.from(text).length * 8 })
  const currentPage = render.pageHeaderFooters.find(entry => entry.sectionId === sectionId && entry.headerStory === 'header')
  assert.ok(currentPage.header.includes('2026/11/22 / 01:04:05 pm'))
  assert.equal(currentPage.headerLayoutLines.flatMap(line => line.glyphs).map(glyph => glyph.text).join(''), currentPage.header)
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: false })
  assert.deepEqual(history(reopened.studio, reopened.document.id), { undo: 0, redo: 0, dirty: false })
})

test('date language provenance and clear fallback survive public history and two internal saves', () => {
  const source = sourceModel('date-language-history'), defaults = stories.documentStoryDefaultStyle(source)
  const sourceFormat = language => ({ version: 1, propertiesXml: '<w:rPr><w:lang w:val="' + language + '"/></w:rPr>',
    namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } })
  const run = (text, language, patch) => docModel.createRun(text, { ...defaults.text, ...patch,
    metadata: { source: { text, style: { language }, sourceFormat: sourceFormat(language) } } })
  const first = source.sectionHeaderFooters.first
  first.header = ''; first.storyContent.header = { runs: [docModel.createRun('', defaults.text)], paragraphs: stories.documentStoryParagraphs('', defaults.paragraph) }
  first.pageFields = [...first.pageFields.filter(field => field.story !== 'header'), {
    story: 'header', field: 'DATE', offset: 0, instruction: 'DATE \\@ "dddd d MMMM yyyy"', cachedResult: 'old cache',
    metadata: { source: { textStyle: { language: 'de-DE' }, sourceFormat: sourceFormat('de-DE') }, officeStoryFieldCache: { version: 1, runs: [
      run('old', 'fr-FR', { color: '#aa0000', bold: true }),
      run(' cache', 'de-DE', { color: '#0000aa', italic: true, href: 'https://example.org/language' }),
    ] } },
  }]
  const opened = open(source), { studio, document, sectionId } = opened
  assert.equal(set(opened, { storyLinkToPrevious: { header: false } }).changed, true)
  const sourceHeader = storySnapshot(resolved(document, 'first'), 'header'), footer = storySnapshot(resolved(document, sectionId), 'footer')
  const update = (target, date) => set(target, { pageFields: fieldService.refreshDocumentStoryDateFields(
    resolved(target.document, target.sectionId), 'header', 'en-US', date, target.document.getModel(),
  ) })
  const headerField = target => resolved(target.document, target.sectionId).pageFields.find(field => field.story === 'header')
  assert.equal(update(opened, new Date(2026, 10, 22, 13, 4, 5)).changed, true)
  assert.equal(headerField(opened).cachedResult, 'dimanche 22 novembre 2026')
  assert.deepEqual(history(studio, document.id), { undo: 2, redo: 0, dirty: true })
  const reopened = open(copy(document.save())), { studio: nextStudio, document: nextDocument } = reopened
  assert.deepEqual(storySnapshot(resolved(nextDocument, sectionId), 'header'), storySnapshot(resolved(document, sectionId), 'header'))
  assert.equal(headerField(reopened).metadata.source.textStyle.language, 'de-DE')
  assert.equal(stories.documentStoryFieldCache(headerField(reopened)).runs[0].metadata.source.style.language, 'fr-FR')
  assert.equal(update(reopened, new Date(2026, 10, 23, 13, 4, 5)).changed, true)
  assert.equal(headerField(reopened).cachedResult, 'lundi 23 novembre 2026')
  const beforeClear = copy(nextDocument.getModel()), input = editor.storyBuffer(resolved(nextDocument, sectionId), 'header').value
  const request = { sectionId, story: 'header', start: 0, end: 1 }
  assert.equal(nextDocument.clearFormatting(reopened.body, request).changed, true)
  const afterClear = copy(nextDocument.getModel()), cleared = headerField(reopened)
  assert.equal(cleared.metadata.officeStoryTextFormatCleared, true)
  assert.ok(stories.documentStoryFieldCache(cleared).runs.every(run => run.metadata.officeStoryTextFormatCleared === true))
  assert.equal(update(reopened, new Date(2026, 10, 23, 13, 4, 5)).changed, true)
  assert.equal(headerField(reopened).cachedResult, 'Monday 23 November 2026')
  const afterRefresh = copy(nextDocument.getModel())
  assert.equal(headerField(reopened).metadata.source.textStyle.language, 'de-DE')
  assert.equal(stories.documentStoryFieldCache(headerField(reopened)).runs[0].metadata.source.style.language, 'fr-FR')
  assert.equal(stories.documentStoryFieldCache(headerField(reopened)).runs.at(-1).href, 'https://example.org/language')
  assert.equal(editor.storyBuffer(resolved(nextDocument, sectionId), 'header').value, input)
  assert.deepEqual(storySnapshot(resolved(nextDocument, 'first'), 'header'), sourceHeader)
  assert.deepEqual(storySnapshot(resolved(nextDocument, sectionId), 'footer'), footer)
  assert.deepEqual(history(nextStudio, nextDocument.id), { undo: 3, redo: 0, dirty: true })
  assert.equal(nextDocument.undo(), true); assert.deepEqual(nextDocument.getModel(), afterClear)
  assert.equal(nextDocument.undo(), true); assert.deepEqual(nextDocument.getModel(), beforeClear)
  assert.equal(nextDocument.redo(), true); assert.equal(nextDocument.redo(), true); assert.deepEqual(nextDocument.getModel(), afterRefresh)
  const second = open(copy(nextDocument.save())), field = headerField(second)
  assert.equal(field.cachedResult, 'Monday 23 November 2026')
  assert.equal(field.metadata.officeStoryTextFormatCleared, true)
  assert.equal(stories.documentStoryFieldCache(field).runs.at(-1).href, 'https://example.org/language')
  assert.equal(update(second, new Date(2026, 10, 23, 13, 4, 5)).changed, false)
  const future = fieldService.refreshDocumentStoryDateFields(resolved(second.document, sectionId), 'header', 'en-US', new Date(2026, 10, 24, 13, 4, 5), second.document.getModel())
  assert.equal(future.find(value => value.story === 'header').cachedResult, 'Tuesday 24 November 2026')
  assert.equal(headerField(second).cachedResult, 'Monday 23 November 2026')
  const render = layoutEngine.layoutDocument(second.document.getModel(), { measure: text => Array.from(text).length * 8 })
  const currentPage = render.pageHeaderFooters.find(entry => entry.sectionId === sectionId && entry.headerStory === 'header')
  assert.equal(currentPage.header, 'Monday 23 November 2026')
  assert.equal(currentPage.headerLayoutLines.flatMap(line => line.glyphs).map(glyph => glyph.text).join(''), currentPage.header)
  assert.deepEqual(history(nextStudio, nextDocument.id), { undo: 3, redo: 0, dirty: false })
  assert.deepEqual(history(second.studio, second.document.id), { undo: 0, redo: 0, dirty: false })
})

test('permissions inspect actual story changes and allow echoed inherited content in another permitted story', () => {
  const source = sourceModel('story-links-denied'), sectionId = source.blocks[1].id
  source.permissionRules = [{ id: 'deny-even-header', scope: 'entity', targetId: sectionId + ':evenPageHeader', action: 'edit', effect: 'deny', userIds: [] },
    { id: 'deny-footers', scope: 'entity', targetId: sectionId + ':footer', action: 'edit', effect: 'deny', userIds: [] }]
  const opened = open(source), { studio, document } = opened
  const original = copy(document.getModel())
  assert.throws(() => set(opened, { storyLinkToPrevious: { evenPageHeader: false } }))
  assert.throws(() => set(opened, { linkToPrevious: false }))
  assert.deepEqual(document.getModel(), original); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(set(opened, { storyLinkToPrevious: { header: false } }).changed, true)
  const beforeFooter = storySnapshot(resolved(document, sectionId), 'footer'), current = resolved(document, sectionId)
  const patch = editor.replaceStoryRange(current, 'header', 1, 1, 'Allowed', document.getModel())
  assert.equal(set(opened, patch).changed, true)
  assert.deepEqual(storySnapshot(resolved(document, sectionId), 'footer'), beforeFooter)
  const edited = copy(document.getModel()), counts = history(studio, document.id)
  const fields = resolved(document, sectionId).pageFields.map(field => field.story === 'footer' ? { ...field, cachedResult: '8', metadata: undefined } : field)
  assert.throws(() => set(opened, { pageFields: fields }))
  assert.deepEqual(document.getModel(), edited); assert.deepEqual(history(studio, document.id), counts)
})

test('invalid association and canceled commands publish neither content nor a history entry', () => {
  const opened = open(), { studio, document } = opened, original = copy(document.getModel())
  for (const patch of [{ storyLinkToPrevious: { unknown: true } }, { storyLinkToPrevious: { header: 'false' } }, { storyLinkToPrevious: null }]) {
    assert.throws(() => set(opened, patch)); assert.deepEqual(document.getModel(), original)
  }
  assert.throws(() => document.executeCommand({ id: 'doc.command.set-section-header-footer', sectionId: 'first', patch: { storyLinkToPrevious: { header: true } } }, opened.body))
  const canceled = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'doc.command.set-section-header-footer') event.cancel = true })
  assert.throws(() => set(opened, { storyLinkToPrevious: { header: false } })); canceled.dispose()
  assert.deepEqual(document.getModel(), original); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
})

test('story edits and history remain isolated when another document is active', () => {
  const first = open(), secondModel = sourceModel('story-links-other')
  first.studio.createDocument(secondModel); first.studio.activateDocument(secondModel.id)
  const second = new FDocument(first.studio, secondModel.id), untouched = copy(second.getModel())
  set(first, { storyLinkToPrevious: { evenPageFooter: false } })
  assert.equal(first.studio.activeDocumentId, second.id); assert.deepEqual(second.getModel(), untouched)
  assert.deepEqual(history(first.studio, second.id), { undo: 0, redo: 0, dirty: false })
  assert.equal(first.document.undo(), true); assert.deepEqual(second.getModel(), untouched)
  assert.equal(first.document.redo(), true); assert.deepEqual(second.getModel(), untouched)
})

test('public field formatting preserves mixed live numeric results through content history and internal save/reopen', () => {
  const source = domain.createBlankDocument(), defaults = stories.documentStoryDefaultStyle(source)
  source.id = 'live-field-history'; source.sectionSetup = { pageNumberStart: 19 }
  source.blocks = [docModel.createParagraph('First'), { ...docModel.createParagraph(), id: 'page-two', type: 'pageBreak', runs: [] }, docModel.createParagraph('Second')]
  source.sectionHeaderFooters = { first: { ...sections.EMPTY_SECTION_HEADER_FOOTER, header: '', pageFields: [{
    story: 'header', field: 'PAGE', offset: 0, instruction: 'PAGE', cachedResult: '19', metadata: {
      retained: { source: 'imported-mixed-cache' }, officeStoryFieldCache: { version: 1, runs: [
        docModel.createRun('1', { ...defaults.text, color: '#aa0000', bold: true }),
        docModel.createRun('9', { ...defaults.text, color: '#0000aa', italic: true, href: 'https://example.org/field' }),
      ] },
    },
  }] } }
  const opened = open(source), { studio, document } = opened, original = copy(document.getModel())
  const measure = { measure: text => Array.from(text).length * 8 }
  const render = api => layoutEngine.layoutDocument(api.getModel(), measure)
  const glyphs = value => value.pageHeaderFooters[1].headerLayoutLines.flatMap(line => line.glyphs)
    .map(glyph => [glyph.text, glyph.offset, glyph.end, glyph.run.color, glyph.run.bold, glyph.run.italic, glyph.run.underline])
  const initial = render(document)
  assert.deepEqual(initial.pages.map(page => page.pageNumber), [19, 20]); assert.equal(initial.pageHeaderFooters[1].header, '20')
  assert.deepEqual(glyphs(initial).map(glyph => glyph.slice(0, 6)), [['2', 0, 1, '#aa0000', true, false], ['0', 0, 1, '#0000aa', false, true]])
  assert.deepEqual(document.getModel(), original); assert.deepEqual(history(studio, document.id), { undo: 0, redo: 0, dirty: false })
  const result = document.executeCommand({ id: 'doc.command.format-text', story: { sectionId: 'first', story: 'header', start: 0, end: 1 }, patch: { underline: true } }, opened.body)
  assert.equal(result.changed, true)
  const formatted = copy(document.getModel()), field = resolved(document, 'first').pageFields[0], rendered = render(document)
  assert.equal(field.cachedResult, '19'); assert.equal(field.instruction, 'PAGE')
  assert.equal(docModel.runText(stories.documentStoryFieldCache(field).runs), '19')
  assert.equal(stories.documentStoryFieldCache(field).runs[1].href, 'https://example.org/field')
  assert.deepEqual(glyphs(rendered), [['2', 0, 1, '#aa0000', true, false, true], ['0', 0, 1, '#0000aa', false, true, true]])
  assert.deepEqual(document.getModel(), formatted); assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: true })
  assert.equal(document.undo(), true); assert.deepEqual(document.getModel(), original)
  assert.deepEqual(glyphs(render(document)), glyphs(initial))
  assert.equal(document.redo(), true); assert.deepEqual(document.getModel(), formatted)
  const saved = copy(document.save()), reopened = open(saved)
  assert.equal(resolved(reopened.document, 'first').pageFields[0].cachedResult, '19')
  assert.deepEqual(glyphs(render(reopened.document)), glyphs(rendered))
  assert.deepEqual(history(studio, document.id), { undo: 1, redo: 0, dirty: false })
  assert.deepEqual(history(reopened.studio, reopened.document.id), { undo: 0, redo: 0, dirty: false })
})
