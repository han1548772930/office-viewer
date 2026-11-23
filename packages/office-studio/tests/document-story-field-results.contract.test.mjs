import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, createRun, runText } from '../src/features/document/model/document-model.ts'
import { createDocumentNumericPageField, resolveDocumentPageFieldResult } from '../src/features/document/model/document-page-field-result.ts'
import { documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, documentStoryParagraphs, isDocumentTextBoundary } from '../src/features/document/model/document-story.ts'
import { EMPTY_SECTION_HEADER_FOOTER, resolveSectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'
import { refreshDocumentStoryFieldResult } from '../src/features/document/services/document-story-fields.ts'
import { documentStoryDisplay } from '../src/features/document/services/story-rendering.ts'
import { storyBuffer } from '../src/features/document/services/story-editor.ts'
import { layoutDocument } from '../src/features/document/layout/document-layout.ts'

// Deferred pure-model and Canvas layout drafts, 2026-11-22. Not executed.
const copy = value => JSON.parse(JSON.stringify(value))
const visible = runs => runs.filter(run => run.text.length).map(run => [run.text, run.color, run.bold, run.italic])
const measurer = { measure: text => Array.from(text).length * 8 }
function fieldModel() {
  const model = createBlankDocument(), defaults = documentStoryDefaultStyle(model)
  model.id = 'dynamic-page-fields'; model.sectionSetup = { pageNumberStart: 19 }
  const boundary = { ...createParagraph(), id: 'manual-page', type: 'pageBreak', runs: [] }
  model.blocks = [createParagraph('First'), boundary, createParagraph('Second')]
  const page = { story: 'header', field: 'PAGE', offset: 0, instruction: 'PAGE', cachedResult: '19', metadata: {
    retained: { values: [1, 2] }, officeStoryFieldCache: { version: 1, runs: [
      createRun('1', { ...defaults.text, bold: true, color: '#aa0000', metadata: { retained: { values: [1, 2] } } }),
      createRun('9', { ...defaults.text, italic: true, color: '#0000aa', href: 'https://example.org/field' }),
    ] },
  } }
  const total = { story: 'header', field: 'NUMPAGES', offset: 3, instruction: 'NUMPAGES', cachedResult: '88', metadata: {
    officeStoryFieldCache: { version: 1, runs: [createRun('8', { ...defaults.text, color: '#00aa00' }), createRun('8', { ...defaults.text, color: '#aa00aa' })] },
  } }
  model.sectionHeaderFooters = { first: { ...EMPTY_SECTION_HEADER_FOOTER, header: ' / ', pageFields: [page, total], storyContent: { header: {
    runs: [createRun(' / ', defaults.text)], paragraphs: documentStoryParagraphs(' / ', defaults.paragraph),
  } } } }
  return { model, page, total, defaults }
}

test('live numeric results retain mixed cached formats, links and independent metadata ownership', () => {
  const { model, page, defaults } = fieldModel(), before = copy(page)
  const displayed = documentStoryFieldDisplayRuns(page, '20', defaults.text)
  assert.deepEqual(visible(displayed), [['2', '#aa0000', true, false], ['0', '#0000aa', false, true]])
  assert.equal(displayed[1].href, 'https://example.org/field')
  displayed[0].color = '#000000'; displayed[0].metadata = { outside: true }
  const sameText = documentStoryFieldDisplayRuns(page, '19', defaults.text)
  sameText[1].href = 'https://example.org/changed'
  sameText[0].metadata.retained.values.push(3)
  assert.deepEqual(page, before)
  const appended = documentStoryFieldDisplayRuns(page, '19😀', defaults.text)
  assert.equal(runText(appended), '19😀')
  assert.equal(appended.at(-1).color, '#0000aa')
  let offset = 0
  for (const run of appended) { offset += run.text.length; assert.equal(isDocumentTextBoundary('19😀', offset), true) }
  const refreshed = refreshDocumentStoryFieldResult(page, '20', model)
  assert.deepEqual(visible(documentStoryFieldCache(refreshed).runs), visible(documentStoryFieldDisplayRuns(page, '20', defaults.text)))
  refreshed.metadata.retained.values.push(3); assert.deepEqual(page, before)
})

test('result replacement preserves Unicode boundaries and unchanged empty cache entities', () => {
  const { defaults } = fieldModel(), field = { story: 'header', field: 'DATE', cachedResult: 'A😀B', metadata: { officeStoryFieldCache: { version: 1, runs: [
    createRun('A', { ...defaults.text, color: '#aa0000' }),
    createRun('😀', { ...defaults.text, color: '#00aa00' }),
    createRun('B', { ...defaults.text, color: '#0000aa' }),
    createRun('', { ...defaults.text, metadata: { entity: { id: 'trailing-cache-entity' } } }),
  ] } } }
  const before = copy(field), displayed = documentStoryFieldDisplayRuns(field, 'A😇B', defaults.text)
  assert.deepEqual(visible(displayed).map(([text, color]) => [text, color]), [['A', '#aa0000'], ['😇', '#00aa00'], ['B', '#0000aa']])
  assert.equal(displayed.at(-1).id, field.metadata.officeStoryFieldCache.runs.at(-1).id)
  displayed.at(-1).metadata.entity.id = 'external'
  assert.deepEqual(field, before)
  const shortened = documentStoryFieldDisplayRuns(field, 'A', defaults.text)
  assert.equal(runText(shortened), 'A')
  let offset = 0
  for (const run of shortened) { offset += run.text.length; assert.equal(isDocumentTextBoundary('A', offset), true) }
})

test('expanded PAGE and NUMPAGES display text retain one input object each without rewriting cached values', () => {
  const { model } = fieldModel(), before = copy(model), section = resolveSectionHeaderFooter(model, 'first')
  const display = documentStoryDisplay(section, 'header', 20, 123, model)
  assert.equal(display.text, '20 / 123'); assert.equal(storyBuffer(section, 'header').value, '\uFFFC / \uFFFC')
  assert.deepEqual(display.inputRange(0, 2), { start: 0, end: 1 })
  assert.deepEqual(display.inputRange(1, 1), { start: 0, end: 0 })
  assert.deepEqual(display.inputRange(5, 8), { start: 4, end: 5 })
  assert.deepEqual(display.inputRange(8, 8), { start: 5, end: 5 })
  assert.deepEqual(visible(display.runs).map(([text, color]) => [text, color]), [
    ['2', '#aa0000'], ['0', '#0000aa'], [' / ', documentStoryDefaultStyle(model).text.color], ['1', '#00aa00'], ['23', '#aa00aa'],
  ])
  assert.deepEqual(model, before)
})

test('Canvas re-pagination paints live numeric formats and keeps cached runs out of model history', () => {
  const { model } = fieldModel(), before = copy(model), layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.pages.map(page => page.pageNumber), [19, 20])
  assert.deepEqual(layout.pageHeaderFooters.map(page => page.header), ['19 / 2', '20 / 2'])
  const pageGlyphs = layout.pageHeaderFooters[1].headerLayoutLines.flatMap(line => line.glyphs)
  assert.deepEqual(pageGlyphs.filter(glyph => glyph.offset === 0 && glyph.end === 1).map(glyph => [glyph.text, glyph.run.color, glyph.run.bold, glyph.run.italic]), [
    ['2', '#aa0000', true, false], ['0', '#0000aa', false, true],
  ])
  assert.ok(pageGlyphs.filter(glyph => glyph.offset === 4).every(glyph => glyph.end === 5))
  assert.deepEqual(model, before)
})

test('a content-locked numeric field remains live during rendering and retains its locked cache', () => {
  const { model } = fieldModel()
  model.sectionHeaderFooters.first.storyContent.header.locks = [{ id: 'locked-page', lock: 'contentLocked', startOffset: 0, endOffset: 1 }]
  const before = copy(model), layout = layoutDocument(model, measurer)
  assert.equal(layout.pageHeaderFooters[1].header, '20 / 2')
  assert.equal(model.sectionHeaderFooters.first.pageFields[0].cachedResult, '19')
  assert.deepEqual(model, before)
})

test('live rendering still rejects corrupt cached ranges and falls back to a modeled style when no cache exists', () => {
  const { page, defaults } = fieldModel(), invalid = copy(page)
  invalid.cachedResult = 'different'
  assert.throws(() => documentStoryFieldDisplayRuns(invalid, '20', defaults.text), /缓存文本/)
  const plain = { story: 'header', field: 'PAGE', cachedResult: '1', textStyle: { ...defaults.text, fontSize: 24, bold: true } }
  const runs = documentStoryFieldDisplayRuns(plain, '20', defaults.text)
  assert.equal(runText(runs), '20'); assert.equal(runs[0].fontSize, 24); assert.equal(runs[0].bold, true)
})

for (const [instruction, pageNumber, pageCount, expected] of [
  ['PAGE', 6, 2, '6'],
  ['PAGE \\*ARABIC', 6, 2, '6'],
  ['PAGE \\*ROMAN', 6, 2, 'VI'],
  ['PAGE \\*roman', 6, 2, 'vi'],
  ['PAGE \\*"roman"', 6, 2, 'vi'],
  ['PAGE \\*roman \\*Upper', 6, 2, 'VI'],
  ['PAGE \\*ROMAN \\*Lower', 6, 2, 'vi'],
  ['PAGE \\*roman \\*ARABIC', 6, 2, '6'],
  ['PAGE \\*ARABIC \\*roman', 6, 2, 'vi'],
  ['PAGE \\*roman \\*Upper \\*ARABIC', 6, 2, '6'],
  ['PAGE \\*ALPHABETIC', 28, 2, 'BB'],
  ['PAGE \\*alphabetic', 28, 2, 'bb'],
  ['PAGE \\*ROMAN \\*MERGEFORMAT', 16, 2, 'XVI'],
  ['PAGE \\*roman \\*CHARFORMAT', 16, 2, 'xvi'],
  ['PAGE \\*ROMAN', 0, 2, '0'],
  ['NUMPAGES', 6, 2, '2'],
  ['NUMPAGES \\*roman', 6, 2, 'ii'],
  ['NUMPAGES \\*ALPHABETIC', 6, 28, 'BB'],
]) test('numeric instruction ' + instruction + ' keeps its ordered switches and source cache', () => {
  const field = { story: 'header', field: instruction.startsWith('NUMPAGES') ? 'NUMPAGES' : 'PAGE', instruction, cachedResult: 'original' }, before = copy(field)
  assert.equal(resolveDocumentPageFieldResult(field, pageNumber, pageCount), expected)
  assert.deepEqual(field, before)
})

test('unsupported numeric instructions and unresolved counts retain cached values', () => {
  for (const instruction of ['PAGE \\*UNKNOWN', 'PAGE \\# "0.00"', 'PAGE \\h', 'PAGE extra', 'PAGE \\*', 'PAGE \\*"roman" trailing', 'DATE']) {
    const field = { story: 'header', field: 'PAGE', instruction, cachedResult: 'retained' }
    assert.equal(resolveDocumentPageFieldResult(field, 6, 2), 'retained')
  }
  assert.equal(resolveDocumentPageFieldResult({ story: 'header', field: 'NUMPAGES', cachedResult: 'pending' }, 6), 'pending')
  assert.equal(resolveDocumentPageFieldResult({ story: 'header', field: 'PAGE', cachedResult: 'invalid' }, Number.MAX_SAFE_INTEGER + 1, 2), 'invalid')
  for (const instruction of ['PAGE \\*ROMAN', 'PAGE \\*ALPHABETIC']) {
    assert.equal(resolveDocumentPageFieldResult({ story: 'header', field: 'PAGE', instruction, cachedResult: 'bounded' }, Number.MAX_SAFE_INTEGER, 2), 'bounded')
  }
  assert.equal(resolveDocumentPageFieldResult({ story: 'header', field: 'DATE', cachedResult: 'date' }, 6, 2), 'date')
})

test('new numeric fields use current logical page and physical count without an invented cache', () => {
  assert.deepEqual(createDocumentNumericPageField('PAGE', 19, 2), { field: 'PAGE', instruction: 'PAGE', cachedResult: '19' })
  assert.deepEqual(createDocumentNumericPageField('PAGE', 0, 2), { field: 'PAGE', instruction: 'PAGE', cachedResult: '0' })
  assert.deepEqual(createDocumentNumericPageField('NUMPAGES', 19, 2), { field: 'NUMPAGES', instruction: 'NUMPAGES', cachedResult: '2' })
  for (const [field, page, total] of [['DATE', 19, 2], ['PAGE', -1, 2], ['PAGE', undefined, 2], ['PAGE', 1.5, 2], ['NUMPAGES', 19, 0], ['NUMPAGES', 19, undefined]]) {
    assert.throws(() => createDocumentNumericPageField(field, page, total))
  }
})

test('formatted numeric field strings and Canvas glyphs use the same value with atomic input ranges', () => {
  const { model, page, total } = fieldModel()
  page.instruction = 'PAGE \\*roman'; total.instruction = 'NUMPAGES \\*ROMAN'
  const before = copy(model), layout = layoutDocument(model, measurer)
  assert.deepEqual(layout.pageHeaderFooters.map(page => page.header), ['xix / II', 'xx / II'])
  for (const entry of layout.pageHeaderFooters) {
    assert.equal(entry.headerLayoutLines.flatMap(line => line.glyphs).map(glyph => glyph.text).join(''), entry.header)
    assert.ok(entry.headerLayoutLines.flatMap(line => line.glyphs).filter(glyph => glyph.offset === 0).every(glyph => glyph.end === 1))
  }
  assert.deepEqual(model, before)
})
