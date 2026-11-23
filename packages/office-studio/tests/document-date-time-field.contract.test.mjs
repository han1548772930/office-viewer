import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankDocument } from '../src/domain/model.ts'
import { createRun, runText } from '../src/features/document/model/document-model.ts'
import { documentStoryDefaultStyle, documentStoryFieldCache, documentStoryParagraphs } from '../src/features/document/model/document-story.ts'
import { createDocumentDateTimeField, formatDocumentDateTimePicture, resolveDocumentDateTimeFieldResult } from '../src/features/document/services/document-date-time-field.ts'
import { refreshDocumentStoryDateFields } from '../src/features/document/services/document-story-fields.ts'
import { EMPTY_SECTION_HEADER_FOOTER } from '../src/features/document/services/section-header-footer.ts'
import { storyBuffer, storyDateField } from '../src/features/document/services/story-editor.ts'
import { normalizeDocument } from '../src/parser/documentNormalizer.ts'

// Deferred explicit-picture, language and cached-run contracts, 2026-11-22.
// No validation was executed; Office regional defaults are not these fixtures.
const date = new Date(2026, 10, 22, 13, 4, 5), copy = value => JSON.parse(JSON.stringify(value))
const field = (instruction, type = 'DATE') => ({ field: type, instruction, cachedResult: 'source cache' })
const languageRun = (text, language, metadata = {}) => createRun(text, {
  ...documentStoryDefaultStyle().text, metadata: { source: { text, style: language === undefined ? {} : { language } }, ...metadata },
})
const languageField = (runs, language, metadata = {}) => ({
  story: 'header', field: 'DATE', offset: 0, instruction: 'DATE \\@ "dddd d MMMM yyyy"',
  cachedResult: runs === undefined ? 'source cache' : runText(runs), metadata: {
    ...(language === undefined ? {} : { source: { textStyle: { language } } }),
    ...(runs === undefined ? {} : { officeStoryFieldCache: { version: 1, runs } }), ...metadata,
  },
})
const refreshLanguageField = (source, locale = 'en-US') => refreshDocumentStoryDateFields(
  { ...EMPTY_SECTION_HEADER_FOOTER, pageFields: [source] }, 'header', locale, date,
)[0]

test('Word date-time markers keep month/minute and 12/24 hour meanings distinct', () => {
  assert.equal(formatDocumentDateTimePicture('d dd ddd dddd M MM MMM MMMM yy yyyy h hh H HH m mm s ss AM/PM am/pm', 'en-US', date),
    '22 22 Sun Sunday 11 11 Nov November 26 2026 1 01 13 13 4 04 5 05 PM pm')
  for (const [picture, expected] of [['MM/mm', '11/04'], ['HH:mm:ss', '13:04:05'], ['hh:mm:ss AM/PM', '01:04:05 PM'], ['yyyyMMdd', '20261122']]) {
    assert.equal(formatDocumentDateTimePicture(picture, 'en-US', date), expected)
  }
  const midnight = new Date(2026, 10, 22, 0, 4, 5), noon = new Date(2026, 10, 22, 12, 4, 5)
  assert.equal(formatDocumentDateTimePicture('h hh H HH AM/PM', 'en-US', midnight), '12 12 0 00 AM')
  assert.equal(formatDocumentDateTimePicture('h hh H HH AM/PM', 'en-US', noon), '12 12 12 12 PM')
})

test('quoted words, doubled apostrophes, escapes and Unicode literals retain their text', () => {
  for (const [picture, expected] of [
    ["yyyy 'o''clock'", "2026 o'clock"],
    ["yyyy '\\\"Hi\\\"'", '2026 "Hi"'],
    ["yyyy '\\\\' 'o''clock'", "2026 \\ o'clock"],
    ['yyyy \\M\\d \\😀', '2026 Md 😀'],
    ["yyyy年M月d日 'd MMMM yyyy' 😀", '2026年11月22日 d MMMM yyyy 😀'],
  ]) assert.equal(formatDocumentDateTimePicture(picture, 'en-US', date), expected)
  assert.equal(resolveDocumentDateTimeFieldResult(field('DATE \\@ "yyyy \'said \\\"Hi\\\"\'"'), 'en-US', date), '2026 said "Hi"')
  assert.equal(resolveDocumentDateTimeFieldResult(field('DATE \\@ yyyy/MM/dd'), 'en-US', date), '2026/11/22')
  assert.equal(resolveDocumentDateTimeFieldResult(field('  date \\@"yyyy/MM/dd" \\*"MERGEFORMAT"  '), 'en-US', date), '2026/11/22')
})

test('month and weekday names follow the supplied language with date-context inflection', () => {
  assert.equal(formatDocumentDateTimePicture('dddd d MMMM yyyy', 'fr-FR', date), 'dimanche 22 novembre 2026')
  assert.equal(formatDocumentDateTimePicture('dddd d MMMM yyyy', 'de-DE', date), 'Sonntag 22 November 2026')
  assert.equal(formatDocumentDateTimePicture('dddd d MMMM yyyy', 'ru-RU', date), 'воскресенье 22 ноября 2026')
  assert.equal(formatDocumentDateTimePicture('MMMM', 'ru-RU', date), 'ноябрь')
  assert.equal(formatDocumentDateTimePicture('yyyy/MM/dd', 'ar-SA-u-ca-islamic-nu-arab', date), '2026/11/22')
  const april = new Date(2026, 3, 20, 13, 4, 5)
  assert.equal(resolveDocumentDateTimeFieldResult(field('DATE \\@ "MMMM" \\*Upper'), 'tr-TR', april), 'NİSAN')
  assert.equal(resolveDocumentDateTimeFieldResult(field('DATE \\@ "MMMM" \\*Upper \\*Lower'), 'tr-TR', april), 'nisan')
})

test('the first nonempty cache run supplies the date language before field provenance and UI locale', () => {
  const source = languageField([languageRun('old', 'fr-FR'), languageRun(' cache', 'de-DE')], 'en-US')
  const before = copy(source), refreshed = refreshLanguageField(source)
  assert.equal(refreshed.cachedResult, 'dimanche 22 novembre 2026')
  assert.deepEqual(refreshed.metadata.source, before.metadata.source)
  assert.equal(documentStoryFieldCache(refreshed).runs[0].metadata.source.style.language, 'fr-FR')
  assert.deepEqual(source, before)
  assert.equal(refreshLanguageField(refreshed), refreshed)
})

test('empty cache runs do not select a language or suppress the first nonempty run language', () => {
  for (const emptyMetadata of [{}, { officeStoryTextFormatCleared: true }]) {
    const source = languageField([languageRun('', 'de-DE', emptyMetadata), languageRun('source cache', 'fr-FR')], 'en-US')
    assert.equal(refreshLanguageField(source).cachedResult, 'dimanche 22 novembre 2026')
    const allEmpty = languageField([languageRun('', 'de-DE', emptyMetadata)], 'fr-FR')
    const expanded = refreshLanguageField(allEmpty)
    assert.equal(expanded.cachedResult, 'dimanche 22 novembre 2026')
    assert.equal(refreshLanguageField(expanded), expanded)
  }
})

test('missing first-result language falls back to field provenance and then UI locale', () => {
  for (const source of [
    languageField(undefined, 'de-DE'), languageField([], 'de-DE'),
    languageField([languageRun('old', undefined), languageRun(' cache', 'fr-FR')], 'de-DE'),
  ]) assert.equal(refreshLanguageField(source).cachedResult, 'Sonntag 22 November 2026')
  assert.equal(refreshLanguageField(languageField([languageRun('source cache')])).cachedResult, 'Sunday 22 November 2026')
})

test('invalid source languages retain the original cache without falling back to a valid UI locale', () => {
  for (const language of ['', 'en_US', 'zz-ZZ']) {
    for (const source of [languageField([languageRun('source cache', language)], 'fr-FR'), languageField(undefined, language)]) {
      const before = copy(source)
      assert.equal(refreshLanguageField(source), source, language)
      assert.deepEqual(source, before)
    }
  }
  const source = languageField([languageRun('source cache', 'fr-FR')], 'de-DE')
  assert.equal(refreshLanguageField(source, 'en_US').cachedResult, 'dimanche 22 novembre 2026')
})

test('a field or first-result clear marker ignores both preserved source languages', () => {
  for (const [fieldMetadata, runMetadata] of [
    [{ officeStoryTextFormatCleared: true }, {}], [{}, { officeStoryTextFormatCleared: true }],
  ]) {
    const source = languageField([languageRun('source cache', 'fr-FR', runMetadata)], 'de-DE', fieldMetadata)
    const before = copy(source), refreshed = refreshLanguageField(source)
    assert.equal(refreshed.cachedResult, 'Sunday 22 November 2026')
    assert.deepEqual(refreshed.metadata.source, source.metadata.source)
    assert.deepEqual(source, before)
    assert.equal(refreshLanguageField(refreshed), refreshed)
  }
  const laterClear = languageField([languageRun('old', 'fr-FR'), languageRun(' cache', 'de-DE', { officeStoryTextFormatCleared: true })], 'en-US')
  assert.equal(refreshLanguageField(laterClear).cachedResult, 'dimanche 22 novembre 2026')
})

test('parsed cache language survives normalization and drives date refresh ahead of the field language', () => {
  const sourceFormat = { version: 1, propertiesXml: '<w:rPr><w:lang w:val="fr-FR"/></w:rPr>',
    namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } }
  const cachedRuns = [{ text: 'old', style: { language: 'fr-FR', color: '#aa0000' }, sourceFormat },
    { text: ' cache', style: { language: 'de-DE', color: '#0000aa' }, link: 'https://example.org/language' }]
  const parsed = { field: 'DATE', offset: 0, instruction: 'DATE \\@ "dddd d MMMM yyyy"', cachedResult: 'old cache',
    textStyle: { language: 'de-DE' }, cachedRuns }
  const model = normalizeDocument({ schemaVersion: 1, parserVersion: 'draft', format: 'docx', fileName: 'language.docx', size: 1,
    model: 'document', diagnostics: {}, payload: { modelType: 'document', data: {
      documentId: 'language-model', blocks: [], sectionLayout: [{ headerSections: [{ referenceType: 'default', storyText: {
        text: '', runs: [{ text: '', style: {} }], paragraphs: [{ startOffset: 0, endOffset: 0, style: {} }], fields: [parsed],
      } }] }],
    } } })
  const section = model.sectionHeaderFooters.first, before = copy(model)
  const refreshed = refreshDocumentStoryDateFields(section, 'header', 'en-US', date, model)[0]
  assert.equal(refreshed.cachedResult, 'dimanche 22 novembre 2026')
  assert.equal(refreshed.metadata.source.textStyle.language, 'de-DE')
  assert.equal(documentStoryFieldCache(refreshed).runs[0].metadata.source.style.language, 'fr-FR')
  assert.equal(documentStoryFieldCache(refreshed).runs.at(-1).metadata.link, 'https://example.org/language')
  assert.equal(storyBuffer({ ...section, pageFields: [refreshed] }, 'header').value, '\uFFFC')
  assert.deepEqual(model, before)
})

for (const locale of ['en-US', 'zh-CN', 'fr-FR', 'de-DE', 'ru-RU', 'ja-JP', 'ar-SA']) {
  for (const format of ['short', 'long', 'time']) test('generated ' + locale + ' ' + format + ' instruction resolves its original cache', () => {
    for (const value of [date, new Date(2024, 1, 29, 0, 4, 5)]) {
      const created = createDocumentDateTimeField(format, locale, value), before = copy(created)
      assert.deepEqual(storyDateField(format, locale, value), created)
      assert.equal(resolveDocumentDateTimeFieldResult(created, locale, value), created.cachedResult)
      assert.deepEqual(created, before)
    }
  })
}

test('invalid dates, languages, field syntax and unsupported pictures retain the source cache', () => {
  for (const instruction of [
    'DATE', 'DATE \\@ ""', 'TIME \\@ "HH:mm:ss"', 'DATE \\@ "yyyy" trailing', 'DATE \\@ "yyyy', 'DATE \\@ "yyyy\\"',
    'DATE \\@ "yyyy" \\@ "MM"', 'DATE \\@ "yyyy" \\*', 'DATE \\@ "yyyy" \\*CHARFORMAT',
    'DATE \\@ "yyyy" \\*Unknown', 'DATE \\@ "yyyy" \\h', 'DATE \\@ "yyyy" \\l en-US', 'DATE \\# "0.00"',
    'DATE \\@ "QQ yyyy"', 'DATE \\@ "yyyy \'open"', 'DATE \\@ "MMMMM"', 'DATE \\@ "yyy"',
    'DATE \\@ "dddddd"', 'DATE \\@ "mmmm"', 'DATE \\@ "HHH"',
  ]) {
    const source = field(instruction), before = copy(source)
    assert.equal(resolveDocumentDateTimeFieldResult(source, 'en-US', date), undefined, instruction)
    assert.deepEqual(source, before)
  }
  for (const locale of ['', 'en_US', 'zz-ZZ']) {
    assert.equal(resolveDocumentDateTimeFieldResult(field('DATE \\@ "yyyy"'), locale, date), undefined)
    assert.throws(() => createDocumentDateTimeField('short', locale, date), /语言/)
  }
  assert.equal(resolveDocumentDateTimeFieldResult(field('DATE \\@ "yyyy"'), 'en-US', new Date(NaN)), undefined)
  assert.throws(() => createDocumentDateTimeField('time', 'en-US', new Date(NaN)), /日期/)
  assert.throws(() => createDocumentDateTimeField('unknown', 'en-US', date), /格式/)
  assert.equal(resolveDocumentDateTimeFieldResult({ field: 'PAGE', instruction: 'DATE \\@ "yyyy"' }, 'en-US', date), undefined)
})

test('custom DATE/TIME refresh updates mixed cache spans only in the selected unlinked story', () => {
  const model = createBlankDocument(), defaults = documentStoryDefaultStyle(model)
  const dateField = { story: 'header', field: 'DATE', offset: 0, instruction: 'DATE \\@ "yyyy/MM/dd" \\*MERGEFORMAT', cachedResult: '2026/11/21', textStyle: defaults.text, metadata: {
    source: { retainedInstruction: 'original' }, officeStoryStyleBaseline: defaults.text, officeStoryFieldCache: { version: 1, runs: [
      createRun('2026/', { ...defaults.text, color: '#aa0000', bold: true }),
      createRun('11/21', { ...defaults.text, color: '#0000aa', italic: true, href: 'https://example.org/date' }),
    ] },
  } }
  const timeField = { story: 'header', field: 'TIME', offset: 3, instruction: 'TIME \\@ "hh:mm:ss AM/PM" \\*Lower', cachedResult: '12:59:58 pm', metadata: {
    officeStoryFieldCache: { version: 1, runs: [
      createRun('12:', { ...defaults.text, color: '#00aa00', bold: true }),
      createRun('59:58 pm', { ...defaults.text, color: '#aa00aa', italic: true, href: 'https://example.org/time' }),
    ] },
  } }
  const other = { story: 'footer', ...field('DATE \\@ "yyyy/MM/dd"'), offset: 0 }, unknown = { story: 'header', ...field('DATE \\h'), offset: 3 }
  const section = { ...EMPTY_SECTION_HEADER_FOOTER, header: ' / ', pageFields: [dateField, timeField, other, unknown], storyContent: { header: {
    runs: [createRun(' / ', defaults.text)], paragraphs: documentStoryParagraphs(' / ', defaults.paragraph),
  } } }
  const before = copy(section), refreshed = refreshDocumentStoryDateFields(section, 'header', 'en-US', date, model)
  assert.equal(refreshed[0].cachedResult, '2026/11/22'); assert.equal(refreshed[1].cachedResult, '01:04:05 pm')
  assert.deepEqual(documentStoryFieldCache(refreshed[0]).runs.map(run => [run.text, run.color, run.bold, run.italic]),
    [['2026/', '#aa0000', true, false], ['11/22', '#0000aa', false, true]])
  assert.equal(documentStoryFieldCache(refreshed[0]).runs[1].href, 'https://example.org/date')
  assert.equal(runText(documentStoryFieldCache(refreshed[1]).runs), '01:04:05 pm')
  assert.equal(documentStoryFieldCache(refreshed[1]).runs.at(-1).href, 'https://example.org/time')
  assert.equal(refreshed[2], other); assert.equal(refreshed[3], unknown)
  assert.equal(refreshed[0].instruction, dateField.instruction); assert.equal(refreshed[0].offset, dateField.offset)
  assert.deepEqual(refreshed[0].metadata.source, dateField.metadata.source)
  assert.equal(storyBuffer({ ...section, pageFields: refreshed }, 'header').value, storyBuffer(section, 'header').value)
  assert.deepEqual(section, before)
  const same = refreshDocumentStoryDateFields({ ...section, pageFields: refreshed }, 'header', 'en-US', date, model)
  assert.ok(same.every((value, index) => value === refreshed[index]))
  assert.ok(refreshDocumentStoryDateFields(section, 'header', 'en_US', date, model).every((value, index) => value === section.pageFields[index]))
  assert.throws(() => refreshDocumentStoryDateFields({ ...section, storyLinkToPrevious: { header: true } }, 'header', 'en-US', date, model), /取消.*链接/)
})
