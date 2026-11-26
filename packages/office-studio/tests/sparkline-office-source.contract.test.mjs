import assert from 'node:assert/strict'
import { test } from 'vitest'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { normalizeOfficeSparklineSource } from '../src/features/sparkline/sparkline-office-source.ts'
import { createSparklineSpec, planSparklineMutation, translateSparklineForMove } from '../src/features/sparkline/sparkline-model.ts'
import { deserializeWorkbookSparklineResource, importWorkbookSparklineResources, serializeWorkbookSparklineResource, SHEET_SPARKLINE_PLUGIN } from '../src/features/sparkline/sparkline-resources.ts'
import { installSparklinePatches, sparklineCell, sparklineJson, sparklineWorkbook } from './fixtures/sparkline-workbench-fixture.mjs'

// Deferred metadata, planner and persistence contracts. No tests or checks were run in this batch.
const source = label => ({ version: 1,
  groupXml: `<x14:sparklineGroup xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" futurePolicy="${label}"><x14:colorSeries theme="4" tint="0.5"/></x14:sparklineGroup>`,
  resolvedColors: { colorSeries: '#FF8080', colorNegative: 'rgba(0,0,255,0.5019607843137255)' },
  themeXml: '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements/></a:theme>',
  indexedColors: ['ff000000', 'ffffffff', 'ff102030'] })
const admittedSource = label => normalizeOfficeSparklineSource(source(label))
function workbook() {
  const book = sparklineWorkbook()
  for (const address of ['G1', 'G2', 'G4', 'G5']) sparklineCell(book.sheets[0], address).sparkline.officeSource = admittedSource(address)
  return book
}
const plan = (book, request) => planSparklineMutation(book.sheets[0], request, { sheets: book.sheets, createGroupId: () => 'new-authored-group' })

test('Office source normalization preserves inert XML, color bindings and a detached uppercase palette', () => {
  const input = source('author'), output = normalizeOfficeSparklineSource(input)
  assert.deepEqual(output, { ...input, indexedColors: ['FF000000', 'FFFFFFFF', 'FF102030'] })
  input.resolvedColors.colorSeries = '#000000'; input.indexedColors[2] = 'FF556677'
  assert.equal(output.resolvedColors.colorSeries, '#FF8080'); assert.equal(output.indexedColors[2], 'FF102030')
})

test('Office source admission rejects unknown versions, fields and invalid palettes before model publication', () => {
  for (const patch of [{ version: 2 }, { future: true }, { resolvedColors: [] }, { resolvedColors: { unknownColor: '#fff' } },
    { indexedColors: ['102030'] }, { indexedColors: new Array(257).fill('FF102030') }, { groupXml: '' }, { themeXml: null }]) {
    assert.throws(() => normalizeOfficeSparklineSource({ ...source('author'), ...patch }))
  }
  assert.equal(normalizeOfficeSparklineSource(undefined), undefined)
})

test('multibyte source XML obeys its UTF8 budget rather than only its UTF16 character count', () => {
  const xml = `<g>${'中'.repeat(44_000)}</g>`
  assert.ok(xml.length < 128 * 1024)
  assert.throws(() => normalizeOfficeSparklineSource({ ...source('author'), groupXml: xml }), /预算/)
})

test('parser DTO becomes an editable sparkline without dropping its original theme or template', () => {
  const raw = createSparklineSpec({ type: 'line', color: '#FF8080', sourceRange: "'Data, O''Neil'!$B$1:E$1", targetRange: 'G1',
    officeSource: source('parser'), groupId: 'parsed', lineWidth: 2.25, points: { negativePoint: { visible: false, color: 'rgba(0,0,255,0.5)' } } })
  const parsed = normalizeWorkbook({ schemaVersion: 1, parserVersion: 'deferred', format: 'xlsx', fileName: 'sparkline-source.xlsx', size: 10, model: 'spreadsheet',
    payload: { modelType: 'spreadsheet', data: { activeSheet: 0, date1904: false, sheets: [{ name: 'Trend', cells: [{ row: 0, col: 6, value: null, display: '', dataType: 'blank', sparkline: raw }] }] } },
    diagnostics: { warnings: [], errors: [], capabilities: ['sparklines'], stats: {} } })
  assert.deepEqual(parsed.sheets[0].cells['0:6'].sparkline.officeSource, admittedSource('parser'))
  assert.equal(parsed.sheets[0].cells['0:6'].sparkline.lineWidth, 2.25)
  assert.equal(parsed.sheets[0].cells['0:6'].sparkline.sourceRange, "'Data, O''Neil'!$B$1:E$1")
})

test('whole-group settings keep each member native template and do not mutate the live candidate', () => {
  const book = workbook(), before = sparklineJson(book)
  const result = plan(book, { kind: 'settings', ranges: ['G1'], settings: { lineWidth: 7.5, points: { firstPoint: { visible: true } } } })
  assert.deepEqual(Object.keys(result.patches), ['G1', 'G2'])
  for (const address of ['G1', 'G2']) {
    assert.deepEqual(result.patches[address].officeSource, admittedSource(address))
    assert.equal(result.patches[address].lineWidth, 7.5)
    assert.notEqual(result.patches[address].officeSource, sparklineCell(book.sheets[0], address).sparkline.officeSource)
  }
  assert.deepEqual(book, before)
})

test('combining members preserves per-member retained Office metadata while applying one author configuration', () => {
  const book = workbook(), result = plan(book, { kind: 'combine', ranges: ['G1', 'G4'], settings: { type: 'line' } })
  assert.equal(result.patches.G1.groupId, 'new-authored-group'); assert.equal(result.patches.G4.groupId, 'new-authored-group')
  assert.deepEqual(result.patches.G1.officeSource, admittedSource('G1')); assert.deepEqual(result.patches.G4.officeSource, admittedSource('G4'))
})

test('uncombining and relative move keep the source template and absolute reference semantics', () => {
  const book = workbook(), result = plan(book, { kind: 'uncombine', ranges: ['G1'] })
  assert.deepEqual(result.patches.G1.officeSource, admittedSource('G1'))
  const input = { ...result.patches.G1, sourceRange: '$B1:E$1' }, shifted = translateSparklineForMove(input, 2, 1)
  assert.equal(shifted.sourceRange, '$B3:F$1'); assert.deepEqual(shifted.officeSource, input.officeSource)
})

test('group source repair and relocation carry native templates in stable member order', () => {
  const book = workbook(), result = plan(book, { kind: 'change-source', ranges: ['G1'], resetType: 'group', sourceRanges: ['B3:E4'], targetRanges: ['H1:H2'] })
  assert.equal(result.patches.G1, null); assert.equal(result.patches.G2, null)
  assert.equal(result.patches.H1.sourceRange, 'B3:E3'); assert.equal(result.patches.H2.sourceRange, 'B4:E4')
  assert.deepEqual(result.patches.H1.officeSource, admittedSource('G1')); assert.deepEqual(result.patches.H2.officeSource, admittedSource('G2'))
})

test('item source repair plus shared settings preserve both moved and unchanged group members', () => {
  const book = workbook(), result = plan(book, { kind: 'change-source', ranges: ['G1'], resetType: 'item', sourceRanges: ['B3:E3'], targetRanges: ['I1'], settings: { lineWidth: 2.25 } })
  assert.deepEqual(result.patches.I1.officeSource, admittedSource('G1')); assert.deepEqual(result.patches.G2.officeSource, admittedSource('G2'))
  assert.equal(result.patches.G2.sourceRange, 'B2:E2'); assert.equal(result.patches.G2.lineWidth, 2.25)
})

test('public plugin resource roundtrip preserves every member Office template and retained source spelling', () => {
  const book = workbook(); sparklineCell(book.sheets[0], 'G1').sparkline.sourceRange = '$B$1:E$1'
  const resource = serializeWorkbookSparklineResource(book), detached = sparklineWorkbook({ withGroups: false })
  const restored = deserializeWorkbookSparklineResource(detached, resource)
  for (const address of ['G1', 'G2', 'G4', 'G5']) assert.deepEqual(restored.data[address].officeSource, admittedSource(address))
  assert.equal(restored.data.G1.sourceRange, '$B$1:E$1')
  importWorkbookSparklineResources(detached, [{ name: SHEET_SPARKLINE_PLUGIN, data: JSON.stringify(resource) }])
  assert.deepEqual(serializeWorkbookSparklineResource(detached), resource)
  const candidate = sparklineJson(detached); installSparklinePatches(candidate.sheets[0], restored.data)
  assert.deepEqual(sparklineCell(candidate.sheets[0], 'G2').sparkline.officeSource, admittedSource('G2'))
})

test('invalid late member metadata rejects a plugin import without partially changing existing groups', () => {
  const book = workbook(), resource = serializeWorkbookSparklineResource(book), before = sparklineJson(book)
  resource[book.id].data.south.sparklines[4][6].officeStudioOfficeSource.version = 2
  assert.throws(() => importWorkbookSparklineResources(book, [{ name: SHEET_SPARKLINE_PLUGIN, data: JSON.stringify(resource) }]))
  assert.deepEqual(book, before)
})
