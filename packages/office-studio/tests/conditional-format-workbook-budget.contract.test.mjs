import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankSheet, createBlankWorkbook } from '../src/domain/model.ts'
import { ensureWorkbookId, workbookFromInput } from '../src/api/model-utils.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { OfficeStudio } from '../src/api/office-studio.ts'
import { assertWorkbookConditionalFormatBudget } from '../src/features/conditional-formatting/conditional-format-workbook-budget.ts'

// Execution is deferred to the user's unified verification pass.
const rule = (vendor = '') => ({ id: 'bounded-rule', type: 'cellIs', ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
  operator: 'greaterThan', formula1: '0', style: { bold: true }, priority: 1, stopIfTrue: false, vendor })
const envelope = sheets => ({ schemaVersion: 1, parserVersion: 'draft', format: 'xlsx', fileName: 'cf-budget.xlsx', size: 10,
  model: 'spreadsheet', payload: { modelType: 'spreadsheet', data: { sheets } },
  diagnostics: { warnings: [], errors: [], stats: {}, capabilities: ['conditionalFormatting'] } })

test('array sheets and snapshot sheet maps count all rules with one shared UTF-8 budget', () => {
  const first = rule('中文🙂'), second = rule('same extension on another sheet')
  const firstLimit = Buffer.byteLength(JSON.stringify([first]), 'utf8')
  assert.doesNotThrow(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: [first] }], firstLimit))
  const arrays = [{ conditionalFormats: [first] }, { conditionalFormats: [] }, { conditionalFormats: [second] }]
  const maps = { first: arrays[0], empty: arrays[1], second: arrays[2] }
  for (const sheets of [arrays, maps]) {
    assert.throws(() => assertWorkbookConditionalFormatBudget(sheets, firstLimit), error => error instanceof RangeError && /CF_SOURCE_TOTAL_LIMIT/.test(error.message))
    assert.doesNotThrow(() => assertWorkbookConditionalFormatBudget(sheets, Buffer.byteLength(JSON.stringify([first, second]), 'utf8')))
  }
})

test('sheet-local rule conversion keys match snapshot serialization while empty sheets add no logical separators', () => {
  const keyed = { toJSON: key => rule(key === '0' ? '文'.repeat(64) : 'short global index') }
  const sheets = [{ conditionalFormats: [keyed] }, { conditionalFormats: [] }, { conditionalFormats: [keyed, keyed] }]
  const represented = sheets.flatMap(sheet => JSON.parse(JSON.stringify({ conditionalFormats: sheet.conditionalFormats })).conditionalFormats)
  const exact = Buffer.byteLength(JSON.stringify(represented), 'utf8')
  assert.ok(exact > Buffer.byteLength(JSON.stringify(sheets.flatMap(sheet => sheet.conditionalFormats)), 'utf8'))
  for (const input of [sheets, { first: sheets[0], empty: sheets[1], last: sheets[2] }]) {
    assert.doesNotThrow(() => assertWorkbookConditionalFormatBudget(input, exact))
    assert.throws(() => assertWorkbookConditionalFormatBudget(input, exact - 1), /CF_SOURCE_TOTAL_LIMIT/)
  }
  const normalized = normalizeWorkbook(envelope(sheets.map((sheet, index) => ({ ...sheet, name: 'Keyed ' + index, cells: [] }))))
  assert.deepEqual(normalized.sheets.flatMap(sheet => sheet.conditionalFormats.map(rule => rule.vendor)), represented.map(rule => rule.vendor))
})

test('rule-array conversion uses its property context once and must retain an array', () => {
  const rules = [rule('raw')]
  const replacement = [{ toJSON: key => rule('保留:' + key) }, undefined]
  let calls = 0
  rules.toJSON = key => { calls++; return key === 'conditionalFormats' ? replacement : [] }
  const represented = JSON.parse(JSON.stringify({ conditionalFormats: rules })).conditionalFormats
  const exact = Buffer.byteLength(JSON.stringify(represented), 'utf8')
  calls = 0
  assert.doesNotThrow(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: rules }], exact))
  assert.equal(calls, 1)
  assert.throws(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: rules }], exact - 1), /CF_SOURCE_TOTAL_LIMIT/)
  rules.toJSON = key => key === 'conditionalFormats' ? [rule('property-context')] : []
  for (const input of [{ sheets: [{ conditionalFormats: rules }] }, { sheets: { first: { conditionalFormats: rules } } }]) {
    assert.deepEqual(workbookFromInput(input).sheets[0].conditionalFormats, [rule('property-context')])
  }
  const normalized = normalizeWorkbook(envelope([{ name: 'Property array', cells: [], conditionalFormats: rules }]))
  assert.equal(normalized.sheets[0].conditionalFormats[0].vendor, 'property-context')
  for (const value of [undefined, null, 'discarded rules', { vendor: 'object' }]) {
    rules.toJSON = () => value
    assert.throws(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: rules }]), /JSON 转换结果必须是数组/)
  }
})

test('present malformed rule lists reject before parser and internal snapshot clones', () => {
  for (const conditionalFormats of [null, {}, 'rules', 42]) {
    const sheets = [{ name: 'Malformed', cells: [], conditionalFormats }]
    assert.throws(() => normalizeWorkbook(envelope(sheets)), /规则必须是数组/)
    assert.throws(() => workbookFromInput({ sheets }), /规则必须是数组/)
    assert.throws(() => workbookFromInput({ sheets: { malformed: sheets[0] } }), /规则必须是数组/)
  }
})

test('partial array/map snapshots preserve conditional rules and isolate retained vendor data', () => {
  const source = rule({ retained: ['original'] })
  for (const input of [{ sheets: [{ id: 'budget-sheet', name: 'Data', conditionalFormats: [source] }] },
    { sheetOrder: ['budget-sheet'], sheets: { 'budget-sheet': { name: 'Data', conditionalFormats: [source] } } }]) {
    const model = workbookFromInput(input)
    assert.deepEqual(model.sheets[0].conditionalFormats, [source])
    model.sheets[0].conditionalFormats[0].vendor.retained.push('isolated')
    assert.deepEqual(source.vendor.retained, ['original'])
  }
})

test('multi-sheet oversized sources fail before any snapshot or parser metadata cloning', () => {
  // Shared strings keep this input compact; serialization would allocate over 256 MiB.
  const source = rule('x'.repeat(1024 * 1024)), rules = Array(129).fill(source)
  let clones = 0
  const metadata = { toJSON: () => { clones++; return { inspected: true } } }
  const sheets = [0, 1].map(index => ({ ...createBlankSheet('Sheet ' + index), conditionalFormats: rules }))
  const snapshot = { ...createBlankWorkbook(), sheets, metadata }
  for (const read of [() => ensureWorkbookId(snapshot), () => workbookFromInput(snapshot),
    () => normalizeWorkbook(envelope(sheets.map(sheet => ({ ...sheet, cells: [], properties: metadata }))))]) {
    assert.throws(read, error => error instanceof RangeError && /CF_SOURCE_TOTAL_LIMIT/.test(error.message))
    assert.equal(clones, 0)
  }
})

test('repeated references count each retained instance while cycles and non-JSON source data reject', () => {
  const shared = rule('shared'), singleLimit = Buffer.byteLength(JSON.stringify([shared]), 'utf8')
  assert.throws(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: [shared, shared] }], singleLimit), /CF_SOURCE_TOTAL_LIMIT/)
  const circular = rule(); circular.vendor = circular
  assert.throws(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: [circular] }]), TypeError)
  assert.throws(() => assertWorkbookConditionalFormatBudget([{ conditionalFormats: [rule(1n)] }]), TypeError)
})

test('cross-sheet command budget rejection preserves the authoritative model, revision, selection and history', t => {
  const workbook = createBlankWorkbook(); workbook.calculationMode = 'manual'
  workbook.sheets.push(createBlankSheet('Second'))
  const studio = new OfficeStudio({ workbook, sheetTextMeasure: () => 0 })
  t.onTestFinished(() => studio.dispose())
  const id = studio.activeWorkbookId, before = studio.getWorkbookModel(id), selection = studio.getSelectionState()
  const revision = studio.getWorkbookRevision(id), stack = studio.editor.history.units[id]
  const undoDepth = stack?.undo.length ?? 0, redoDepth = stack?.redo.length ?? 0
  const source = rule('x'.repeat(1024 * 1024)), rules = Array(129).fill(source)
  assert.throws(() => studio.commitWorkbook(id, 'Cross-sheet oversized conditional rules', draft => {
    for (const sheet of draft.sheets) sheet.conditionalFormats = rules
  }), /CF_SOURCE_TOTAL_LIMIT/)
  assert.equal(studio.getWorkbookModel(id), before); assert.equal(studio.getWorkbookRevision(id), revision)
  assert.deepEqual(studio.getSelectionState(), selection)
  assert.equal(studio.editor.history.units[id]?.undo.length ?? 0, undoDepth)
  assert.equal(studio.editor.history.units[id]?.redo.length ?? 0, redoDepth)
  assert.equal(before.modified, false)
})
