import assert from 'node:assert/strict'
import { test } from 'vitest'
import { conditionalFormatOfficeBaseline, normalizeConditionalFormatOfficeSource } from '../src/features/conditional-formatting/conditional-format-office-source.ts'
import { assertConditionalFormatModelBudget, conditionalFormatOfficeFragmentByteLimit,
  conditionalFormatSerializedSize, conditionalFormatUtf8Size } from '../src/features/conditional-formatting/conditional-format-source-budget.ts'

// These declarations are drafts; execution is deferred to the user's unified verification pass.
const utf8 = value => new TextEncoder().encode(value).byteLength
const jsonBytes = value => utf8(JSON.stringify(value))
const totalLimit = error => error instanceof RangeError && /CF_SOURCE_TOTAL_LIMIT/.test(error.message)
const invalidJson = error => error instanceof TypeError && /CF_SOURCE_XML_INVALID/.test(error.message)
const rule = (id = 'source-budget-rule') => ({ id, type: 'cellIs',
  ranges: [{ startRow: 0, startCol: 0, endRow: 3, endCol: 0 }], operator: 'greaterThan', formula1: '0',
  style: { bold: true }, priority: 17, stopIfTrue: false })
const source = (id = 'source-budget-rule') => ({ version: 1, format: 'xlsx', originalId: id, storage: 'base',
  ruleXml: '<cfRule xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" type="cellIs" operator="greaterThan" priority="17"><formula>0</formula></cfRule>',
  groupXml: '<conditionalFormatting xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" sqref="A1:A4"/>',
  baseline: conditionalFormatOfficeBaseline(rule(id)) })

test('UTF-8 byte counts match TextEncoder for Unicode, surrogate pairs and lone surrogates at capped boundaries', () => {
  const samples = ['', '\u0000a\u007f', '\u0080\u07ff\u0800\uffff', '中文🙂', '\ud800', '\udc00',
    '\ud800X\udc00', '\ud83d\ude42\ud800', '\ud800\ud800\udc00']
  for (const sample of samples) {
    const expected = utf8(sample)
    assert.equal(conditionalFormatUtf8Size(sample), expected)
    for (let limit = 0; limit <= expected + 1; limit++) {
      assert.equal(conditionalFormatUtf8Size(sample, limit), expected > limit ? limit + 1 : expected)
    }
  }
})

test('serialized bytes match JSON.stringify for escapes, omitted fields, sparse arrays and nonfinite extension numbers', () => {
  const samples = [null, true, false, 0, -0, 1e-7, 1e21, Number.NaN, Number.POSITIVE_INFINITY,
    '"\\\b\t\n\f\r\u0000中文🙂\ud800\udc00\ud800',
    { '中文🙂\n': '"\\\u001f', missing: undefined, fn: () => 1, symbol: Symbol('ignored'),
      vendor: { nan: Number.NaN, positive: Number.POSITIVE_INFINITY, negative: Number.NEGATIVE_INFINITY } },
    [undefined, () => 1, Symbol('array null'), , { optional: undefined, value: '保留🙂' }]]
  for (const sample of samples) {
    const expected = jsonBytes(sample)
    assert.equal(conditionalFormatSerializedSize(sample), expected)
    assert.equal(conditionalFormatSerializedSize(sample, expected), expected)
    assert.throws(() => conditionalFormatSerializedSize(sample, expected - 1), totalLimit)
  }
})

test('pure toJSON conversions retain keyed values and omit properties with the same representation as JSON.stringify', () => {
  const shared = { text: '共享🙂' }
  const keyedFunction = () => 'not invoked'; keyedFunction.toJSON = key => `函数:${key}🙂`
  const sample = { stamp: new Date('2026-11-23T00:00:00.000Z'),
    keyed: { toJSON: key => `键:${key}` }, omitted: { toJSON: () => undefined },
    convertedFunction: keyedFunction, array: [{ toJSON: key => `索引:${key}` }, { toJSON: () => undefined }, keyedFunction], first: shared, second: shared,
    dictionary: Object.assign(Object.create(null), { '空原型': true }) }
  assert.equal(conditionalFormatSerializedSize(sample), jsonBytes(sample))
  assert.throws(() => conditionalFormatSerializedSize(sample, jsonBytes(sample) - 1), totalLimit)
  assert.equal(conditionalFormatSerializedSize({ toJSON: key => `根:${key}` }), jsonBytes({ toJSON: key => `根:${key}` }))
  assert.equal(conditionalFormatSerializedSize(keyedFunction), jsonBytes(keyedFunction))
  const propertyArray = []; propertyArray.toJSON = key => key === 'conditionalFormats' ? '文'.repeat(64) : []
  const serializedProperty = JSON.parse(JSON.stringify({ conditionalFormats: propertyArray })).conditionalFormats
  const propertyBytes = jsonBytes(serializedProperty)
  assert.equal(conditionalFormatSerializedSize(propertyArray, propertyBytes, 'conditionalFormats'), propertyBytes)
  assert.throws(() => conditionalFormatSerializedSize(propertyArray, propertyBytes - 1, 'conditionalFormats'), totalLimit)
})

test('aggregate budgets include array brackets and commas at the exact serialized boundary', () => {
  assert.doesNotThrow(() => assertConditionalFormatModelBudget([], 2))
  assert.throws(() => assertConditionalFormatModelBudget([], 1), totalLimit)
  const keyed = { toJSON: key => ({ id: key, vendor: key ? '文'.repeat(64) : 'short root' }) }
  for (const values of [[{ rule: '甲' }, { rule: '🙂' }], [keyed, keyed]]) {
    const exact = jsonBytes(values)
    assert.doesNotThrow(() => assertConditionalFormatModelBudget(values, exact))
    assert.throws(() => assertConditionalFormatModelBudget(values, exact - 1), totalLimit)
  }
})

test('two individually bounded Office sources reject a shared budget without changing either input', () => {
  const first = { ...source('first'), vendor: { payload: '第一份🙂' } }
  const second = { ...source('second'), vendor: { payload: '第二份中文' } }
  const values = [first, second], before = JSON.stringify(values)
  const limit = Math.max(jsonBytes([first]), jsonBytes([second]))
  for (const item of values) assert.doesNotThrow(() => assertConditionalFormatModelBudget([item], limit))
  assert.throws(() => assertConditionalFormatModelBudget(values, limit), totalLimit)
  assert.equal(JSON.stringify(values), before)
  assert.equal(values[0], first); assert.equal(values[1], second)
  assert.doesNotThrow(() => assertConditionalFormatModelBudget(values, jsonBytes(values)))
})

test('iterable rules spend one cumulative budget before a later source can be accepted', () => {
  const first = { ...rule('first'), officeSource: source('first') }
  const second = { ...rule('second'), officeSource: source('second'), vendor: '后一份🙂' }
  function* values() { yield first; yield second }
  const before = JSON.stringify([first, second])
  assert.throws(() => assertConditionalFormatModelBudget(values(), jsonBytes([first])), totalLimit)
  assert.doesNotThrow(() => assertConditionalFormatModelBudget(values(), jsonBytes([first, second])))
  assert.equal(JSON.stringify([first, second]), before)
})

test('an XML fragment below the UTF-16 character limit rejects when its actual UTF-8 bytes exceed 16 MiB', () => {
  const original = source(), payload = '汉'.repeat(Math.floor(conditionalFormatOfficeFragmentByteLimit / 3) + 1)
  original.ruleXml = original.ruleXml.replace('</cfRule>', `<!--${payload}--></cfRule>`)
  assert.ok(original.ruleXml.length < conditionalFormatOfficeFragmentByteLimit)
  assert.ok(Buffer.byteLength(original.ruleXml, 'utf8') > conditionalFormatOfficeFragmentByteLimit)
  const xml = original.ruleXml, baseline = original.baseline
  assert.throws(() => normalizeConditionalFormatOfficeSource(original),
    error => error instanceof RangeError && /CF_SOURCE_FRAGMENT_LIMIT/.test(error.message))
  assert.equal(original.ruleXml, xml); assert.equal(original.baseline, baseline)
})

test('originalId accepts exactly 4096 UTF-8 bytes and rejects a shorter Unicode string over that byte budget', () => {
  const original = source(), boundary = `${'汉'.repeat(1365)}a`, overflow = '汉'.repeat(1366)
  assert.equal(utf8(boundary), 4096); assert.ok(overflow.length < 4096); assert.equal(utf8(overflow), 4098)
  assert.equal(normalizeConditionalFormatOfficeSource({ ...original, originalId: boundary }).originalId, boundary)
  assert.throws(() => normalizeConditionalFormatOfficeSource({ ...original, originalId: overflow }), TypeError)
  assert.equal(original.originalId, 'source-budget-rule')
})

test('version 1 source extensions remain counted and isolated while unknown baseline fields are filtered', () => {
  const original = { ...source(), vendor: { labels: ['中文🙂'], optional: undefined },
    baseline: { ...source().baseline, vendor: { uninspectable: true }, style: { bold: true, vendor: 'excluded' } } }
  const before = JSON.stringify(original), normalized = normalizeConditionalFormatOfficeSource(original)
  assert.equal(normalized.version, 1); assert.deepEqual(normalized.vendor, { labels: ['中文🙂'] })
  assert.equal(normalized.baseline.vendor, undefined); assert.deepEqual(normalized.baseline.style, { bold: true })
  assert.equal(conditionalFormatSerializedSize(normalized), jsonBytes(normalized))
  assert.doesNotThrow(() => assertConditionalFormatModelBudget([normalized], jsonBytes([normalized])))
  const withoutExtensions = { ...normalized }; delete withoutExtensions.vendor
  assert.throws(() => assertConditionalFormatModelBudget([normalized], jsonBytes([withoutExtensions])), totalLimit)
  normalized.vendor.labels.push('copy'); normalized.baseline.ranges[0].endRow = 9
  assert.equal(JSON.stringify(original), before)
})

test('cyclic, BigInt and excessively deep source extensions fail before source normalization clones the envelope', () => {
  const cyclic = {}; cyclic.self = cyclic
  const deep = {}; let cursor = deep
  for (let depth = 0; depth < 130; depth++) { cursor.child = {}; cursor = cursor.child }
  for (const vendor of [cyclic, { value: 1n }, deep]) {
    const original = { ...source(), vendor }, baseline = original.baseline, xml = original.ruleXml
    assert.throws(() => normalizeConditionalFormatOfficeSource(original), invalidJson)
    assert.equal(original.vendor, vendor); assert.equal(original.baseline, baseline); assert.equal(original.ruleXml, xml)
  }
})

test('non-JSON domain values reject rather than being counted as an empty object or absent root', () => {
  for (const value of [undefined, Object.create({ inherited: true }), new Map([['vendor', 1]]), /cf/u]) {
    assert.throws(() => conditionalFormatSerializedSize(value), invalidJson)
  }
})
