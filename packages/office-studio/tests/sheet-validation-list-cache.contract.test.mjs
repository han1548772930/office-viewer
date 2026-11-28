import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { SheetValidationListReadCache } from '../src/features/sheet-validation/validation-list-cache.ts'
import { serializeValidationListItems } from '../src/features/sheet-validation/validation-engine.ts'

// Implementation-phase retention/permission drafts. No test runner has been executed.
const area = (row, col, endRow = row, endCol = col) => ({ startRow: row, startCol: col, endRow, endCol })
function sample(limits) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'list-cache'; workbook.activeSheetId = sheet.id = 'data'; sheet.name = 'Data'; workbook.calculationMode = 'manual'
  sheet.cells['0:3'] = createCell(0, 3, 'Red'); sheet.cells['1:3'] = createCell(1, 3, 'Blue')
  sheet.cells['0:4'] = createCell(0, 4, true)
  const rule = { id: 'choices', type: 'listMultiple', ranges: [area(0, 0)],
    formula1: serializeValidationListItems(['Red', 'Green, dark', 'red', 'It"s blue']), formula2: '#ff0000,#008000,#ffffff,#0000ff' }
  const cache = new SheetValidationListReadCache(limits)
  return { workbook, sheet, rule, cache, read: (value = rule, user = 'local', region = 'zh-CN') => cache.read(value, workbook, sheet.id, user, region) }
}

test('inline results survive unrelated snapshot revisions and retain candidate spelling, CSV labels and colors', () => {
  const h = sample(), before = cloneModel(h.workbook), first = h.read()
  assert.deepEqual(first, [{ value: 'Red', label: 'Red', color: '#ff0000' },
    { value: 'Green, dark', label: 'Green, dark', color: '#008000' }, { value: 'It"s blue', label: 'It"s blue', color: '#0000ff' }])
  const revised = cloneModel(h.workbook); revised.sheets[0].cells['20:10'] = createCell(20, 10, 'unrelated revision')
  const second = h.cache.read(cloneModel(h.rule), revised, h.sheet.id, 'local', 'zh-CN')
  assert.equal(second, first); assert.equal(h.cache.retained.rules, 1); assert.deepEqual(h.workbook, before)
  assert.equal(Object.isFrozen(first), true); assert.equal(first.every(Object.isFrozen), true)
  assert.throws(() => first.push({ value: 'injected', label: 'injected' }), TypeError)
  assert.throws(() => { first[0].label = 'injected' }, TypeError)
})

test.each(['labels', 'colors', 'type', 'id'])('a changed %s definition cannot reuse old inline candidates', change => {
  const h = sample(), first = h.read(), next = cloneModel(h.rule)
  if (change === 'labels') next.formula1 = serializeValidationListItems(['Approved', 'Green, dark'])
  if (change === 'colors') next.formula2 = '#0000ff,#ff0000'
  if (change === 'type') next.type = 'list'
  if (change === 'id') next.id = 'another-rule'
  const second = h.read(next)
  assert.notEqual(second, first)
  if (change === 'labels') assert.equal(second[0].label, 'Approved')
  if (change === 'colors') assert.equal(second[0].color, '#0000ff')
  assert.equal(h.read(next), second)
})

test.each(['unit', 'sheet', 'user', 'region'])('a %s change away and back permanently evicts the old scope', change => {
  const h = sample(), first = h.read(), other = cloneModel(h.workbook)
  if (change === 'unit') { other.id = 'another-unit'; h.cache.read(h.rule, other, h.sheet.id, 'local', 'zh-CN') }
  if (change === 'sheet') h.cache.read(h.rule, other, 'another-sheet', 'local', 'zh-CN')
  if (change === 'user') h.read(h.rule, 'peer')
  if (change === 'region') h.read(h.rule, 'local', 'de-DE')
  const reopened = h.read()
  assert.notEqual(reopened, first); assert.deepEqual(reopened, first); assert.equal(h.cache.retained.rules, 1)
})

test('the rule/byte LRU evicts the least recently used signature while preserving a touched entry', () => {
  const h = sample({ maxRules: 2, maxBytes: 4096 }), a = { ...h.rule, id: 'a' }, b = { ...h.rule, id: 'b' }, c = { ...h.rule, id: 'c' }
  const first = h.read(a), second = h.read(b)
  assert.equal(h.read(a), first); h.read(c); assert.equal(h.read(a), first)
  assert.notEqual(h.read(b), second); assert.equal(h.cache.retained.rules, 2); assert.ok(h.cache.retained.estimatedBytes <= 4096)
  h.cache.clear(); assert.deepEqual(h.cache.retained, { rules: 0, estimatedBytes: 0 })
})

test.each([{ maxBytes: 1 }, { maxOptionsPerRule: 2 }])('admission failure %j skips retention without truncating or mutating the complete candidate list', limits => {
  const h = sample(limits), first = h.read(), next = h.read()
  assert.equal(first.length, 3); assert.deepEqual(first, next); assert.notEqual(first, next)
  assert.equal(Object.isFrozen(first), true); assert.equal(first.every(Object.isFrozen), true)
  assert.deepEqual(h.cache.retained, { rules: 0, estimatedBytes: 0 })
})

test.each(['=$D$1:$D$2', '=IF($E$1,$D$1,$D$2)'])('referenced %s sources re-read live values and View permissions instead of persisting a stale result', formula => {
  const h = sample(), rule = { ...h.rule, formula1: formula }, first = h.read(rule)
  assert.equal(first[0].label, 'Red')
  h.sheet.cells['0:3'].value = 'Approved'; assert.equal(h.read(rule)[0].label, 'Approved')
  h.sheet.rangeProtectionRules = [{ id: 'private-source', ranges: [area(0, 3)], ownerUserId: 'peer', allowedUserIds: [],
    viewState: 'noOneElseCanView', editState: 'onlyMe' }]
  assert.deepEqual(h.read(rule), [])
  assert.equal(h.read(rule, 'peer')[0].label, 'Approved')
  h.sheet.rangeProtectionRules = []; h.sheet.cells['0:4'].value = false
  if (formula.startsWith('=IF')) assert.equal(h.read(rule)[0].label, 'Blue')
  assert.deepEqual(h.cache.retained, { rules: 0, estimatedBytes: 0 }); assert.equal(Object.isFrozen(first), true)
})

test('a cache clear during candidate extraction cannot republish the old signature', () => {
  const h = sample(); let reads = 0
  Object.defineProperty(h.rule, 'formula2', { configurable: true, get: () => { if (++reads === 2) h.cache.clear(); return '#ff0000' } })
  const old = h.read(); assert.equal(old[0].color, '#ff0000'); assert.equal(h.cache.retained.rules, 0)
  Object.defineProperty(h.rule, 'formula2', { configurable: true, writable: true, value: '#008000' })
  const next = h.read(); assert.notEqual(next, old); assert.equal(next[0].color, '#008000'); assert.equal(h.cache.retained.rules, 1)
})

test('a scope switch inside the signature getter cannot retain the interrupted result in the successor scope', () => {
  const h = sample(), foreign = { ...h.rule, id: 'foreign', formula1: '"Other"' }; let once = true, retained
  Object.defineProperty(h.rule, 'formula2', { get: () => {
    if (once) { once = false; retained = h.cache.read(foreign, h.workbook, h.sheet.id, 'peer', 'de-DE') }
    return '#ff0000'
  } })
  assert.equal(h.read()[0].label, 'Red')
  assert.equal(h.cache.read(foreign, h.workbook, h.sheet.id, 'peer', 'de-DE'), retained)
  assert.equal(h.cache.retained.rules, 1)
})

test.each([{ maxRules: 0 }, { maxBytes: -1 }, { maxOptionsPerRule: 1.5 }, { maxRules: Number.MAX_SAFE_INTEGER + 1 }])('invalid cache budget %j fails before retaining data', limits => {
  assert.throws(() => new SheetValidationListReadCache(limits), RangeError)
})
