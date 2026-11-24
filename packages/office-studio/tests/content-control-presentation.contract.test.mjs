import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-11 implementation drafts. These assertions have not been run or accepted.
let server, models, modelUtils, controls, presentation, properties, commandState
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/content-control-presentation', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  models = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  presentation = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-control-presentation.ts')
  properties = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-control-properties.ts')
  commandState = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/content-control-command-state.ts')
})
afterAll(async () => { await server?.close() })

function fixture(text = 'NAME') {
  const model = models.createBlankDocument()
  model.id = 'content-control-presentation'; model.blocks[0].id = 'p'
  model.blocks[0].runs[0].text = text; model.contentControls = []
  return model
}
const text = model => model.blocks[0].runs.map(run => run.text).join('')
const info = (model, id = 'control') => controls.getDocumentContentControl(model, id)
function add(model, input = {}) {
  return controls.createDocumentContentControl(model, { id: 'control', type: 'text', blockId: 'p', startOffset: 0, endOffset: text(model).length, ...input })
}
function rejected(model, action, message) {
  const before = structuredClone(model)
  assert.throws(action, message)
  assert.deepEqual(model, before, 'a rejected control edit must not publish partial text, fields or properties')
}

test('checkbox creation paints a Unicode symbol, applies its font and shifts a later control', () => {
  const model = fixture('___|NEXT')
  add(model, { id: 'next', startOffset: 4, endOffset: 8 })
  add(model, { type: 'checkbox', endOffset: 3, dataBinding: { xpath: '/checked', dirty: false } })
  assert.equal(text(model), '☐|NEXT')
  assert.deepEqual(controls.contentControlRange(info(model)), { blockId: 'p', startOffset: 0, endOffset: 1 })
  assert.deepEqual(controls.contentControlRange(info(model, 'next')), { blockId: 'p', startOffset: 2, endOffset: 6 })
  assert.equal(model.blocks[0].runs[0].fontFamily, 'Segoe UI Symbol')
  controls.setDocumentContentControlChecked(model, 'control', true)
  assert.equal(text(model), '☒|NEXT')
  assert.equal(info(model).value, 'true'); assert.equal(info(model).checkbox.checked, true)
  assert.equal(info(model).dataBinding.dirty, true)
})

test('supplementary checkbox code points update UTF-16 ranges and remain reversible by later values', () => {
  const model = fixture('___|NEXT')
  add(model, { id: 'next', startOffset: 4, endOffset: 8 })
  add(model, { type: 'checkbox', endOffset: 3, checked: true })
  controls.updateDocumentContentControl(model, 'control', { checkbox: { checked: true,
    checkedState: { value: '1F7E2', font: 'Segoe UI Emoji' }, uncheckedState: { value: '2610', font: 'Segoe UI Symbol' } } })
  assert.equal(text(model), '🟢|NEXT'); assert.equal(info(model).endOffset, 2)
  assert.equal(info(model, 'next').startOffset, 3); assert.equal(model.blocks[0].runs[0].fontFamily, 'Segoe UI Emoji')
  controls.setDocumentContentControlChecked(model, 'control', false)
  assert.equal(text(model), '☐|NEXT'); assert.equal(info(model).endOffset, 1)
  assert.equal(info(model, 'next').startOffset, 2)
})

test('all four fixed-version checkbox presets preserve their exact character pair', () => {
  const expected = { Cross: ['☒', '☐'], Check: ['☑', '☐'], Minimal: ['✓', '□'], Round: ['●', '○'] }
  for (const [name, glyphs] of Object.entries(expected)) {
    const model = fixture('_')
    add(model, { type: 'checkbox', checkbox: { checked: false, ...presentation.CONTENT_CONTROL_CHECKBOX_STYLES[name] } })
    assert.equal(text(model), glyphs[1])
    controls.setDocumentContentControlChecked(model, 'control', true)
    assert.equal(text(model), glyphs[0])
    assert.equal(presentation.contentControlCheckboxStyle(info(model)), name)
  }
})

test('checkbox style recognition uses actual symbol/font pairs and ignores stale preset metadata', () => {
  const model = fixture('_'); add(model, { type: 'checkbox' })
  assert.equal(presentation.contentControlCheckboxStyle(info(model)), 'Cross')
  controls.updateDocumentContentControl(model, 'control', { checkbox: { checked: false,
    checkedState: { value: '0x002611', font: 'Custom Symbols' }, uncheckedState: { value: '2610', font: 'Custom Symbols' } }, metadata: { checkboxStyle: 'Cross' } })
  assert.equal(presentation.contentControlCheckboxStyle(info(model)), undefined)
  assert.equal(model.blocks[0].runs[0].fontFamily, 'Custom Symbols')
})

test('a wrapper lock allows property updates and protects removal independently', () => {
  const model = fixture()
  add(model, { lock: 'sdtLocked' })
  controls.updateDocumentContentControl(model, 'control', { alias: 'Customer', tag: 'customer' })
  assert.equal(info(model).alias, 'Customer'); assert.equal(info(model).lock, 'sdtLocked')
  rejected(model, () => controls.removeDocumentContentControl(model, 'control'), /wrapper is locked/)
  controls.setDocumentContentControlValue(model, 'control', 'Mira')
  assert.equal(text(model), 'Mira')
})

test('content locks reject checkbox symbols and date display changes, and an explicit unlock can apply them', () => {
  const model = fixture('_')
  add(model, { type: 'checkbox', lock: 'contentLocked' })
  rejected(model, () => controls.setDocumentContentControlChecked(model, 'control', true), /content is locked/)
  rejected(model, () => controls.updateDocumentContentControl(model, 'control', { checkbox: { checked: false,
    ...presentation.CONTENT_CONTROL_CHECKBOX_STYLES.Round } }), /content is locked/)
  controls.updateDocumentContentControl(model, 'control', { lock: 'unlocked', checkbox: { checked: false,
    ...presentation.CONTENT_CONTROL_CHECKBOX_STYLES.Round } })
  assert.equal(text(model), '○')
  const date = fixture('2026-09-03')
  add(date, { type: 'date', dateValue: '2026-09-03', lock: 'contentLocked' })
  rejected(date, () => controls.updateDocumentContentControl(date, 'control', { dateFormat: 'dd MMM yyyy' }), /content is locked/)
  controls.updateDocumentContentControl(date, 'control', { alias: 'Date' })
  assert.equal(info(date).alias, 'Date'); assert.equal(text(date), '2026-09-03')
})

test('creating a checkbox beneath a locked ancestor cannot replace that ancestor content', () => {
  const model = fixture('ABCD')
  add(model, { id: 'group', type: 'group', lock: 'contentLocked' })
  rejected(model, () => add(model, { type: 'checkbox', parentId: 'group', startOffset: 1, endOffset: 3 }), /content is locked: group/)
})

test('text and rich text values store the same LF-normalized content as their document runs', () => {
  for (const type of ['text', 'richText']) {
    const model = fixture('_')
    add(model, { type, multiLine: true })
    controls.setDocumentContentControlValue(model, 'control', 'A\r\nB\rC')
    assert.equal(text(model), 'A\nB\nC'); assert.equal(info(model).value, 'A\nB\nC')
    assert.equal(info(model).endOffset, 5)
  }
  const model = fixture('_')
  add(model)
  rejected(model, () => controls.setDocumentContentControlValue(model, 'control', 'A\nB'), /多行/)
})

test('UTF-16 ranges cannot split a surrogate pair or contain non-integer coordinates', () => {
  for (const [startOffset, endOffset] of [[1, 2], [2, 3], [0.5, 3], [0, NaN], [0, Infinity]]) {
    const model = fixture('A😀B')
    rejected(model, () => add(model, { startOffset, endOffset }), /UTF-16|range is invalid/)
  }
  const model = fixture('A😀B')
  add(model, { startOffset: 1, endOffset: 3 })
  assert.equal(controls.contentControlText(model, info(model)), '😀')
})

test('date storage uses strict Gregorian ISO while display supports localized names, literals and weekday tokens', () => {
  assert.equal(presentation.formatContentControlDate('2026-09-03', 'ddd, dd MMM yyyy', 'en-US'), 'Thu, 03 Sep 2026')
  assert.equal(presentation.formatContentControlDate('2026-09-03', 'dd MMM yyyy', 'fr-FR'), '03 sept. 2026')
  assert.equal(presentation.formatContentControlDate('2026-09-03', '"On" dddd, MMMM d, yyyy', 'en-US'), 'On Thursday, September 3, 2026')
  assert.equal(presentation.formatContentControlDate('2026-09-03', "yyyy 'year' MM\\M dd", 'en-US'), '2026 year 09M 03')
  assert.equal(presentation.formatContentControlDate('2026-09-03', 'yyyy-MM-dd', 'en-US', 'thai'), '2569-09-03')
  for (const invalid of ['2026-02-30', '2026-9-3', '0000-01-01', '2026-13-01', '2026-02-28T00:00:00Z']) {
    assert.equal(presentation.validContentControlIsoDate(invalid), false)
    assert.throws(() => presentation.formatContentControlDate(invalid), /YYYY-MM-DD/)
  }
  assert.equal(presentation.validContentControlIsoDate('0001-01-01'), true)
})

test('date format, locale and calendar validation rejects malformed updates atomically', () => {
  const model = fixture('2026-09-03')
  add(model, { type: 'date', dateValue: '2026-09-03' })
  for (const patch of [{ dateFormat: 'yyyy "' }, { dateFormat: 'yyyy\\' }, { dateLocale: 'bad_locale!' }, { dateCalendar: 'unknown-calendar' }]) {
    rejected(model, () => controls.updateDocumentContentControl(model, 'control', patch), /日期/)
  }
  controls.updateDocumentContentControl(model, 'control', { dateFormat: 'dd MMM yyyy', dateLocale: 'fr-FR' })
  assert.equal(text(model), '03 sept. 2026'); assert.equal(info(model).dateValue, '2026-09-03')
})

test('list edits preserve values, update display labels and retain caller-defined ordering', () => {
  const model = fixture('Alpha')
  add(model, { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }, { value: 'B', displayText: 'Beta' }] })
  const items = [{ value: 'B', displayText: 'Beta' }, { value: 'A', displayText: 'Renamed' }, { value: 'C', displayText: 'Renamed' }]
  controls.updateDocumentContentControl(model, 'control', { items })
  assert.equal(text(model), 'Renamed'); assert.equal(info(model).value, 'A')
  assert.deepEqual(info(model).items.map(item => item.value), ['B', 'A', 'C'])
  controls.setDocumentContentControlValue(model, 'control', 'C')
  assert.equal(text(model), 'Renamed'); assert.equal(info(model).value, 'C')
  rejected(model, () => controls.updateDocumentContentControl(model, 'control', { items: items.slice(0, 2) }), /not present/)
  rejected(model, () => controls.updateDocumentContentControl(model, 'control', { items: [{ value: 'C' }, { value: 'C' }] }), /重复/)
  rejected(model, () => controls.updateDocumentContentControl(model, 'control', { items: [] }), /至少/)
})

test('combo boxes retain free values when an option is removed and can select an option label later', () => {
  const model = fixture('Free')
  add(model, { type: 'comboBox', value: 'Free', items: [{ value: 'A', displayText: 'Alpha' }] })
  controls.setDocumentContentControlValue(model, 'control', 'Other')
  assert.equal(text(model), 'Other'); assert.equal(info(model).value, 'Other')
  controls.updateDocumentContentControl(model, 'control', { items: [{ value: 'B', displayText: 'Beta' }] })
  assert.equal(text(model), 'Other')
  controls.setDocumentContentControlValue(model, 'control', 'B')
  assert.equal(text(model), 'Beta'); assert.equal(info(model).value, 'B')
})

test('input nested metadata, list items and getter snapshots cannot mutate persisted control properties', () => {
  const model = fixture('A'), input = { type: 'dropDownList', value: 'A', items: [{ value: 'A', displayText: 'Alpha' }],
    metadata: { nested: { retained: ['original'] } }, dataBinding: { xpath: '/value' } }
  add(model, input)
  input.items[0].displayText = 'Changed'; input.metadata.nested.retained.push('external'); input.dataBinding.xpath = '/changed'
  assert.equal(info(model).items[0].displayText, 'Alpha')
  assert.deepEqual(info(model).metadata.nested.retained, ['original']); assert.equal(info(model).dataBinding.xpath, '/value')
  const snapshot = info(model); snapshot.items[0].value = 'X'; snapshot.metadata.nested.retained.push('snapshot')
  assert.equal(info(model).items[0].value, 'A'); assert.deepEqual(info(model).metadata.nested.retained, ['original'])
})

test('runtime property DTOs reject invalid enums, injected range fields, non-JSON and invalid checkbox states before publication', () => {
  const model = fixture('_')
  add(model, { type: 'checkbox' })
  const cycle = {}; cycle.self = cycle
  for (const patch of [{ appearance: 'invisible' }, { lock: 'locked' }, { startOffset: 0 }, { color: '#fff' },
    { metadata: { callback: () => {} } }, { metadata: cycle }, { metadata: new Date() },
    { checkbox: { checked: 'true' } }, { checkbox: { checked: false, uncheckedState: { value: 'D800' } } },
    { checkbox: { checked: false, uncheckedState: { value: '110000' } } }, { checkbox: { checked: false, uncheckedState: { font: '' } } }]) {
    rejected(model, () => controls.updateDocumentContentControl(model, 'control', patch), /内容控件|颜色|复选框/)
  }
})

test('property drafts submit only edited properties and detect concurrent edits to those properties or positions', () => {
  const model = fixture('NAME')
  add(model, { alias: 'Customer', tag: 'original', metadata: { source: 'Office' } })
  const baseline = info(model), draft = { alias: 'Customer', tag: 'new-tag', metadata: structuredClone(baseline.metadata) }
  const patch = properties.contentControlPropertyDraftPatch(baseline, draft)
  assert.deepEqual(patch, { tag: 'new-tag' })
  const current = { ...structuredClone(baseline), value: 'other', alias: 'Other user title' }
  assert.doesNotThrow(() => properties.assertContentControlPropertyDraftCurrent(baseline, current, patch))
  assert.throws(() => properties.assertContentControlPropertyDraftCurrent(baseline, { ...current, tag: 'other-tag' }, patch), /属性已变化/)
  assert.throws(() => properties.assertContentControlPropertyDraftCurrent(baseline, { ...current, endOffset: 3 }, patch), /位置或类型/)
  properties.validateContentControlPropertyPatch(patch, current)
})

test('value capability reads actual document permissionRules, content locks and command types', () => {
  const model = fixture('_')
  add(model, { type: 'checkbox' })
  const command = { id: 'doc.command.set-content-control-checked', params: { controlId: 'control', checked: true } }
  const permission = { documentId: model.id, userId: 'editor' }
  assert.equal(commandState.contentControlValueCommandState(model, command, { permission }).available, true)
  model.permissionRules = [{ id: 'deny-p', scope: 'paragraph', targetId: 'p', action: 'edit', effect: 'deny', userIds: ['editor'] }]
  assert.equal(commandState.contentControlValueCommandState(model, command, { permission }).available, false)
  assert.equal(commandState.contentControlValueCommandState(model, command, { permission: { ...permission, userId: 'other' } }).available, true)
  model.permissionRules = []
  assert.equal(commandState.contentControlValueCommandState(model, command, { disabled: true }).available, false)
  controls.updateDocumentContentControl(model, 'control', { lock: 'contentLocked' })
  assert.equal(commandState.contentControlValueCommandState(model, command, { permission }).available, false)
  assert.equal(commandState.contentControlValueCommandState(model, { id: 'doc.command.select-content-control-item', params: { controlId: 'control', value: 'A' } }).available, false)
})

test('checkbox replacement in a table cell updates the authoritative nested paragraph instead of a stale reference', () => {
  const model = fixture('_')
  const table = modelUtils.createParagraph(); table.id = 'table'; table.type = 'table'; table.runs = []
  table.metadata = { table: { rows: [{ id: 'row', cells: [{ id: 'cell', blocks: [model.blocks[0]], style: {}, rowSpan: 1, colSpan: 1 }] }] } }
  model.blocks = [table]
  controls.createDocumentContentControl(model, { id: 'control', type: 'checkbox', blockId: 'p', startOffset: 0, endOffset: 1 })
  controls.setDocumentContentControlChecked(model, 'control', true)
  const paragraph = model.blocks[0].metadata.table.rows[0].cells[0].blocks[0]
  assert.equal(paragraph.runs.map(run => run.text).join(''), '☒')
  assert.equal(paragraph.runs[0].fontFamily, 'Segoe UI Symbol')
  assert.equal(controls.contentControlText(model, info(model)), '☒')
})
