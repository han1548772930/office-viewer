import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankSheet, createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { createSheetDataValidationRule } from '../src/features/sheet-validation/validation-engine.ts'
import { planSheetDataValidationCommand, SHEET_DATA_VALIDATION_COMMANDS, sheetDataValidationCommandInput } from '../src/features/sheet-validation/validation-rule-command.ts'
import { validationPanelSettings, validationRuleFromPanel } from '../src/features/sheet-validation/validation-panel-state.ts'
import { COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, sheetValidationListOptions } from '../src/features/sheet-validation/validation-list-input.ts'

// Implementation-phase contracts. The real OfficeStudio/SSR harness and Office reopen samples have not been run.
let vite, OfficeStudio, FOffice
beforeAll(async () => {
  vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;({ OfficeStudio, FOffice } = await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))
}, { timeout: 60_000 })
afterAll(async () => { await vite?.close() })
const area = (row = 0, column = 0, endRow = row, endCol = column) => ({ startRow: row, endRow, startCol: column, endCol })
function fixture(configure = () => {}) {
  const model = createBlankWorkbook(), sheet = model.sheets[0]
  model.id = 'validation-rules'; model.activeSheetId = sheet.id = 'data'; model.calculationMode = 'manual'
  sheet.rowCount = 100; sheet.columnCount = 30; sheet.cells['0:0'] = createCell(0, 0, 'Red'); sheet.cells['2:0'] = createCell(2, 0, 42)
  sheet.dataValidations = [createSheetDataValidationRule({ id: 'list', type: 'list', formula1: '"Red,Green"', formula2: '#ff0000,', metadata: { retained: ['rule'] } }, [area(0, 0, 3, 1)]),
    createSheetDataValidationRule({ id: 'number', type: 'whole', operator: 'between', formula1: '1', formula2: '100' }, [area(8, 0, 12, 1)])]
  const other = createBlankSheet('Other'); other.id = 'other'; model.sheets.push(other); configure(model, sheet)
  const studio = new OfficeStudio({ workbook: model, currentUserId: 'local-user' }), book = FOffice.newAPI(studio).getActiveWorkbook()
  const current = () => studio.getWorkbookModel(model.id), count = () => studio.editor.history.units[model.id]?.undo.length ?? 0
  const target = () => ({ unitId: model.id, sheetId: 'data', expectedRevision: studio.getWorkbookRevision(model.id),
    expectedHistoryGeneration: studio.editor.history.captureSavepoint(model.id).generation, expectedUserId: studio.currentUserId,
    expectedRulesSignature: JSON.stringify(current().sheets[0].dataValidations) })
  const execute = (action, payload, overrides = {}) => studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS[action], { ...target(), ...payload, ...overrides })
  return { studio, book, current, count, target, execute, rules: () => current().sheets[0].dataValidations, dispose: () => studio.dispose() }
}

test('setting one discontinuous rule clips overlapping coverage without changing cells and creates one isolated content history entry', () => {
  const h = fixture()
  try {
    const before = cloneModel(h.current()), rule = createSheetDataValidationRule({ id: 'date', type: 'date', metadata: { showTime: true } }, [area(0, 0), area(2, 0)])
    const input = Object.freeze({ ...h.target(), rule: Object.freeze(cloneModel(rule)) }), frozenBefore = JSON.stringify(input)
    assert.equal(h.studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.set, input), true); assert.equal(h.count(), 1); assert.equal(JSON.stringify(input), frozenBefore)
    const added = cloneModel(h.current()); assert.deepEqual(added.sheets[0].cells, before.sheets[0].cells)
    assert.deepEqual(h.rules().find(item => item.id === 'date').ranges, [area(0, 0), area(2, 0)])
    for (const row of [0, 2]) assert.equal(h.rules().find(item => item.id === 'list').ranges.some(range => row >= range.startRow && row <= range.endRow && range.startCol === 0), false)
    h.book.undo(); assert.deepEqual(h.current().sheets[0].dataValidations, before.sheets[0].dataValidations); h.book.redo(); assert.deepEqual(h.current().sheets[0].dataValidations, added.sheets[0].dataValidations)
    const reopened = new OfficeStudio({ workbook: h.book.save() })
    try { assert.deepEqual(reopened.getWorkbookModel(added.id).sheets[0].dataValidations, added.sheets[0].dataValidations) } finally { reopened.dispose() }
  } finally { h.dispose() }
})

test('replacing a rule moves its complete coverage while retaining ID, colors, metadata and unrelated models', () => {
  const h = fixture()
  try {
    const before = cloneModel(h.current()), rule = { ...cloneModel(h.rules()[0]), ranges: [area(4, 3, 5, 3)] }
    assert.equal(h.execute('replace', { rule }), true); assert.equal(h.count(), 1)
    const saved = h.rules().find(item => item.id === 'list'); assert.deepEqual(saved.ranges, rule.ranges); assert.equal(saved.formula1, rule.formula1)
    assert.equal(saved.formula2, '#ff0000,'); assert.deepEqual(saved.metadata, { retained: ['rule'] }); assert.deepEqual(h.rules().find(item => item.id === 'number'), before.sheets[0].dataValidations[1])
    assert.deepEqual(h.current().sheets[0].cells, before.sheets[0].cells); assert.deepEqual(h.current().sheets[1], before.sheets[1])
  } finally { h.dispose() }
})

test.each(['list', 'listMultiple'])('INDEX %s authoring, candidate selection and source edits share actual commands, colors, history and internal reopen', type => {
  const h = fixture((_model, sheet) => {
    for (const [row, value] of [[0, 'Red'], [1, 'Green, dark'], [2, 'Blue'], [3, 'Red']]) sheet.cells[`${row}:3`] = createCell(row, 3, value)
  })
  try {
    const before = cloneModel(h.current()), settings = validationPanelSettings(undefined, type)
    settings.listMode = 'reference'; settings.listReference = '=INDEX($D$1:$D$4,0,1)'; settings.formula2 = '#ff0000,#008000,#0000ff'
    const rule = validationRuleFromPanel(settings, [area(5, 5)], { id: 'index-source', metadata: { retained: ['index'] }, renderMode: 'custom' })
    assert.equal(h.execute('set', { rule }), true); assert.equal(h.count(), 1); assert.deepEqual(h.current().sheets[0].cells, before.sheets[0].cells)
    const after = cloneModel(h.current()), sheet = h.book.getActiveSheet()
    assert.deepEqual(sheet.getCellValidationOptions(5, 5), ['Red', 'Green, dark', 'Blue'])
    assert.deepEqual(sheetValidationListOptions(h.rules().find(item => item.id === 'index-source'), h.current(), 'data', 'local-user').map(item => item.color), ['#ff0000', '#008000', '#0000ff'])
    h.book.undo(); assert.deepEqual(h.current(), before); h.book.redo(); assert.deepEqual(h.current(), after)
    const selected = h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { unitId: before.id, sheetId: 'data', row: 5, column: 5,
      ruleId: 'index-source', values: ['Green, dark'] })
    assert.equal(selected.accepted, true); assert.equal(h.current().sheets[0].cells['5:5'].value, type === 'listMultiple' ? '"Green, dark"' : 'Green, dark')
    sheet.getRange('D3').setValue('Approved'); assert.deepEqual(sheet.getCellValidationOptions(5, 5), ['Red', 'Green, dark', 'Approved'])
    const saved = h.book.save(), reopened = new OfficeStudio({ workbook: saved })
    try {
      const restored = FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet()
      assert.deepEqual(restored.getDataValidations(), h.rules()); assert.deepEqual(restored.getCellValidationOptions(5, 5), ['Red', 'Green, dark', 'Approved'])
      assert.equal(restored.getRange('F6').getValue(), h.current().sheets[0].cells['5:5'].value)
    } finally { reopened.dispose() }
    h.current().sheets[0].rangeProtectionRules = [{ id: 'hidden-index', ranges: [area(2, 3)], ownerUserId: 'peer', allowedUserIds: [],
      viewState: 'noOneElseCanView', editState: 'onlyMe' }]
    assert.deepEqual(sheet.getCellValidationOptions(5, 5), []); const unchanged = cloneModel(h.current()), count = h.count()
    assert.throws(() => h.studio.executeCommandSync(COMMIT_SHEET_VALIDATION_LIST_INPUT_COMMAND, { unitId: before.id, sheetId: 'data', row: 5, column: 5,
      ruleId: 'index-source', values: ['Green, dark'] }), /候选|当前列表/)
    assert.deepEqual(h.current(), unchanged); assert.equal(h.count(), count)
  } finally { h.dispose() }
})

test('identical normalized rules and missing removals are no-ops, while real removal is one reversible transaction', () => {
  const h = fixture()
  try {
    const before = cloneModel(h.current()), rule = cloneModel(h.rules()[0]); rule.metadata = { retained: ['rule'] }
    assert.equal(h.execute('replace', { rule }), false); assert.equal(h.execute('remove', { ruleId: 'missing' }), false); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    assert.equal(h.execute('remove', { ruleId: 'list' }), true); assert.equal(h.count(), 1); assert.deepEqual(h.rules().map(rule => rule.id), ['number'])
    h.book.undo(); assert.deepEqual(h.rules(), before.sheets[0].dataValidations)
  } finally { h.dispose() }
})

test('type-qualified clearing clips only requested date coverage and retains the list, number and outside date ranges', () => {
  const h = fixture((_model, sheet) => { sheet.dataValidations.push(createSheetDataValidationRule({ id: 'date', type: 'date' }, [area(20, 0, 24, 0)])) })
  try {
    const before = cloneModel(h.rules())
    assert.equal(h.execute('clearTypes', { ranges: [area(21, 0, 23, 0)], types: ['date'] }), true)
    assert.deepEqual(h.rules().find(rule => rule.id === 'date').ranges, [area(20, 0), area(24, 0)])
    assert.deepEqual(h.rules().filter(rule => rule.id !== 'date'), before.filter(rule => rule.id !== 'date')); assert.equal(h.count(), 1)
    assert.equal(h.execute('clearTypes', { ranges: [area(21, 0, 23, 0)], types: ['date'] }), false); assert.equal(h.count(), 1)
    assert.equal(h.execute('clear', { ranges: [area(20, 0, 24, 0)] }), true); assert.equal(h.rules().some(rule => rule.id === 'date'), false)
  } finally { h.dispose() }
})

test.each(['oldView', 'newView', 'oldEdit', 'newEdit'])('%s denial blocks replacement over both the original and proposed full ranges', kind => {
  const h = fixture((_model, sheet) => { sheet.rangeProtectionRules = [{ id: 'deny', ranges: [kind.startsWith('old') ? area(3, 1) : area(4, 3)], ownerUserId: 'owner', allowedUserIds: [],
    viewState: kind.endsWith('View') ? 'noOneElseCanView' : 'othersCanView', editState: kind.endsWith('Edit') ? 'onlyMe' : 'othersCanEdit' }] })
  try {
    const before = cloneModel(h.current()), rule = { ...cloneModel(h.rules()[0]), ranges: [area(4, 3, 5, 3)] }
    assert.throws(() => h.execute('replace', { rule }), /查看|权限|保护|编辑/); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    if (kind.startsWith('old')) { assert.throws(() => h.execute('remove', { ruleId: 'list' }), /查看|权限|保护|编辑/); assert.deepEqual(h.current(), before) }
  } finally { h.dispose() }
})

test('clearing one accessible portion preserves hidden rule metadata and untouched coverage instead of exposing or rewriting it', () => {
  const h = fixture((_model, sheet) => { sheet.rangeProtectionRules = [{ id: 'deny', ranges: [area(3, 1)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'onlyMe' }] })
  try {
    const original = cloneModel(h.rules()[0]); assert.equal(h.execute('clear', { ranges: [area(0, 0)] }), true)
    const saved = h.rules().find(rule => rule.id === 'list'); assert.equal(saved.formula1, original.formula1); assert.equal(saved.formula2, original.formula2); assert.deepEqual(saved.metadata, original.metadata)
    assert.ok(saved.ranges.some(range => range.startRow <= 3 && range.endRow >= 3 && range.startCol <= 1 && range.endCol >= 1)); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test.each(['revision', 'history', 'user', 'rules', 'host'])('stale %s metadata prevents even a nominal no-op from publishing', kind => {
  const h = fixture()
  try {
    const target = h.target()
    if (kind === 'revision') target.expectedRevision++
    else if (kind === 'history') target.expectedHistoryGeneration++
    else if (kind === 'user') target.expectedUserId = 'other'
    else if (kind === 'rules') target.expectedRulesSignature = '[]'
    else { target.host = 'workbench'; h.book.getSheetBySheetId('other').activate() }
    const before = cloneModel(h.current())
    assert.throws(() => h.studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.remove, { ...target, ruleId: 'missing' }), /更改/)
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('inactive API workbooks retain independent rules/history and cancellation does not mutate either workbook', () => {
  const h = fixture()
  try {
    const peer = FOffice.newAPI(h.studio).createWorkbook({ id: 'validation-peer' }), beforePeer = cloneModel(h.studio.getWorkbookModel(peer.getId())), before = cloneModel(h.current())
    const subscription = h.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === SHEET_DATA_VALIDATION_COMMANDS.remove) event.cancel = true })
    assert.throws(() => h.execute('remove', { ruleId: 'list' }), /canceled/); subscription.dispose(); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
    assert.equal(h.execute('remove', { ruleId: 'list' }, { host: 'api' }), true); assert.equal(h.count(), 1); assert.deepEqual(h.studio.getWorkbookModel(peer.getId()), beforePeer)
  } finally { h.dispose() }
})

test('a changed payload getter is detected after authority capture and is never merged into a prepared rule candidate', () => {
  const h = fixture()
  try {
    const old = cloneModel(h.rules()[0]), input = { ...h.target() }, reads = {}
    const rule = { ...cloneModel(old), ranges: [area(4, 3)] }
    Object.defineProperty(input, 'rule', { get() { reads.rule = (reads.rule ?? 0) + 1; h.rules()[0].metadata.external = true; return rule }, enumerable: true })
    assert.throws(() => h.studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.replace, input), /更改/); assert.equal(reads.rule, 1)
    assert.deepEqual(h.rules()[0].ranges, old.ranges); assert.equal(h.rules()[0].metadata.external, true); assert.equal(h.count(), 0)
    const fields = { ...h.target(), ruleId: 'missing' }, frozen = {}
    for (const [field, value] of Object.entries(fields)) Object.defineProperty(frozen, field, { get() { reads[field] = (reads[field] ?? 0) + 1; return value }, enumerable: true })
    assert.equal(h.studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.remove, Object.freeze(frozen)), false)
    for (const field of Object.keys(fields)) assert.equal(reads[field], 1)
  } finally { h.dispose() }
})

test.each(['planning', 'mutator', 'prepareView'])('View revoked during %s rejects a prepared definition without an edit-history record', stage => {
  const h = fixture()
  try {
    const old = cloneModel(h.rules()), payload = { ...h.target(), rule: { ...cloneModel(h.rules()[0]), ranges: [area(4, 3)] } }
    const revoke = () => { h.current().sheets[0].rangeProtectionRules = [{ id: 'deny', ranges: [area(4, 3)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }] }
    if (stage === 'planning') {
      const original = h.studio.assertCanViewRange.bind(h.studio); let calls = 0
      h.studio.assertCanViewRange = (...args) => { original(...args); if (++calls === 1) revoke() }
    } else {
      const original = h.studio.commitWorkbook.bind(h.studio)
      h.studio.commitWorkbook = (id, label, mutate, options) => original(id, label, candidate => { if (stage === 'mutator') revoke(); mutate(candidate) },
        { ...options, prepareView: candidate => { if (stage === 'prepareView') revoke(); options.prepareView(candidate) } })
    }
    assert.throws(() => h.studio.executeCommandSync(SHEET_DATA_VALIDATION_COMMANDS.replace, payload), /更改|查看|保护|权限/)
    assert.deepEqual(h.rules(), old); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})

test('an imported unedited formula outside the current parser grammar survives coverage edits, while an edited malformed expression is rejected', () => {
  const h = fixture((_model, sheet) => { sheet.dataValidations[0] = { id: 'vendor', type: 'custom', ranges: [area()], formula1: '=AddIn!FutureFunction()', metadata: { retained: 'Office formula' } } })
  try {
    const original = cloneModel(h.rules()[0]); assert.equal(h.execute('replace', { rule: { ...original, ranges: [area(2, 0)] } }), true)
    assert.equal(h.rules().find(rule => rule.id === 'vendor').formula1, original.formula1); assert.deepEqual(h.rules().find(rule => rule.id === 'vendor').metadata, original.metadata)
    const before = cloneModel(h.current()); assert.throws(() => h.execute('replace', { rule: { ...cloneModel(h.rules().find(rule => rule.id === 'vendor')), formula1: '=1+' } }))
    assert.deepEqual(h.current(), before); assert.equal(h.count(), 1)
  } finally { h.dispose() }
})

test('malformed DTOs reject atomically and pure planning leaves both frozen input and source unchanged', () => {
  const h = fixture()
  try {
    const malformed = [['clear', { ranges: [] }], ['clearTypes', { ranges: [area()], types: ['unknown'] }], ['replace', { rule: { id: 'bad', type: 'whole', ranges: [area()], formula1: 7 } }],
      ['remove', { ruleId: '' }], ['clear', { ranges: [area(-1)] }], ['remove', { ruleId: 'list', expectedRevision: -1 }], ['remove', { ruleId: 'list', host: 'unowned' }]]
    const before = cloneModel(h.current())
    for (const [action, payload] of malformed) { assert.throws(() => h.execute(action, payload)); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0) }
    const payload = Object.freeze({ ...h.target(), rule: Object.freeze({ ...cloneModel(h.rules()[0]), ranges: [area(2, 3)] }) }), frozenBefore = JSON.stringify(payload)
    const plan = planSheetDataValidationCommand(h.current().sheets[0], sheetDataValidationCommandInput('replace', payload))
    assert.equal(plan.changed, true); assert.equal(JSON.stringify(payload), frozenBefore); assert.deepEqual(h.current(), before)
    assert.throws(() => sheetDataValidationCommandInput('unknown', { ...h.target(), ranges: [area()], types: ['date'] }), /动作/)
  } finally { h.dispose() }
})

test('canonicalizing a padded rule ID still checks the complete original protected coverage', () => {
  const h = fixture((_model, sheet) => { sheet.rangeProtectionRules = [{ id: 'private-old', ranges: [area(3, 1)], ownerUserId: 'owner', allowedUserIds: [], viewState: 'noOneElseCanView', editState: 'othersCanEdit' }] })
  try {
    const before = cloneModel(h.current()), rule = { ...cloneModel(h.rules()[0]), id: ' list ', ranges: [area(4, 3)] }
    assert.throws(() => h.execute('replace', { rule }), /查看|权限|保护/); assert.deepEqual(h.current(), before); assert.equal(h.count(), 0)
  } finally { h.dispose() }
})
