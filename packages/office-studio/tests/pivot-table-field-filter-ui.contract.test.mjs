import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { pivotChartFieldButtonsIds as ids, pivotChartFieldButtonsWorkbook } from './fixtures/pivot-chart-field-buttons-fixture.mjs'

// Deferred ordinary-report contracts. These real algorithms/Facades/controller drafts have never run.
let server, api, state, members, conditions, panels, owners, hooks
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, state, members, conditions, panels, owners, hooks] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-table-field-filter-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-button-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-filter-condition.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-owner.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/views/use-pivot-table-field-filter.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup(options = {}) {
  const studio = new api.OfficeStudio({ workbook: pivotChartFieldButtonsWorkbook(options), currentUserId: 'local-user' })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), target = { studio, unitId: book.getId(), sheetId: ids.host, pivotId: ids.primary }
  const model = () => studio.getWorkbookModel(target.unitId), report = () => model().pivotTables.find(value => value.id === target.pivotId)
  return { ...target, target, book, model, report, field: name => report().fields.find(field => field.name === name),
    pivot: () => book.getPivotTables().find(value => value.getPivotTableId() === target.pivotId),
    panel: () => panels.createPivotPanelDraft(model(), ids.host, 'D3:H7', target.pivotId),
    history: () => studio.editor.history.units[target.unitId]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
const owner = f => owners.capturePivotChartPanelOwner(f.studio, f.unitId, f.sheetId)
const context = f => ({ ...owners.pivotChartPanelOwnerContext(owner(f)), hostSheetId: f.sheetId })
const total = f => f.pivot().getData().slice(1).flatMap(row => row.slice(1)).reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0)
async function memberDraft(f, name = 'Region') { const field = f.field(name), pivot = f.pivot(); return members.createPivotChartFieldFilterDraft(await pivot.getFieldItems(field.id), pivot.getManualFilterItemState(field.id)) }
function conditionDraft(f, name = 'Date') {
  const field = f.field(name), pivot = f.pivot()
  return conditions.createPivotChartFieldConditionsDraft(pivot.getFilterCapabilities(field.id), pivot.getDateFilter(field.id), pivot.getLabelConditionFilter(field.id))
}
function controller(f) {
  const scope = effectScope(), calls = [], options = shallowReactive({ ...f.target, panelDraft: f.panel(), owner: owner(f),
    active: true, sourceReady: true, busy: false, restoreFocus: fieldId => calls.push(['focus', fieldId]),
    applied(fieldId, action, changed) {
      state.synchronizePivotTableFieldFilterDraft(options.panelDraft, f.report(), fieldId, action)
      options.owner = owner(f); calls.push(['applied', fieldId, action, changed])
    } })
  const controls = scope.run(() => hooks.usePivotTableFieldFilter(options))
  return { options, controls, calls, stop: () => scope.stop(), reread() { options.panelDraft = f.panel(); options.owner = owner(f) } }
}
async function changeOwner(f, change) {
  if (change === 'source') f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)
  if (change === 'history') f.studio.editor.history.clear(f.unitId)
  if (change === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
  if (change === 'host') f.book.setActiveSheet(ids.other)
  await nextTick()
}

test('ordinary snapshot uses actual local field IDs, shared badges and defensive metadata without history', () => {
  const f = setup()
  try {
    const before = json(f.book.save()), snapshot = state.readPivotTableFieldFilterSnapshot(f.target, f.field('Region').id)
    assert.equal(snapshot.field.fieldId, f.field('Region').id); assert.equal(snapshot.field.label, 'Region'); assert.equal(snapshot.editable, true)
    snapshot.field.field.displayName = 'Changed'; snapshot.options.repeatRowLabels = true
    assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['sourceSheet', 'sourceRange', 'area', 'sourceColumn'])('changed %s draft cannot reuse a committed field catalog', change => {
  const f = setup()
  try {
    const draft = f.panel(), id = f.field('Region').id, field = draft.fields.find(field => field.fieldId === id)
    if (change === 'sourceSheet') draft.sourceSheetId = ids.other
    if (change === 'sourceRange') draft.sourceRange = 'D3:H6'
    if (change === 'area') field.area = 'column'
    if (change === 'sourceColumn') field.sourceColumn++
    assert.equal(state.pivotTableFieldFilterDraftCompatible(f.model(), draft, f.pivotId, id), false); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('pending destination, display name, unused field and options stay compatible and survive a filter merge', () => {
  const f = setup()
  try {
    const draft = f.panel(), id = f.field('Region').id
    draft.targetAddress = 'J30'; draft.fields.find(field => field.fieldId === id).displayName = 'Pending region'; draft.options.repeatRowLabels = true
    const before = json(draft); assert.equal(state.pivotTableFieldFilterDraftCompatible(f.model(), draft, f.pivotId, id), true)
    state.synchronizePivotTableFieldFilterDraft(draft, f.report(), id, 'applyFilter'); assert.deepEqual(draft, before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['sourceRange', 'sourceSheetId', 'targetAddress', 'targetSheetId', 'options', 'fields'])('panel signature tracks %s while the field request waits', key => {
  const f = setup()
  try {
    const draft = f.panel(), before = state.pivotTableFieldFilterPanelSignature(draft)
    if (key === 'options') draft.options.repeatRowLabels = true
    else if (key === 'fields') draft.fields[0].displayName = 'Pending caption'
    else draft[key] = `${draft[key]}-changed`
    assert.notEqual(state.pivotTableFieldFilterPanelSignature(draft), before)
  } finally { f.dispose() }
})

test.each(['missing', 'Amount'])('missing or value field %s does not open an ordinary member editor', name => {
  const f = setup()
  try { assert.throws(() => state.readPivotTableFieldFilterSnapshot(f.target, f.field(name)?.id ?? 'missing')); assert.equal(f.history(), 0) }
  finally { f.dispose() }
})

const partialCases = [[{ partialDate: true }, 'Date', '2026', 100], [{ typedMembers: true }, 'Region', '1', 30]]
test.each(partialCases)('ordinary partial %j retains underlying raw identity through manual confirmation/history/reopen', async (options, name, key) => {
  const f = setup(options); let reopened
  try {
    const draft = await memberDraft(f, name), before = json(f.book.save()); assert.deepEqual(draft.partialKeys, [key]); assert.deepEqual(draft.selectedKeys, [])
    await state.applyPivotTableFieldFilterAction(f.pivot(), f.field(name).id, 'applyFilter', context(f), draft)
    assert.equal(total(f), 10); assert.equal(f.history(), 1)
    if (name === 'Region') assert.deepEqual(f.field(name).recordSelection.values, [1])
    const saved = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
    reopened = new api.OfficeStudio({ workbook: saved }); const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotTables().find(value => value.getPivotTableId() === ids.primary)
    assert.deepEqual(restored.getManualFilterItemState(f.field(name).id).partialItems, [key])
  } finally { reopened?.dispose(); f.dispose() }
})

test.each(partialCases)('checking ordinary partial %j selects every raw member in that group', async (options, name, key, expected) => {
  const f = setup(options)
  try {
    const draft = members.setPivotChartFilterSelection(await memberDraft(f, name), key, true)
    await state.applyPivotTableFieldFilterAction(f.pivot(), f.field(name).id, 'applyFilter', context(f), draft)
    assert.equal(total(f), expected); assert.deepEqual(f.pivot().getManualFilterItemState(f.field(name).id).partialItems, [])
  } finally { f.dispose() }
})

test('empty ordinary member draft rejects before any command or output publication', async () => {
  const f = setup()
  try {
    const draft = members.setPivotChartFilterAll(await memberDraft(f), false)
    await assert.rejects(() => state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Region').id, 'applyFilter', context(f), draft), /至少选择/)
    assert.equal(f.history(), 0); assert.equal(total(f), 100)
  } finally { f.dispose() }
})

test('ordinary manual filter updates real output and linked chart in one context-bearing command', async () => {
  const f = setup(), events = []
  try {
    f.studio.events.addEvent('CommandExecuted', event => events.push(event)); const before = json(f.book.save()), source = json(f.model().sheets.find(sheet => sheet.id === ids.source).cells)
    const draft = members.setPivotChartFilterSelection(await memberDraft(f), 'South', false)
    await state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Region').id, 'applyFilter', context(f), draft)
    assert.equal(total(f), 40); assert.equal(f.history(), 1); assert.deepEqual(f.model().sheets.find(sheet => sheet.id === ids.source).cells, source)
    const changed = events.filter(event => event.id === 'sheet.command.set-pivot-filter'); assert.equal(changed.length, 1)
    assert.equal(changed[0].params.hostSheetId, ids.host); assert.equal(changed[0].params.expectedUserId, 'local-user')
    assert.equal(f.book.getPivotChartById(ids.linked).getData().series.flatMap(series => series.values).reduce((a, value) => a + (typeof value === 'number' ? value : 0), 0), 40)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test('ordinary select-all clears member only; clear-all removes each allowed filter category', async () => {
  const f = setup()
  try {
    const id = f.field('Region').id; await f.pivot().setOptions({ allowMultipleFilters: true })
    await f.pivot().setLabelManualFilter(id, ['North']); await f.pivot().setLabelConditionFilter(id, { operator: 'captionEqual', expected: 'North' })
    await f.pivot().setValueFilter(id, { operator: 'valueGreaterThanOrEqual', expected: 35, valueFieldId: f.field('Amount').id })
    await state.applyPivotTableFieldFilterAction(f.pivot(), id, 'applyFilter', context(f), members.setPivotChartFilterAll(await memberDraft(f), true))
    assert.ok(f.field('Region').labelFilter); assert.ok(f.field('Region').valueFilter); assert.equal(f.field('Region').groupedItems, undefined)
    const before = json(f.book.save()), history = f.history(); await state.applyPivotTableFieldFilterAction(f.pivot(), id, 'clearFilter', context(f))
    for (const key of ['selectedItems', 'groupedItems', 'partialItems', 'recordSelection', 'dateFilter', 'labelFilter', 'valueFilter']) assert.equal(f.field('Region')[key], undefined)
    assert.equal(total(f), 100); assert.equal(f.history(), history + 1); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each([
  ['dateBetween', [46055, 46073], 50], ['dateEqual', 46055, 20], ['dateNewerThan', 46055, 70], ['dateNewerThanOrEqual', 46055, 90],
  ['dateNotBetween', [46055, 46073], 50], ['dateNotEqual', 46055, 80], ['dateOlderThan', 46055, 10], ['dateOlderThanOrEqual', 46055, 30],
])('ordinary typed date %s uses shared semantics, transaction/history and saved state', async (operator, expected, sum) => {
  const f = setup()
  try {
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { operator,
      first: Array.isArray(expected) ? expected[0] : expected, second: Array.isArray(expected) ? expected[1] : undefined })
    const before = json(f.book.save()); await state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Date').id, 'applyDateFilter', context(f), undefined, draft)
    assert.equal(total(f), sum); assert.deepEqual(f.pivot().getDateFilter(f.field('Date').id).expected, expected); assert.equal(f.history(), 1)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each([
  ['captionBeginsWith', 'N', '', 40], ['captionBetween', 'North', 'South', 100], ['captionContains', 'or', '', 40], ['captionEndsWith', 'orth', '', 40],
  ['captionEqual', 'North', '', 40], ['captionGreaterThan', 'North', '', 60], ['captionGreaterThanOrEqual', 'North', '', 100],
  ['captionLessThan', 'South', '', 40], ['captionLessThanOrEqual', 'South', '', 100], ['captionNotBeginsWith', 'N', '', 60],
  ['captionNotBetween', 'North', 'South', 0], ['captionNotContains', 'or', '', 60], ['captionNotEndsWith', 'orth', '', 60], ['captionNotEqual', 'North', '', 60],
])('ordinary caption %s is a real report condition', async (operator, first, second, sum) => {
  const f = setup()
  try {
    const draft = conditions.setPivotChartLabelConditionDraft(conditionDraft(f, 'Region'), { operator, first, second })
    await state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Region').id, 'applyLabelCondition', context(f), undefined, draft)
    assert.equal(total(f), sum); assert.equal(f.pivot().getLabelConditionFilter(f.field('Region').id).operator, operator); assert.equal(f.history(), 1)
  } finally { f.dispose() }
})

test.each([false, true])('ordinary date1904=%s consumes the true capability and preserves fractional operands', async date1904 => {
  const f = setup({ date1904, dateSerials: true })
  try {
    const serial = f.model().sheets.find(sheet => sheet.id === ids.source).cells['4:5'].value
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { first: serial })
    assert.equal(draft.capabilities.date1904, date1904); await state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Date').id, 'applyDateFilter', context(f), undefined, draft)
    assert.equal(total(f), 20); assert.equal(f.pivot().getDateFilter(f.field('Date').id).expected, date1904 ? 44593 : 46055)
  } finally { f.dispose() }
})

test('ordinary existing compatibility day survives typed draft and JSON reopen without calendar normalization', async () => {
  const f = setup({ compatibilityDate: true }); let restored
  try {
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { first: 60.25 })
    await state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Date').id, 'applyDateFilter', context(f), undefined, draft)
    assert.equal(total(f), 40); restored = new api.OfficeStudio({ workbook: json(f.book.save()) })
    const pivot = api.FOffice.newAPI(restored).getActiveWorkbook().getPivotTables().find(pivot => pivot.getPivotTableId() === ids.primary)
    assert.equal(pivot.getDateFilter(f.field('Date').id).expected, 60.25)
  } finally { restored?.dispose(); f.dispose() }
})

test.each(['missing', 'reversed', 'nonDate'])('ordinary invalid %s date draft is rejected before content publication', async mode => {
  const f = setup()
  try {
    const name = mode === 'nonDate' ? 'Region' : 'Date', draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f, name), mode === 'missing' ? {}
      : mode === 'reversed' ? { operator: 'dateBetween', first: 46073, second: 46055 } : { first: 46055 })
    await assert.rejects(() => state.applyPivotTableFieldFilterAction(f.pivot(), f.field(name).id, 'applyDateFilter', context(f), undefined, draft)); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each([['sortAscending', 'ascending'], ['sortDescending', 'descending'], ['clearSort', undefined]])('ordinary %s merges only committed sort into a pending panel draft', async (action, direction) => {
  const f = setup()
  try {
    if (action === 'clearSort') await f.pivot().setLabelSort(f.field('Region').id, { type: 'descending' }, context(f))
    const draft = f.panel(); draft.targetAddress = 'K30'; draft.options.repeatRowLabels = true
    const before = json(f.book.save()), history = f.history(); await state.applyPivotTableFieldFilterAction(f.pivot(), f.field('Region').id, action, context(f))
    state.synchronizePivotTableFieldFilterDraft(draft, f.report(), f.field('Region').id, action)
    assert.equal(draft.fields.find(field => field.fieldId === f.field('Region').id).sort, direction)
    assert.equal(draft.targetAddress, 'K30'); assert.equal(draft.options.repeatRowLabels, true); assert.equal(f.history(), history + 1)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each([{ readonly: true }, { protectedOutput: true }])('ordinary viewable protection %j permits catalog reads but prevents content actions', async options => {
  const f = setup(options)
  try { assert.equal(state.readPivotTableFieldFilterSnapshot(f.target, f.field('Region').id).editable, false); assert.deepEqual((await memberDraft(f)).items, ['North', 'South']); assert.equal(f.history(), 0) }
  finally { f.dispose() }
})

test.each([{ hiddenSource: true }, { hiddenOutput: true }])('ordinary hidden source/output %j cannot expose a field snapshot or catalog', options => {
  const f = setup(options)
  try { assert.throws(() => state.readPivotTableFieldFilterSnapshot(f.target, f.field('Region').id)); assert.throws(() => f.pivot().getFieldItems(f.field('Region').id)); assert.equal(f.history(), 0) }
  finally { f.dispose() }
})

test.each([null, 'true', 1])('ordinary multiple-filter draft refuses non-boolean %j', value => {
  const f = setup()
  try { assert.throws(() => state.setPivotTablePanelMultipleFilters(f.panel(), value)); assert.equal(f.history(), 0) }
  finally { f.dispose() }
})

test.each([true, false])('ordinary configure multiple=%s shares category cleanup and repeated configuration is a no-op', async enabled => {
  const f = setup()
  try {
    const id = f.field('Region').id; await f.pivot().setOptions({ allowMultipleFilters: true }); await f.pivot().setLabelManualFilter(id, ['North'])
    await f.pivot().setLabelConditionFilter(id, { operator: 'captionEqual', expected: 'North' })
    const draft = f.panel(), before = json(f.book.save()), history = f.history(); state.setPivotTablePanelMultipleFilters(draft, enabled)
    const input = { ...panels.preparePivotPanelInput(f.model(), draft, f.pivotId), ...context(f) }
    api.FOffice.newAPI(f.studio).executeCommandSync(panels.CONFIGURE_PIVOT_COMMAND, input)
    assert.equal(f.pivot().getOptions().allowMultipleFilters, enabled)
    if (enabled) { assert.ok(f.field('Region').groupedItems); assert.ok(f.field('Region').labelFilter); assert.equal(f.history(), history) }
    else { assert.equal(f.field('Region').groupedItems, undefined); assert.equal(f.field('Region').labelFilter, undefined); assert.equal(total(f), 100); assert.equal(f.history(), history + 1) }
    const saved = json(f.book.save()), after = f.history(), repeat = f.panel()
    state.setPivotTablePanelMultipleFilters(repeat, enabled)
    api.FOffice.newAPI(f.studio).executeCommandSync(panels.CONFIGURE_PIVOT_COMMAND, { ...panels.preparePivotPanelInput(f.model(), repeat, f.pivotId), ...context(f) })
    assert.equal(f.history(), after); assert.deepEqual(json(f.book.save()), saved)
    if (!enabled) { f.book.undo(); assert.deepEqual(json(f.book.save()), before) }
  } finally { f.dispose() }
})

test.each(['source', 'history', 'user', 'host'])('late ordinary member read rejects changed %s owner', async change => {
  const f = setup(), mounted = controller(f), original = api.FPivotTable.prototype.getFieldItems; let release
  api.FPivotTable.prototype.getFieldItems = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
  try {
    const task = mounted.controls.open(f.field('Region').id); await changeOwner(f, change); release(); await task
    assert.equal(mounted.controls.fieldId.value, ''); assert.equal(mounted.controls.draft.value, undefined); assert.equal(mounted.controls.conditions.value, undefined)
  } finally { api.FPivotTable.prototype.getFieldItems = original; release?.(); mounted.stop(); f.dispose() }
})

test('old ordinary member finally cannot replace a newer catalog or clear its loading state', async () => {
  const f = setup(), mounted = controller(f), queue = [], original = api.FPivotTable.prototype.getFieldItems
  api.FPivotTable.prototype.getFieldItems = async function(...args) { await new Promise(resolve => queue.push(resolve)); return original.apply(this, args) }
  try {
    const first = mounted.controls.open(f.field('Region').id), second = mounted.controls.open(f.field('Segment').id)
    queue[0](); await first; assert.equal(mounted.controls.fieldId.value, f.field('Segment').id); assert.equal(mounted.controls.loading.value, true)
    queue[1](); await second; assert.deepEqual(mounted.controls.draft.value.items, ['A', 'B']); assert.equal(mounted.controls.loading.value, false)
  } finally { api.FPivotTable.prototype.getFieldItems = original; queue.forEach(release => release()); mounted.stop(); f.dispose() }
})

const mutationMethods = { manual: ['setLabelManualFilter', 'applyFilter', 'Region'], date: ['setDateFilter', 'applyDateFilter', 'Date'],
  label: ['setLabelConditionFilter', 'applyLabelCondition', 'Region'], sort: ['setLabelSort', 'sortDescending', 'Region'] }
test.each(Object.keys(mutationMethods).flatMap(kind => ['source', 'history', 'user', 'host'].map(change => [kind, change])))
  ('pending ordinary %s rejects a changed %s context without partial output/history', async (kind, change) => {
    const f = setup(), mounted = controller(f), [method, action, name] = mutationMethods[kind], original = api.FPivotTable.prototype[method]; let release
    api.FPivotTable.prototype[method] = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
    try {
      await mounted.controls.open(f.field(name).id)
      if (kind === 'manual') mounted.controls.select('South', false)
      if (kind === 'date') mounted.controls.dateChange({ first: 46055 })
      if (kind === 'label') mounted.controls.labelChange({ first: 'North' })
      const task = mounted.controls.mutate(action); await changeOwner(f, change); const before = json(f.book.save()), history = f.history()
      release(); assert.equal(await task, false); assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), history)
    } finally { api.FPivotTable.prototype[method] = original; release?.(); mounted.stop(); f.dispose() }
  })

test('old ordinary edit finally cannot clear a new busy request after history reread', async () => {
  const f = setup(), mounted = controller(f), queue = [], original = api.FPivotTable.prototype.setLabelManualFilter
  api.FPivotTable.prototype.setLabelManualFilter = async function(...args) { await new Promise(resolve => queue.push(resolve)); return original.apply(this, args) }
  try {
    await mounted.controls.open(f.field('Region').id); mounted.controls.select('South', false); const first = mounted.controls.mutate('applyFilter')
    f.studio.editor.history.clear(f.unitId); mounted.reread(); await nextTick(); await mounted.controls.open(f.field('Segment').id)
    mounted.controls.select('B', false); const second = mounted.controls.mutate('applyFilter'); queue[0](); assert.equal(await first, false)
    assert.equal(mounted.controls.busy.value, true); assert.equal(mounted.controls.fieldId.value, f.field('Segment').id)
    queue[1](); assert.equal(await second, true); assert.equal(total(f), 30); assert.equal(f.history(), 1)
  } finally { api.FPivotTable.prototype.setLabelManualFilter = original; queue.forEach(release => release()); mounted.stop(); f.dispose() }
})

test('ordinary readonly controller permits search while members and conditions stay unchanged', async () => {
  const f = setup({ readonly: true }), mounted = controller(f)
  try {
    await mounted.controls.open(f.field('Region').id); mounted.controls.search('North'); mounted.controls.select('South', false); mounted.controls.labelChange({ first: 'North' })
    assert.equal(mounted.controls.draft.value.search, 'North'); assert.deepEqual(mounted.controls.draft.value.selectedKeys, ['North', 'South'])
    assert.equal(mounted.controls.conditions.value.label.first, ''); assert.equal(await mounted.controls.mutate('applyFilter'), false); assert.equal(f.history(), 0)
  } finally { mounted.stop(); f.dispose() }
})

test('ordinary synthetic composition preserves text and blocks content/Escape until a stale input ends', async () => {
  const f = setup(), mounted = controller(f)
  try {
    await mounted.controls.open(f.field('Region').id); mounted.controls.compositionStart(); mounted.controls.search('中')
    mounted.controls.keydown({ key: 'Escape', keyCode: 229, isComposing: true, stopPropagation() {}, preventDefault() { throw new Error('Composition Escape cannot close') } })
    assert.equal(await mounted.controls.mutate('applyFilter'), false); f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)
    assert.equal(mounted.controls.stale.value, true); assert.equal(mounted.controls.draft.value.search, '中')
    mounted.controls.compositionEnd(); await nextTick(); assert.equal(mounted.controls.fieldId.value, '')
  } finally { mounted.stop(); f.dispose() }
})

test('ordinary success restores its field trigger while a superseded close cannot steal focus', async () => {
  const f = setup(), mounted = controller(f)
  try {
    await mounted.controls.open(f.field('Region').id); mounted.controls.select('South', false); assert.equal(await mounted.controls.mutate('applyFilter'), true); await nextTick()
    assert.deepEqual(mounted.calls.at(-1), ['focus', f.field('Region').id]); await mounted.controls.open(f.field('Segment').id)
    mounted.controls.close(); mounted.options.active = false; await nextTick(); assert.equal(mounted.calls.filter(call => call[0] === 'focus').length, 1)
  } finally { mounted.stop(); f.dispose() }
})

test('ordinary controller disposal rejects late members and clears session listeners/state', async () => {
  const f = setup(), mounted = controller(f), original = api.FPivotTable.prototype.getFieldItems; let release
  api.FPivotTable.prototype.getFieldItems = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
  try {
    const task = mounted.controls.open(f.field('Region').id); mounted.stop(); release(); await task
    assert.equal(mounted.controls.fieldId.value, ''); assert.equal(mounted.controls.snapshot.value, undefined); assert.equal(mounted.controls.loading.value, false)
  } finally { api.FPivotTable.prototype.getFieldItems = original; release?.(); f.dispose() }
})
