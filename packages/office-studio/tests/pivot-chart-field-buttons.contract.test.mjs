import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { pivotChartFieldButtonsIds as ids, pivotChartFieldButtonsWorkbook, pivotChartFieldButtonsReport,
  pivotChartFieldButtonField } from './fixtures/pivot-chart-field-buttons-fixture.mjs'

// Deferred contracts against real sparse cells, report algorithms, Facades and controller sessions.
// No hook, server, test, type check, build or benchmark has been executed.
let server, api, buttons, conditions, owners, controllers
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, buttons, conditions, owners, controllers] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-button-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-filter-condition.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-owner.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/views/use-pivot-chart-field-buttons.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup(options = {}, chartId = ids.owned) {
  const studio = new api.OfficeStudio({ workbook: pivotChartFieldButtonsWorkbook(options), currentUserId: 'local-user' })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), unitId = book.getId()
  const target = { studio, unitId, sheetId: ids.host, chartId }
  return { ...target, target, book, pivot: () => book.getPivotChartById(chartId), model: () => studio.getWorkbookModel(unitId),
    report: () => pivotChartFieldButtonsReport(studio.getWorkbookModel(unitId), chartId),
    field: (name, area) => pivotChartFieldButtonField(studio.getWorkbookModel(unitId), chartId, name, area),
    history: () => studio.editor.history.units[unitId]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
const commandContext = f => owners.pivotChartPanelOwnerContext(owners.capturePivotChartPanelOwner(f.studio, f.unitId, f.sheetId))
const total = f => f.pivot().getData().series.flatMap(series => series.values).reduce((sum, value) => sum + (typeof value === 'number' ? value : 0), 0)
async function memberDraft(f, name = 'Region') {
  const field = f.field(name), pivot = f.pivot(), items = await pivot.getFieldItems(field.id)
  return buttons.createPivotChartFieldFilterDraft(items, pivot.getManualFilterItemState(field.id))
}
function conditionDraft(f, name = 'Date') {
  const field = f.field(name), pivot = f.pivot()
  return conditions.createPivotChartFieldConditionsDraft(pivot.getFilterCapabilities(field.id), pivot.getDateFilter(field.id), pivot.getLabelConditionFilter(field.id))
}
function mountController(f, extras = {}) {
  const calls = [], scope = effectScope(), options = shallowReactive({ ...f.target, active: true, busy: false,
    restoreFocus: () => calls.push('focus'), openSettings: id => { calls.push(id); return true }, ...extras })
  const controls = scope.run(() => controllers.usePivotChartFieldButtons(options))
  return { controls, options, calls, stop: () => scope.stop() }
}

test('four sections expose real local IDs, authored order and defensive settings without model/history changes', () => {
  const f = setup()
  try {
    const before = json(f.book.save()), snapshot = buttons.readPivotChartFieldButtonsSnapshot(f.target)
    assert.deepEqual(snapshot.sections.map(section => section.area), ['filter', 'column', 'row', 'value'])
    assert.deepEqual(snapshot.sections.flatMap(section => section.fields.map(field => field.label)), ['Date', 'Segment', 'Region', 'Amount'])
    assert.ok(snapshot.sections.flatMap(section => section.fields).every(field => field.fieldId.startsWith('ui-owned-')))
    snapshot.sections[0].fields[0].field.displayName = 'Local mutation'
    assert.equal(f.field('Date').displayName, undefined); assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('repeated source/value instances remain separate buttons in authored positions', () => {
  const f = setup({ overflow: true })
  try {
    const values = buttons.readPivotChartFieldButtonsSnapshot(f.target).sections.find(section => section.area === 'value').fields
    assert.deepEqual(values.map(field => field.label), ['Amount', 'Amount average', 'Amount count'])
    assert.equal(new Set(values.map(field => field.fieldId)).size, 3); assert.equal(new Set(values.map(field => field.field.sourceColumn)).size, 1)
    assert.deepEqual(values.map(field => field.field.aggregation), ['sum', 'avg', 'count'])
  } finally { f.dispose() }
})

test.each([
  [{ selectedItems: ['North'] }, true], [{ groupedItems: ['North'] }, true], [{ recordSelection: { mode: 'include', values: [1] } }, true],
  [{ recordSelection: { mode: 'exclude', values: [1] } }, true], [{ recordSelection: { mode: 'exclude', values: [] } }, false],
  [{ recordSelection: { mode: 'all' } }, false], [{ dateFilter: { operator: 'dateEqual', expected: 1 } }, true],
  [{ labelFilter: { operator: 'captionEqual', expected: 'North' } }, true],
  [{ valueFilter: { operator: 'valueGreaterThan', expected: 5, valueFieldId: 'amount' } }, true],
])('field badges reflect effective manual/typed/date/label/value state %j', (state, filtered) => {
  const f = setup()
  try {
    const field = { ...json(f.field('Region')), ...state }
    const button = buttons.pivotChartFieldButtonSections({ fields: [field] }).find(section => section.area === 'row').fields[0]
    assert.equal(button.filtered, filtered)
  } finally { f.dispose() }
})

test.each([false, true])('grouped checklist reads a year rather than raw date strings (partial=%s)', async partialDate => {
  const f = setup({ grouped: true, partialDate })
  try {
    const draft = await memberDraft(f, 'Date')
    assert.deepEqual(draft.items, ['2026']); assert.deepEqual(draft.selectedKeys, partialDate ? [] : ['2026'])
    assert.deepEqual(draft.partialKeys, partialDate ? ['2026'] : []); assert.equal(total(f), partialDate ? 10 : 100)
  } finally { f.dispose() }
})

test.each([['Region', { typedMembers: true }, '1', 10], ['Date', { partialDate: true }, '2026', 10]])
  ('unchanged partial %s selection preserves raw scalar membership through command/history/reopen', async (name, options, key, expected) => {
    const f = setup(options); let reopened
    try {
      const draft = await memberDraft(f, name), field = f.field(name), before = json(f.book.save())
      assert.deepEqual(draft.partialKeys, [key]); assert.equal(draft.selectedKeys.length, 0)
      const edit = buttons.pivotChartFieldFilterEdit(field.id, draft); assert.deepEqual(edit.partialItems, [key]); assert.equal(edit.isAll, false)
      assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), field.id, 'applyFilter', commandContext(f), draft), true)
      assert.equal(total(f), expected); assert.equal(f.history(), 1); assert.equal(f.field(name).recordSelection.mode, 'include')
      if (name === 'Region') assert.deepEqual(f.field(name).recordSelection.values, [1])
      const saved = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
      reopened = new api.OfficeStudio({ workbook: saved }); const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(ids.owned)
      assert.deepEqual(restored.getManualFilterItemState(field.id).partialItems, [key]); assert.equal(restored.getData().series.flatMap(series => series.values).filter(value => typeof value === 'number').reduce((a, b) => a + b, 0), expected)
    } finally { reopened?.dispose(); f.dispose() }
  })

test.each([['Region', { typedMembers: true }, '1', 30], ['Date', { partialDate: true }, '2026', 100]])
  ('clicking partial %s to checked deliberately selects all underlying members', async (name, options, key, expected) => {
    const f = setup(options)
    try {
      const draft = buttons.setPivotChartFilterSelection(await memberDraft(f, name), key, true)
      assert.deepEqual(draft.partialKeys, []); assert.ok(draft.selectedKeys.includes(key))
      assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field(name).id, 'applyFilter', commandContext(f), draft), true)
      assert.equal(total(f), expected)
    } finally { f.dispose() }
  })

test('unchecking the only partial group produces a disabled empty draft and no content publication', async () => {
  const f = setup({ partialDate: true })
  try {
    const draft = buttons.setPivotChartFilterSelection(await memberDraft(f, 'Date'), '2026', false)
    assert.deepEqual(draft.partialKeys, []); assert.deepEqual(draft.selectedKeys, [])
    assert.throws(() => buttons.pivotChartFieldFilterEdit(f.field('Date').id, draft), /至少选择/); assert.equal(total(f), 10); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('search and an off-page edit preserve all other selections and partial membership', () => {
  const items = Array.from({ length: 130 }, (_, index) => `Region ${String(index).padStart(3, '0')}`)
  let draft = buttons.createPivotChartFieldFilterDraft(items, { checkedItems: items.slice(1), partialItems: [items[0]] })
  draft = buttons.setPivotChartFilterSearch(draft, '129'); assert.deepEqual(buttons.pivotChartFilterVisibleMembers(draft), [items[129]])
  draft = buttons.setPivotChartFilterSelection(draft, items[129], false)
  assert.equal(draft.selectedKeys.length, 128); assert.deepEqual(draft.partialKeys, [items[0]])
  assert.equal(buttons.setPivotChartFilterSearch(draft, '').selectedKeys.length, 128)
  const all = buttons.setPivotChartFilterAll(draft, true); assert.deepEqual(all.partialKeys, []); assert.equal(all.selectedKeys.length, 130)
})

test('blank members keep their empty key while search uses the accessible display label', () => {
  const draft = buttons.setPivotChartFilterSearch(buttons.createPivotChartFieldFilterDraft(['', 'North'], { checkedItems: ['', 'North'], partialItems: [] }), '空白')
  assert.deepEqual(buttons.pivotChartFilterVisibleMembers(draft), ['']); assert.equal(buttons.pivotChartFilterMemberLabel(''), '（空白）')
})

test.each([
  [['North', 'North'], { checkedItems: ['North'], partialItems: [] }], [[1], { checkedItems: [], partialItems: [] }],
  [['North'], { checkedItems: ['South'], partialItems: [] }], [['North'], { checkedItems: [], partialItems: ['South'] }],
  [['North'], { checkedItems: ['North', 'North'], partialItems: [] }], [['North'], { checkedItems: ['North'], partialItems: ['North'] }],
])('malformed catalog/state is rejected rather than reinterpreting raw member identity %j', (items, state) => {
  assert.throws(() => buttons.createPivotChartFieldFilterDraft(items, state))
})

test.each([ids.owned, ids.linked])('manual field action edits %s through one real command with history and source preservation', async chartId => {
  const f = setup({}, chartId)
  try {
    const before = json(f.book.save()), source = json(f.model().sheets.find(sheet => sheet.id === ids.source).cells), commands = []
    f.studio.events.addEvent('CommandExecuted', event => commands.push(event))
    const draft = buttons.setPivotChartFilterSelection(await memberDraft(f), 'South', false)
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Region').id, 'applyFilter', commandContext(f), draft), true)
    assert.equal(total(f), 40); assert.equal(f.history(), 1); assert.deepEqual(f.model().sheets.find(sheet => sheet.id === ids.source).cells, source)
    assert.equal(commands.filter(event => event.id === 'sheet.command.edit-pivot-chart-report').length, 1)
    if (chartId === ids.owned) assert.deepEqual(f.model().sheets.find(sheet => sheet.id === ids.host).cells, before.sheets.find(sheet => sheet.id === ids.host).cells)
    else assert.notDeepEqual(f.model().sheets.find(sheet => sheet.id === ids.host).cells, before.sheets.find(sheet => sheet.id === ids.host).cells)
    const saved = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
  } finally { f.dispose() }
})

test('select-all clears only manual members; clear-all removes manual, label and value categories', async () => {
  const f = setup()
  try {
    await f.pivot().applyReportEdit({ type: 'setOptions', options: { allowMultipleFilters: true } })
    await f.pivot().setLabelManualFilter(f.field('Region').id, ['North'])
    await f.pivot().setLabelConditionFilter(f.field('Region').id, { operator: 'captionEqual', expected: 'North' })
    await f.pivot().setValueFilter(f.field('Region').id, { operator: 'valueGreaterThanOrEqual', expected: 35, valueFieldId: f.field('Amount').id })
    let draft = buttons.setPivotChartFilterAll(await memberDraft(f), true)
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Region').id, 'applyFilter', commandContext(f), draft), true)
    assert.equal(f.field('Region').selectedItems, undefined); assert.ok(f.field('Region').labelFilter); assert.ok(f.field('Region').valueFilter); assert.equal(total(f), 40)
    const beforeClear = json(f.book.save()), history = f.history()
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Region').id, 'clearFilter', commandContext(f)), true)
    for (const key of ['selectedItems', 'groupedItems', 'partialItems', 'recordSelection', 'labelFilter', 'dateFilter', 'valueFilter']) assert.equal(f.field('Region')[key], undefined)
    assert.equal(total(f), 100); assert.equal(f.history(), history + 1); f.book.undo(); assert.deepEqual(json(f.book.save()), beforeClear)
  } finally { f.dispose() }
})

test.each([['sortAscending', 'ascending', ['North', 'South']], ['sortDescending', 'descending', ['South', 'North']], ['clearSort', undefined, ['North', 'South']]])
  ('sort action %s changes real category order and remains undoable', async (action, direction, categories) => {
    const f = setup()
    try {
      if (action === 'clearSort') await f.pivot().setLabelSort(f.field('Region').id, { type: 'descending' })
      const before = json(f.book.save()), history = f.history()
      assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Region').id, action, commandContext(f)), true)
      assert.equal(f.field('Region').sort, direction); assert.deepEqual(f.pivot().getData().categories, categories)
      assert.equal(f.history(), history + 1); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    } finally { f.dispose() }
  })

test('remove targets one repeated value instance and preserves its source sibling and chart appearance', async () => {
  const f = setup({ overflow: true })
  try {
    const field = f.report().fields.find(field => field.displayName === 'Amount average'), before = json(f.book.save()), options = json(f.pivot().getInfo().options)
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), field.id, 'remove', commandContext(f)), true)
    assert.equal(f.report().fields.filter(field => field.area === 'value').length, 2); assert.equal(f.report().fields.some(candidate => candidate.id === field.id), false)
    assert.ok(f.report().fields.some(candidate => candidate.displayName === 'Amount count')); assert.deepEqual(f.pivot().getInfo().options, options)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each([[{ readonly: true }, ids.owned, false], [{ protectedChart: true }, ids.owned, false], [{ protectedOutput: true }, ids.linked, false]])
  ('viewable protection %j allows reading but disables report content edits', (options, chartId, editable) => {
    const f = setup(options, chartId)
    try { const snapshot = buttons.readPivotChartFieldButtonsSnapshot(f.target); assert.equal(snapshot.editable, editable); assert.equal(snapshot.sections.length, 4); assert.equal(f.history(), 0) }
    finally { f.dispose() }
  })

test.each([[{ hiddenSource: true }, ids.owned], [{ hiddenSource: true }, ids.linked], [{ hiddenOutput: true }, ids.linked]])
  ('unviewable source/output %j does not expose fields or register content', (options, chartId) => {
    const f = setup(options, chartId)
    try { assert.throws(() => buttons.readPivotChartFieldButtonsSnapshot(f.target)); assert.equal(f.history(), 0) }
    finally { f.dispose() }
  })

test.each([[true, true], [true, false], [false, true], [false, false]])('single controls metadata drives field visibility (fields=%s expand=%s)', async (showFieldButtons, showExpandCollapseButtons) => {
  const f = setup(), mounted = mountController(f)
  try {
    await mounted.controls.openField(f.field('Region').id)
    await f.pivot().setControls({ showFieldButtons, showExpandCollapseButtons }); await nextTick()
    assert.equal(mounted.controls.visible.value, showFieldButtons)
    assert.deepEqual(mounted.controls.snapshot.value.controls, { showFieldButtons, showExpandCollapseButtons })
    if (!showFieldButtons) { assert.equal(mounted.controls.openAnchor.value, ''); assert.equal(mounted.controls.draft.value, undefined) }
    assert.equal(f.pivot().getInfo().options?.pivotChart?.showFieldButtons, undefined)
  } finally { mounted.stop(); f.dispose() }
})

test.each([
  ['dateBetween', '2026-02-02', '2026-02-20', 50], ['dateEqual', '2026-02-02', undefined, 20], ['dateNewerThan', '2026-02-02', undefined, 70],
  ['dateNewerThanOrEqual', '2026-02-02', undefined, 90], ['dateNotBetween', '2026-02-02', '2026-02-20', 50], ['dateNotEqual', '2026-02-02', undefined, 80],
  ['dateOlderThan', '2026-02-02', undefined, 10], ['dateOlderThanOrEqual', '2026-02-02', undefined, 30],
])('date condition %s selects the corresponding real source records', async (operator, first, second, expected) => {
  const f = setup()
  try {
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { operator,
      first: conditions.pivotChartFilterDateFromCalendarInput(first, false), second: second ? conditions.pivotChartFilterDateFromCalendarInput(second, false) : undefined })
    const before = json(f.book.save())
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Date').id, 'applyDateFilter', commandContext(f), undefined, draft), true)
    assert.equal(total(f), expected); assert.equal(f.pivot().getDateFilter(f.field('Date').id).operator, operator); assert.equal(f.history(), 1)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each([
  ['captionBeginsWith', 'N', '', 40], ['captionBetween', 'North', 'South', 100], ['captionContains', 'or', '', 40], ['captionEndsWith', 'orth', '', 40],
  ['captionEqual', 'North', '', 40], ['captionGreaterThan', 'North', '', 60], ['captionGreaterThanOrEqual', 'North', '', 100],
  ['captionLessThan', 'South', '', 40], ['captionLessThanOrEqual', 'South', '', 100], ['captionNotBeginsWith', 'N', '', 60],
  ['captionNotBetween', 'North', 'South', 0], ['captionNotContains', 'or', '', 60], ['captionNotEndsWith', 'orth', '', 60], ['captionNotEqual', 'North', '', 60],
])('label condition %s uses the shared caption semantics and real records', async (operator, first, second, expected) => {
  const f = setup()
  try {
    const draft = conditions.setPivotChartLabelConditionDraft(conditionDraft(f, 'Region'), { operator, first, second })
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Region').id, 'applyLabelCondition', commandContext(f), undefined, draft), true)
    assert.equal(total(f), expected); assert.equal(f.pivot().getLabelConditionFilter(f.field('Region').id).operator, operator); assert.equal(f.history(), 1)
  } finally { f.dispose() }
})

test.each([false, true])('whole-day=%s preserves exact serial operands and compares time as specified', async useWholeDay => {
  const f = setup({ dateSerials: true })
  try {
    const serial = conditions.pivotChartFilterDateFromCalendarInput('2026-02-20', false)
    f.book.getSheetBySheetId(ids.source).getRange('F3').setValue(serial); f.book.getSheetBySheetId(ids.source).getRange('F4').setValue(serial + .5)
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { first: serial, useWholeDay })
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Date').id, 'applyDateFilter', commandContext(f), undefined, draft), true)
    assert.equal(total(f), useWholeDay ? 60 : 40); assert.equal(f.pivot().getDateFilter(f.field('Date').id).expected, serial)
  } finally { f.dispose() }
})

test.each([false, true])('date calendar adapter and field capability honor date1904=%s', async date1904 => {
  const f = setup({ date1904, dateSerials: true })
  try {
    const epoch = conditions.pivotChartFilterDateFromCalendarInput('1904-01-01', date1904); assert.equal(epoch, date1904 ? 0 : 1462)
    assert.deepEqual(f.pivot().getFilterCapabilities(f.field('Date').id), { supportsDateFilter: true, date1904 })
    const first = conditions.pivotChartFilterDateFromCalendarInput('2026-02-02 00:00:00.000', date1904)
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { first })
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Date').id, 'applyDateFilter', commandContext(f), undefined, draft), true)
    assert.equal(total(f), 20); assert.equal(conditions.pivotChartFilterDateCalendarInput(first, date1904), '2026-02-02 00:00:00.000')
  } finally { f.dispose() }
})

test('Excel compatibility day and fractional operands survive typed condition draft and JSON reopen', async () => {
  const f = setup({ compatibilityDate: true }); let reopened
  try {
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f), { first: 60.25 })
    assert.equal(conditions.pivotChartFilterDateCalendarInput(60.25, false), '1900-02-29 06:00:00.000')
    assert.equal(conditions.pivotChartFilterDateFromCalendarInput('1900-02-29 06:00:00.000', false), 60.25)
    assert.equal(conditions.pivotChartFilterDateFromCalendarInput('1900-02-29', true), undefined)
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), f.field('Date').id, 'applyDateFilter', commandContext(f), undefined, draft), true); assert.equal(total(f), 40)
    const saved = json(f.book.save()); reopened = new api.OfficeStudio({ workbook: saved })
    const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(ids.owned)
    assert.equal(restored.getDateFilter(f.field('Date').id).expected, 60.25); assert.equal(conditionDraft(f).date.first, 60.25)
  } finally { reopened?.dispose(); f.dispose() }
})

test('date capability cannot be guessed from a renamed field or numeric member display', async () => {
  const f = setup({ typedMembers: true })
  try {
    await f.pivot().setFieldSetting(f.field('Region').id, { displayName: 'Date' })
    const draft = conditions.setPivotChartDateConditionDraft(conditionDraft(f, 'Region'), { first: 60 })
    assert.equal(draft.capabilities.supportsDateFilter, false); assert.throws(() => conditions.pivotChartDateConditionFilter(draft), /不支持日期/)
    assert.equal(f.pivot().getFilterCapabilities(f.field('Amount').id).supportsDateFilter, false)
  } finally { f.dispose() }
})

test.each(['2026-02-30', '2026-01-01 24:00:00', '2026-01-01 12:60:00', '2026-01-01 12:00:60', 'invalid'])
  ('invalid calendar input %s stays invalid instead of normalizing into another date', value => {
    assert.equal(conditions.pivotChartFilterDateFromCalendarInput(value, false), undefined)
  })

test('missing or reversed date bounds reject the UI draft before content is published', () => {
  const f = setup()
  try {
    const base = conditionDraft(f); assert.throws(() => conditions.pivotChartDateConditionFilter(base), /有效日期/)
    const reversed = conditions.setPivotChartDateConditionDraft(base, { operator: 'dateBetween', first: 2, second: 1 })
    assert.throws(() => conditions.pivotChartDateConditionFilter(reversed), /结束日期/); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['date', 'label'])('clearing %s condition preserves other allowed filter categories and can be undone', async kind => {
  const f = setup()
  try {
    await f.pivot().applyReportEdit({ type: 'setOptions', options: { allowMultipleFilters: true } })
    const field = f.field('Date'), all = await memberDraft(f, 'Date')
    const selected = buttons.setPivotChartFilterSelection(all, '2026-01-15', false)
    await buttons.applyPivotChartFieldButtonAction(f.pivot(), field.id, 'applyFilter', commandContext(f), selected)
    await f.pivot().setDateFilter(field.id, { operator: 'dateOlderThanOrEqual', expected: conditions.pivotChartFilterDateFromCalendarInput('2026-02-20', false) })
    await f.pivot().setLabelConditionFilter(field.id, { operator: 'captionContains', expected: '2026-' })
    const before = json(f.book.save()), history = f.history()
    assert.equal(await buttons.applyPivotChartFieldButtonAction(f.pivot(), field.id, kind === 'date' ? 'clearDateFilter' : 'clearLabelCondition', commandContext(f)), true)
    assert.ok(f.field('Date').groupedItems); assert.equal(kind === 'date' ? f.field('Date').dateFilter : f.field('Date').labelFilter, undefined)
    assert.ok(kind === 'date' ? f.field('Date').labelFilter : f.field('Date').dateFilter); assert.equal(f.history(), history + 1)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test('older member completion never replaces a newer field or clears its loading state', async () => {
  const f = setup(), mounted = mountController(f), queue = [], original = api.FSheetPivotChart.prototype.getFieldItems
  api.FSheetPivotChart.prototype.getFieldItems = async function(...args) { await new Promise(resolve => queue.push(resolve)); return original.apply(this, args) }
  try {
    const first = mounted.controls.openField(f.field('Region').id), second = mounted.controls.openField(f.field('Segment').id)
    assert.equal(queue.length, 2); queue[0](); await first
    assert.equal(mounted.controls.field.value.fieldId, f.field('Segment').id); assert.equal(mounted.controls.loading.value, true); assert.equal(mounted.controls.draft.value, undefined)
    queue[1](); await second; assert.equal(mounted.controls.loading.value, false); assert.deepEqual(mounted.controls.draft.value.items, ['A', 'B'])
    assert.equal(f.history(), 0)
  } finally { api.FSheetPivotChart.prototype.getFieldItems = original; queue.forEach(resolve => resolve()); mounted.stop(); f.dispose() }
})

test('old content finally cannot clear a newer busy request after history generation changes', async () => {
  const f = setup(), mounted = mountController(f), queue = [], original = api.FSheetPivotChart.prototype.applyReportEdit
  api.FSheetPivotChart.prototype.applyReportEdit = async function(...args) { await new Promise(resolve => queue.push(resolve)); return original.apply(this, args) }
  try {
    await mounted.controls.openField(f.field('Region').id); mounted.controls.select('South', false); const first = mounted.controls.mutate('applyFilter')
    f.studio.editor.history.clear(f.unitId); await nextTick(); await mounted.controls.openField(f.field('Segment').id)
    mounted.controls.select('B', false); const second = mounted.controls.mutate('applyFilter'); assert.equal(queue.length, 2)
    queue[0](); assert.equal(await first, false); assert.equal(mounted.controls.busy.value, true); assert.equal(mounted.controls.field.value.fieldId, f.field('Segment').id)
    queue[1](); assert.equal(await second, true); assert.equal(mounted.controls.busy.value, false); assert.equal(total(f), 30); assert.equal(f.history(), 1)
  } finally { api.FSheetPivotChart.prototype.applyReportEdit = original; queue.forEach(resolve => resolve()); mounted.stop(); f.dispose() }
})

test.each(['source', 'history', 'user', 'sheet'])('pending member read rejects a changed %s owner', async change => {
  const f = setup(), mounted = mountController(f), original = api.FSheetPivotChart.prototype.getFieldItems; let release
  api.FSheetPivotChart.prototype.getFieldItems = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
  try {
    const task = mounted.controls.openField(f.field('Region').id)
    if (change === 'source') f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)
    if (change === 'history') f.studio.editor.history.clear(f.unitId)
    if (change === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
    if (change === 'sheet') f.book.setActiveSheet(ids.other)
    release(); await task; assert.equal(mounted.controls.draft.value, undefined); assert.equal(mounted.controls.conditions.value, undefined)
    assert.equal(mounted.controls.loading.value, false); assert.equal(f.field('Region').selectedItems, undefined)
  } finally { api.FSheetPivotChart.prototype.getFieldItems = original; release?.(); mounted.stop(); f.dispose() }
})

test.each(['source', 'history', 'user', 'sheet'])('pending manual publication rejects a changed %s context without partial edits', async change => {
  const f = setup(), mounted = mountController(f), original = api.FSheetPivotChart.prototype.applyReportEdit; let release
  api.FSheetPivotChart.prototype.applyReportEdit = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
  try {
    await mounted.controls.openField(f.field('Region').id); mounted.controls.select('South', false); const task = mounted.controls.mutate('applyFilter')
    if (change === 'source') f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)
    if (change === 'history') f.studio.editor.history.clear(f.unitId)
    if (change === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
    if (change === 'sheet') f.book.setActiveSheet(ids.other)
    const beforeRelease = json(f.book.save()), history = f.history(); release(); assert.equal(await task, false)
    assert.deepEqual(json(f.book.save()), beforeRelease); assert.equal(f.history(), history); assert.equal(mounted.controls.busy.value, false)
  } finally { api.FSheetPivotChart.prototype.applyReportEdit = original; release?.(); mounted.stop(); f.dispose() }
})

test('readonly controller permits search/settings while refusing member and condition mutations', async () => {
  const f = setup({ readonly: true }), mounted = mountController(f)
  try {
    await mounted.controls.openField(f.field('Region').id); mounted.controls.search('South'); mounted.controls.select('North', false)
    mounted.controls.labelChange({ first: 'South' }); assert.deepEqual(mounted.controls.draft.value.selectedKeys, ['North', 'South'])
    assert.equal(mounted.controls.draft.value.search, 'South'); assert.equal(mounted.controls.conditions.value.label.first, '')
    assert.equal(await mounted.controls.mutate('applyFilter'), false); await mounted.controls.settings()
    assert.deepEqual(mounted.calls, [f.field('Region').id]); assert.equal(f.history(), 0)
  } finally { mounted.stop(); f.dispose() }
})

test('composition blocks Escape/content actions and preserves a stale input until composition ends', async () => {
  const f = setup(), mounted = mountController(f)
  try {
    await mounted.controls.openField(f.field('Region').id); mounted.controls.compositionStart(); mounted.controls.search('中')
    mounted.controls.keydown({ key: 'Escape', keyCode: 229, isComposing: true, stopPropagation() {}, preventDefault() { throw new Error('Composing Escape must not cancel') } })
    assert.ok(mounted.controls.openAnchor.value); assert.equal(await mounted.controls.mutate('remove'), false)
    f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11); assert.equal(mounted.controls.stale.value, true)
    assert.ok(mounted.controls.openAnchor.value); assert.equal(mounted.controls.draft.value.search, '中')
    mounted.controls.compositionEnd(); await nextTick(); assert.equal(mounted.controls.openAnchor.value, '')
  } finally { mounted.stop(); f.dispose() }
})

test('successful action restores focus only to its still-current owner', async () => {
  const f = setup(), mounted = mountController(f)
  try {
    await mounted.controls.openField(f.field('Region').id); mounted.controls.select('South', false)
    assert.equal(await mounted.controls.mutate('applyFilter'), true); await nextTick(); assert.deepEqual(mounted.calls, ['focus'])
    await mounted.controls.openField(f.field('Segment').id); mounted.controls.close(); mounted.options.active = false; await nextTick()
    assert.deepEqual(mounted.calls, ['focus']); assert.equal(f.history(), 1)
  } finally { mounted.stop(); f.dispose() }
})

test('disposing a controller rejects old member completion and removes its interaction state', async () => {
  const f = setup(), mounted = mountController(f), original = api.FSheetPivotChart.prototype.getFieldItems; let release
  api.FSheetPivotChart.prototype.getFieldItems = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
  try {
    const task = mounted.controls.openField(f.field('Region').id); mounted.stop(); release(); await task
    assert.equal(mounted.controls.openAnchor.value, ''); assert.equal(mounted.controls.snapshot.value, undefined); assert.equal(mounted.controls.draft.value, undefined)
    assert.equal(mounted.controls.loading.value, false); assert.equal(f.history(), 0)
  } finally { api.FSheetPivotChart.prototype.getFieldItems = original; release?.(); mounted.stop(); f.dispose() }
})
