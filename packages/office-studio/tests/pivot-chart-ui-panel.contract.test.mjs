import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { pivotChartUIIds as ids, pivotChartUIRange as sourceRange, pivotChartUIOptions as reportOptions,
  pivotChartUIWorkbook, pivotChartUIReport } from './fixtures/pivot-chart-ui-fixture.mjs'

// Deferred contracts. No test, Vite hook, type check, build or browser below has been run.
let server, api, panel, fields, owners, hooks, vue, ribbon
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, panel, fields, owners, hooks, vue, ribbon] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-draft.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-owner.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/views/use-sheet-chart-panels.ts'),
    server.ssrLoadModule('vue'),
    server.ssrLoadModule('/packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const model = f => f.studio.getWorkbookModel(f.unitId)
const undoCount = f => f.studio.editor.history.units[f.unitId]?.undo.length ?? 0
function setup(options = {}, currentUserId = 'local-user') {
  const studio = new api.OfficeStudio({ workbook: pivotChartUIWorkbook(options), currentUserId }), book = api.FOffice.newAPI(studio).getActiveWorkbook()
  return { studio, book, unitId: book.getId(), sheet: book.getSheetBySheetId(ids.host), dispose: () => studio.dispose() }
}
function draftFor(f, chartId = ids.owned) {
  const draft = panel.createPivotChartPanelDraft(model(f), ids.host, { chartId })
  const source = panel.readPivotChartPanelSource(model(f), draft, ids.host, f.studio.currentUserId)
  draft.fieldsConfig = panel.pivotChartFieldsFromReport(source.report); return { draft, source }
}
function freshDraft(f) {
  const draft = panel.createPivotChartPanelDraft(model(f), ids.host, { sourceKind: 'range', sourceSheetId: ids.source, sourceRange })
  const source = panel.readPivotChartPanelSource(model(f), draft, ids.host, f.studio.currentUserId)
  for (const [id, area] of [['0', 'row'], ['1', 'value']]) draft.fieldsConfig = fields.addPivotChartDraftField(draft.fieldsConfig, source.fields.find(field => field.fieldId === id), area)
  draft.fieldsConfig.options = { ...reportOptions }; return { draft, source }
}
const fieldBySource = (config, id) => config.fields.find(field => field.sourceFieldId === id)
const owner = f => owners.capturePivotChartPanelOwner(f.studio, f.unitId, ids.host)
const commandContext = f => owners.pivotChartPanelOwnerContext(owner(f))

test('fresh source reads real nonzero-offset metadata without reserving reports, charts, caches or cells', () => {
  const f = setup({ existing: false })
  try {
    const before = json(model(f)), baseline = undoCount(f), { draft, source } = freshDraft(f)
    assert.deepEqual(source.fields.map(field => [field.fieldId, field.fieldName, field.sourceColumn]), [
      ['0', 'Region', 3], ['1', 'Amount', 4], ['2', 'Date', 5], ['3', 'Segment', 6], ['4', 'Score', 7],
    ])
    assert.equal(draft.sourceRange, 'D3:H7'); assert.equal(source.report.outputRange, undefined)
    assert.equal(source.report.generatedCellKeys, undefined); assert.deepEqual(json(model(f)), before); assert.equal(undoCount(f), baseline)
  } finally { f.dispose() }
})

for (const [chartId, kind] of [[ids.owned, 'range'], [ids.linked, 'pivotTable']]) test(`${kind} authoring reads the actual owner and detached field configuration`, () => {
  const f = setup()
  try {
    const before = json(model(f)), { draft, source } = draftFor(f, chartId)
    assert.equal(draft.sourceKind, kind); assert.equal(draft.sourceSheetId, ids.source); assert.equal(draft.sourceRange, 'D3:H7')
    assert.deepEqual(draft.fieldsConfig.fields, pivotChartUIReport(model(f), chartId).fields)
    assert.deepEqual(source.fields.map(field => field.fieldId), ['0', '1', '2', '3', '4'])
    draft.fieldsConfig.fields[0].displayName = 'Draft only'; source.report.options.valuePosition = 'row'
    assert.deepEqual(json(model(f)), before); assert.equal(undoCount(f), 0)
  } finally { f.dispose() }
})

test('source navigation keeps the currently edited range and clears only source-specific chart mappings', () => {
  const f = setup()
  try {
    const { draft } = draftFor(f), author = json(draft.settings.config.options), originalFields = json(draft.fieldsConfig)
    draft.sourceRange = 'D3:G7'; draft.settings.config.options.mapping = { categoryIndexes: [2], valueIndexes: [3] }
    draft.settings.config.categoryField = 2; draft.settings.config.valueFields = [3]; draft.settings.config.seriesToCategory = { 3: 2 }
    const input = panel.pivotChartPanelInputFromDraft(draft)
    assert.deepEqual(input.sourceRange, { startRow: 2, endRow: 6, startCol: 3, endCol: 6 })
    const fresh = panel.createPivotChartPanelDraft(model(f), ids.host, { ...input, chartId: undefined })
    assert.equal(fresh.sourceRange, 'D3:G7'); panel.resetPivotChartDraftMapping(draft)
    assert.equal(draft.settings.config.categoryField, 0); assert.deepEqual(draft.settings.config.valueFields, [])
    assert.deepEqual(draft.settings.config.seriesToCategory, {}); assert.equal(draft.settings.config.options.mapping, undefined)
    assert.deepEqual(draft.settings.config.options, author); assert.deepEqual(draft.fieldsConfig, originalFields)
  } finally { f.dispose() }
})

test('a source needs header plus data, belongs to this workbook and cannot replace existing chart ownership', () => {
  const f = setup()
  try {
    const { draft } = draftFor(f)
    draft.sourceKind = 'pivotTable'; draft.pivotTableId = ids.primary
    assert.throws(() => panel.readPivotChartPanelSource(model(f), draft, ids.host, f.studio.currentUserId), /归属/)
    draft.chartId = ''; draft.sourceKind = 'range'; draft.sourceRange = 'D3:H3'
    assert.throws(() => panel.readPivotChartPanelSource(model(f), draft, ids.host, f.studio.currentUserId), /标题|数据/)
    draft.sourceRange = 'D3:H7'; draft.sourceSheetId = 'missing'
    assert.throws(() => panel.readPivotChartPanelSource(model(f), draft, ids.host, f.studio.currentUserId), /来源工作表/)
    assert.equal(undoCount(f), 0)
  } finally { f.dispose() }
})

for (const [name, options, chartId] of [
  ['fresh hidden source', { existing: false, hiddenSource: true }, undefined],
  ['owned hidden source', { hiddenSource: true }, ids.owned],
  ['linked hidden output', { hiddenOutput: true }, ids.linked],
]) test(`${name} refuses field/member exposure before creating a draft preview`, () => {
  const f = setup(options, 'peer')
  try {
    const before = json(model(f)), draft = panel.createPivotChartPanelDraft(model(f), ids.host,
      chartId ? { chartId } : { sourceKind: 'range', sourceSheetId: ids.source, sourceRange })
    assert.throws(() => panel.readPivotChartPanelSource(model(f), draft, ids.host, 'peer'), /无法访问/)
    assert.deepEqual(json(model(f)), before); assert.equal(undoCount(f), 0)
  } finally { f.dispose() }
})

for (const [name, options, chartId] of [
  ['worksheet edit protection', { readonly: true }, ids.owned],
  ['chart range edit protection', { protectedChart: true }, ids.owned],
  ['linked output edit protection', { protectedOutput: true }, ids.linked],
]) test(`${name} leaves readable metadata available while disabling configuration`, () => {
  const f = setup(options, 'peer')
  try {
    const { draft, source } = draftFor(f, chartId), chart = model(f).charts.find(value => value.id === chartId)
    assert.equal(panel.canViewPivotChartPanelSource(model(f), source.report, ids.host, chart, 'peer', draft.sourceKind === 'pivotTable'), true)
    assert.equal(panel.canConfigurePivotChartPanel(model(f), source.report, ids.host, chart, 'peer', draft.sourceKind === 'pivotTable'), false)
    assert.equal(undoCount(f), 0)
  } finally { f.dispose() }
})

test('adding an available field resolves canonical identity while moving a local instance instead of duplicating it', () => {
  const f = setup()
  try {
    const { draft, source } = draftFor(f), before = json(draft.fieldsConfig), localId = fieldBySource(draft.fieldsConfig, '0').id
    const next = fields.addPivotChartDraftField(draft.fieldsConfig, source.fields[0], 'column')
    assert.equal(next.fields.filter(field => field.sourceFieldId === '0').length, 1)
    assert.equal(fieldBySource(next, '0').id, localId); assert.equal(fieldBySource(next, '0').area, 'column')
    assert.deepEqual(draft.fieldsConfig, before)
  } finally { f.dispose() }
})

test('same-source field instances move and reorder by exact local ID without affecting their sibling', () => {
  const config = { fields: [
    { id: 'local-a', sourceFieldId: 'source-a', sourceColumn: 3, name: 'Region', area: 'row', position: 0 },
    { id: 'local-b', sourceFieldId: 'source-a', sourceColumn: 3, name: 'Region copy', area: 'row', position: 1 },
    { id: 'local-c', sourceFieldId: 'source-c', sourceColumn: 4, name: 'Amount', area: 'column', position: 0 },
  ] }
  const moved = fields.movePivotChartDraftField(config, 'local-b', 'column', 0)
  assert.deepEqual(fields.pivotChartFieldsInArea(moved, 'row').map(field => [field.id, field.position]), [['local-a', 0]])
  assert.deepEqual(fields.pivotChartFieldsInArea(moved, 'column').map(field => [field.id, field.position]), [['local-b', 0], ['local-c', 1]])
  const reordered = fields.movePivotChartDraftField(moved, 'local-b', 'column', 1)
  assert.deepEqual(fields.pivotChartFieldsInArea(reordered, 'column').map(field => [field.id, field.position]), [['local-c', 0], ['local-b', 1]])
  assert.equal(config.fields[1].area, 'row'); assert.equal(config.fields[1].position, 1)
})

test('moving a filtered category to values removes incompatible filter/group/sort state and preserves the original draft', () => {
  const config = { fields: [{ id: 'r', sourceFieldId: '0', sourceColumn: 3, name: 'Region', area: 'row', position: 0,
    selectedItems: [], sort: { type: 'asc' }, manualGroups: [{ name: 'Group', items: ['North'] }],
    numberGroup: { start: 0, end: 10, step: 1 }, dateGroupType: ['month'], valueFilter: { type: 'greaterThan', value: 1 } }] }
  const next = fields.movePivotChartDraftField(config, 'r', 'value', 0)
  assert.equal(next.fields[0].aggregation, 'sum')
  for (const key of ['selectedItems', 'sort', 'manualGroups', 'numberGroup', 'dateGroupType', 'valueFilter']) assert.equal(Object.hasOwn(next.fields[0], key), false)
  assert.deepEqual(config.fields[0].selectedItems, []); assert.equal(config.fields[0].area, 'row')
})

test('removal renumbers each field area and missing IDs or invalid destinations reject without changing the original', () => {
  const config = { fields: ['a', 'b', 'c'].map((id, position) => ({ id, sourceFieldId: id, sourceColumn: position, name: id, area: 'row', position })) }
  assert.deepEqual(fields.removePivotChartDraftFields(config, ['b']).fields.map(field => [field.id, field.position]), [['a', 0], ['c', 1]])
  assert.throws(() => fields.removePivotChartDraftFields(config, ['missing']), /不存在/)
  assert.throws(() => fields.movePivotChartDraftField(config, 'b', 'row', 3), /超出/)
  assert.deepEqual(config.fields.map(field => [field.id, field.position]), [['a', 0], ['b', 1], ['c', 2]])
})

test('empty selected labels and clearing the label filter remain different author values', () => {
  const config = { fields: [{ id: 'r', sourceFieldId: '0', sourceColumn: 3, name: 'Region', area: 'row', position: 0 }] }
  const empty = fields.setPivotChartDraftField(config, 'r', { selectedItems: [] }), cleared = fields.setPivotChartDraftField(empty, 'r', { selectedItems: undefined })
  assert.deepEqual(empty.fields[0].selectedItems, []); assert.equal(Object.hasOwn(cleared.fields[0], 'selectedItems'), false)
  assert.deepEqual(fields.setPivotChartDraftField(config, 'r', { selectedItems: ['North', 'North', 'South'] }).fields[0].selectedItems, ['North', 'South'])
  assert.equal(Object.hasOwn(config.fields[0], 'selectedItems'), false)
})

test('real preview calculation uses the detached configured report and does not publish generated worksheet cells', async () => {
  const f = setup({ existing: false })
  try {
    const before = json(model(f)), { draft, source } = freshDraft(f), config = panel.pivotChartPreviewConfig(draft.settings.config, draft.fieldsConfig, source, ids.host)
    assert.equal(config.pivotTableId, undefined); assert.deepEqual(config.ownedPivot.fields, draft.fieldsConfig.fields)
    assert.equal(config.ownedPivot.outputRange, undefined); assert.equal(config.ownedPivot.generatedCellKeys, undefined)
    const data = await f.sheet.getChartSettingsData(config)
    assert.deepEqual(data.categories, ['North', 'South']); assert.deepEqual(data.series[0].values, [40, 60])
    assert.deepEqual(json(model(f)), before); assert.equal(undoCount(f), 0)
  } finally { f.dispose() }
})

test('range creation inserts a real owned chart with initial fields and one content history entry', async () => {
  const f = setup({ existing: false })
  try {
    const cells = json(model(f).sheets.map(sheet => [sheet.id, sheet.cells])), { draft } = freshDraft(f)
    draft.title = 'Regional revenue'; const info = panel.pivotChartCreateInfo(draft, model(f), ids.host, undefined, f.studio.currentUserId)
    assert.equal(info.source.kind, 'range'); assert.equal(info.source.sourceRangeInfo.range.startColumn, 3)
    const chart = await f.sheet.insertPivotChart(info, commandContext(f))
    assert.equal(chart.getInfo().title, 'Regional revenue'); assert.ok(chart.getInfo().ownedPivot); assert.equal(chart.getInfo().pivotTableId, undefined)
    assert.deepEqual(chart.getFieldsConfig().fields.map(field => [field.sourceFieldId, field.area]), [['0', 'row'], ['1', 'value']])
    assert.equal(model(f).pivotTables.length, 0); assert.equal(undoCount(f), 1)
    assert.deepEqual(json(model(f).sheets.map(sheet => [sheet.id, sheet.cells])), cells)
  } finally { f.dispose() }
})

for (const chartId of [ids.owned, ids.linked]) test(`${chartId} configure retains new fields through the visual authoring pass and records one undoable transaction`, async () => {
  const f = setup(); let reopened
  try {
    const before = json(f.book.save()), baseline = undoCount(f), { draft } = draftFor(f, chartId), fieldId = fieldBySource(draft.fieldsConfig, '1').id
    draft.title = 'Average by region'; draft.fieldsConfig = fields.setPivotChartDraftField(draft.fieldsConfig, fieldId, { aggregation: 'average', displayName: 'Mean revenue', format: '0.00' })
    const input = panel.pivotChartConfigureInput(draft)
    for (const key of ['source', 'pivotTableId', 'ownedPivot']) assert.equal(Object.hasOwn(input.patch, key), false)
    assert.equal(await f.sheet.getPivotChart(chartId).configure(input, commandContext(f)), true)
    const target = f.sheet.getPivotChart(chartId), report = pivotChartUIReport(model(f), chartId)
    assert.equal(target.getInfo().title, 'Average by region'); assert.equal(report.fields.find(field => field.id === fieldId).aggregation, 'average')
    assert.equal(report.fields.find(field => field.id === fieldId).displayName, 'Mean revenue'); assert.equal(undoCount(f), baseline + 1)
    const data = await f.sheet.getChartSettingsData(target.getInfo())
    assert.deepEqual(data.categories, ['North', 'South']); assert.deepEqual(data.series[0].values, [20, 30])
    assert.deepEqual(target.getInfo().options.threeD, before.charts.find(chart => chart.id === chartId).options.threeD)
    const after = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), after)
    reopened = new api.OfficeStudio({ workbook: after }); const reopenedBook = api.FOffice.newAPI(reopened).getActiveWorkbook()
    assert.equal(reopenedBook.getPivotChartById(chartId).getFieldsConfig().fields.find(field => field.id === fieldId).aggregation, 'average')
  } finally { reopened?.dispose(); f.dispose() }
})

test('reapplying unchanged field and visual configuration is accepted without revision, dirty or history changes', async () => {
  const f = setup()
  try {
    const { draft } = draftFor(f); await f.sheet.getPivotChart(ids.owned).configure(panel.pivotChartConfigureInput(draft), commandContext(f))
    f.studio.editor.history.markSaved(f.unitId)
    const current = draftFor(f).draft, before = json(f.book.save()), revision = f.studio.getWorkbookRevision(f.unitId), baseline = undoCount(f)
    assert.equal(await f.sheet.getPivotChart(ids.owned).configure(panel.pivotChartConfigureInput(current), commandContext(f)), true)
    assert.deepEqual(json(f.book.save()), before); assert.equal(f.studio.getWorkbookRevision(f.unitId), revision); assert.equal(undoCount(f), baseline)
    assert.equal(f.studio.editor.history.isDirty(f.unitId), false)
  } finally { f.dispose() }
})

for (const [aggregation, expected] of [
  ['sum', [40, 60]], ['count', [2, 2]], ['countNums', [2, 2]], ['average', [20, 30]], ['max', [30, 40]], ['min', [10, 20]],
  ['product', [300, 800]], ['stdDev', [Math.sqrt(200), Math.sqrt(200)]], ['stdDevp', [10, 10]], ['var', [200, 200]], ['varp', [100, 100]],
]) test(`the ${aggregation} UI option produces its actual numeric aggregation through configure`, async () => {
  const f = setup()
  try {
    const { draft } = draftFor(f), valueId = fieldBySource(draft.fieldsConfig, '1').id
    draft.fieldsConfig = fields.setPivotChartDraftField(draft.fieldsConfig, valueId, { aggregation })
    await f.sheet.getPivotChart(ids.owned).configure(panel.pivotChartConfigureInput(draft), commandContext(f))
    const data = await f.sheet.getChartSettingsData(f.sheet.getPivotChart(ids.owned).getInfo())
    assert.deepEqual(data.categories, ['North', 'South'])
    assert.equal(data.series[0].values.length, expected.length)
    data.series[0].values.forEach((value, index) => assert.ok(Math.abs(value - expected[index]) < 1e-9, `${aggregation}: ${value} != ${expected[index]}`))
  } finally { f.dispose() }
})

test('real label edits preserve empty selection, clearing semantics and the source-order member list', async () => {
  const f = setup()
  try {
    const { draft } = draftFor(f), categoryId = fieldBySource(draft.fieldsConfig, '0').id, target = f.sheet.getPivotChart(ids.owned)
    assert.deepEqual(target.getFieldItems(categoryId), ['North', 'South'])
    draft.fieldsConfig = fields.setPivotChartDraftField(draft.fieldsConfig, categoryId, { selectedItems: [] })
    await target.configure(panel.pivotChartConfigureInput(draft), commandContext(f))
    assert.deepEqual(target.getFieldsConfig().fields.find(field => field.id === categoryId).selectedItems, [])
    const cleared = draftFor(f).draft
    cleared.fieldsConfig = fields.setPivotChartDraftField(cleared.fieldsConfig, categoryId, { selectedItems: undefined })
    await target.configure(panel.pivotChartConfigureInput(cleared), commandContext(f))
    assert.equal(Object.hasOwn(target.getFieldsConfig().fields.find(field => field.id === categoryId), 'selectedItems'), false)
    assert.deepEqual(target.getFieldItems(categoryId), ['North', 'South'])
  } finally { f.dispose() }
})

for (const [name, steer] of [
  ['worksheet', f => f.book.setActiveSheet(ids.other)],
  ['user', f => Object.defineProperty(f.studio, 'currentUserId', { value: 'peer', configurable: true })],
  ['history generation', f => f.studio.editor.history.clear(f.unitId)],
  ['content revision', f => f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)],
]) test(`captured field publication refuses a changed ${name} without replacing the newer model`, async () => {
  const f = setup()
  try {
    const captured = owner(f), { draft } = draftFor(f); draft.title = 'Old pending draft'; await steer(f)
    const before = json(model(f)), baseline = undoCount(f)
    assert.equal(owners.pivotChartPanelOwnerCurrent(captured), false)
    assert.throws(() => owners.pivotChartPanelOwnerContext(captured), /更改|改变|重新读取/)
    assert.deepEqual(json(model(f)), before); assert.equal(undoCount(f), baseline)
  } finally { f.dispose() }
})

test('an owner accepts only its own one-revision publication and rejects closed workbooks', async () => {
  const f = setup()
  try {
    const captured = owner(f), { draft } = draftFor(f); draft.title = 'Published'
    await f.sheet.getPivotChart(ids.owned).configure(panel.pivotChartConfigureInput(draft), owners.pivotChartPanelOwnerContext(captured))
    assert.equal(owners.pivotChartPanelOwnerCurrent(captured), false); assert.equal(owners.pivotChartPanelOwnerCurrent(captured, true), true)
    f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(12)
    assert.equal(owners.pivotChartPanelOwnerCurrent(captured, true), false)
    const next = owner(f); f.studio.dispose(); assert.equal(owners.pivotChartPanelOwnerCurrent(next, true), false)
  } finally { f.dispose() }
})

test('reread synchronizes the hook owner while old session callbacks cannot modify a newer panel', async () => {
  const f = setup(), scope = vue.effectScope(); let refreshes = 0, focuses = 0, panels
  try {
    scope.run(() => { panels = hooks.useSheetChartPanels({ studio: () => f.studio, unitId: () => f.unitId, sheetId: () => ids.host,
      sheets: () => model(f).sheets, active: () => true, canRestoreFocus: () => true,
      refresh: () => { refreshes++ }, focus: () => { focuses++ }, notify: message => assert.fail(message) }) })
    assert.equal(panels.openPivotChart({ chartId: ids.owned }), true); const first = panels.pivotChartContext.value
    f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(12)
    first.ownerChanged(owner(f)); first.applied(); assert.equal(refreshes, 1)
    first.setOpen(false); first.closed(); assert.equal(focuses, 1)
    assert.equal(panels.openPivotChart({ chartId: ids.linked }), true); const second = panels.pivotChartContext.value
    first.setBusy(true); first.ownerChanged(owner(f)); first.applied(); first.closed()
    assert.equal(panels.pivotChartContext.value, second); assert.equal(panels.pivotChartPanelOpen.value, true); assert.equal(refreshes, 1); assert.equal(focuses, 1)
  } finally { scope.stop(); f.dispose() }
})

for (const chartId of [ids.owned, ids.linked]) test(`${chartId} refresh/remove use real commands and retain source cells plus independent linked reports`, async () => {
  const f = setup()
  try {
    const target = f.sheet.getPivotChart(chartId), sourceCells = json(model(f).sheets.find(sheet => sheet.id === ids.source).cells)
    assert.equal(await target.refresh(commandContext(f)), true); const before = json(f.book.save()), baseline = undoCount(f)
    assert.equal(await target.remove(commandContext(f)), true); assert.equal(f.sheet.getPivotChart(chartId), null)
    assert.ok(model(f).pivotTables.find(report => report.id === ids.primary)); assert.equal(undoCount(f), baseline + 1)
    assert.deepEqual(json(model(f).sheets.find(sheet => sheet.id === ids.source).cells), sourceCells)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test('the existing Ribbon root is available without an existing PivotTable and keeps its declared grid location', () => {
  const tabs = ribbon.createSheetRibbon({ hasSheet: true, canUndo: false, canRedo: false, canMerge: false, canUnmerge: false,
    canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, pivotChartAvailable: true, dataValidationAvailable: true, canSort: true,
    fontFamily: 'Arial', fontSize: 11, numberFormat: 'General', gridlines: true, frozen: false, zoom: 100, fullscreen: false })
  const matches = tabs.flatMap(tab => tab.groups.flatMap(group => group.items)).filter(item => item.id === 'sheet.operation.open-pivot-chart')
  assert.equal(matches.length, 1); assert.equal(matches[0].disabled, false)
  assert.deepEqual(matches[0].gridLayout, { row: 2, column: 1, showLabel: true })
})
