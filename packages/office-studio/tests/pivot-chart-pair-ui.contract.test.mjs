import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { pivotChartUIIds as ids, pivotChartUIRange as sourceRange, pivotChartUIOptions as reportOptions,
  pivotChartUIWorkbook, pivotChartUIReport } from './fixtures/pivot-chart-ui-fixture.mjs'

// Independent paired-creation/presentation UI contracts. All hooks and cases remain unexecuted.
let server, api, panel, fields, owners
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, panel, fields, owners] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-draft.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-owner.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value)), both = { showFieldButtons: true, showExpandCollapseButtons: true }
function setup(existing = false) {
  const studio = new api.OfficeStudio({ workbook: pivotChartUIWorkbook({ existing }), currentUserId: 'local-user' })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), unitId = book.getId()
  return { studio, book, unitId, host: book.getSheetBySheetId(ids.host), model: () => studio.getWorkbookModel(unitId),
    history: () => studio.editor.history.units[unitId]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
function fresh(f, host = ids.host) {
  const draft = panel.createPivotChartPanelDraft(f.model(), host, { sourceKind: 'range', sourceSheetId: ids.source, sourceRange })
  const source = panel.readPivotChartPanelSource(f.model(), draft, host, 'local-user')
  for (const [id, area] of [['0', 'row'], ['1', 'value']]) draft.fieldsConfig = fields.addPivotChartDraftField(draft.fieldsConfig, source.fields.find(field => field.fieldId === id), area)
  draft.fieldsConfig.options = { ...reportOptions }; return { draft, source }
}
function paired(draft, target = ids.host) { draft.createMode = 'withPivotTable'; draft.pivotTableSheetId = target; draft.pivotTableAddress = 'B9'; return draft }
const commandContext = f => owners.pivotChartPanelOwnerContext(owners.capturePivotChartPanelOwner(f.studio, f.unitId, ids.host))
const info = (f, draft) => panel.pivotChartCreateInfo(draft, f.model(), ids.host, undefined, 'local-user')
const options = (f, draft) => panel.pivotChartCreateOptions(draft, f.model(), ids.host, 'local-user')
const protect = (sheet, range, hidden = false) => { sheet.rangeProtectionRules = [{ id: 'pair-ui-rule', permissionId: 'pair-ui-rule', description: '',
  ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: hidden ? 'noOneElseCanView' : 'othersCanView', ranges: [range] }] }

test('new range authoring exposes the fixed two modes and defaults to owned chart-only metadata', () => {
  const f = setup()
  try {
    const before = json(f.model()), { draft } = fresh(f)
    assert.deepEqual(panel.PIVOT_CHART_CREATION_MODES, [{ value: 'chartOnly', label: '仅创建透视图' }, { value: 'withPivotTable', label: '同时创建透视表和透视图' }])
    assert.equal(draft.createMode, 'chartOnly'); assert.equal(options(f, draft), undefined); assert.equal(info(f, draft).source.kind, 'range')
    assert.deepEqual(f.model(), before)
  } finally { f.dispose() }
})

test.each([[ids.host, 'A1'], [ids.source, 'J3']])('default table anchor in %s is %s and stays outside the source', (host, address) => {
  const f = setup()
  try { const { draft } = fresh(f, host); assert.equal(draft.pivotTableSheetId, host); assert.equal(draft.pivotTableAddress, address) }
  finally { f.dispose() }
})

test('reread input retains destination/mode without putting command destination inside chart appearance', () => {
  const f = setup()
  try {
    const { draft } = fresh(f); paired(draft, ids.other); draft.pivotTableAddress = '$C$12'
    const input = panel.pivotChartPanelInputFromDraft(draft), reloaded = panel.createPivotChartPanelDraft(f.model(), ids.host, input)
    assert.equal(reloaded.createMode, 'withPivotTable'); assert.equal(reloaded.pivotTableSheetId, ids.other); assert.equal(reloaded.pivotTableAddress, '$C$12')
    const request = info(f, draft); assert.equal(request.source.kind, 'range'); assert.equal(request.createMode, undefined); assert.equal(request.pivotTableAddress, undefined)
    assert.deepEqual(options(f, draft), { pivotTable: { anchor: { unitId: f.unitId, subUnitId: ids.other, row: 11, col: 2 } } })
    assert.equal(f.model().pivotTables.length, 0); assert.equal(f.model().charts.length, 0)
  } finally { f.dispose() }
})

test.each([ids.host, ids.other])('UI helper creates linked table/chart in %s with one command, history and reopen', async target => {
  const f = setup(); let reopened
  try {
    const { draft } = fresh(f); paired(draft, target); const before = json(f.book.save()), source = json(f.model().sheets.find(sheet => sheet.id === ids.source).cells)
    const commands = []; f.studio.events.addEvent('CommandExecuted', event => commands.push(event))
    const chart = await f.host.insertPivotChart(info(f, draft), commandContext(f), options(f, draft))
    const report = f.model().pivotTables[0]; assert.equal(chart.getPivotTableReference().pivotTableId, report.id)
    assert.equal(report.anchor.subUnitId, target); assert.equal(chart.getSubUnitId(), ids.host); assert.equal(chart.getInfo().ownedPivot, undefined)
    assert.deepEqual(chart.getData().series[0].values, [40, 60]); assert.equal(f.history(), 1)
    assert.equal(commands.filter(command => command.id === 'sheet.command.add-pivot-chart').length, 1)
    assert.deepEqual(f.model().sheets.find(sheet => sheet.id === ids.source).cells, source)
    const saved = json(f.book.save()); f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
    reopened = new api.OfficeStudio({ workbook: saved }); const restored = api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(chart.getId())
    assert.deepEqual(restored.getData().series[0].values, [40, 60]); assert.equal(restored.getPivotTableReference().pivotTableId, report.id)
  } finally { reopened?.dispose(); f.dispose() }
})

test('chart-only metadata still creates an owned report without any PivotTable output', async () => {
  const f = setup()
  try {
    const { draft } = fresh(f), before = json(f.model().sheets)
    const chart = await f.host.insertPivotChart(info(f, draft), commandContext(f), options(f, draft))
    assert.equal(chart.getSourceKind(), 'range'); assert.ok(chart.getInfo().ownedPivot); assert.equal(f.model().pivotTables.length, 0)
    assert.deepEqual(f.model().sheets, before); assert.equal(f.history(), 1)
  } finally { f.dispose() }
})

test('new authoring from an existing report remains linked and ignores a requested paired mode', async () => {
  const f = setup(true)
  try {
    const draft = panel.createPivotChartPanelDraft(f.model(), ids.host, { pivotTableId: ids.primary, sourceKind: 'pivotTable', createMode: 'withPivotTable' })
    const source = panel.readPivotChartPanelSource(f.model(), draft, ids.host, 'local-user'); draft.fieldsConfig = panel.pivotChartFieldsFromReport(source.report)
    const count = f.model().pivotTables.length; assert.equal(draft.createMode, 'chartOnly'); assert.equal(options(f, draft), undefined)
    const chart = await f.host.insertPivotChart(info(f, draft), commandContext(f), options(f, draft))
    assert.equal(chart.getPivotTableReference().pivotTableId, ids.primary); assert.equal(f.model().pivotTables.length, count)
  } finally { f.dispose() }
})

test.each(['', 'A1:B2', 'A0', 'XFE1', 'A1048577', 'Report!A1', 'D3'])('invalid target %s remains an editable error before any publication', address => {
  const f = setup()
  try {
    const { draft } = fresh(f); paired(draft, address === 'D3' ? ids.source : ids.host); draft.pivotTableAddress = address
    const before = json(f.model()); assert.throws(() => options(f, draft)); assert.ok(panel.pivotChartCreateOptionsError(draft, f.model(), ids.host, 'local-user'))
    assert.deepEqual(f.model(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['worksheet', 'anchor', 'hidden-target'])('paired %s permission is checked before building destination metadata', boundary => {
  const f = setup()
  try {
    const { draft } = fresh(f); paired(draft); const sheet = f.model().sheets.find(sheet => sheet.id === ids.host)
    if (boundary === 'worksheet') sheet.worksheetProtection = { id: 'pair-sheet', permissionId: 'pair-sheet', unitType: 'worksheet', ownerUserId: 'peer',
      allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
    else protect(sheet, { startRow: 8, endRow: 8, startCol: 1, endCol: 1 }, boundary === 'hidden-target')
    const before = json(f.model()); assert.throws(() => options(f, draft), /所选/); assert.deepEqual(f.model(), before)
  } finally { f.dispose() }
})

test.each(['occupied-anchor', 'expanded-output'])('%s is rejected by the real command after advisory anchor validation without partial resources', async boundary => {
  const f = setup()
  try {
    const { draft } = fresh(f); paired(draft); const request = info(f, draft), metadata = options(f, draft)
    if (boundary === 'occupied-anchor') f.host.getRange('B9').setValue('Occupied')
    else protect(f.model().sheets.find(sheet => sheet.id === ids.host), { startRow: 10, endRow: 10, startCol: 2, endCol: 2 })
    const before = json(f.model()), baseline = f.history(); assert.ok(options(f, draft))
    await assert.rejects(f.host.insertPivotChart(request, commandContext(f), metadata))
    assert.deepEqual(f.model(), before); assert.equal(f.history(), baseline); assert.equal(f.model().pivotTables.length, 0); assert.equal(f.model().charts.length, 0)
  } finally { f.dispose() }
})

test.each([ids.host, ids.other])('paired preview uses the prospective %s anchor while computing only a detached report', async target => {
  const f = setup()
  try {
    const { draft, source } = fresh(f); paired(draft, target); const before = json(f.model()), sourceBefore = json(source.report)
    const config = panel.pivotChartPreviewConfig(draft.settings.config, draft.fieldsConfig, source, ids.host, options(f, draft).pivotTable)
    assert.deepEqual(config.ownedPivot.anchor, { unitId: f.unitId, subUnitId: target, row: 8, col: 1 })
    assert.equal(config.ownedPivot.outputRange, undefined); assert.equal(config.ownedPivot.generatedCellKeys, undefined); assert.equal(config.ownedPivot.axisOutput, undefined)
    const data = await f.host.getChartSettingsData(config); assert.deepEqual(data.series[0].values, [40, 60])
    assert.deepEqual(source.report, sourceBefore); assert.deepEqual(f.model(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('the two presentation defaults stay true without creating authored state for an unchanged click', () => {
  const draft = panel.createPivotChartDraft(), before = json(draft.settings)
  assert.deepEqual(draft.controls, both); panel.setPivotChartDraftControls(draft, both)
  assert.deepEqual(draft.settings, before); assert.equal(draft.settings.config.controls, undefined)
})

test.each([ids.owned, ids.linked].flatMap(chartId => [[false, true], [true, false], [false, false]].map(([a, b]) => [chartId, a, b])))
  ('%s presentation [%s,%s] is one configure and never changes the report authority', async (chartId, a, b) => {
    const f = setup(true); let reopened
    try {
      const draft = panel.createPivotChartPanelDraft(f.model(), ids.host, { chartId })
      draft.fieldsConfig = panel.pivotChartFieldsFromReport(pivotChartUIReport(f.model(), chartId))
      const report = json(pivotChartUIReport(f.model(), chartId)), sourceCells = json(f.model().sheets.find(sheet => sheet.id === ids.source).cells), commands = []
      f.studio.events.addEvent('CommandExecuted', event => commands.push(event)); panel.setPivotChartDraftControls(draft, { showFieldButtons: a, showExpandCollapseButtons: b })
      const request = panel.pivotChartConfigureInput(draft); assert.deepEqual(request.patch.controls, { showFieldButtons: a, showExpandCollapseButtons: b })
      await f.book.getPivotChartById(chartId).configure(request, commandContext(f))
      assert.deepEqual(f.book.getPivotChartById(chartId).getControls(), request.patch.controls); assert.deepEqual(pivotChartUIReport(f.model(), chartId), report)
      assert.deepEqual(f.model().sheets.find(sheet => sheet.id === ids.source).cells, sourceCells); assert.equal(f.history(), 1)
      assert.equal(commands.filter(command => command.id === 'sheet.command.configure-pivot-chart').length, 1)
      assert.equal(commands.filter(command => command.id === 'sheet.command.update-pivot-chart-controls').length, 0)
      const saved = json(f.book.save()); f.book.undo(); assert.deepEqual(f.book.getPivotChartById(chartId).getControls(), both)
      f.book.redo(); assert.deepEqual(f.book.getPivotChartById(chartId).getControls(), request.patch.controls)
      reopened = new api.OfficeStudio({ workbook: saved }); assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getPivotChartById(chartId).getControls(), request.patch.controls)
    } finally { reopened?.dispose(); f.dispose() }
  })

test.each([{ showFieldButtons: 'false' }, { showExpandCollapseButtons: undefined }, { extra: true }])('invalid presentation %j is refused without modifying any draft fields', patch => {
  const draft = panel.createPivotChartDraft(), before = json(draft.settings)
  assert.throws(() => panel.setPivotChartDraftControls(draft, patch)); assert.deepEqual(draft.settings, before)
})

test('paired creation stores both presentation switches with initial fields in its one content command', async () => {
  const f = setup()
  try {
    const { draft } = fresh(f); paired(draft); panel.setPivotChartDraftControls(draft, { showFieldButtons: false, showExpandCollapseButtons: false })
    const chart = await f.host.insertPivotChart(info(f, draft), commandContext(f), options(f, draft))
    assert.deepEqual(chart.getControls(), { showFieldButtons: false, showExpandCollapseButtons: false })
    assert.deepEqual(chart.getData().series[0].values, [40, 60]); assert.equal(f.history(), 1)
  } finally { f.dispose() }
})

test('an existing chart cannot be reinterpreted as a paired range creation by destination metadata', () => {
  const f = setup(true)
  try {
    const draft = panel.createPivotChartPanelDraft(f.model(), ids.host, { chartId: ids.owned }); paired(draft)
    const before = json(f.model()); assert.throws(() => options(f, draft), /尚未创建/); assert.deepEqual(f.model(), before)
  } finally { f.dispose() }
})
