import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { pivotChartFieldButtonsIds as ids } from './fixtures/pivot-chart-field-buttons-fixture.mjs'
import { pivotValuesUiWorkbook } from './fixtures/pivot-values-ui-fixture.mjs'

// Deferred UI/Facade/controller contracts. No checks, dev server or browser have run.
let server, api, ui, engine, panel, chartDraft, buttons, owners, hooks
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, ui, engine, panel, chartDraft, buttons, owners, hooks] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-values-position-ui.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-engine.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-draft.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-field-button-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/services/pivot-chart-panel-owner.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/pivot/views/use-pivot-chart-field-buttons.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function setup(ownership = 'owned', options = {}) {
  const studio = new api.OfficeStudio({ workbook: pivotValuesUiWorkbook(options), currentUserId: 'local-user' })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), chartId = ownership === 'owned' ? ids.owned : ids.linked
  const target = { studio, unitId: book.getId(), sheetId: ids.host, chartId }
  const model = () => studio.getWorkbookModel(target.unitId), pivot = () => book.getPivotChartById(chartId)
  const report = () => model().charts.find(chart => chart.id === chartId)?.ownedPivot ?? model().pivotTables.find(report => report.id === ids.primary)
  return { ...target, target, book, model, pivot, report, table: () => book.getPivotTables().find(pivot => pivot.getPivotTableId() === ids.primary),
    tableDraft: () => panel.createPivotPanelDraft(model(), ids.host, 'D3:H7', ids.primary),
    history: () => studio.editor.history.units[target.unitId]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
const source = f => json(f.model().sheets.find(sheet => sheet.id === ids.source).cells)
const context = f => owners.pivotChartPanelOwnerContext(owners.capturePivotChartPanelOwner(f.studio, f.unitId, f.sheetId))
const byArea = (fields, area) => fields.filter(field => field.area === area).sort((a, b) => a.position - b.position)
function assertRendered(f, report, position, index) {
  const output = engine.calculatePivotReport(f.model(), report), effective = index < 0 ? 2 : Math.min(index, 2)
  if (position === 'row') {
    const labels = ['Region', 'Date']; labels.splice(effective, 0, 'Values')
    assert.deepEqual(output.values[output.headerRows - 1].slice(0, 3), labels)
  } else assert.deepEqual([...new Set(output.values[effective].slice(2))].sort(), ['Amount average', 'Amount sum'])
  return output.values
}
function controller(f) {
  const calls = [], scope = effectScope(), options = shallowReactive({ ...f.target, active: true, busy: false,
    openSettings: () => false, restoreFocus: () => calls.push('focus') })
  const controls = scope.run(() => hooks.usePivotChartFieldButtons(options))
  return { calls, options, controls, stop: () => scope.stop() }
}
async function changeOwner(f, kind) {
  if (kind === 'source') f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)
  if (kind === 'history') f.studio.editor.history.clear(f.unitId)
  if (kind === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
  if (kind === 'host') f.book.setActiveSheet(ids.other)
  await nextTick()
}

test.each(['row', 'column', 'none'].flatMap(position => [-1, 0, 1, 2, 99].map(index => [position, index])))
  ('UI %s/%s preserves raw state and uses the shared effective gap without a source mutation', (position, index) => {
    const f = setup('owned', { valuePosition: position, valueIndex: index })
    try {
      const before = json(f.book.save()), state = ui.readPivotValuesUiState(f.report().options, f.report().fields)
      assert.deepEqual({ position: state.position, index: state.index }, f.pivot().getValuesPosition())
      assert.equal(state.effectiveIndex, index === -1 ? 2 : Math.min(index, 2)); assert.equal(state.area, position === 'row' ? 'row' : 'column')
      const entries = ui.pivotValuesUiEntries(byArea(f.report().fields, state.area), state, state.area)
      assert.equal(entries.filter(item => item.kind === 'values').length, position === 'none' ? 0 : 1)
      if (position !== 'none') assert.equal(entries[state.effectiveIndex].kind, 'values')
      assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), 0)
    } finally { f.dispose() }
  })

test.each([0, 1, 2, 3])('%s measures show a special UI item only for multiple authored value instances', count => {
  const f = setup()
  try {
    const fields = f.report().fields.filter(field => field.area !== 'value')
    for (let index = 0; index < count; index++) fields.push({ ...f.report().fields.find(field => field.area === 'value'), id: `local-value-${index}` })
    assert.equal(ui.readPivotValuesUiState(f.report().options, fields).visible, count > 1); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('an authored field named Values remains distinct from the UI marker and is never replaced by a synthetic source field', () => {
  const f = setup()
  try {
    const fields = byArea(f.report().fields, 'column').map((field, index) => ({ ...field, id: index ? field.id : 'Values', displayName: 'ΣValues' }))
    const state = ui.readPivotValuesUiState(f.report().options, f.report().fields), entries = ui.pivotValuesUiEntries(fields, state, 'column')
    assert.equal(entries[0].field.id, 'Values'); assert.equal(entries[2].kind, 'values')
    assert.equal(f.pivot().getSourceFieldsInfo().length, 5); assert.equal(f.report().fields.length, 6); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each([
  { valuePosition: 'row', options: { valuePosition: 'column' } }, { valueIndex: 0, options: { valueIndex: 1 } },
  { valuePosition: 'row', valueIndex: 1, options: { valuePosition: 'column', valueIndex: 1 } },
  { valuePosition: 'column', valueIndex: 1, options: { valuePosition: 'column', valueIndex: -1 } },
])('conflicting public Values aliases %j reject before a model or history edit', config => {
  const f = setup()
  try { assert.throws(() => ui.pivotChartValuesUiOptions({ fields: f.report().fields, ...config }), /冲突/); assert.equal(f.history(), 0) }
  finally { f.dispose() }
})

test.each([-2, 1.5, NaN, '1', Infinity])('invalid raw UI index %s cannot enter either panel draft', index => {
  const f = setup(), draft = f.tableDraft()
  try {
    const before = json(draft); assert.throws(() => ui.setPivotTableValuesUiDraft(draft, 'row', index))
    assert.throws(() => ui.setPivotChartValuesUiDraft(f.pivot().getFieldsConfig(), 'row', index)); assert.deepEqual(draft, before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['row', 'column'].flatMap(position => [0, 1, 2, -1].map(index => [position, index])))
  ('ordinary %s/%s draft commits a real axis insertion in one configure and preserves history/reopen', (position, index) => {
    const f = setup('linked'), draft = f.tableDraft(); let restored
    try {
      const before = json(f.book.save()), cells = source(f); draft.targetAddress = 'K30'; ui.setPivotTableValuesUiDraft(draft, position, index)
      assert.deepEqual(json(f.book.save()), before)
      api.FOffice.newAPI(f.studio).executeCommandSync(panel.CONFIGURE_PIVOT_COMMAND,
        { ...panel.preparePivotPanelInput(f.model(), draft, ids.primary), ...context(f), hostSheetId: ids.host })
      assert.equal(f.history(), 1); assert.deepEqual(source(f), cells); assert.deepEqual(f.table().getValuesPosition(), { position, index })
      assertRendered(f, f.report(), position, index); const saved = json(f.book.save())
      f.book.undo(); assert.deepEqual(json(f.book.save()), before); f.book.redo(); assert.deepEqual(json(f.book.save()), saved)
      restored = new api.OfficeStudio({ workbook: saved }); assert.deepEqual(api.FOffice.newAPI(restored).getActiveWorkbook().getPivotTables().find(pivot => pivot.getPivotTableId() === ids.primary).getValuesPosition(), { position, index })
    } finally { restored?.dispose(); f.dispose() }
  })

test.each(['owned', 'linked'].flatMap(ownership => ['row', 'column'].flatMap(position => [0, 1, 2, -1].map(index => [ownership, position, index]))))
  ('%s chart %s/%s synchronizes public aliases and actual report layout through one configure', async (ownership, position, index) => {
    const f = setup(ownership)
    try {
      const before = json(f.book.save()), cells = source(f), config = ui.setPivotChartValuesUiDraft(f.pivot().getFieldsConfig(), position, index)
      assert.deepEqual(ui.pivotChartValuesUiOptions(config), { valuePosition: position, valueIndex: index }); assert.deepEqual(json(f.book.save()), before)
      await f.pivot().configure({ fieldsConfig: config }, context(f)); assert.deepEqual(f.pivot().getValuesPosition(), { position, index })
      assertRendered(f, f.report(), position, index); assert.deepEqual(source(f), cells)
      const changed = position !== 'column' || index !== -1; assert.equal(f.history(), changed ? 1 : 0)
      if (changed) { f.book.undo(); assert.deepEqual(json(f.book.save()), before) }
    } finally { f.dispose() }
  })

test.each(['table', 'owned', 'linked'])('%s None retains measures and column fallback while repeated None remains a no-op', async kind => {
  const f = setup(kind === 'owned' ? 'owned' : 'linked')
  try {
    const pivot = kind === 'table' ? f.table() : f.pivot(), before = json(f.book.save())
    await pivot.updateValuePosition('column', 1, context(f)); const columns = assertRendered(f, f.report(), 'column', 1)
    const history = f.history(); await pivot.updateValuePosition('none', 1, context(f))
    assert.equal(f.report().fields.filter(field => field.area === 'value').length, 2); assert.deepEqual(assertRendered(f, f.report(), 'none', 1), columns)
    assert.equal(f.history(), history + 1); const saved = json(f.book.save()); await pivot.updateValuePosition('none', 1, context(f))
    assert.equal(f.history(), history + 1); assert.deepEqual(json(f.book.save()), saved); f.book.undo(); f.book.undo(); assert.deepEqual(json(f.book.save()), before)
  } finally { f.dispose() }
})

test.each(['row', 'column'])('%s oversized index is visible as a retained choice and survives draft/config/JSON reopen', async position => {
  const f = setup('owned', { valuePosition: position, valueIndex: 99 }); let restored
  try {
    const state = ui.readPivotValuesUiState(f.report().options, f.report().fields)
    assert.ok(ui.pivotValuesUiIndexOptions(state).some(option => option.value === 99 && option.label.includes('当前末位')))
    const saved = json(f.book.save()), before = assertRendered(f, f.report(), position, 99)
    restored = new api.OfficeStudio({ workbook: saved }); const report = restored.getWorkbookModel(ids.unit).charts.find(chart => chart.id === ids.owned).ownedPivot
    assert.equal(report.options.valueIndex, 99); assert.deepEqual(engine.calculatePivotReport(restored.getWorkbookModel(ids.unit), report).values, before); assert.equal(f.history(), 0)
  } finally { restored?.dispose(); f.dispose() }
})

test.each(['row', 'column'])('ordinary %s dimension reordering preserves the shared Values gap and main draft', position => {
  const f = setup('linked'), draft = f.tableDraft()
  try {
    draft.targetAddress = 'K30'; ui.setPivotTableValuesUiDraft(draft, position, 1); const previous = ui.pivotTableValuesUiFieldOrder(draft)
    const first = draft.fields.find(field => field.area === position), sibling = draft.fields.find(field => field.area === position && field !== first)
    const left = draft.fields.indexOf(first), right = draft.fields.indexOf(sibling); draft.fields[left] = sibling; draft.fields[right] = first
    ui.adjustPivotTableValuesUiFieldEdit(draft, previous, first.fieldId); assert.equal(draft.options.valueIndex, 0); assert.equal(draft.targetAddress, 'K30'); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['row', 'column'])('chart %s dimension reordering updates both aliases through the shared gap rule', position => {
  const f = setup()
  try {
    const config = ui.setPivotChartValuesUiDraft(f.pivot().getFieldsConfig(), position, 1), first = byArea(config.fields, position)[0]
    const next = chartDraft.movePivotChartDraftField(config, first.id, position, 1)
    assert.equal(next.valueIndex, 0); assert.equal(next.options.valueIndex, 0); assert.equal(config.valueIndex, 1); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['row', 'column'])('%s multiple dimension removal retains only the surviving gap without creating a Values field', position => {
  const f = setup()
  try {
    const config = ui.setPivotChartValuesUiDraft(f.pivot().getFieldsConfig(), position, 1), fields = byArea(config.fields, position)
    const after = chartDraft.removePivotChartDraftFields(config, [fields[0].id, fields[1].id])
    assert.equal(after.valueIndex, 0); assert.equal(after.fields.filter(field => field.area === 'value').length, 2); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s actual field-button snapshot exposes Values separately from local field IDs', ownership => {
  const f = setup(ownership)
  try {
    const snapshot = buttons.readPivotChartFieldButtonsSnapshot(f.target); assert.equal(snapshot.values.visible, true); assert.equal(snapshot.values.measureCount, 2)
    assert.equal(snapshot.sections.flatMap(section => section.fields).length, f.report().fields.length)
    assert.equal(snapshot.sections.flatMap(section => section.fields).some(field => field.fieldId === 'values'), false); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['owned', 'linked'])('%s Values controller commits one real command and restores focus after a no-op or direction edit', async ownership => {
  const f = setup(ownership), mounted = controller(f), commands = []
  try {
    f.studio.events.addEvent('CommandExecuted', event => commands.push(event)); mounted.controls.openValues()
    assert.equal(await mounted.controls.mutateValues('row', 0), true); await nextTick()
    assert.equal(f.history(), 1); assert.deepEqual(f.pivot().getValuesPosition(), { position: 'row', index: 0 }); assert.deepEqual(mounted.calls, ['focus'])
    assert.equal(commands.filter(command => command.id === 'sheet.command.edit-pivot-chart-report').length, 1)
    assert.equal(commands.at(-1).params.edit.type, 'updateValuePosition'); assert.equal(commands.at(-1).params.expectedUserId, 'local-user')
    mounted.controls.openValues(); assert.equal(await mounted.controls.mutateValues('row', 0), true); assert.equal(f.history(), 1)
  } finally { mounted.stop(); f.dispose() }
})

test.each(['owned', 'linked'].flatMap(ownership => ['source', 'history', 'user', 'host'].map(change => [ownership, change])))
  ('pending %s Values rejects a changed %s owner without partial output or history', async (ownership, change) => {
    const f = setup(ownership), mounted = controller(f), original = api.FSheetPivotChart.prototype.updateValuePosition; let release
    api.FSheetPivotChart.prototype.updateValuePosition = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
    try {
      mounted.controls.openValues(); const task = mounted.controls.mutateValues('row', 0); await changeOwner(f, change)
      const before = json(f.book.save()), history = f.history(); release(); assert.equal(await task, false)
      assert.deepEqual(json(f.book.save()), before); assert.equal(f.history(), history)
    } finally { api.FSheetPivotChart.prototype.updateValuePosition = original; release?.(); mounted.stop(); f.dispose() }
  })

test('old Values finally cannot clear a newer busy request after a history generation changes', async () => {
  const f = setup(), mounted = controller(f), queue = [], original = api.FSheetPivotChart.prototype.updateValuePosition
  api.FSheetPivotChart.prototype.updateValuePosition = async function(...args) { await new Promise(resolve => queue.push(resolve)); return original.apply(this, args) }
  try {
    mounted.controls.openValues(); const first = mounted.controls.mutateValues('row', 0); f.studio.editor.history.clear(f.unitId); await nextTick()
    mounted.controls.openValues(); const second = mounted.controls.mutateValues('column', 0); queue[0](); assert.equal(await first, false)
    assert.equal(mounted.controls.busy.value, true); queue[1](); assert.equal(await second, true); assert.equal(f.history(), 1)
  } finally { api.FSheetPivotChart.prototype.updateValuePosition = original; queue.forEach(release => release()); mounted.stop(); f.dispose() }
})

test.each([{ readonly: true }, { protectedOutput: true }, { protectedChart: true }])('viewable protection %j exposes the position but cannot edit Values', async options => {
  const f = setup(options.protectedOutput ? 'linked' : 'owned', options), mounted = controller(f)
  try { mounted.controls.openValues(); assert.equal(mounted.controls.snapshot.value.editable, false); assert.equal(await mounted.controls.mutateValues('row', 0), false); assert.equal(f.history(), 0) }
  finally { mounted.stop(); f.dispose() }
})

test.each([{ hiddenSource: true }, { hiddenOutput: true }, { hiddenSource: true, hiddenOutput: true }])('hidden source/output %j cannot expose a Values snapshot', options => {
  const f = setup('linked', options)
  try { assert.throws(() => buttons.readPivotChartFieldButtonsSnapshot(f.target)); assert.equal(f.history(), 0) }
  finally { f.dispose() }
})

test('Values session respects composition Escape and drops a source-stale owner when composition ends', async () => {
  const f = setup(), mounted = controller(f)
  try {
    mounted.controls.openValues(); mounted.controls.compositionStart()
    mounted.controls.keydown({ key: 'Escape', keyCode: 229, isComposing: true, stopPropagation() {}, preventDefault() { throw new Error('Composition cannot close') } })
    assert.equal(await mounted.controls.mutateValues('row', 0), false); f.book.getSheetBySheetId(ids.source).getRange('E4').setValue(11)
    assert.equal(mounted.controls.stale.value, true); mounted.controls.compositionEnd(); await nextTick(); assert.equal(mounted.controls.openAnchor.value, '')
  } finally { mounted.stop(); f.dispose() }
})

test('a superseded Values close cannot steal focus and disposal clears a pending Values session', async () => {
  const f = setup(), mounted = controller(f), original = api.FSheetPivotChart.prototype.updateValuePosition; let release
  api.FSheetPivotChart.prototype.updateValuePosition = async function(...args) { await new Promise(resolve => { release = resolve }); return original.apply(this, args) }
  try {
    mounted.controls.openValues(); mounted.controls.close(); mounted.options.active = false; await nextTick(); assert.deepEqual(mounted.calls, [])
    mounted.options.active = true; await nextTick(); mounted.controls.openValues(); const task = mounted.controls.mutateValues('row', 0)
    mounted.stop(); release(); await task; assert.equal(mounted.controls.snapshot.value, undefined); assert.equal(mounted.controls.busy.value, false)
  } finally { api.FSheetPivotChart.prototype.updateValuePosition = original; release?.(); f.dispose() }
})
