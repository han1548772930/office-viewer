import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'
import { effectScope, nextTick, shallowReactive } from 'vue'
import { timelineSettingsUiWorkbook, timelineUiIds as ids } from './fixtures/timeline-settings-ui-fixture.mjs'
import { timelineFirst, timelineReportTotals } from './fixtures/timeline-filter-fixture.mjs'

// Deferred real Facade/model/command/history/controller contracts. No checks have run.
let server, api, settings, dates, panel, hooks
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, settings, dates, panel, hooks] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/timeline-settings-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/timeline-range-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-panel-state.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/slicer/use-timeline-date-filter.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const optional = ['caption', 'style', 'scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar']
const operators = ['equal', 'notEqual', 'newerThan', 'newerThanOrEqual', 'olderThan', 'olderThanOrEqual', 'between', 'notBetween']
function setup(options = {}) {
  const studio = new api.OfficeStudio({ workbook: options.model ?? timelineSettingsUiWorkbook(options), currentUserId: 'local-user' })
  const book = api.FOffice.newAPI(studio).getActiveWorkbook(), target = { workbookId: book.getId(), sheetId: ids.host }, commands = []
  studio.events.addEvent('CommandExecuted', event => commands.push(event))
  const model = () => studio.getWorkbookModel(book.getId()), control = (id = ids.primary) => book.getTimeline(id)
  const state = () => json(book.save()), history = () => studio.editor.history.units[book.getId()]?.undo.length ?? 0
  const snapshot = (id = ids.primary) => settings.readPanelTimelineSettings(studio, target, id)
  const filter = (id = ids.primary) => dates.readPanelTimelineDateFilter(studio, target, id)
  const draft = (id = ids.primary) => settings.createTimelinePanelEditDraft(panel.readSlicerPanelState(studio, target).controls.find(value => value.info.id === id).info, snapshot(id).settings)
  return { studio, book, target, commands, model, control, state, history, snapshot, filter, draft, first: () => timelineFirst(model()), dispose: () => studio.dispose() }
}
function retainedDisplay(f) {
  const model = f.model()
  return json({ sheets: model.sheets.map(sheet => ({ id: sheet.id, cells: sheet.cells })), pivots: model.pivotTables, charts: model.charts,
    caches: model.slicerCaches, sources: model.slicers.map(({ id, sourceFieldId, pivotTableId, cacheId }) => ({ id, sourceFieldId, pivotTableId, cacheId })) })
}
const rangeDraft = (startDate, endDate, useWholeDay) => ({ mode: 'range', range: { startDate, endDate,
  ...(useWholeDay !== undefined ? { useWholeDay } : {}) }, condition: { operator: 'equal' } })
const conditionDraft = (operator, first, second, useWholeDay) => ({ mode: 'condition', range: {}, condition: { operator, first,
  ...(second !== undefined ? { second } : {}), ...(useWholeDay !== undefined ? { useWholeDay } : {}) } })
const clearDraft = () => ({ mode: 'none', range: {}, condition: { operator: 'equal' } })
const result = promise => promise.then(value => value, () => false)
function gate(f, kind) {
  const method = { settings: 'updateSettings', presentation: 'updatePresentation', range: 'setDateRange', bounds: 'getDateBounds' }[kind]
  const pending = [], original = method ? api.FSheetTimeline.prototype[method] : f.studio.executeCommand
  if (method) api.FSheetTimeline.prototype[method] = async function(...args) {
    await new Promise(resolve => pending.push(resolve)); return original.apply(this, args)
  }
  else f.studio.executeCommand = async function(id, params) {
    if (id === (kind === 'condition' ? 'sheet.command.set-timeline-date-condition' : 'sheet.command.set-timeline-date-range')) await new Promise(resolve => pending.push(resolve))
    return original.call(this, id, params)
  }
  return { pending, execute: (id, params) => original.call(f.studio, id, params),
    release(index = 0) { pending.splice(index, 1)[0]?.() },
    restore() { if (method) api.FSheetTimeline.prototype[method] = original; else f.studio.executeCommand = original; pending.splice(0).forEach(resolve => resolve()) } }
}
function controller(f, id = ids.primary) {
  const calls = [], scope = effectScope(), options = shallowReactive({ studio: f.studio, ...f.target, timelineId: id, open: true,
    changed: id => calls.push(['changed', id]), close: owner => { calls.push(['close', owner]); options.open = false } })
  const controls = scope.run(() => hooks.useTimelineDateFilter(options))
  return { calls, options, controls, stop: () => scope.stop() }
}
async function settle() { await nextTick(); await new Promise(resolve => setImmediate(resolve)); await nextTick() }
async function changeOwner(f, kind) {
  if (kind === 'source') f.book.getSheetBySheetId(ids.source).getRange('F4').setValue(11)
  if (kind === 'history') f.studio.editor.history.clear(f.target.workbookId)
  if (kind === 'user') Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
  if (kind === 'host') f.book.setActiveSheet(ids.other)
  await settle()
}
function editFor(f, kind) {
  if (kind === 'settings') return settings.applyPanelTimelineSettings(f.snapshot(), { caption: 'Pending Timeline caption' })
  return dates.applyPanelTimelineDateFilter(f.filter(), kind === 'range' ? rangeDraft(f.first() + .5, f.first() + .75)
    : kind === 'condition' ? conditionDraft('equal', f.first() + .5) : clearDraft())
}

test.each([ids.primary, ids.shared])('%s draft reads real numeric Timeline settings without inventing optional author values', id => {
  const f = setup({ viewDefaults: true })
  try {
    const before = f.state(), snapshot = f.snapshot(id), draft = f.draft(id)
    assert.equal(snapshot.settings.level, 2); assert.equal(snapshot.settings.selectionLevel, 2)
    assert.equal(draft.level, 'month'); assert.equal(draft.selectionLevel, 'month'); assert.equal(draft.captionFromName, true)
    assert.equal(draft.caption, snapshot.settings.name); assert.equal(draft.scrollPosition, undefined)
    assert.deepEqual(settings.buildPanelTimelineSettingsPatch(draft, f.draft(id)), {})
    assert.deepEqual(settings.buildPanelTimelineGeometryPatch(draft, f.draft(id)), {})
    snapshot.settings.name = 'Detached read'; assert.notEqual(f.control(id).getSettings().name, snapshot.settings.name)
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(optional)('clearing Timeline optional %s preserves own undefined in its minimal patch', key => {
  const f = setup({ authoredSettings: true })
  try {
    const original = f.draft(), draft = f.draft()
    if (key === 'caption') draft.captionFromName = true; else delete draft[key]
    const patch = settings.buildPanelTimelineSettingsPatch(draft, original)
    assert.deepEqual(Object.keys(patch), [key]); assert.equal(Object.hasOwn(patch, key), true); assert.equal(patch[key], undefined)
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('all ten display properties commit one formal settings command while shared filtering, source and geometry remain intact', async () => {
  const f = setup({ range: { startDate: 46023.25, endDate: 46024 } }), before = f.state(), stable = retainedDisplay(f), shared = f.control(ids.shared).getModel()
  try {
    const patch = { name: 'Timeline display', caption: 'Calendar title', level: 2, selectionLevel: 1, style: 'TimelineStyleDark2',
      scrollPosition: f.first() + .123456789, showHeader: false, showSelectionLabel: false, showTimeLevel: false, showHorizontalScrollbar: false }
    assert.equal(await settings.applyPanelTimelineSettings(f.snapshot(), patch), true); assert.deepEqual(f.control().getSettings(), patch)
    assert.equal(f.history(), 1); assert.equal(f.commands.filter(event => event.id === 'sheet.command.update-timeline-settings').length, 1)
    assert.deepEqual(f.control(ids.shared).getModel(), shared); assert.deepEqual(retainedDisplay(f), stable)
    assert.deepEqual(f.control().getModel().drawingGeometry, before.slicers.find(view => view.id === ids.primary).drawingGeometry)
    f.book.undo(); assert.deepEqual(f.state(), before); f.book.redo(); assert.deepEqual(f.control().getSettings(), patch)
  } finally { f.dispose() }
})

test('mixed display and affine geometry use one presentation transaction with numeric-to-local grain adaptation', async () => {
  const f = setup(), before = f.state(), stable = retainedDisplay(f), original = f.draft(), draft = f.draft()
  try {
    draft.captionFromName = false; draft.caption = 'Mixed calendar'; draft.level = 'month'; draft.selectionLevel = 'quarter'
    draft.layout.width += 40; draft.layout.height += 20; draft.drawing.anchorType = 'none'; draft.drawing.keepRatio = true; draft.lockedPosition = true
    const patch = settings.buildPanelTimelineSettingsPatch(draft, original), geometry = settings.buildPanelTimelineGeometryPatch(draft, original)
    assert.deepEqual(Object.keys(geometry).sort(), ['anchorType', 'keepRatio', 'layout', 'lockedPosition'])
    assert.equal(await settings.applyPanelTimelineSettings(f.snapshot(), patch, geometry), true); assert.equal(f.history(), 1)
    const event = f.commands.find(command => command.id === 'sheet.command.update-slicer-presentation')
    assert.equal(event.params.patch.level, 'month'); assert.equal(event.params.patch.selectionLevel, 'quarter')
    assert.equal(f.commands.filter(command => command.id === 'sheet.command.update-timeline-settings').length, 0)
    assert.equal(f.control().getSettings().level, 2); assert.equal(f.control().getModel().lockedPosition, true)
    assert.deepEqual(retainedDisplay(f), stable); f.book.undo(); assert.deepEqual(f.state(), before)
  } finally { f.dispose() }
})

test('inherited Timeline caption follows its renamed view and leaves canonical source/cache names unchanged', async () => {
  const f = setup({ captionInherited: true }), draft = f.draft(), original = f.draft(), stable = retainedDisplay(f)
  try {
    draft.name = 'Renamed calendar'; const patch = settings.buildPanelTimelineSettingsPatch(draft, original)
    assert.deepEqual(patch, { name: 'Renamed calendar' }); assert.equal(await settings.applyPanelTimelineSettings(f.snapshot(), patch), true)
    assert.equal(f.control().getSettings().caption, undefined); assert.equal(f.control().getSettings().name, draft.name)
    assert.deepEqual(retainedDisplay(f), stable)
  } finally { f.dispose() }
})

test('Timeline defaults clear all seven optional keys, retain name, preserve no-op history and survive JSON reopen', async () => {
  const f = setup({ authoredSettings: true }), original = f.draft(), draft = f.draft(); let reopened
  try {
    draft.captionFromName = true; draft.level = draft.selectionLevel = 'month'; for (const key of optional.filter(key => key !== 'caption')) delete draft[key]
    const patch = settings.buildPanelTimelineSettingsPatch(draft, original)
    for (const key of optional) { assert.equal(Object.hasOwn(patch, key), true); assert.equal(patch[key], undefined) }
    assert.equal(await settings.applyPanelTimelineSettings(f.snapshot(), patch), true); assert.equal(f.history(), 1)
    const saved = f.control().getSettings(); assert.deepEqual(saved, { name: original.name, level: 2, selectionLevel: 2 })
    assert.equal(await settings.applyPanelTimelineSettings(f.snapshot(), {}), true); assert.equal(f.history(), 1)
    reopened = new api.OfficeStudio({ workbook: json(f.book.save()), currentUserId: 'local-user' })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getTimeline(ids.primary).getSettings(), saved)
  } finally { reopened?.dispose(); f.dispose() }
})

test.each(['duplicate', 'empty-caption', 'wrong-style', 'invalid-level'])('%s display validation refuses an atomic draft before history or source changes', async kind => {
  const f = setup(), before = f.state(), original = f.draft(), draft = f.draft()
  try {
    if (kind === 'duplicate') draft.name = ` ${f.control(ids.shared).getSettings().name.toUpperCase()} `
    if (kind === 'empty-caption') { draft.captionFromName = false; draft.caption = '' }
    if (kind === 'wrong-style') draft.style = 'SlicerStyleLight1'
    if (kind === 'invalid-level') draft.level = 'week'
    assert.ok(settings.timelineSettingsPanelDraftError(f.snapshot(), draft, original))
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each([false, true])('date1904=%s bounds are inclusive original typed Date bounds independently of selection', async date1904 => {
  const f = setup({ date1904 }), first = f.first()
  try {
    await f.control().setDateRange({ startDate: first + .5, endDate: first + .75 }); const before = f.state(), history = f.history()
    const bounds = await dates.readPanelTimelineDateBounds(f.filter())
    assert.deepEqual(bounds, { minDate: first + .25, maxDate: first + 31.5, dateSystem: date1904 ? 'date1904' : 'date1900' })
    const draft = dates.fillTimelinePanelSourceDays(rangeDraft(first + .5, first + .75), bounds)
    assert.deepEqual(draft.range, { startDate: first, endDate: first + 32 }); assert.equal(bounds.maxDate, first + 31.5)
    assert.deepEqual(f.state(), before); assert.equal(f.history(), history)
  } finally { f.dispose() }
})

test.each(['cached-date', 'raw-number', 'raw-text'])('%s original field capability uses stored Date identity rather than its heading', kind => {
  const workbook = timelineSettingsUiWorkbook({ noDates: true }), source = workbook.sheets.find(sheet => sheet.id === ids.source), first = timelineFirst(workbook)
  for (let row = 3; row <= 13; row += 1) source.cells[`${row}:4`].value = kind === 'raw-text' ? '2026-01-01T12:00:00' : first + .5
  if (kind === 'cached-date') { source.cells['12:4'].formula = '=DATE(2026,1,1)+0.5'; source.cells['12:4'].cachedDataType = 'date' }
  const f = setup({ model: workbook })
  try {
    const fields = panel.readSlicerPanelState(f.studio, f.target).pivots.filter(source => source.kind === 'pivot').flatMap(source => source.fields)
    assert.equal(fields.filter(field => field.id === 'date-source').every(field => field.controls.includes('timeline')), kind === 'cached-date')
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['slicer', 'timeline'].flatMap(kind => ['pivot', 'ownedPivotChart'].map(source => [kind, source])))
  ('insert panel %s/%s supplies formal Info, source references, view values and CSS placement in one command', async (kind, source) => {
    const f = setup(), before = f.state(), reportId = f.model().pivotTables[0].id
    try {
      const id = await panel.insertPanelSlicer(f.studio, f.target, { type: kind, sourceFieldId: 'date-source',
        ...(source === 'pivot' ? { pivotTableId: reportId } : { ownedPivotChartId: 'owned-chart' }),
        name: `Inserted ${kind} ${source}`, caption: 'Inserted calendar source', anchor: { row: 100, column: 10, rowOffset: 2, columnOffset: 3 },
        size: { width: 350, height: 160 }, ...(kind === 'timeline' ? { level: 'year', selectionLevel: 'day' } : { columnCount: 2, rowHeight: 28 }) })
      const control = kind === 'timeline' ? f.book.getTimeline(id) : f.book.getSlicer(id), commandId = kind === 'timeline' ? 'sheet.command.insert-timeline' : 'sheet.command.insert-slicer'
      assert.ok(control); assert.equal(f.history(), 1); const event = f.commands.find(command => command.id === commandId)
      assert.equal(event.params.info.dataSource.kind, 'pivot'); assert.equal(event.params.info.dataSource.sourceFieldId, 'date-source')
      assert.equal(Object.hasOwn(event.params.info, 'type'), false); assert.equal(Object.hasOwn(event.params.info, 'pivotTableId'), false)
      assert.equal(control.getSettings().name, `Inserted ${kind} ${source}`)
      if (kind === 'timeline') { assert.equal(control.getSettings().level, 0); assert.equal(control.getSettings().selectionLevel, 3) }
      else { assert.equal(control.getSettings().columnCount, 2); assert.equal(control.getSettings().rowHeight, 28) }
      assert.equal(control.getDrawing().transform.width, 350); assert.equal(control.getDrawing().transform.height, 160)
      if (source === 'ownedPivotChart') assert.deepEqual(event.params.info.dataSource.ownedPivotCharts, [{ subUnitId: ids.host, pivotChartId: 'owned-chart' }])
      else assert.deepEqual(event.params.info.dataSource.pivotTables, [{ subUnitId: ids.other, pivotTableId: reportId }])
      f.book.undo(); assert.deepEqual(f.state(), before)
    } finally { f.dispose() }
  })

test.each([false, true])('date1904=%s exact half-open range includes cached Date and excludes colliding raw values and the end endpoint', async date1904 => {
  const f = setup({ date1904 }), first = f.first(), before = f.state(), sharedSettings = f.control(ids.shared).getSettings()
  try {
    const draft = rangeDraft(first + .5, first + .75, false)
    assert.equal(await dates.applyPanelTimelineDateFilter(f.filter(), draft), true); assert.equal(f.history(), 1)
    assert.deepEqual(f.control().getDateRange(), draft.range); assert.deepEqual(f.control(ids.shared).getDateRange(), draft.range)
    assert.deepEqual(f.control(ids.shared).getSettings(), sharedSettings); assert.deepEqual(timelineReportTotals(f.model()), [80, 80, 80])
    const event = f.commands.find(command => command.id === 'sheet.command.set-timeline-date-range')
    assert.deepEqual(Object.keys(event.params.range).sort(), ['endDate', 'startDate', 'useWholeDay'])
    assert.equal(Object.hasOwn(event.params.range, 'level'), false); assert.equal(f.control().getDateCondition(), undefined)
    f.book.undo(); assert.deepEqual(f.state(), before); f.book.redo(); assert.deepEqual(timelineReportTotals(f.model()), [80, 80, 80])
  } finally { f.dispose() }
})

test.each([false, true])('date1904=%s calendar input converts the same day while untouched fractional operands remain exact', date1904 => {
  const f = setup({ date1904 }), first = f.first(), author = first + .123456789
  try {
    const draft = rangeDraft(author, first + .75, false), state = dates.buildTimelinePanelDateFilter(draft)
    assert.equal(state.range.startDate, author)
    assert.equal(dates.timelinePanelDateFromCalendarInput('2026-01-01T12:00:00', date1904), first + .5)
    assert.equal(dates.timelinePanelDateCalendarInput(first + .5, date1904), '2026-01-01 12:00:00.000')
    assert.equal(dates.timelinePanelDateCompatibilityDay(60.25, date1904), !date1904)
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('Excel 1900 compatibility day remains a valid calendar operand and retains its exact author range', () => {
  const f = setup({ compatibilityDay: true })
  try {
    const snapshot = f.filter(), draft = dates.createTimelineDateFilterDraft(snapshot)
    assert.equal(dates.timelinePanelDateCalendarInput(60.25, false), '1900-02-29 06:00:00.000')
    assert.equal(dates.timelinePanelDateFromCalendarInput('1900-02-29T06:00:00.000', false), 60.25)
    assert.deepEqual(dates.buildTimelinePanelDateFilter(draft), snapshot.state); assert.equal(draft.range.endDate, 60.75)
    assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test('whole-day mode retains author fractions and uses floor/ceil endpoints while keeping the exclusive end contract', async () => {
  const f = setup(), first = f.first()
  try {
    const draft = rangeDraft(first + .5, first + .75, true)
    assert.equal(await dates.applyPanelTimelineDateFilter(f.filter(), draft), true)
    assert.deepEqual(f.control().getDateRange(), draft.range); assert.deepEqual(timelineReportTotals(f.model()), [110, 110, 110])
    assert.equal(dates.timelinePanelDateFilterError(rangeDraft(first, first)), 'Timeline range startDate must precede endDate')
  } finally { f.dispose() }
})

test.each(operators)('condition %s replaces the author range using the shared Pivot date intent, not a range/member approximation', async operator => {
  const f = setup(), first = f.first(), totals = { equal: 80, notEqual: 210, newerThan: 200, newerThanOrEqual: 280,
    olderThan: 10, olderThanOrEqual: 90, between: 110, notBetween: 180 }
  try {
    await f.control().setDateRange({ startDate: first + .25, endDate: first + .75 })
    const before = f.state(), baseline = f.history(), between = operator === 'between' || operator === 'notBetween'
    const draft = conditionDraft(operator, between ? first + .25 : first + .5, between ? first + .75 : undefined, false)
    assert.equal(await dates.applyPanelTimelineDateFilter(f.filter(), draft), true); assert.equal(f.history(), baseline + 1)
    const condition = { operator, expected: between ? [first + .25, first + .75] : first + .5, useWholeDay: false }
    assert.deepEqual(f.control().getDateCondition(), condition); assert.deepEqual(f.control(ids.shared).getDateCondition(), condition)
    assert.equal(f.control().getDateRange(), undefined); assert.equal(typeof f.control().setDateCondition, 'undefined')
    assert.deepEqual(timelineReportTotals(f.model()), [totals[operator], totals[operator], totals[operator]])
    const event = f.commands.find(command => command.id === 'sheet.command.set-timeline-date-condition')
    assert.deepEqual(event.params.condition, condition); assert.equal(Object.hasOwn(event.params, 'range'), false)
    f.book.undo(); assert.deepEqual(f.state(), before)
  } finally { f.dispose() }
})

test.each(['range', 'condition'])('clearing a retained %s uses the shared range command with undefined and no empty-member substitute', async kind => {
  const f = setup(), first = f.first()
  try {
    await dates.applyPanelTimelineDateFilter(f.filter(), kind === 'range' ? rangeDraft(first + .5, first + .75) : conditionDraft('equal', first + .5))
    const before = f.state(), baseline = f.history()
    assert.equal(await dates.applyPanelTimelineDateFilter(f.filter(), clearDraft()), true); assert.equal(f.history(), baseline + 1)
    assert.equal(f.control().getDateRange(), undefined); assert.equal(f.control().getDateCondition(), undefined)
    assert.equal(f.control(ids.shared).getDateRange(), undefined); assert.equal(f.control(ids.shared).getDateCondition(), undefined)
    const event = f.commands.filter(command => command.id === 'sheet.command.set-timeline-date-range').at(-1)
    assert.equal(Object.hasOwn(event.params, 'range'), true); assert.equal(event.params.range, undefined)
    assert.equal(f.commands.some(command => command.id === 'sheet.command.set-slicer-items'), false)
    f.book.undo(); assert.deepEqual(f.state(), before)
  } finally { f.dispose() }
})

test.each(['missing-range', 'reversed-range', 'equal-range', 'missing-operand', 'missing-pair', 'reversed-pair'])('%s invalid filter refuses before history or partial report changes', async kind => {
  const f = setup(), first = f.first(), before = f.state()
  try {
    const draft = kind === 'missing-range' ? rangeDraft(undefined, first + 1) : kind === 'reversed-range' ? rangeDraft(first + 1, first)
      : kind === 'equal-range' ? rangeDraft(first, first) : kind === 'missing-operand' ? conditionDraft('equal', undefined)
        : kind === 'missing-pair' ? conditionDraft('between', first) : conditionDraft('notBetween', first + 1, first)
    assert.ok(dates.timelinePanelDateFilterError(draft)); await assert.rejects(dates.applyPanelTimelineDateFilter(f.filter(), draft))
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['noDates', 'hiddenSource'])('%s source query failure preserves retained range/condition and independent display settings', async kind => {
  const first = timelineFirst(timelineSettingsUiWorkbook()), f = setup({ [kind]: true, range: { startDate: first + .5, endDate: first + .75 } })
  try {
    const snapshot = f.filter(), before = f.control().getDateRange()
    assert.deepEqual(snapshot.state, { range: before }); assert.ok(f.snapshot().settings)
    if (kind === 'noDates') {
      const bounds = await dates.readPanelTimelineDateBounds(snapshot); assert.equal(bounds.minDate, undefined); assert.equal(bounds.maxDate, undefined)
    } else await assert.rejects(dates.readPanelTimelineDateBounds(snapshot))
    assert.deepEqual(f.control().getDateRange(), before); assert.equal(f.control().getDateCondition(), undefined); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['settings', 'range', 'condition'])('viewable read-only %s permits retained reads and refuses edits without content/history', async kind => {
  const f = setup({ readonly: true }), before = f.state()
  try { assert.ok(f.snapshot().settings); assert.deepEqual(f.filter().state, {}); assert.equal(await editFor(f, kind), false)
    assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})

test.each(['settings', 'range', 'condition', 'clear'].flatMap(kind => ['source', 'history', 'user', 'host', 'closed', 'disposed'].map(change => [kind, change])))
  ('pending %s revokes changed %s at the real dispatch without modifying the replacement owner', async (kind, change) => {
    const first = timelineFirst(timelineSettingsUiWorkbook()), f = setup({ range: { startDate: first + .25, endDate: first + 1 } }), delayed = gate(f, kind)
    let active = true
    try {
      const pending = result(kind === 'settings' ? settings.applyPanelTimelineSettings(f.snapshot(), { caption: 'Late owner' }, {}, () => active)
        : dates.applyPanelTimelineDateFilter(f.filter(), kind === 'range' ? rangeDraft(first + .5, first + .75)
          : kind === 'condition' ? conditionDraft('equal', first + .5) : clearDraft(), () => active))
      assert.equal(delayed.pending.length, 1)
      if (change === 'closed') active = false; else if (change === 'disposed') f.dispose(); else await changeOwner(f, change)
      const before = change === 'disposed' ? undefined : f.state(), baseline = change === 'disposed' ? undefined : f.history()
      delayed.release(); assert.equal(await pending, false)
      if (before) { assert.deepEqual(f.state(), before); assert.equal(f.history(), baseline) }
    } finally { delayed.restore(); if (change !== 'disposed') f.dispose() }
  })

test.each(['settings', 'range', 'condition', 'clear'].flatMap(kind => ['history', 'user'].map(change => [kind, change])))
  ('%s owner guard rechecks %s changed by a later real BeforeCommandExecute listener', async (kind, change) => {
    const first = timelineFirst(timelineSettingsUiWorkbook()), f = setup({ range: { startDate: first + .25, endDate: first + 1 } }), delayed = gate(f, kind)
    let listener
    try {
      const pending = result(editFor(f, kind)), id = { settings: 'sheet.command.update-timeline-settings', range: 'sheet.command.set-timeline-date-range',
        condition: 'sheet.command.set-timeline-date-condition', clear: 'sheet.command.set-timeline-date-range' }[kind]
      listener = f.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id !== id) return
        if (change === 'history') f.studio.editor.history.clear(f.target.workbookId)
        else Object.defineProperty(f.studio, 'currentUserId', { configurable: true, value: 'peer' })
      })
      const before = f.state(); delayed.release(); assert.equal(await pending, false); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
    } finally { listener?.dispose(); delayed.restore(); f.dispose() }
  })

test.each(['range', 'condition', 'clear'])('date controller %s accepts exactly its own publish and closes after the real shared command', async kind => {
  const first = timelineFirst(timelineSettingsUiWorkbook()), f = setup(kind === 'clear' ? { range: { startDate: first + .25, endDate: first + 1 } } : {}), c = controller(f)
  try {
    await settle(); assert.equal(c.controls.boundsLoading.value, false)
    if (kind === 'clear') c.controls.setMode('none')
    else if (kind === 'range') { c.controls.setMode('range'); c.controls.setRange({ startDate: first + .5, endDate: first + .75 }) }
    else { c.controls.setMode('condition'); c.controls.setCondition({ operator: 'equal', first: first + .5 }) }
    assert.equal(c.controls.canApply.value, true); assert.equal(await c.controls.apply(), true); assert.equal(f.history(), 1)
    assert.equal(c.calls.filter(call => call[0] === 'changed').length, 1); assert.equal(c.calls.filter(call => call[0] === 'close').length, 1)
    assert.equal(c.controls.busy.value, false)
  } finally { c.stop(); f.dispose() }
})

test.each(['reread', 'history'])('old bounds and finally preserve a fresh pending source read after %s', async kind => {
  const f = setup(), delayed = gate(f, 'bounds'), c = controller(f)
  try {
    assert.equal(delayed.pending.length, 1)
    if (kind === 'history') await changeOwner(f, 'history')
    c.controls.read(); assert.equal(delayed.pending.length, 2); const readSession = c.controls.readSession.value
    delayed.release(); await settle(); assert.equal(c.controls.bounds.value, undefined); assert.equal(c.controls.boundsLoading.value, true)
    assert.equal(c.controls.readSession.value, readSession)
    delayed.release(); await settle(); assert.equal(c.controls.boundsLoading.value, false); assert.equal(c.controls.bounds.value.minDate, f.first() + .25)
    assert.equal(c.controls.boundsError.value, ''); assert.equal(f.history(), 0)
  } finally { delayed.restore(); c.stop(); f.dispose() }
})

test.each(['history', 'external-range'])('old clear params identity and finally preserve a newer pending clear after %s', async kind => {
  const first = timelineFirst(timelineSettingsUiWorkbook()), f = setup({ range: { startDate: first + .25, endDate: first + 1 } }), c = controller(f), delayed = gate(f, 'clear')
  try {
    await settle(); c.controls.setMode('none'); const old = c.controls.apply(); assert.equal(delayed.pending.length, 1)
    if (kind === 'history') await changeOwner(f, 'history')
    else { await delayed.execute('sheet.command.set-timeline-date-range', { unitId: f.target.workbookId, slicerId: ids.primary,
      range: { startDate: first + .5, endDate: first + .75 } }); await settle() }
    assert.equal(c.controls.stale.value, true); assert.equal(c.controls.busy.value, false); const baseline = f.history()
    c.controls.read(); await settle(); c.controls.setMode('none'); const fresh = c.controls.apply(); assert.equal(delayed.pending.length, 2)
    delayed.release(); assert.equal(await old, false); assert.equal(c.controls.busy.value, true); assert.equal(c.controls.stale.value, false); assert.equal(f.history(), baseline)
    delayed.release(); assert.equal(await fresh, true); assert.equal(f.history(), baseline + 1); assert.equal(f.control().getDateRange(), undefined)
  } finally { delayed.restore(); c.stop(); f.dispose() }
})

test.each(['range', 'condition'])('stale %s composition retains local input, blocks Enter/apply/reread and waits for explicit refresh', async kind => {
  const f = setup(), c = controller(f), first = f.first()
  try {
    await settle(); c.controls.setMode(kind)
    if (kind === 'range') c.controls.setRange({ startDate: first + .5, endDate: first + .75 })
    else c.controls.setCondition({ operator: 'equal', first: first + .5 })
    c.controls.compositionStart(); await changeOwner(f, 'source'); const session = c.controls.readSession.value
    if (kind === 'range') c.controls.setRange({ startDate: first + .6 }); else c.controls.setCondition({ first: first + .6 })
    c.controls.read(); assert.equal(c.controls.readSession.value, session); assert.equal(await c.controls.apply(), false)
    c.controls.compositionEnd(); assert.equal(c.controls.stale.value, true)
    assert.equal(kind === 'range' ? c.controls.draft.value.range.startDate : c.controls.draft.value.condition.first, first + .6)
    assert.equal(c.calls.filter(call => call[0] === 'changed').length, 0)
    c.controls.read(); await settle(); assert.equal(c.controls.stale.value, false); assert.equal(c.controls.draft.value.mode, 'none')
  } finally { c.stop(); f.dispose() }
})

test.each(['closed', 'disposed'])('date controller %s revokes a delayed range and emits no changed callback', async kind => {
  const f = setup(), c = controller(f), delayed = gate(f, 'range'), before = f.state()
  try {
    await settle(); c.controls.setMode('range'); c.controls.setRange({ startDate: f.first() + .5, endDate: f.first() + .75 })
    const pending = c.controls.apply(); if (kind === 'closed') c.controls.close(); else c.stop()
    delayed.release(); assert.equal(await pending, false); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
    assert.equal(c.calls.filter(call => call[0] === 'changed').length, 0)
  } finally { delayed.restore(); c.stop(); f.dispose() }
})

test('ordinary Slicer cannot be read as formal Timeline settings or a Timeline date filter', () => {
  const f = setup(), before = f.state()
  try {
    assert.equal(f.book.getTimeline(ids.slicer), null); assert.ok(f.book.getSlicer(ids.slicer)); assert.equal(f.book.getSlicer(ids.primary), null)
    assert.throws(() => f.snapshot(ids.slicer)); assert.throws(() => f.filter(ids.slicer)); assert.deepEqual(f.state(), before); assert.equal(f.history(), 0)
  } finally { f.dispose() }
})
