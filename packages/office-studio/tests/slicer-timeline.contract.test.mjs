import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let api
let panel
let renderer
let layoutApi
let tableApi
let viewAdapter
let drawingApi

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  panel = await server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-panel-state.ts')
  renderer = await server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-renderer.ts')
  layoutApi = await server.ssrLoadModule('/packages/office-studio/src/rendering/spreadsheet/layout.ts')
  tableApi = await server.ssrLoadModule('/packages/office-studio/src/features/sheet-tables/sheet-tables.ts')
  viewAdapter = await server.ssrLoadModule('/packages/office-studio/src/services/excelViewAdapter.ts')
  drawingApi = await server.ssrLoadModule('/packages/office-studio/src/features/slicer/slicer-drawing.ts')
})

afterAll(async () => { await server?.close() })

async function setup(rows, { date1904 = false, labels = [0] } = {}) {
  const model = api.createBlankWorkbook()
  model.date1904 = date1904
  const studio = new api.OfficeStudio({ workbook: model })
  const officeApi = api.FOffice.newAPI(studio)
  const workbook = officeApi.getActiveWorkbook()
  const source = workbook.getActiveSheet()
  const output = workbook.insertSheet('Summary')
  const lastColumn = String.fromCharCode(64 + rows[0].length)
  source.getRange(`A1:${lastColumn}${rows.length}`).setValues(rows)
  const pivot = await workbook.addPivotTable({
    unitId: workbook.getId(), subUnitId: source.getSheetId(), sheetName: source.getSheetName(),
    range: { startRow: 0, startColumn: 0, endRow: rows.length - 1, endColumn: rows[0].length - 1 },
  }, api.PositionTypeEnum.Existing, { unitId: workbook.getId(), subUnitId: output.getSheetId(), row: 0, col: 0 })
  for (const field of labels) await pivot.addField(field, 'row')
  await pivot.addField(rows[0].length - 1, 'value')
  return { studio, officeApi, workbook, source, output, pivot, target: { workbookId: workbook.getId(), sheetId: output.getSheetId() } }
}

function boundField(studio, workbook, fieldId = '0') {
  return studio.getWorkbookModel(workbook.getId()).pivotTables[0].fields.find((field) => field.id === fieldId && field.area !== 'value')
}

test('panel date conversion shares Excel 1900 and 1904 semantics including serial 60', () => {
  for (const [serial, calendar] of [[-1, '1899-12-30'], [0, '1899-12-31'], [1, '1900-01-01'], [59, '1900-02-28'], [60, '1900-02-29'], [61, '1900-03-01'], [40729, '2011-07-05']]) {
    assert.equal(panel.timelineInputToSerial(calendar, false), serial)
    assert.equal(panel.timelineValueToInput(serial, false), calendar)
  }
  assert.equal(panel.timelineInputToSerial('1904-01-01', true), 0)
  assert.equal(panel.timelineValueToInput(0, true), '1904-01-01')
  assert.equal(panel.timelineInputToSerial('1903-12-31', true), -1)
  assert.equal(panel.timelineValueToInput(-1, true), '1903-12-31')
  assert.equal(panel.timelineValueToInput(60, true), '1904-03-01')
  assert.equal(panel.timelineValueToInput(0.75, false), '1899-12-31')
  assert.equal(panel.timelineValueToInput(-0.25, false), '1899-12-30')
  assert.equal(panel.timelineInputToSerial('1904-02-29', true), 59)
  assert.equal(panel.timelineInputToSerial('2011-07-05', true), 39267)
  assert.equal(panel.timelineInputToSerial('0001-01-01', true), -695055)
  assert.equal(panel.timelineValueToInput(-695055, true), '0001-01-01')
  assert.equal(panel.timelineInputToSerial('  ', true), undefined)
  for (const text of ['2026-02-30', '2100-02-29', '1900-02-29', '2026-13-01', '2026-01-00', '0000-01-01', '2026-1-1', 'not a date']) {
    assert.throws(() => panel.timelineInputToSerial(text, true), (error) => error.code === 'INVALID_TIMELINE_DATE')
  }
  for (const value of [null, undefined, '', false, true, Number.NaN, Number.POSITIVE_INFINITY, 1e100, '2026-02-30', '2026-2-30', '2026-02-30T12:00:00Z']) {
    assert.equal(panel.timelineValueToInput(value, false), '')
  }
  assert.equal(panel.timelineValueToInput('2026-02-28T23:30:00-02:00', false), '2026-03-01')
  const oldTimezone = process.env.TZ
  try {
    for (const timezone of ['Asia/Shanghai', 'America/Los_Angeles']) {
      process.env.TZ = timezone
      assert.equal(panel.timelineValueToInput('2026-02-28T23:30:00', false), '2026-02-28')
    }
  } finally {
    if (oldTimezone === undefined) delete process.env.TZ
    else process.env.TZ = oldTimezone
  }
})

test('timeline keeps February 28, Excel February 29 and March 1 as separate whole days', async () => {
  const { studio, workbook, output, pivot, target } = await setup([
    ['Date', 'Amount'], [1, 1], [59, 2], [59.75, 3], [60, 4], [60.75, 5], [61, 6], ['2026-02-30', 7],
  ])
  try {
    const id = await panel.insertPanelSlicer(studio, target, { pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'timeline' })
    const timeline = output.getSlicer(id)
    assert.deepEqual(timeline.getItems(), ['1', '59', '59.75', '60', '60.75', '61'])
    await panel.applyPanelTimelineRange(studio, target, id, { start: '1900-02-28', end: '1900-02-28', level: 'day' })
    assert.deepEqual(timeline.getInfo().selectedItems, ['59', '59.75'])
    await panel.applyPanelTimelineRange(studio, target, id, { start: '1900-02-29', end: '1900-02-29', level: 'day' })
    assert.deepEqual(timeline.getInfo().selectedItems, ['60', '60.75'])
    assert.deepEqual(boundField(studio, workbook).selectedItems, ['60', '60.75'])
    await timeline.setTimelineRange({ start: 60, end: 60, level: 'month' })
    assert.deepEqual(timeline.getInfo().selectedItems, ['59', '59.75', '60', '60.75'])
    assert.equal(panel.readSlicerPanelState(studio, target).controls[0].info.timelineRange.start, 60)
  } finally { studio.dispose() }
})

test('timeline day/month/quarter/year bounds are inclusive with open and empty ranges', async () => {
  const { studio, output, pivot, target } = await setup([
    ['Date', 'Amount'], ['2025-12-31', 1], ['2026-01-01', 2], ['2026-02-28', 3],
    ['2026-03-31T23:59:59Z', 4], ['2026-04-01', 5], ['2026-12-31', 6], ['2027-01-01', 7],
  ])
  try {
    const timeline = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'timeline' })
    for (const [input, expected] of [
      [{ start: '2026-02-01', end: '2026-02-28', level: 'month' }, ['2026-02-28']],
      [{ start: '2026-02-12', end: '2026-03-02', level: 'quarter' }, ['2026-01-01', '2026-02-28', '2026-03-31T23:59:59Z']],
      [{ start: '2026-06-15', end: '2026-06-15', level: 'year' }, ['2026-01-01', '2026-02-28', '2026-03-31T23:59:59Z', '2026-04-01', '2026-12-31']],
      [{ start: '', end: '2026-01-01', level: 'day' }, ['2025-12-31', '2026-01-01']],
      [{ start: '2027-01-01', end: '', level: 'day' }, ['2027-01-01']],
      [{ start: '2030-01-01', end: '2030-01-01', level: 'day' }, []],
    ]) {
      await panel.applyPanelTimelineRange(studio, target, timeline.getId(), input)
      assert.deepEqual(timeline.getInfo().selectedItems, expected)
    }
    await panel.clearPanelSlicer(studio, target, timeline.getId())
    assert.equal(timeline.getInfo().selectedItems, undefined)
    assert.equal(timeline.getInfo().timelineRange, undefined)
  } finally { studio.dispose() }
})

test('1904 numeric dates and panel input select the same leap-day rows', async () => {
  const { studio, output, pivot, target } = await setup([
    ['Date', 'Amount'], [0, 1], [58, 2], [59, 3], [59.9, 4], [60, 5],
  ], { date1904: true })
  try {
    const timeline = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'timeline' })
    await panel.applyPanelTimelineRange(studio, target, timeline.getId(), { start: '1904-02-29', end: '1904-02-29', level: 'day' })
    assert.deepEqual(timeline.getInfo().selectedItems, ['59', '59.9'])
    await timeline.setTimelineRange({ start: 0, end: 0, level: 'month' })
    assert.deepEqual(timeline.getInfo().selectedItems, ['0'])
    assert.equal(panel.readSlicerPanelState(studio, target).date1904, true)
  } finally { studio.dispose() }
})

test('multiple slicers intersect, distinguish no selection from clear and survive history/save/reopen', async () => {
  const { studio, workbook, output, pivot, target } = await setup([
    ['Region', 'Period', 'Amount'], ['East', 'Q1', 100], ['East', 'Q2', 150], ['West', 'Q1', 80], ['North', 'Q2', 120],
  ], { labels: [0, 1] })
  let restored
  try {
    const regionId = await panel.insertPanelSlicer(studio, target, { pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer', caption: '区域' })
    const period = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '1', type: 'slicer' })
    const region = output.getSlicer(regionId)
    await panel.applyPanelSlicerItems(studio, target, regionId, ['West', 'East', 'West'])
    assert.deepEqual(region.getInfo().selectedItems, ['West', 'East'])
    await period.setItems(['Q1'])
    assert.equal(pivot.getData().at(-1).at(-1), 180)
    await panel.applyPanelSlicerItems(studio, target, regionId, [])
    assert.deepEqual(boundField(studio, workbook).selectedItems, [])
    assert.deepEqual(panel.readSlicerPanelState(studio, target).controls[0].selectedItems, [])
    assert.equal(pivot.getData().flat().some((value) => typeof value === 'number'), false)
    workbook.undo()
    assert.deepEqual(region.getInfo().selectedItems, ['West', 'East'])
    workbook.redo()
    assert.deepEqual(region.getInfo().selectedItems, [])
    await panel.clearPanelSlicer(studio, target, regionId)
    assert.equal(region.getInfo().selectedItems, undefined)
    assert.deepEqual(period.getInfo().selectedItems, ['Q1'])
    assert.equal(pivot.getData().at(-1).at(-1), 180)
    await panel.applyPanelSlicerItems(studio, target, regionId, ['East'])
    const saved = JSON.parse(JSON.stringify(workbook.save()))
    restored = new api.OfficeStudio({ workbook: saved })
    const reopened = api.FOffice.newAPI(restored).getActiveWorkbook()
    assert.deepEqual(reopened.getSlicer(regionId).getInfo().selectedItems, ['East'])
    assert.deepEqual(reopened.getSlicer(regionId).getItems(), ['East', 'West', 'North'])
    assert.equal(reopened.getPivotTables()[0].getData().at(-1).at(-1), 100)
    await panel.removePanelSlicer(studio, target, regionId)
    assert.equal(workbook.getSlicer(regionId), null)
    assert.deepEqual(boundField(studio, workbook).selectedItems, ['East'])
    workbook.undo()
    assert.deepEqual(workbook.getSlicer(regionId).getInfo().selectedItems, ['East'])
    assert.deepEqual(boundField(studio, workbook).selectedItems, ['East'])
  } finally { restored?.dispose(); studio.dispose() }
})

test('table slicer filters mixed scalar and blank values without replacing worksheet filters', async () => {
  const model = api.createBlankWorkbook()
  const studio = new api.OfficeStudio({ workbook: model })
  try {
    const workbook = api.FOffice.newAPI(studio).getActiveWorkbook()
    const sheet = workbook.getActiveSheet()
    const other = workbook.insertSheet('Other')
    sheet.getRange('A1:B5').setValues([
      ['Category', 'Amount'], ['Alpha', 1], ['', 2], [4, 4], [false, 5],
    ])
    sheet.getRange('D1:E3').setValues([
      ['Status', 'Value'], ['keep', 1], ['hide', 2],
    ])
    sheet.setAutoFilter({ startRow: 0, startCol: 3, endRow: 2, endCol: 4 })
    sheet.setFilterCriteria(3, { kind: 'values', selectedItems: ['keep'] })
    other.getRange('A1:B2').setValues([['Flag', 'Value'], ['visible', 1]])
    other.setAutoFilter({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 })
    other.setFilterCriteria(0, { kind: 'values', selectedItems: [] })

    const current = studio.getWorkbookModel(workbook.getId())
    const withTable = tableApi.createSheetTable(current, {
      sheetId: sheet.getSheetId(), range: { startRow: 0, startCol: 0, endRow: 4, endCol: 1 }, name: 'Categories', id: 'table-categories',
    })
    studio.mutateWorkbook(workbook.getId(), 'sheet.command.test-create-table', 'Create table fixture', (draft) => {
      draft.tables = withTable.tables
    })
    const slicer = await sheet.insertSlicer({ tableId: 'table-categories', sourceFieldId: withTable.tables[0].columns[0].id, type: 'slicer' })

    assert.deepEqual(slicer.getItems(), ['Alpha', '', '4', 'false'])
    await panel.applyPanelSlicerItems(studio, { workbookId: workbook.getId(), sheetId: sheet.getSheetId() }, slicer.getId(), ['Alpha', '', '4'])
    const filteredModel = studio.getWorkbookModel(workbook.getId())
    const tableFilter = filteredModel.tables.find(({ id }) => id === 'table-categories').autoFilter
    assert.deepEqual(tableFilter.filteredRows, [4])
    assert.deepEqual(filteredModel.sheets.find(({ id }) => id === sheet.getSheetId()).autoFilter.filteredRows, [2])
    assert.deepEqual(filteredModel.sheets.find(({ id }) => id === other.getSheetId()).autoFilter.filteredRows, [1])
    assert.deepEqual(viewAdapter.toExcelWorkbookView(filteredModel).sheets.find(({ id }) => id === sheet.getSheetId()).filteredRows, [2, 4])
    assert.deepEqual(panel.readSlicerPanelState(studio, { workbookId: workbook.getId(), sheetId: sheet.getSheetId() }).controls[0].selectedItems, ['Alpha', '', '4'])
    await panel.removePanelSlicer(studio, { workbookId: workbook.getId(), sheetId: sheet.getSheetId() }, slicer.getId())
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).tables.find(({ id }) => id === 'table-categories').autoFilter.filteredRows, [4])
  } finally { studio.dispose() }
})

test('invalid or canceled slicer commands do not mutate model or add history', async () => {
  const { studio, officeApi, workbook, output, pivot, target } = await setup([
    ['Region', 'Amount'], ['East', 100], ['West', 80],
  ])
  try {
    const slicer = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer' })
    const baseline = JSON.stringify(studio.getWorkbookModel(workbook.getId()))
    await assert.rejects(() => slicer.setItems(['Missing']), /unknown item/)
    await assert.rejects(() => panel.applyPanelSlicerItems(studio, target, slicer.getId(), ['Missing']), (error) => error.code === 'UNKNOWN_SLICER_ITEM')
    await assert.rejects(() => slicer.setTimelineRange({ start: 1, level: 'day' }), /require selected items/)
    await assert.rejects(() => panel.insertPanelSlicer(studio, target, { pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'timeline' }), (error) => error.code === 'SOURCE_FIELD_UNAVAILABLE')
    const subscription = officeApi.addEvent('BeforeSlicerChange', (event) => { event.cancel = true })
    await assert.rejects(() => slicer.setItems(['East']), /canceled/)
    subscription.dispose()
    assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId())), baseline)
    // The most recent successful command is still creation, not a failed edit.
    workbook.undo()
    assert.equal(workbook.getSlicer(slicer.getId()), null)
  } finally { studio.dispose() }
})

test('invalid timeline boundaries reject before writing pivot filters or history', async () => {
  const { studio, workbook, output, pivot, target } = await setup([
    ['Date', 'Amount'], ['2026-02-28', 10], ['2026-03-01', 20],
  ])
  try {
    const timeline = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'timeline' })
    const baseline = JSON.stringify(studio.getWorkbookModel(workbook.getId()))
    for (const range of [
      { start: '2026-02-30', level: 'day' }, { end: '2026-02-30T12:00:00Z', level: 'month' },
      { start: 1e100, level: 'day' }, { start: true, level: 'day' },
      { start: '2026-03-01', end: '2026-02-28', level: 'day' }, { level: 'week' },
    ]) await assert.rejects(() => timeline.setTimelineRange(range), /valid dates|start must not exceed|Invalid timeline level/)
    await assert.rejects(() => panel.applyPanelTimelineRange(studio, target, timeline.getId(), { start: '2026-02-30', end: '', level: 'day' }), (error) => error.code === 'INVALID_TIMELINE_DATE')
    await assert.rejects(() => panel.applyPanelTimelineRange(studio, target, timeline.getId(), { start: '2026-03-01', end: '2026-02-28', level: 'day' }), (error) => error.code === 'INVALID_TIMELINE_ORDER')
    assert.equal(JSON.stringify(studio.getWorkbookModel(workbook.getId())), baseline)
    workbook.undo()
    assert.equal(workbook.getSlicer(timeline.getId()), null)
  } finally { studio.dispose() }
})

test('panel selection follows authoritative pivot state and rejects controls from another sheet', async () => {
  const { studio, workbook, source, output, pivot, target } = await setup([
    ['Region', 'Amount'], ['East', 100], ['West', 80],
  ])
  try {
    const slicer = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer' })
    await pivot.setLabelManualFilter('0', ['West'])
    assert.deepEqual(panel.readSlicerPanelState(studio, target).controls[0].selectedItems, ['West'])
    const foreignTarget = { workbookId: workbook.getId(), sheetId: source.getSheetId() }
    await assert.rejects(() => panel.clearPanelSlicer(studio, foreignTarget, slicer.getId()), (error) => error.code === 'SLICER_NOT_FOUND')
    assert.deepEqual(boundField(studio, workbook).selectedItems, ['West'])
    assert.equal(panel.readSlicerPanelState(studio, foreignTarget).controls.length, 0)
    assert.throws(() => panel.readSlicerPanelState(studio, { ...target, sheetId: 'gone' }), (error) => error.code === 'WORKSHEET_NOT_FOUND')
  } finally { studio.dispose() }
})

test('slicer insertion exposes only supported controls and presentation settings use undoable commands', async () => {
  const { studio, workbook, output, pivot, target } = await setup([
    ['Region', 'Date', 'Amount'], ['East', '2026-01-02', 100], ['West', '2026-02-03', 80],
  ], { labels: [0, 1] })
  try {
    const before = panel.readSlicerPanelState(studio, target)
    assert.deepEqual(before.pivots[0].fields.map(({ controls }) => controls), [['slicer'], ['slicer', 'timeline'], ['slicer']])
    await assert.rejects(() => panel.insertPanelSlicer(studio, target, {
      pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'timeline',
    }), (error) => error.code === 'UNSUPPORTED_SOURCE_FIELD')

    const slicerId = await panel.insertPanelSlicer(studio, target, {
      pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer',
    })
    await panel.applyPanelSlicerPresentation(studio, target, slicerId, {
      caption: 'Region filter', columnCount: 2, rowHeight: 30, showCaption: false,
      lockedPosition: true, style: 'SlicerStyleLight2', layout: { left: 240, top: 260, width: 280, height: 320 },
    })
    assert.deepEqual(
      (({ caption, columnCount, rowHeight, showCaption, lockedPosition, style, layout }) => ({ caption, columnCount, rowHeight, showCaption, lockedPosition, style, layout }))(output.getSlicer(slicerId).getInfo()),
      { caption: 'Region filter', columnCount: 2, rowHeight: 30, showCaption: false, lockedPosition: true, style: 'SlicerStyleLight2', layout: { left: 240, top: 260, width: 280, height: 320 } },
    )
    workbook.undo()
    assert.equal(output.getSlicer(slicerId).getInfo().columnCount, 1)
    assert.equal(output.getSlicer(slicerId).getInfo().caption, 'Region')
    assert.deepEqual(output.getSlicer(slicerId).getInfo().layout, { left: 200, top: 200, width: 220, height: 280 })
    workbook.redo()
    assert.equal(output.getSlicer(slicerId).getInfo().style, 'SlicerStyleLight2')

    const timelineId = await panel.insertPanelSlicer(studio, target, {
      pivotTableId: pivot.getPivotTableId(), sourceFieldId: '1', type: 'timeline',
    })
    await panel.applyPanelSlicerPresentation(studio, target, timelineId, {
      level: 'year', selectionLevel: 'quarter', showHeader: false,
      showSelectionLabel: false, showTimeLevel: true, showHorizontalScrollbar: false,
      style: 'TimelineStyleDark1',
    })
    assert.deepEqual(
      (({ level, selectionLevel, showHeader, showSelectionLabel, showTimeLevel, showHorizontalScrollbar, style }) => ({ level, selectionLevel, showHeader, showSelectionLabel, showTimeLevel, showHorizontalScrollbar, style }))(output.getSlicer(timelineId).getInfo()),
      { level: 'year', selectionLevel: 'quarter', showHeader: false, showSelectionLabel: false, showTimeLevel: true, showHorizontalScrollbar: false, style: 'TimelineStyleDark1' },
    )
  } finally { studio.dispose() }
})

test('slicers sharing a source field mirror filter state and workbook drawing anchors', async () => {
  const { studio, workbook, output, pivot, target } = await setup([
    ['Region', 'Amount'], ['East', 100], ['West', 80], ['North', 60],
  ])
  try {
    const first = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer' })
    const second = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer' })
    await first.setItems(['East'])
    assert.deepEqual(second.getInfo().selectedItems, ['East'])
    assert.deepEqual(panel.readSlicerPanelState(studio, target).controls.map(({ selectedItems }) => selectedItems), [['East'], ['East']])

    await panel.applyPanelSlicerPresentation(studio, target, first.getId(), { anchorType: 'both', keepRatio: true })
    assert.equal(first.getInfo().drawing.anchorType, 'both')
    assert.equal(first.getInfo().drawing.keepRatio, true)
    const sheetView = () => viewAdapter.toExcelWorkbookView(studio.getWorkbookModel(workbook.getId())).sheets.find(({ id }) => id === output.getSheetId())
    const bounds = () => drawingApi.resolveSlicerBounds(first.getInfo(), new layoutApi.SpreadsheetGeometry(sheetView()))
    const original = bounds()
    studio.mutateWorkbook(workbook.getId(), 'sheet.command.test-resize-row', 'Resize anchored row', (draft) => {
      draft.sheets.find(({ id }) => id === output.getSheetId()).rowHeights[0] = 52
    })
    const bothAnchored = bounds()
    assert.equal(bothAnchored.top, original.top + 26)
    assert.equal(bothAnchored.width, original.width)
    assert.equal(bothAnchored.height, original.height)

    await panel.applyPanelSlicerPresentation(studio, target, first.getId(), { anchorType: 'position', keepRatio: false })
    assert.equal(first.getInfo().drawing.anchorType, 'position')
    assert.equal(first.getInfo().drawing.keepRatio, false)
    const positionAnchored = bounds()
    studio.mutateWorkbook(workbook.getId(), 'sheet.command.test-resize-row', 'Resize preceding row', (draft) => {
      draft.sheets.find(({ id }) => id === output.getSheetId()).rowHeights[0] = 64
    })
    const movedPosition = bounds()
    assert.equal(movedPosition.top, positionAnchored.top + 12)
    assert.equal(movedPosition.width, positionAnchored.width)
    assert.equal(movedPosition.height, positionAnchored.height)

    await panel.applyPanelSlicerPresentation(studio, target, first.getId(), { anchorType: 'none' })
    const absolute = bounds()
    studio.mutateWorkbook(workbook.getId(), 'sheet.command.test-resize-row', 'Resize absolute row', (draft) => {
      draft.sheets.find(({ id }) => id === output.getSheetId()).rowHeights[0] = 76
    })
    assert.deepEqual(bounds(), absolute)

    await panel.applyPanelSlicerPresentation(studio, target, first.getId(), { arrange: 'front' })
    assert.ok(first.getInfo().drawing.zIndex > second.getInfo().drawing.zIndex)
    await panel.applyPanelSlicerPresentation(studio, target, first.getId(), { arrange: 'back' })
    assert.ok(first.getInfo().drawing.zIndex < second.getInfo().drawing.zIndex)
  } finally { studio.dispose() }
})

test('slicer canvas uses Pro pixel defaults, continuous timeline periods, and matching hit targets', async () => {
  const { studio, workbook, output, pivot } = await setup([
    ['Region', 'Date'], ['East', '2026-01-15'], ['West', '2026-03-02'],
  ], { labels: [0, 1] })
  try {
    const slicer = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '0', type: 'slicer' })
    const timeline = await output.insertSlicer({ pivotTableId: pivot.getPivotTableId(), sourceFieldId: '1', type: 'timeline' })
    assert.deepEqual(slicer.getInfo().layout, { left: 200, top: 200, width: 220, height: 280 })
    assert.deepEqual(timeline.getInfo().layout, { left: 200, top: 200, width: 420, height: 140 })
    const periods = renderer.timelinePeriods(['2026-01-15', '2026-03-02'], 'month', false)
    assert.equal(periods.length, 3)
    assert.deepEqual(periods.map(({ label }) => label), ['1月', '2月', '3月'])
    const sheetView = api.toExcelWorkbookView(studio.getWorkbookModel(workbook.getId())).sheets.find(sheet => sheet.id === output.getSheetId())
    const geometry = new layoutApi.SpreadsheetGeometry(sheetView)
    const controls = [slicer, timeline].map(control => ({ model: control.getInfo(), items: control.getItems() }))
    const layouts = renderer.slicerObjectLayouts(geometry, controls, false)
    assert.deepEqual(layouts.map(({ rect }) => rect), [
      { x: 200, y: 200, width: 220, height: 280 },
      { x: 200, y: 200, width: 420, height: 140 },
    ])
    assert.equal(renderer.hitTestSlicerObjects(layouts, layouts[0].clear.x + 2, layouts[0].clear.y + 2).type, 'clear')
    assert.equal(renderer.slicerPalette('SlicerStyleLight1', false).selected, '#3F83F8')
  } finally { studio.dispose() }
})
