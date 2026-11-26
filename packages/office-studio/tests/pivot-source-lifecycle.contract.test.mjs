import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCell } from '../src/domain/model.ts'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { mapSheetAxisIndex, transformSheetAxisRange } from '../src/domain/sheet-axis-structure.ts'
import { calculatePivotReport, getPivotSourceFields, pivotOutputValues, renderPivotTable } from '../src/features/pivot/services/pivot-engine.ts'
import { getPivotSourceAvailability, invalidatePivotSourceContentRanges, reconcilePivotSourceCaches, transformPivotSourceReport } from '../src/features/pivot/services/pivot-source-lifecycle.ts'
import { synchronizeLinkedPivotChartReports } from '../src/features/pivot/services/pivot-chart-report-lifecycle.ts'
import { planWorkbookAxisStructureCommand } from '../src/features/sheet-workbench/sheet-axis-structure-command.ts'
import { planWorkbookCellShiftCommand } from '../src/features/sheet-workbench/sheet-cell-shift-command.ts'
import { planWorksheetCopy } from '../src/features/sheet-workbench/sheet-worksheet-copy.ts'
import { planWorksheetInsert, planWorksheetRemove, planWorksheetRename } from '../src/features/sheet-workbench/sheet-worksheet-lifecycle.ts'
import { applySheetClipboardPastePlan, captureSheetClipboard, createSheetClipboardPastePlan } from '../src/features/sheet-workbench/sheet-clipboard.ts'
import { pivotSourceLifecycleWorkbook, allLifecycleReports, lifecycleRange } from './fixtures/pivot-source-lifecycle-fixture.mjs'

// Deferred drafts: no test runner, browser, Office sample, build or WASM check has been executed for this batch.
const json = value => JSON.parse(JSON.stringify(value))
const axis = (workbook, edit, sheetId = 'data') => planWorkbookAxisStructureCommand(workbook, sheetId, [edit], 'owner').workbook
const sourceCache = workbook => workbook.pivotCaches.find(cache => cache.id === 'source-cache')
const field = (report, id) => report.fields.find(field => field.id === id)
const sum = (workbook, report) => {
  const result = calculatePivotReport(workbook, report), index = report.fields.filter(field => field.area === 'row').length
  return result.values.slice(result.headerRows).reduce((total, row) => total + Number(row[index] ?? 0), 0)
}
function currentOutput(workbook, report) {
  const calculated = calculatePivotReport(workbook, report)
  assert.deepEqual(pivotOutputValues(workbook, report), calculated.values)
  assert.deepEqual(report.axisOutput, { rowAxis: calculated.rowAxis, columnAxis: calculated.columnAxis })
  for (const key of report.generatedCellKeys) assert.equal(workbook.sheets.find(sheet => sheet.id === report.anchor.subUnitId).cells[key].metadata.pivotTableId, report.id)
}
function paste(workbook, { sourceSheetId = 'data', sourceRange = lifecycleRange, targetSheetId = 'paste', row = 10, col = 8, copyType = 'copy' } = {}) {
  const snapshot = captureSheetClipboard(workbook, { unitId: workbook.id, sheetId: sourceSheetId, range: sourceRange }, { userId: 'owner', copyType })
  const plan = createSheetClipboardPastePlan(workbook, { unitId: workbook.id, sheetId: targetSheetId,
    range: { startRow: row, endRow: row, startCol: col, endCol: col } }, snapshot, 'all', { userId: 'owner' })
  assert.equal(plan.result.accepted, true)
  const next = structuredClone(workbook); applySheetClipboardPastePlan(next, plan); return next
}

test('a detached source transform retires matrix metadata while preserving report-local hierarchy and canonical provenance', () => {
  const original = pivotSourceLifecycleWorkbook(), before = json(original), edit = { axis: 'column', kind: 'insert', start: 4, count: 1 }
  const next = transformPivotSourceReport(original, original.pivotTables[0], { sheetId: 'data',
    range: range => transformSheetAxisRange(range, edit)[0], column: column => mapSheetAxisIndex(column, edit), point: (row, col) => ({ row, col }) })
  assert.equal(next.axisOutput, undefined); assert.deepEqual(next.drillDownMap, {})
  assert.deepEqual(next.collapseInfo, before.pivotTables[0].collapseInfo)
  assert.equal(field(next, 'segment-local').sourceColumn, 5); assert.equal(field(next, 'segment-local').sourceFieldId, 'segment-source')
  assert.equal(next.revision, 1); assert.deepEqual(json(original), before)
})

test.each([0, 2])('insertion before the source at column %s preserves every local instance and one shared cache identity', start => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'column', kind: 'insert', start, count: 2 })
  for (const report of allLifecycleReports(next)) {
    assert.equal(report.pivotCacheId, 'source-cache'); assert.deepEqual(report.fields.map(field => field.id), original.pivotTables[0].fields.map(field => field.id))
    assert.equal(field(report, 'amount-local').sourceColumn, 7); assert.equal(field(report, 'amount-average-local').sourceColumn, 7)
    assert.equal(field(report, 'amount-average-local').sourceFieldId, 'amount-source'); assert.equal(sum(next, report), 100)
  }
  assert.equal(next.pivotCaches.length, 1); assert.equal(sourceCache(next).revision, 1)
  for (const report of next.pivotTables) currentOutput(next, report)
  assert.equal(next.charts.find(chart => chart.ownedPivot).source.sheetId, 'data')
})

test('an interior source insertion allocates one unconfigured canonical field shared by all report owners', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'column', kind: 'insert', start: 5, count: 1 })
  const added = sourceCache(next).fields.find(field => field.sourceColumn === 5)
  assert.equal(added.id, 'source-field:1'); assert.equal(sourceCache(next).nextFieldId, 1)
  for (const report of allLifecycleReports(next)) {
    assert.equal(field(report, 'amount-local').sourceColumn, 6)
    assert.ok(report.fields.every(field => field.sourceFieldId !== added.id)); assert.equal(sum(next, report), 100)
    assert.equal(getPivotSourceFields(next, report).find(field => field.sourceColumn === 5).fieldId, added.id)
  }
})

test('copying a source column creates a new cache field without silently adding another report value instance', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'column', kind: 'copy', start: 5, count: 1, destination: 6 })
  const added = sourceCache(next).fields.find(field => field.sourceColumn === 6)
  assert.notEqual(added.id, 'amount-source'); assert.equal(added.name, 'Amount')
  assert.ok(allLifecycleReports(next).every(report => report.fields.filter(field => field.area === 'value').length === 2))
  assert.equal(sum(next, next.pivotTables[0]), 100); currentOutput(next, next.pivotTables[0])
})

test.each([{ start: 5, destination: 3 }, { start: 3, destination: 6 }])('internal column permutation %j preserves hierarchy keys, value rules and observable aggregates', edit => {
  const original = pivotSourceLifecycleWorkbook()
  for (const report of allLifecycleReports(original)) field(report, 'amount-local').showDataAs = { type: 'runningTotal', baseFieldId: 'region-local', baseItem: '' }
  for (const report of original.pivotTables) renderPivotTable(original, report)
  const expected = calculatePivotReport(original, original.pivotTables[0]).values
  const next = axis(original, { axis: 'column', kind: 'move', count: 1, ...edit })
  assert.deepEqual(next.pivotTables[0].source.range, lifecycleRange)
  for (const report of allLifecycleReports(next)) {
    assert.deepEqual(report.collapseInfo, original.pivotTables[0].collapseInfo)
    assert.equal(field(report, 'amount-local').showDataAs.baseFieldId, 'region-local')
    assert.deepEqual(calculatePivotReport(next, report).values, expected); assert.equal(report.revision, 1)
  }
  assert.equal(sourceCache(next).revision, 1); currentOutput(next, next.pivotTables[0])
})

test('deleting a canonical value source retires both SUM and AVG instances and clears linked output as one resource change', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'column', kind: 'delete', start: 5, count: 1 })
  for (const report of allLifecycleReports(next)) assert.ok(report.fields.every(field => field.sourceFieldId !== 'amount-source'))
  assert.ok(sourceCache(next).fields.every(field => field.id !== 'amount-source'))
  for (const report of next.pivotTables) { assert.equal(report.outputRange, undefined); assert.deepEqual(report.generatedCellKeys, []); currentOutput(next, report) }
  const linked = next.charts.find(chart => chart.id === 'linked-a')
  assert.deepEqual(linked.source.range, { startRow: 12, endRow: 12, startCol: 1, endCol: 1 }); assert.equal(linked.revision, 1)
})

test.each([3, 4])('deleting hierarchy source column %s removes local state and resets dangling Show Values As bases', start => {
  const original = pivotSourceLifecycleWorkbook(), base = start === 3 ? 'region-local' : 'segment-local'
  for (const report of allLifecycleReports(original)) field(report, 'amount-local').showDataAs = { type: 'runningTotal', baseFieldId: base, baseItem: '' }
  const next = axis(original, { axis: 'column', kind: 'delete', start, count: 1 })
  for (const report of allLifecycleReports(next)) {
    assert.equal(report.collapseInfo, undefined); assert.equal(field(report, 'amount-local').showDataAs, undefined)
    assert.equal(report.fields.filter(field => field.area === 'row')[0].position, 0); assert.equal(sum(next, report), 100)
  }
  if (start === 3) { assert.deepEqual(next.slicers, []); assert.deepEqual(next.slicerCaches, []) }
  else assert.equal(next.slicers[0].sourceFieldId, 'region-source')
})

test('structural deletion clears a surviving value filter whose selected measure instance was removed', () => {
  const original = pivotSourceLifecycleWorkbook()
  for (const report of allLifecycleReports(original)) {
    report.fields.push({ id: 'other-local', sourceFieldId: 'other-source', sourceColumn: 6, name: 'Other', area: 'value', position: 2, aggregation: 'sum' })
    field(report, 'segment-local').valueFilter = { operator: 'valueGreaterThan', expected: 0, valueFieldId: 'amount-local' }
  }
  const next = axis(original, { axis: 'column', kind: 'delete', start: 5, count: 1 })
  assert.ok(allLifecycleReports(next).every(report => field(report, 'segment-local').valueFilter === undefined))
  assert.ok(allLifecycleReports(next).every(report => field(report, 'other-local').position === 0))
  currentOutput(next, next.pivotTables[0]); assert.equal(sum(next, next.pivotTables[0]), 1000)
})

test('retired canonical field IDs are not reused by a later interior insertion', () => {
  const original = pivotSourceLifecycleWorkbook(), removed = axis(original, { axis: 'column', kind: 'delete', start: 5, count: 1 })
  const inserted = axis(removed, { axis: 'column', kind: 'insert', start: 5, count: 1 })
  assert.ok(sourceCache(inserted).fields.every(field => field.id !== 'amount-source'))
  assert.equal(sourceCache(inserted).fields.find(field => field.sourceColumn === 5).id, 'source-field:1')
  assert.equal(sourceCache(inserted).revision, 2)
})

test.each(['insert', 'delete', 'copy'])('row %s updates record membership and every report owner with one cache revision', kind => {
  const original = pivotSourceLifecycleWorkbook(), edit = { axis: 'row', kind, start: 3, count: 1, ...(kind === 'copy' ? { destination: 4 } : {}) }
  const next = axis(original, edit), expected = kind === 'delete' ? 90 : kind === 'copy' ? 110 : 100
  assert.equal(sourceCache(next).revision, 1)
  for (const report of allLifecycleReports(next)) { assert.equal(sum(next, report), expected); assert.equal(report.revision, 1) }
  for (const report of next.pivotTables) currentOutput(next, report)
})

test('an internal row move still invalidates the cache when the source rectangle stays equal', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'row', kind: 'move', start: 3, count: 1, destination: 7 })
  assert.deepEqual(sourceCache(next).source.range, lifecycleRange); assert.equal(sourceCache(next).revision, 1)
  assert.equal(next.pivotTables[0].revision, 1); assert.equal(sum(next, next.pivotTables[0]), 100)
  currentOutput(next, next.pivotTables[0])
})

test.each([3, 4])('a no-op row move to %s preserves source/cache revisions and hierarchy metadata', destination => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'row', kind: 'move', start: 3, count: 1, destination })
  assert.deepEqual(json(next.pivotCaches), json(original.pivotCaches)); assert.deepEqual(json(next.pivotTables), json(original.pivotTables))
  assert.deepEqual(json(next.charts), json(original.charts))
})

test('disconnected caches follow the same mapping while authored cache identities never coalesce by range', () => {
  const original = pivotSourceLifecycleWorkbook({ disconnected: true }), next = axis(original, { axis: 'column', kind: 'insert', start: 5, count: 1 })
  assert.equal(next.pivotCaches.length, 2)
  const detached = next.pivotCaches.find(cache => cache.id === 'disconnected-cache')
  assert.equal(detached.fields.find(field => field.sourceColumn === 5).id, 'source-field:5')
  assert.equal(sourceCache(next).fields.find(field => field.sourceColumn === 5).id, 'source-field:1')
  assert.equal(detached.revision, 1); assert.equal(sourceCache(next).revision, 1)
})

test('an output-only row insertion updates both linked chart sources and rebuilds drill-down coordinates', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'row', kind: 'insert', start: 10, count: 2 }, 'reports')
  assert.equal(sourceCache(next).revision, 0); assert.equal(next.pivotTables[0].anchor.row, 14)
  for (const linked of next.charts.filter(chart => chart.pivotTableId === 'report-a')) {
    assert.deepEqual(linked.source.range, next.pivotTables[0].outputRange); assert.equal(linked.revision, 1)
  }
  assert.equal(next.charts.find(chart => chart.ownedPivot).revision, 0); currentOutput(next, next.pivotTables[0])
})

test('an owned chart host row insertion does not turn the owned report into a worksheet output', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'row', kind: 'insert', start: 0, count: 2 }, 'dashboard')
  const owned = next.charts.find(chart => chart.ownedPivot).ownedPivot
  assert.deepEqual(json(owned), json(original.charts.find(chart => chart.ownedPivot).ownedPivot))
  assert.equal(owned.outputRange, undefined); assert.equal(owned.generatedCellKeys, undefined); assert.equal(owned.axisOutput, undefined)
})

test.each(['right', 'left'])('whole-band cell shift %s uses the same field and cache mapping as an axis edit', direction => {
  const original = pivotSourceLifecycleWorkbook(), next = planWorkbookCellShiftCommand(original, 'data',
    { startRow: 2, endRow: 6, startCol: 5, endCol: 5 }, direction, 'owner').workbook
  assert.equal(sourceCache(next).revision, 1)
  if (direction === 'right') { assert.equal(field(next.pivotTables[0], 'amount-local').sourceColumn, 6); assert.equal(sum(next, next.pivotTables[0]), 100) }
  else assert.ok(allLifecycleReports(next).every(report => report.fields.every(field => field.sourceFieldId !== 'amount-source')))
  for (const report of next.pivotTables) currentOutput(next, report)
})

test('source and output range refusals leave the complete original workbook untouched', () => {
  const original = pivotSourceLifecycleWorkbook(), before = json(original)
  assert.throws(() => axis(original, { axis: 'row', kind: 'delete', start: 2, count: 5 }), /透视数据源/)
  assert.throws(() => planWorkbookCellShiftCommand(original, 'reports', { startRow: 13, endRow: 13, startCol: 1, endCol: 1 }, 'up', 'owner'), /透视输出|透视生成格/)
  assert.deepEqual(json(original), before)
})

test('a cache cannot publish inconsistent source ranges from two newly diverged owners', () => {
  const original = pivotSourceLifecycleWorkbook(), before = json(original), candidate = structuredClone(original)
  candidate.pivotTables[1].source.range.endRow++
  assert.throws(() => reconcilePivotSourceCaches(original, candidate), /同一透视缓存出现不同数据源/)
  assert.deepEqual(json(original), before)
})

test('copying a self-contained sheet forks one source cache, preserves local collapse keys and isolates original records', () => {
  const original = pivotSourceLifecycleWorkbook({ selfContained: true }), before = json(original)
  const next = planWorksheetCopy(original, 'data', { userId: 'owner', targetSheetId: 'copy', targetSheetName: 'Data copy' }).workbook
  const copied = next.pivotTables.filter(report => report.anchor.subUnitId === 'copy'), owned = next.charts.find(chart => chart.sheetId === 'copy' && chart.ownedPivot).ownedPivot
  assert.equal(copied.length, 2); assert.notEqual(copied[0].pivotCacheId, 'source-cache'); assert.equal(copied[0].pivotCacheId, copied[1].pivotCacheId)
  assert.equal(owned.pivotCacheId, copied[0].pivotCacheId); assert.equal(owned.source.subUnitId, 'copy')
  for (const report of [...copied, owned]) { assert.deepEqual(report.collapseInfo, original.pivotTables[0].collapseInfo); assert.equal(field(report, 'amount-local').sourceFieldId, 'amount-source') }
  copied.forEach(report => currentOutput(next, report)); assert.equal(owned.axisOutput, undefined)
  next.sheets.find(sheet => sheet.id === 'copy').cells['3:5'].value = 999
  assert.deepEqual(json(original), before); assert.equal(next.sheets.find(sheet => sheet.id === 'data').cells['3:5'].value, 10)
})

test.each(['reports', 'dashboard'])('copying host %s retains external source identity and remaps each copied ownership reference', sheetId => {
  const original = pivotSourceLifecycleWorkbook(), next = planWorksheetCopy(original, sheetId, { userId: 'owner', targetSheetId: 'copy' }).workbook
  assert.equal(next.pivotCaches.length, 1); assert.equal(sourceCache(next).source.subUnitId, 'data')
  const copiedCharts = next.charts.filter(chart => chart.sheetId === 'copy')
  assert.ok(copiedCharts.length)
  for (const chart of copiedCharts) if (chart.ownedPivot) {
    assert.equal(chart.ownedPivot.source.subUnitId, 'data'); assert.equal(chart.ownedPivot.pivotCacheId, 'source-cache')
    assert.equal(chart.source.sheetId, 'data'); assert.equal(chart.ownedPivot.generatedCellKeys, undefined)
  } else {
    const report = next.pivotTables.find(report => report.id === chart.pivotTableId)
    assert.equal(report.source.subUnitId, 'data'); assert.equal(chart.source.sheetId, report.anchor.subUnitId)
    if (sheetId === 'reports') assert.notEqual(report.id, 'report-a')
  }
})

test('copying only a data source sheet leaves external reports bound to the original stable sheet ID', () => {
  const original = pivotSourceLifecycleWorkbook(), next = planWorksheetCopy(original, 'data', { userId: 'owner', targetSheetId: 'copy' }).workbook
  assert.equal(next.pivotTables.length, 2); assert.equal(next.charts.length, 3); assert.equal(next.pivotCaches.length, 1)
  assert.ok(allLifecycleReports(next).every(report => report.source.subUnitId === 'data'))
})

test('source rename preserves IDs and matrix metadata, and every owned/linked observer advances once', () => {
  const original = pivotSourceLifecycleWorkbook({ disconnected: true }), next = planWorksheetRename(original, 'data', 'Data 2026', 'owner').workbook
  assert.ok(next.pivotCaches.every(cache => cache.source.sheetName === 'Data 2026' && cache.revision === 1))
  for (const report of allLifecycleReports(next)) {
    assert.equal(report.source.subUnitId, 'data'); assert.equal(report.pivotCacheId, 'source-cache'); assert.equal(report.revision, 1)
    assert.equal(field(report, 'region-local').sourceFieldId, 'region-source')
  }
  assert.deepEqual(next.pivotTables[0].axisOutput, original.pivotTables[0].axisOutput)
  assert.ok(next.charts.every(chart => chart.revision === 1))
  const noOp = planWorksheetRename(next, 'data', 'Data 2026', 'owner')
  assert.equal(noOp.changed, false); assert.deepEqual(json(noOp.workbook), json(next))
})

test('source removal retains unresolved identities and the last linked matrix without claiming owned cached-record calculation', () => {
  const original = pivotSourceLifecycleWorkbook({ disconnected: true }), before = json(original), next = planWorksheetRemove(original, 'data', 'owner').workbook
  assert.equal(next.charts.length, 3); assert.equal(next.pivotCaches.length, 2)
  for (const report of allLifecycleReports(next)) {
    assert.equal(report.source.subUnitId, 'data'); assert.equal(report.pivotCacheId, 'source-cache')
    assert.equal(getPivotSourceAvailability(next, report), 'worksheet-missing'); assert.equal(report.axisOutput, undefined)
    assert.throws(() => calculatePivotReport(next, report), /source|worksheet|belong|来源|工作表/i)
  }
  const host = next.sheets.find(sheet => sheet.id === 'reports')
  assert.deepEqual(json(host.cells), before.sheets.find(sheet => sheet.id === 'reports').cells)
  assert.ok(next.charts.every(chart => chart.revision === 1)); assert.deepEqual(json(original), before)
})

test('creating another sheet with the removed source name never rebinds report or cache identity', () => {
  const original = pivotSourceLifecycleWorkbook(), removed = planWorksheetRemove(original, 'data', 'owner').workbook
  const next = planWorksheetInsert(removed, { userId: 'owner', name: 'Data', sheet: { id: 'replacement-data' } }).workbook
  assert.equal(next.sheets.find(sheet => sheet.id === 'replacement-data').name, 'Data')
  assert.ok(allLifecycleReports(next).every(report => report.source.subUnitId === 'data' && getPivotSourceAvailability(next, report) === 'worksheet-missing'))
  assert.equal(sourceCache(next).source.subUnitId, 'data')
})

test.each(['reports', 'dashboard'])('host removal %s retires actual owned/linked resources and keeps surviving source/cache owners', sheetId => {
  const original = pivotSourceLifecycleWorkbook(), next = planWorksheetRemove(original, sheetId, 'owner').workbook
  assert.ok(next.charts.every(chart => chart.sheetId !== sheetId)); assert.ok(next.pivotTables.every(report => report.anchor.subUnitId !== sheetId))
  assert.equal(sourceCache(next).source.subUnitId, 'data')
  if (sheetId === 'reports') { assert.deepEqual(next.charts.map(chart => chart.id), ['owned-chart']); assert.deepEqual(next.slicerCaches[0].connections.pivotTables, []) }
  else { assert.equal(next.charts.some(chart => chart.ownedPivot), false); assert.equal(next.pivotTables.length, 2) }
})

test('saved renamed, copied and unresolved-source snapshots reopen with independent source/cache and local hierarchy identities', () => {
  let current = pivotSourceLifecycleWorkbook({ selfContained: true })
  current = planWorksheetCopy(current, 'data', { userId: 'owner', targetSheetId: 'copy' }).workbook
  current = planWorksheetRename(current, 'data', 'Source renamed', 'owner').workbook
  current = planWorksheetRemove(current, 'data', 'owner').workbook
  const saved = json(current), reopened = workbookFromInput(saved)
  assert.deepEqual(json(reopened.pivotCaches), saved.pivotCaches); assert.deepEqual(json(reopened.pivotTables), saved.pivotTables)
  assert.deepEqual(json(reopened.charts), saved.charts)
  reopened.pivotTables[0].collapseInfo['region-local'].North = false
  assert.deepEqual(json(current), saved)
})

test('a full source range cut moves shared and disconnected cache metadata and refreshes report hosts without aliasing original sheets', () => {
  const original = pivotSourceLifecycleWorkbook({ disconnected: true }), before = json(original)
  const next = paste(original, { copyType: 'cut' })
  for (const report of allLifecycleReports(next)) {
    assert.equal(report.source.subUnitId, 'paste'); assert.equal(report.source.range.startRow, 10); assert.equal(report.source.range.startCol, 8)
    assert.equal(field(report, 'region-local').sourceFieldId, 'region-source'); assert.equal(field(report, 'amount-local').sourceColumn, 10)
    assert.equal(sum(next, report), 100); assert.deepEqual(report.collapseInfo, before.pivotTables[0].collapseInfo)
  }
  assert.ok(next.pivotCaches.every(cache => cache.source.subUnitId === 'paste' && cache.revision === 1))
  next.pivotTables.forEach(report => currentOutput(next, report)); assert.deepEqual(json(original), before)
})

test('copying a complete linked output rebinds its generated cells and publishes fresh axis metadata for the copied report', () => {
  const original = pivotSourceLifecycleWorkbook(), before = json(original)
  const next = paste(original, { sourceSheetId: 'reports', sourceRange: original.pivotTables[0].outputRange, row: 40, col: 2 })
  const copied = next.pivotTables.find(report => report.anchor.subUnitId === 'paste')
  assert.ok(copied); assert.notEqual(copied.id, 'report-a'); assert.equal(copied.pivotCacheId, 'source-cache')
  assert.deepEqual(copied.collapseInfo, original.pivotTables[0].collapseInfo); currentOutput(next, copied)
  assert.deepEqual(json(original), before)
})

test('a copied stale matrix retires old generated cells when the fresh source result has fewer rows', () => {
  const original = pivotSourceLifecycleWorkbook(), source = original.sheets.find(sheet => sheet.id === 'data'), old = original.pivotTables[0]
  for (let row = 3; row <= 6; row++) { source.cells[`${row}:3`].value = 'North'; source.cells[`${row}:4`].value = 'Retail' }
  const before = json(original), next = paste(original, { sourceSheetId: 'reports', sourceRange: old.outputRange, row: 40, col: 2 })
  const copied = next.pivotTables.find(report => report.anchor.subUnitId === 'paste'), target = next.sheets.find(sheet => sheet.id === 'paste')
  currentOutput(next, copied); assert.ok(copied.outputRange.endRow - copied.outputRange.startRow < old.outputRange.endRow - old.outputRange.startRow)
  const freshKeys = new Set(copied.generatedCellKeys)
  for (const key of old.generatedCellKeys) {
    const [row, col] = key.split(':').map(Number), projected = `${40 + row - old.outputRange.startRow}:${2 + col - old.outputRange.startCol}`
    if (!freshKeys.has(projected)) assert.equal(target.cells[projected], undefined)
  }
  assert.deepEqual(json(original), before)
})

test('pasting changed source values refreshes real linked and owned aggregates in the same candidate', () => {
  const original = pivotSourceLifecycleWorkbook(), before = json(original), donor = original.sheets.find(sheet => sheet.id === 'paste')
  donor.cells['0:0'] = createCell(0, 0, 90)
  const next = paste(original, { sourceSheetId: 'paste', sourceRange: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    targetSheetId: 'data', row: 3, col: 5 })
  assert.equal(sourceCache(next).revision, 1); assert.ok(allLifecycleReports(next).every(report => sum(next, report) === 180))
  next.pivotTables.forEach(report => currentOutput(next, report))
  assert.deepEqual(json(original.sheets.find(sheet => sheet.id === 'reports')), before.sheets.find(sheet => sheet.id === 'reports'))
})

test('unchanged source content and presentation-only clipboard scopes do not invalidate cached report state', () => {
  const original = pivotSourceLifecycleWorkbook(), next = structuredClone(original), before = json(next)
  next.sheets[0].cells['3:5'].style.font.bold = true
  invalidatePivotSourceContentRanges(original, next, [{ sheetId: 'data', range: lifecycleRange }])
  assert.deepEqual(json(next.pivotCaches), before.pivotCaches); assert.deepEqual(json(next.pivotTables), before.pivotTables)
  assert.deepEqual(json(next.charts), before.charts)
})

test('owned and linked chart revision synchronization is idempotent within one source transaction', () => {
  const original = pivotSourceLifecycleWorkbook(), next = axis(original, { axis: 'column', kind: 'insert', start: 5, count: 1 }), before = json(next.charts)
  synchronizeLinkedPivotChartReports(original, next); synchronizeLinkedPivotChartReports(original, next)
  assert.deepEqual(json(next.charts), before); assert.ok(next.charts.every(chart => chart.revision === 1))
})
