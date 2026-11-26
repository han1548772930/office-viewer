import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { normalizeWorkbook } from '../src/parser/workbookNormalizer.ts'
import { toExcelWorkbookView, applyExcelWorkbookView } from '../src/services/excelViewAdapter.ts'
import { planWorksheetCopy } from '../src/features/sheet-workbench/sheet-worksheet-copy.ts'
import { planWorksheetRemove, planWorksheetRename } from '../src/features/sheet-workbench/sheet-worksheet-lifecycle.ts'
import { captureSheetClipboard, createSheetClipboardPastePlan, applySheetClipboardPastePlan } from '../src/features/sheet-workbench/sheet-clipboard.ts'
import { assertSheetClipboardWorkbookResources } from '../src/features/sheet-workbench/sheet-clipboard-resource-validation.ts'
import { reconcileSlicerConnectionResources } from '../src/features/slicer/slicer-connections-lifecycle.ts'
import { normalizeSlicerCaches, getSlicerConnections, getSlicerConnectionReports } from '../src/features/slicer/slicer-connections.ts'
import { transformWorkbookAxisResources } from '../src/features/sheet-workbench/sheet-axis-structure-resources.ts'
import { slicerConnectionsWorkbook } from './fixtures/slicer-connections-lifecycle-fixture.mjs'

// Deferred contracts: no tests, Vite services, browsers, WASM or Office files have been executed.
const json = value => JSON.parse(JSON.stringify(value))
const captureRange = { startRow: 0, endRow: 29, startCol: 0, endCol: 12 }
const targetRange = { startRow: 40, endRow: 40, startCol: 4, endCol: 4 }
const capture = workbook => captureSheetClipboard(workbook, { unitId: workbook.id, sheetId: workbook.activeSheetId, range: captureRange }, { userId: 'owner' })
const planPaste = (workbook, snapshot) => createSheetClipboardPastePlan(workbook, { unitId: workbook.id, sheetId: 'paste', range: targetRange }, snapshot, 'all', { userId: 'owner' })

test('saved canonical input reopens shared source/control caches, owned reports and independent presentation without aliases', () => {
  const original = slicerConnectionsWorkbook({ stableFields: true }), saved = json(original), reopened = workbookFromInput(saved)
  assert.deepEqual(json(reopened.slicerCaches), saved.slicerCaches)
  assert.deepEqual(json(reopened.pivotCaches), saved.pivotCaches)
  assert.deepEqual(json(reopened.charts[0].ownedPivot), saved.charts[0].ownedPivot)
  reopened.slicerCaches[0].connections.pivotTables[0].pivotTableId = 'mutated'
  reopened.charts[0].ownedPivot.fields[0].selectedItems.push('South')
  assert.deepEqual(original, saved)
})

test('public partial workbook input retains cache arrays through normalization', () => {
  const source = slicerConnectionsWorkbook(), input = { ...json(source) }; delete input.kind
  const reopened = workbookFromInput(input)
  assert.deepEqual(reopened.slicerCaches[0].connections, source.slicerCaches[0].connections)
  assert.deepEqual(reopened.pivotCaches[0].fields, source.pivotCaches[0].fields)
})

test('view conversion retains source fields and isolates independent owned reports, affine authors and cache connections', () => {
  const original = slicerConnectionsWorkbook(), before = json(original), view = toExcelWorkbookView(original)
  const restored = applyExcelWorkbookView(original, view)
  assert.deepEqual(json(restored.slicerCaches), before.slicerCaches); assert.deepEqual(json(restored.pivotCaches), before.pivotCaches)
  view.charts[0].ownedPivot.fields[0].selectedItems.push('South'); view.charts[0].ownedPivot.source.range.endRow = 100
  view.slicers[0].drawingGeometry.matrix.a = 99; view.slicerCaches[0].connections.pivotTables.pop()
  assert.deepEqual(json(original), before); assert.deepEqual(json(restored.charts[0].ownedPivot), before.charts[0].ownedPivot)
})

test('native parser normalization retains one explicit cache identity across two reports and source-field provenance', () => {
  const original = slicerConnectionsWorkbook({ owned: false }), raw = original.pivotTables.map(pivot => ({ ...json(pivot),
    source: { ...json(pivot.source), subUnitId: 'Data', pivotCacheId: 'native-shared-cache' }, pivotCacheId: 'native-shared-cache',
    anchor: { ...pivot.anchor, subUnitId: 'Reports' } }))
  const workbook = normalizeWorkbook({ schemaVersion: 1, parserVersion: 'contract', format: 'xlsx', fileName: 'shared.xlsx',
    payload: { modelType: 'workbook', data: { sheets: [{ name: 'Data', cells: Object.values(original.sheets[0].cells) }, { name: 'Reports', cells: [] }], pivotTables: raw } }, diagnostics: [] })
  assert.equal(workbook.pivotTables.length, 2); assert.equal(workbook.pivotCaches.length, 1)
  assert.equal(workbook.pivotTables[0].pivotCacheId, 'native-shared-cache')
  assert.equal(workbook.pivotTables[1].source.pivotCacheId, 'native-shared-cache')
  assert.equal(workbook.pivotTables[1].fields[0].sourceFieldId, '0')
})

test('whole-sheet copy remaps every report and owned target, forks one shared control cache and keeps source data cache', () => {
  const source = slicerConnectionsWorkbook({ stableFields: true }), before = json(source)
  const next = planWorksheetCopy(source, 'reports', { userId: 'owner', targetSheetId: 'copy' }).workbook
  const views = next.slicers.filter(view => view.hostSheetId === 'copy'), cache = next.slicerCaches.find(cache => cache.id === views[0].cacheId)
  assert.equal(views.length, 2); assert.equal(views[0].cacheId, views[1].cacheId); assert.notEqual(cache.id, 'control-cache')
  assert.equal(cache.source.pivotCacheId, 'source-cache'); assert.equal(next.pivotCaches.length, 1)
  assert.ok(cache.connections.pivotTables.every(ref => ref.subUnitId === 'copy' && !before.pivotTables.some(pivot => pivot.id === ref.pivotTableId)))
  assert.equal(cache.connections.ownedPivotCharts[0].subUnitId, 'copy')
  assert.notEqual(cache.connections.ownedPivotCharts[0].pivotChartId, 'owned-chart')
  assert.deepEqual(json(next.slicerCaches.find(cache => cache.id === 'control-cache')), before.slicerCaches[0])
  assert.ok(next.pivotTables.filter(pivot => pivot.anchor.subUnitId === 'copy').every(pivot => pivot.fields[0].sourceFieldId === 'region-source'))
  assert.deepEqual(json(source), before)
})

test('copying a self-contained source sheet creates a new source cache for every copied report without coalescing the original', () => {
  const source = slicerConnectionsWorkbook({ singleSheet: true }), before = json(source)
  const next = planWorksheetCopy(source, 'data', { userId: 'owner', targetSheetId: 'copy' }).workbook
  const controls = next.slicers.filter(view => view.hostSheetId === 'copy'), cache = next.slicerCaches.find(cache => cache.id === controls[0].cacheId)
  assert.notEqual(cache.source.pivotCacheId, 'source-cache')
  const copiedSource = next.pivotCaches.find(source => source.id === cache.source.pivotCacheId)
  assert.equal(copiedSource.source.subUnitId, 'copy'); assert.deepEqual(copiedSource.fields, source.pivotCaches[0].fields)
  assert.ok(next.pivotTables.filter(pivot => pivot.anchor.subUnitId === 'copy').every(pivot => pivot.pivotCacheId === copiedSource.id))
  assert.equal(next.charts.find(chart => chart.sheetId === 'copy').ownedPivot.pivotCacheId, copiedSource.id)
  assert.deepEqual(json(source), before)
})

test('legacy report-local field hints copy to canonical source IDs and complete source cache metadata', () => {
  const source = slicerConnectionsWorkbook({ singleSheet: true, stableFields: true })
  source.pivotCaches = []; source.slicerCaches = []
  for (const report of [...source.pivotTables, source.charts[0].ownedPivot]) {
    delete report.pivotCacheId
    report.fields.forEach(field => { delete field.sourceFieldId })
  }
  source.slicers.forEach(view => { delete view.cacheId; view.sourceFieldId = source.pivotTables[0].fields[0].id })
  const before = json(source), next = planWorksheetCopy(source, 'data', { userId: 'owner', targetSheetId: 'copy' }).workbook
  const view = next.slicers.find(view => view.hostSheetId === 'copy'), cache = next.slicerCaches.find(cache => cache.id === view.cacheId)
  assert.equal(view.sourceFieldId, '0'); assert.equal(cache.source.sourceFieldId, '0')
  assert.equal(next.pivotCaches.find(cache => cache.id === next.pivotTables.find(report => report.anchor.subUnitId === 'copy').pivotCacheId).fields.length, 3)
  assert.deepEqual(json(source), before)
})

test('copying an explicitly disconnected view preserves its shared cache identity and omitted connection shape', () => {
  const source = slicerConnectionsWorkbook({ owned: false })
  source.pivotTables = []; source.charts = []; delete source.slicerCaches[0].connections
  source.slicers.forEach(view => { delete view.pivotTableId })
  const next = planWorksheetCopy(source, 'reports', { userId: 'owner', targetSheetId: 'copy' }).workbook
  assert.equal(next.slicerCaches.length, 1); assert.equal(next.slicerCaches[0].connections, undefined)
  assert.ok(next.slicers.filter(view => view.hostSheetId === 'copy').every(view => view.cacheId === 'control-cache'))
})

test('copying only view presentation reuses unchanged report/source connections and authoritative current selection', () => {
  const source = slicerConnectionsWorkbook(), before = json(source)
  for (const view of source.slicers) view.hostSheetId = 'paste'
  const next = planWorksheetCopy(source, 'paste', { userId: 'owner', targetSheetId: 'copy' }).workbook
  const views = next.slicers.filter(view => view.hostSheetId === 'copy')
  assert.ok(views.every(view => view.cacheId === 'control-cache')); assert.equal(next.slicerCaches.length, 1)
  assert.deepEqual(next.slicerCaches[0].connections, before.slicerCaches[0].connections)
  assert.deepEqual(views.map(view => view.selectedItems), [['North'], ['North']])
})

test('worksheet deletion removes one report connection while keeping views bound to the remaining owned report', () => {
  const source = slicerConnectionsWorkbook(), before = json(source)
  source.slicers.forEach(view => { view.hostSheetId = 'paste' })
  source.charts[0].sheetId = 'paste'; source.charts[0].ownedPivot.anchor.subUnitId = 'paste'
  source.slicerCaches[0].connections.ownedPivotCharts[0].subUnitId = 'paste'
  const next = planWorksheetRemove(source, 'reports', 'owner').workbook
  assert.equal(next.slicers.length, 2); assert.deepEqual(next.slicerCaches[0].connections.pivotTables, [])
  assert.deepEqual(next.slicerCaches[0].connections.ownedPivotCharts, [{ subUnitId: 'paste', pivotChartId: 'owned-chart' }])
  assert.ok(next.slicers.every(view => view.pivotTableId === undefined && view.ownedPivotChartId === 'owned-chart'))
  assert.deepEqual(next.slicerCaches[0].selectedItems, before.slicerCaches[0].selectedItems)
})

test('a fully disconnected control retains source, selection and metadata after its last report is deliberately removed', () => {
  const source = slicerConnectionsWorkbook({ owned: false })
  source.slicers.forEach(view => { view.hostSheetId = 'paste' })
  const next = planWorksheetRemove(source, 'reports', 'owner').workbook
  assert.equal(next.slicers.length, 2); assert.deepEqual(next.slicerCaches[0].connections.pivotTables, [])
  assert.deepEqual(next.slicerCaches[0].selectedItems, ['North']); assert.equal(next.pivotCaches[0].source.subUnitId, 'data')
  assert.deepEqual(getSlicerConnections(next, next.slicers[0]), { pivotTables: [] })
})

test('renaming a source worksheet updates retained source cache and every independent owned report', () => {
  const source = slicerConnectionsWorkbook(), next = planWorksheetRename(source, 'data', 'Renamed', 'owner').workbook
  assert.equal(next.pivotCaches[0].source.sheetName, 'Renamed')
  assert.ok(next.pivotTables.every(pivot => pivot.source.sheetName === 'Renamed'))
  assert.equal(next.charts[0].ownedPivot.source.sheetName, 'Renamed')
})

test('source cache and report column edits preserve nonnumeric source IDs through an insertion before the range', () => {
  const original = slicerConnectionsWorkbook({ stableFields: true }), next = structuredClone(original)
  transformWorkbookAxisResources(next, original, original.sheets[0], { axis: 'column', kind: 'insert', start: 0, count: 2 })
  assert.deepEqual(next.pivotCaches[0].fields.map(field => [field.id, field.sourceColumn]), [['region-source', 2], ['1', 3], ['2', 4]])
  assert.ok(next.pivotTables.every(pivot => pivot.fields[0].sourceFieldId === 'region-source' && pivot.fields[0].sourceColumn === 2))
  assert.equal(next.charts[0].ownedPivot.fields[0].sourceColumn, 2)
  assert.equal(next.slicers[0].sourceFieldId, 'region-source')
})

test('disconnected cache-only metadata still transforms when no report exists', () => {
  const original = slicerConnectionsWorkbook({ owned: false }), next = structuredClone(original)
  original.pivotTables = []; original.charts = []; original.slicerCaches[0].connections = { pivotTables: [] }
  next.pivotTables = []; next.charts = []; next.slicerCaches[0].connections = { pivotTables: [] }
  reconcileSlicerConnectionResources(original, next, { sheetId: 'data', range: range => ({ ...range, startCol: range.startCol + 3, endCol: range.endCol + 3 }), column: column => column + 3 })
  assert.equal(next.pivotCaches[0].source.range.startCol, 3)
  assert.deepEqual(next.pivotCaches[0].fields.map(field => field.sourceColumn), [3, 4, 5])
  assert.deepEqual(next.slicerCaches[0].selectedItems, ['North'])
})

test('deleting the actual bound source column removes its controls instead of retargeting them to the next column', () => {
  const original = slicerConnectionsWorkbook(), next = structuredClone(original)
  transformWorkbookAxisResources(next, original, original.sheets[0], { axis: 'column', kind: 'delete', start: 0, count: 1 })
  assert.deepEqual(next.slicers, []); assert.deepEqual(next.slicerCaches, [])
  assert.ok(next.pivotCaches[0].fields.every(field => field.id !== '0'))
})

test('deleted source field IDs are never reused when a new column is inserted at the same position', () => {
  const original = slicerConnectionsWorkbook(), removed = structuredClone(original)
  transformWorkbookAxisResources(removed, original, original.sheets[0], { axis: 'column', kind: 'delete', start: 0, count: 1 })
  const inserted = structuredClone(removed)
  transformWorkbookAxisResources(inserted, removed, removed.sheets[0], { axis: 'column', kind: 'insert', start: 0, count: 1 })
  assert.ok(inserted.pivotCaches[0].fields.every(field => field.id !== '0'))
  // Insertion at the source boundary shifts the existing range, so use an interior insertion to create a new source column.
  const expanded = structuredClone(inserted)
  transformWorkbookAxisResources(expanded, inserted, inserted.sheets[0], { axis: 'column', kind: 'insert', start: 2, count: 1 })
  const created = expanded.pivotCaches[0].fields.find(field => field.sourceColumn === 2)
  assert.match(created.id, /^source-field:\d+$/); assert.ok(expanded.pivotCaches[0].nextFieldId > 0)
})

test('disconnected cache metadata allocates monotonic IDs for inserted source columns without any report', () => {
  const original = slicerConnectionsWorkbook({ owned: false })
  original.pivotTables = []; original.charts = []; original.slicerCaches[0].connections = { pivotTables: [] }
  original.pivotCaches[0].nextFieldId = 7
  const next = structuredClone(original)
  transformWorkbookAxisResources(next, original, original.sheets[0], { axis: 'column', kind: 'insert', start: 1, count: 1 })
  assert.equal(next.pivotCaches[0].fields.find(field => field.sourceColumn === 1).id, 'source-field:8')
  assert.equal(next.pivotCaches[0].nextFieldId, 8)
})

test('unknown unavailable original connections are retained through unrelated lifecycle changes', () => {
  const source = slicerConnectionsWorkbook(), next = structuredClone(source)
  source.slicerCaches[0].connections.pivotTables.push({ subUnitId: 'unavailable', pivotTableId: 'lost-report' })
  next.slicerCaches[0].connections = structuredClone(source.slicerCaches[0].connections)
  reconcileSlicerConnectionResources(source, next)
  assert.ok(next.slicerCaches[0].connections.pivotTables.some(ref => ref.pivotTableId === 'lost-report'))
})

test('unrelated reconciliation preserves an already invalid source and mismatched connection host for explicit later refusal', () => {
  const source = slicerConnectionsWorkbook()
  source.slicerCaches[0].connections.pivotTables[0].subUnitId = 'unavailable'
  source.pivotCaches[0].source.range.endRow = 90
  const next = structuredClone(source)
  reconcileSlicerConnectionResources(source, next)
  assert.equal(next.pivotCaches[0].source.range.endRow, 90)
  assert.equal(next.slicerCaches[0].connections.pivotTables[0].subUnitId, 'unavailable')
  assert.throws(() => getSlicerConnectionReports(next, next.slicers[0]))
})

test('range clipboard captures cache metadata once, remaps all copied report refs and preserves both view matrices', () => {
  const source = slicerConnectionsWorkbook(), before = json(source), snapshot = capture(source), plan = planPaste(source, snapshot)
  assert.equal(snapshot.workbookResources.slicerCaches.length, 1); assert.equal(snapshot.workbookResources.pivotCaches.length, 1)
  assert.equal(plan.result.accepted, true); assert.deepEqual(json(source), before)
  const next = structuredClone(source); applySheetClipboardPastePlan(next, plan)
  const views = next.slicers.filter(view => view.hostSheetId === 'paste'), cache = next.slicerCaches.find(cache => cache.id === views[0].cacheId)
  assert.equal(views.length, 2); assert.equal(views[0].cacheId, views[1].cacheId)
  assert.equal(cache.source.pivotCacheId, 'source-cache'); assert.ok(cache.connections.pivotTables.every(ref => ref.subUnitId === 'paste'))
  assert.equal(cache.connections.ownedPivotCharts[0].subUnitId, 'paste')
  assert.equal(views[0].drawingGeometry.matrix.a, before.slicers[0].drawingGeometry.matrix.a)
  assert.equal(views[0].drawingGeometry.width, before.slicers[0].drawingGeometry.width)
})

test('a cross-workbook copy with every source and report in the copied region creates local caches and no foreign refs', () => {
  const source = slicerConnectionsWorkbook({ singleSheet: true }), snapshot = capture(source), destination = slicerConnectionsWorkbook()
  destination.id = 'destination-unit'; destination.pivotTables = []; destination.charts = []; destination.slicers = []; destination.pivotCaches = []; destination.slicerCaches = []
  const plan = planPaste(destination, snapshot), next = structuredClone(destination); applySheetClipboardPastePlan(next, plan)
  const control = next.slicers[0], cache = next.slicerCaches.find(cache => cache.id === control.cacheId)
  assert.equal(control.unitId, 'destination-unit'); assert.equal(cache.unitId, 'destination-unit')
  assert.notEqual(cache.source.pivotCacheId, 'source-cache')
  assert.ok(next.pivotTables.every(pivot => pivot.source.unitId === 'destination-unit' && pivot.source.subUnitId === 'paste'))
  assert.ok(next.pivotCaches[0].fields.every(field => field.sourceColumn >= targetRange.startCol))
})

test('cutting a disconnected cache source moves its retained source and field columns without replacing cache identity', () => {
  const source = slicerConnectionsWorkbook({ owned: false })
  source.pivotTables = []; source.charts = []; source.slicerCaches[0].connections = { pivotTables: [] }
  const snapshot = captureSheetClipboard(source, { unitId: source.id, sheetId: 'data', range: source.pivotCaches[0].source.range }, { userId: 'owner', copyType: 'cut' })
  const before = json(source), plan = planPaste(source, snapshot), next = structuredClone(source)
  assert.equal(plan.result.accepted, true); applySheetClipboardPastePlan(next, plan)
  assert.equal(next.pivotCaches[0].id, 'source-cache'); assert.equal(next.pivotCaches[0].source.subUnitId, 'paste')
  assert.deepEqual(next.pivotCaches[0].fields.map(field => field.sourceColumn), [4, 5, 6])
  assert.equal(next.slicerCaches[0].source.pivotCacheId, 'source-cache'); assert.deepEqual(next.slicerCaches[0].selectedItems, ['North'])
  assert.deepEqual(json(source), before)
})

test('an unresolved foreign connection fails the whole detached paste without changing the destination', () => {
  const source = slicerConnectionsWorkbook(), snapshot = capture(source), destination = slicerConnectionsWorkbook()
  destination.id = 'destination-unit'; destination.pivotTables = []; destination.charts = []; destination.slicers = []; destination.pivotCaches = []; destination.slicerCaches = []
  const before = json(destination)
  assert.throws(() => planPaste(destination, snapshot), /数据源|连接|报告|来源/)
  assert.deepEqual(json(destination), before)
})

test('copied cache connection replacement cannot silently omit an uncopied incompatible report', () => {
  const source = slicerConnectionsWorkbook({ singleSheet: true }), before = json(source)
  source.pivotTables[1].anchor.subUnitId = 'reports'; source.slicerCaches[0].connections.pivotTables[1].subUnitId = 'reports'
  const current = json(source)
  assert.throws(() => planWorksheetCopy(source, 'data', { userId: 'owner', targetSheetId: 'copy' }), /不同透视缓存|所有连接报告/)
  assert.deepEqual(json(source), current); assert.equal(before.pivotTables.length, source.pivotTables.length)
})

for (const [name, corrupt] of [
  ['foreign cache owner', resource => { resource.slicerCaches[0].unitId = 'foreign' }],
  ['wrong source cache', resource => { resource.slicerCaches[0].source.pivotCacheId = 'foreign' }],
  ['duplicate report ref', resource => { resource.slicerCaches[0].connections.pivotTables.push(structuredClone(resource.slicerCaches[0].connections.pivotTables[0])) }],
  ['non-finite source column', resource => { resource.pivotCaches[0].fields[0].sourceColumn = Infinity }],
  ['invalid source field sequence', resource => { resource.pivotCaches[0].nextFieldId = -1 }],
  ['conflicting owned report', resource => { resource.charts[0].pivotTableId = 'report-a' }],
  ['forged field identity', resource => { resource.slicerCaches[0].source.sourceFieldId = 'absent' }],
]) test(`clipboard cache/ref validation rejects ${name} before the authority is changed`, () => {
  const source = slicerConnectionsWorkbook(), snapshot = capture(source), before = json(source)
  corrupt(snapshot.workbookResources)
  assert.throws(() => planPaste(source, snapshot)); assert.deepEqual(json(source), before)
})

test('retained source cache IDs stay distinct even when two authored caches have equal ranges', () => {
  const source = slicerConnectionsWorkbook({ owned: false })
  source.pivotTables[1].pivotCacheId = 'another-cache'
  source.pivotCaches.push({ ...structuredClone(source.pivotCaches[0]), id: 'another-cache' })
  source.slicerCaches[0].connections.pivotTables.pop()
  normalizeSlicerCaches(source)
  const next = planWorksheetCopy(source, 'reports', { userId: 'owner', targetSheetId: 'copy' }).workbook
  assert.equal(new Set(next.pivotCaches.map(cache => cache.id)).size, 2)
  assert.equal(next.pivotTables.find(pivot => pivot.anchor.subUnitId === 'copy' && pivot.fields[0].sourceFieldId === '0').pivotCacheId, 'source-cache')
  assertSheetClipboardWorkbookResources(capture(source).workbookResources)
})
