import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createPinia } from 'pinia'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChartDefinition } from '../src/features/charts/chart-engine.ts'
import { CHART_OBJECT_COMMANDS, CHART_OBJECT_MENU, CHART_SHEET_METADATA_KEY, PASTE_COPIED_CHART_COMMAND,
  chartObjectCapabilities, chartObjectShortcut, chartObjectUIHost, registerChartObjectUIHost, resolveChartObjectTarget } from '../src/features/charts/chart-object-commands.ts'
import { captureChartObjectClipboard, chartObjectClipboardContent, chartObjectClipboardIdFromContent,
  chartObjectPasteDefinition, createChartObjectId } from '../src/features/charts/chart-object-clipboard.ts'
import { chartObjectClipboardStore } from '../src/features/charts/chart-object-clipboard-store.ts'
import { pruneChartImageResources } from '../src/features/charts/chart-image-resources.ts'
import { chartResourceFixture } from './fixtures/chart-resource-platform-fixture.mjs'

// Written only. These are model/control contracts, not accepted browser or image samples.
const json = value => JSON.parse(JSON.stringify(value))
function fixture(type = 'line') {
  const workbook = createBlankWorkbook(), source = workbook.sheets[0]
  workbook.id = 'objects-book'; source.id = 'source'; workbook.activeSheetId = source.id
  source.cells = Object.fromEntries([['Label', 'Value'], ['North', 7], ['South', 4]]
    .flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const destination = { ...json(source), id: 'destination', name: 'Destination', cells: {} }
  workbook.sheets.push(destination)
  const chart = createChartDefinition(workbook, source.id, { type, title: '<Main & chart>', subtitle: 'Sub',
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 } }, categoryField: 0, valueFields: [1],
    position: { x: 64, y: 48 }, size: { width: 480, height: 320 }, palette: ['#123456', '#abcdef'],
    options: { titleStyle: { rotate: 30, manualLayout: { left: 0.2, top: 0.3 } }, plotArea: { left: 0.1, top: 0.1, width: 0.8, height: 0.8 } },
    legend: { visible: true, position: 'right' }, yAxis: { scaleType: 'linear', min: 0, max: 10 } }, 'original')
  workbook.charts = [chart]
  const target = () => ({ unitId: workbook.id, sheetId: source.id, chartId: chart.id, expectedRevision: 4, snapshot: JSON.stringify(chart) })
  const studio = { activeWorkbookId: workbook.id, currentUserId: 'local-user',
    getWorkbookModel: id => id === workbook.id ? workbook : undefined, getWorkbookRevision: () => 4 }
  return { workbook, source, destination, chart, target, studio }
}
function protection(viewState = 'othersCanView') {
  return { id: 'protected', permissionId: 'permission', unitType: 'worksheet', ownerUserId: 'owner',
    allowedUserIds: [], editState: 'onlyMe', viewState }
}

test('copy captures a detached author and only referenced image resources without changing content', () => {
  const { workbook, chart, target } = fixture('wordCloud'), image = chartResourceFixture(), unused = chartResourceFixture('unused')
  chart.options.wordCloud = { maskImageId: image.id }; workbook.metadata = { chartImageResources: { [image.id]: image, unused } }
  const before = JSON.stringify(workbook), copied = captureChartObjectClipboard(workbook, target())
  assert.equal(JSON.stringify(workbook), before); assert.deepEqual(Object.keys(copied.images), [image.id])
  copied.chart.options.titleStyle.manualLayout.left = 0.9; copied.chart.palette[0] = '#000000'; copied.images[image.id].name = 'changed'
  assert.equal(chart.options.titleStyle.manualLayout.left, 0.2); assert.equal(chart.palette[0], '#123456'); assert.equal(image.name, 'mask.png')
})

test('same-workbook cross-sheet paste preserves source, titles, axes and styles while assigning a fresh identity', () => {
  const { workbook, chart, destination, target } = fixture(), copied = captureChartObjectClipboard(workbook, target())
  const original = json(chart), created = chartObjectPasteDefinition(workbook, copied,
    { unitId: workbook.id, sheetId: destination.id, newChartId: 'pasted', clipboardId: copied.id })
  assert.equal(created.id, 'pasted'); assert.equal(created.sheetId, destination.id); assert.equal(created.source.sheetId, chart.sheetId)
  assert.equal(created.revision, 0); assert.deepEqual(created.position, { x: 88, y: 72 })
  for (const key of ['type', 'title', 'subtitle', 'size', 'palette', 'options', 'legend', 'categoryField', 'valueFields', 'xAxis', 'yAxis']) {
    assert.deepEqual(json(created[key]), original[key], key)
  }
  created.options.titleStyle.manualLayout.left = 0.7; assert.deepEqual(json(chart), original)
  assert.equal(workbook.charts.length, 1); assert.equal(copied.chart.sheetId, chart.sheetId)
})

test('paste rejects changed authors/resources, foreign units, duplicate IDs and mismatched clipboard tokens before mutation', () => {
  const { workbook, chart, destination, target } = fixture('wordCloud'), image = chartResourceFixture()
  chart.options.wordCloud = { maskImageId: image.id }; workbook.metadata = { chartImageResources: { [image.id]: image } }
  const copied = captureChartObjectClipboard(workbook, target()), input = { unitId: workbook.id, sheetId: destination.id, newChartId: 'pasted' }
  const before = JSON.stringify(workbook)
  assert.throws(() => chartObjectPasteDefinition(workbook, copied, { ...input, unitId: 'foreign' }), /跨工作簿/)
  assert.throws(() => chartObjectPasteDefinition(workbook, copied, { ...input, newChartId: chart.id }), /已经存在/)
  assert.throws(() => chartObjectPasteDefinition(workbook, copied, { ...input, clipboardId: 'replaced' }), /剪贴板已更新/)
  assert.throws(() => chartObjectPasteDefinition(workbook, copied, { ...input, position: { x: NaN, y: 0 } }), RangeError)
  assert.equal(JSON.stringify(workbook), before)
  chart.title = 'New author'; assert.throws(() => chartObjectPasteDefinition(workbook, copied, input), /源图表已更改/)
  chart.title = copied.chart.title; workbook.metadata.chartImageResources[image.id] = { ...image, name: 'replaced' }
  assert.throws(() => chartObjectPasteDefinition(workbook, copied, input), /源图片资源已更改/)
  workbook.charts = []; assert.throws(() => chartObjectPasteDefinition(workbook, copied, input), /源图表已更改或删除/)
})

test('copied image references keep the existing immutable resource until the last authored chart is removed', () => {
  const { workbook, chart, destination, target } = fixture('wordCloud'), image = chartResourceFixture()
  chart.options.wordCloud = { maskImageId: image.id }; workbook.metadata = { chartImageResources: { [image.id]: image } }
  const copied = captureChartObjectClipboard(workbook, target())
  workbook.charts.push(chartObjectPasteDefinition(workbook, copied, { unitId: workbook.id, sheetId: destination.id, newChartId: 'pasted' }))
  workbook.charts = workbook.charts.filter(item => item.id !== chart.id); pruneChartImageResources(workbook)
  assert.deepEqual(workbook.metadata.chartImageResources[image.id], image)
  workbook.charts = []; pruneChartImageResources(workbook); assert.equal(workbook.metadata.chartImageResources, undefined)
})

test('both HTML and plain fallbacks carry provenance, escape labels and reject conflicting tokens', () => {
  const { workbook, target } = fixture(), copied = captureChartObjectClipboard(workbook, target()), content = chartObjectClipboardContent(copied)
  assert.equal(chartObjectClipboardIdFromContent({ html: content.html }), copied.id)
  assert.equal(chartObjectClipboardIdFromContent({ text: content.plain }), copied.id)
  assert.ok(content.html.includes('&lt;Main &amp; chart&gt;')); assert.ok(!content.html.includes('<Main'))
  assert.equal(chartObjectClipboardIdFromContent({ html: content.html, text: '[office-chart:another]' }), undefined)
  assert.equal(chartObjectClipboardIdFromContent({ text: '<Main & chart>' }), undefined)
  assert.match(createChartObjectId(), /^[A-Za-z0-9:_-]+$/)
})

test('view and edit permissions apply to the host and source, while the versioned primary chart stays undeletable', () => {
  const { workbook, chart, source, destination } = fixture()
  source.worksheetProtection = protection()
  let state = chartObjectCapabilities(workbook, chart, 'local-user')
  assert.equal(state.edit.disabled, true); assert.equal(state.delete.disabled, true); assert.equal(state.copy.disabled, false); assert.equal(state.exportImage.disabled, false)
  chart.sheetId = destination.id; chart.source.sheetId = source.id; source.worksheetProtection = protection('noOneElseCanView')
  state = chartObjectCapabilities(workbook, chart, 'local-user'); assert.ok(Object.values(state).every(entry => entry.disabled))
  source.worksheetProtection = undefined; destination.metadata = { [CHART_SHEET_METADATA_KEY]: { version: 1, primaryChartId: chart.id } }
  state = chartObjectCapabilities(workbook, chart, 'local-user')
  assert.equal(state.delete.hidden, true); assert.equal(state.delete.disabled, true); assert.equal(state.copy.disabled, false)
})

test('command target aliases, revisions and host disposal protect the originating object and primary fallback', () => {
  const { workbook, source, chart, studio, target } = fixture()
  assert.equal(resolveChartObjectTarget(studio, { unitId: workbook.id, subUnitId: source.id, drawingId: chart.id }).chartId, chart.id)
  assert.throws(() => resolveChartObjectTarget(studio, { chartId: chart.id, drawingId: 'other' }), TypeError)
  assert.throws(() => resolveChartObjectTarget(studio, { chartId: chart.id, sheetId: source.id, subUnitId: 'other' }), TypeError)
  assert.throws(() => resolveChartObjectTarget(studio, { ...target(), expectedRevision: 3 }), /内容已更新/)
  const owned = { ...target(), hostSessionId: 'host-1' }, host = { sessionId: 'host-1', target: () => owned,
    current: value => value.chartId === chart.id, currentSheet: () => true, openPanel: () => true, exportImage: async () => true, writeClipboard: async () => true }
  const registration = registerChartObjectUIHost(studio, host)
  assert.equal(chartObjectUIHost(studio, owned), host); assert.equal(resolveChartObjectTarget(studio).chartId, chart.id)
  registration.dispose(); registration.dispose(); assert.equal(chartObjectUIHost(studio), undefined)
  assert.throws(() => resolveChartObjectTarget(studio, owned), /宿主已切换/)
  source.metadata = { [CHART_SHEET_METADATA_KEY]: { version: 1, primaryChartId: chart.id } }
  assert.equal(resolveChartObjectTarget(studio).chartId, chart.id)
})

test('clipboard tickets refuse out-of-order copies and closing a pending source invalidates its publication', () => {
  const { workbook, target } = fixture(), copied = captureChartObjectClipboard(workbook, target()), store = chartObjectClipboardStore(createPinia())
  try {
    const first = store.beginCopy(workbook.id), second = store.beginCopy(workbook.id)
    assert.equal(store.publish(first, copied, true), false); assert.equal(store.publish(second, copied, false), true)
    copied.chart.title = 'Mutable caller'; assert.notEqual(store.clipboard.chart.title, copied.chart.title)
    const pending = store.beginCopy('pending-source'); store.clearUnit('pending-source')
    assert.equal(store.publish(pending, copied, true), false); assert.equal(store.clipboard, undefined)
  } finally { store.clear(); store.$dispose() }
})

test('object shortcuts respect IME and modifiers and use the same four fixed command IDs as the menu', () => {
  const key = (value, overrides = {}) => ({ key: value, ctrlKey: false, metaKey: false, altKey: false, shiftKey: false, isComposing: false, keyCode: 0, ...overrides })
  assert.equal(chartObjectShortcut(key('c', { ctrlKey: true })), CHART_OBJECT_COMMANDS.copy)
  assert.equal(chartObjectShortcut(key('v', { metaKey: true })), PASTE_COPIED_CHART_COMMAND)
  assert.equal(chartObjectShortcut(key('F2')), CHART_OBJECT_COMMANDS.edit)
  for (const overrides of [{ isComposing: true }, { keyCode: 229 }, { altKey: true }, { shiftKey: true }]) assert.equal(chartObjectShortcut(key('c', { ctrlKey: true, ...overrides })), undefined)
  assert.deepEqual(CHART_OBJECT_MENU.map(entry => entry.order), [0, 1, 3, 4])
})

test('an oversized imported author is rejected before publishing a detached clipboard DTO', () => {
  const { workbook, chart, target } = fixture()
  chart.options.vendorPayload = 'x'.repeat(8_388_609)
  assert.throws(() => captureChartObjectClipboard(workbook, target()), /32 MiB/)
  assert.equal(workbook.charts.length, 1); assert.equal(chart.id, 'original')
})
