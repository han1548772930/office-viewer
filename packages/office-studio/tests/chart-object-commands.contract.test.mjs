import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { chartResourceFixture, chartResourceGoldenTransport, withChartResourcePlatform } from './fixtures/chart-resource-platform-fixture.mjs'

// Deferred facade/command/history drafts. No runner, Vite server or WASM was started in this batch.
let server, api, commands, clipboard
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-object-commands.ts')
  clipboard = await server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-object-clipboard-store.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const store = studio => clipboard.chartObjectClipboardStore(studio.pinia)
const history = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
function deferred() { let resolve; const promise = new Promise(value => { resolve = value }); return { promise, resolve } }
async function setup() {
  const studio = new api.OfficeStudio(), facade = api.FOffice.newAPI(studio), workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:B3').setValues([['Label', 'Value'], ['North', 7], ['South', 4]])
  const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:B3').setValueFields([1]).setTitle('Original').setSubtitle('Sub')
    .setOptions({ titleStyle: { rotate: 30, manualLayout: { left: 0.2, top: 0.3 } }, plotArea: { left: 0.15, top: 0.2, width: 0.7, height: 0.6 } }).build())
  const target = () => ({ unitId: workbook.getId(), sheetId: sheet.getSheetId(), chartId: chart.getChartId(),
    expectedRevision: studio.getWorkbookRevision(workbook.getId()), snapshot: JSON.stringify(studio.getWorkbookModel(workbook.getId()).charts.find(item => item.id === chart.getChartId())) })
  return { studio, facade, workbook, sheet, chart, target }
}

test('focused copy is a non-content command; paste adds a fresh author in one history entry and survives undo/redo/JSON reopening', async () => {
  const { studio, workbook, sheet, chart, target } = await setup()
  let reopened
  try {
    const id = workbook.getId(), before = json(workbook.save()), revision = studio.getWorkbookRevision(id), count = history(studio, id)
    assert.equal(await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, target()), true)
    assert.deepEqual(json(workbook.save()), before); assert.equal(studio.getWorkbookRevision(id), revision); assert.equal(history(studio, id), count)
    const token = store(studio).clipboard.id
    assert.equal(await studio.executeCommand(commands.PASTE_COPIED_CHART_COMMAND,
      { unitId: id, sheetId: sheet.getSheetId(), expectedRevision: revision, clipboardId: token, newChartId: 'pasted-object' }), true)
    const pasted = sheet.getChartById('pasted-object').getInfo(), original = chart.getInfo()
    assert.equal(pasted.id, 'pasted-object'); assert.notEqual(pasted.id, original.id); assert.equal(pasted.source.sheetId, sheet.getSheetId())
    assert.deepEqual(pasted.options, original.options); assert.deepEqual(pasted.size, original.size); assert.equal(pasted.subtitle, original.subtitle)
    assert.equal(studio.getWorkbookRevision(id), revision + 1); assert.equal(history(studio, id), count + 1)
    workbook.undo(); assert.equal(sheet.getChartById('pasted-object'), null); assert.deepEqual(chart.getInfo(), original)
    workbook.redo(); assert.deepEqual(sheet.getChartById('pasted-object').getInfo(), pasted)
    reopened = new api.OfficeStudio({ workbook: json(workbook.save()) })
    const again = api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById('pasted-object')
    assert.deepEqual(again.getInfo(), pasted)
  } finally { reopened?.dispose(); studio.dispose() }
})

test('same-workbook cross-sheet paste keeps live source bindings and rejects unsupported cross-workbook import without losing its clipboard', async () => {
  const { studio, facade, workbook, sheet, chart, target } = await setup()
  try {
    const destination = workbook.insertSheet('Destination')
    await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, target())
    const token = store(studio).clipboard.id
    assert.equal(await studio.executeCommand(commands.PASTE_COPIED_CHART_COMMAND,
      { unitId: workbook.getId(), sheetId: destination.getSheetId(), clipboardId: token, newChartId: 'across-sheet' }), true)
    const pasted = destination.getChartById('across-sheet').getInfo()
    assert.equal(pasted.source.sheetId, sheet.getSheetId()); assert.equal(pasted.sheetId, destination.getSheetId()); assert.equal(pasted.title, chart.getInfo().title)
    const foreign = facade.createWorkbook({ id: 'foreign-object-book' }), before = json(foreign.save()), revision = studio.getWorkbookRevision(foreign.getId())
    await assert.rejects(studio.executeCommand(commands.PASTE_COPIED_CHART_COMMAND,
      { unitId: foreign.getId(), sheetId: foreign.getActiveSheet().getSheetId(), newChartId: 'foreign-copy' }), /跨工作簿|工作表/)
    assert.deepEqual(json(foreign.save()), before); assert.equal(studio.getWorkbookRevision(foreign.getId()), revision); assert.equal(store(studio).clipboard.id, token)
  } finally { studio.dispose() }
})

test('the four focused roots and legacy removal share permission/identity checks and one Before/Executed event per action', async () => {
  const { studio, workbook, chart, target } = await setup(), events = [], panels = [], images = []
  let registration
  try {
    const owned = { ...target(), hostSessionId: 'public-host' }
    registration = commands.registerChartObjectUIHost(studio, { sessionId: 'public-host', target: () => owned,
      current: value => value.unitId === owned.unitId && value.chartId === owned.chartId, currentSheet: () => true,
      openPanel: value => { panels.push(value); return true }, exportImage: async value => { images.push(value); return true }, writeClipboard: async () => true })
    studio.events.addEvent('BeforeCommandExecute', event => events.push(['before', event.id]))
    studio.events.addEvent('CommandExecuted', event => events.push(['executed', event.id]))
    const revision = studio.getWorkbookRevision(workbook.getId())
    assert.equal(studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, owned), true)
    assert.equal(await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.exportImage, owned), true)
    assert.equal(panels[0].chartId, chart.getChartId()); assert.equal(images[0].hostSessionId, 'public-host'); assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
    assert.deepEqual(events.map(entry => entry[1]), [commands.CHART_OBJECT_COMMANDS.edit, commands.CHART_OBJECT_COMMANDS.edit,
      commands.CHART_OBJECT_COMMANDS.exportImage, commands.CHART_OBJECT_COMMANDS.exportImage])
    assert.equal(studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.delete, owned), true)
    assert.equal(workbook.getActiveSheet().getChartById(chart.getChartId()), null); workbook.undo()
    assert.equal(await chart.remove(), true)
  } finally { registration?.dispose(); studio.dispose() }
})

test('invalid/stale target parameters and a command veto leave author, history and clipboard unchanged', async () => {
  const { studio, workbook, chart, target } = await setup()
  let veto
  try {
    const id = workbook.getId(), before = json(workbook.save()), revision = studio.getWorkbookRevision(id), count = history(studio, id)
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, { ...target(), drawingId: 'other' }), TypeError)
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.delete, { ...target(), expectedRevision: revision - 1 }), /内容已更新/)
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.delete, { ...target(), snapshot: '{}' }), /内容已更改/)
    veto = studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commands.CHART_OBJECT_COMMANDS.copy) event.cancel = true })
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, target()), /canceled/)
    assert.equal(store(studio).clipboard, undefined); assert.deepEqual(json(workbook.save()), before); assert.equal(history(studio, id), count); assert.equal(chart.getInfo().title, 'Original')
  } finally { veto?.dispose(); studio.dispose() }
})

test('asynchronous roots refuse sync execution before dispatch; edit protection permits copying and unviewable sources deny it', async () => {
  const { studio, workbook, sheet, target } = await setup(), events = []
  try {
    studio.events.addEvent('BeforeCommandExecute', event => events.push(event.id))
    for (const id of [commands.CHART_OBJECT_COMMANDS.copy, commands.CHART_OBJECT_COMMANDS.exportImage, commands.PASTE_COPIED_CHART_COMMAND]) {
      assert.throws(() => studio.executeCommandSync(id, target()), /异步/)
    }
    assert.deepEqual(events, [])
    const model = studio.getWorkbookModel(workbook.getId()), source = model.sheets.find(item => item.id === sheet.getSheetId())
    source.worksheetProtection = { id: 'guard', permissionId: 'guard-permission', unitType: 'worksheet', ownerUserId: 'other',
      allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
    const revision = studio.getWorkbookRevision(workbook.getId()), count = history(studio, workbook.getId())
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.delete, target()), /保护|编辑/)
    assert.equal(await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, target()), true)
    source.worksheetProtection.viewState = 'noOneElseCanView'
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, target()), /查看|权限/)
    await assert.rejects(studio.executeCommand(commands.CHART_OBJECT_COMMANDS.exportImage, target()), /查看|权限/)
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision); assert.equal(history(studio, workbook.getId()), count)
  } finally { studio.dispose() }
})

test('the chart-sheet primary fallback can be copied but both focused and facade deletion refuse it', async () => {
  const { studio, workbook, sheet, chart, target } = await setup()
  try {
    studio.getWorkbookModel(workbook.getId()).sheets.find(item => item.id === sheet.getSheetId()).metadata = {
      [commands.CHART_SHEET_METADATA_KEY]: { version: 1, primaryChartId: chart.getChartId() },
    }
    const before = json(workbook.save()), revision = studio.getWorkbookRevision(workbook.getId())
    assert.equal(await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, { unitId: workbook.getId(), subUnitId: sheet.getSheetId() }), true)
    assert.equal(await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.delete, target()), false)
    assert.equal(await chart.remove(), false)
    assert.deepEqual(json(workbook.save()), before); assert.equal(studio.getWorkbookRevision(workbook.getId()), revision)
  } finally { studio.dispose() }
})

test('a valid upstream section route reaches the host; invalid combinations leave navigation, author and history unchanged', async () => {
  const { studio, workbook, target } = await setup(), opened = []
  let registration
  try {
    const owned = { ...target(), hostSessionId: 'route-host' }, revision = studio.getWorkbookRevision(workbook.getId()), count = history(studio, workbook.getId()), before = json(workbook.save())
    registration = commands.registerChartObjectUIHost(studio, { sessionId: 'route-host', target: () => owned, current: () => true, currentSheet: () => true,
      openPanel: (value, route) => { opened.push({ value, route }); return true }, exportImage: async () => true, writeClipboard: async () => true })
    assert.equal(studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...owned, route: { tab: 'style', section: 'legend' } }), true)
    assert.deepEqual(opened[0].route, { tab: 'style', section: 'legend' }); assert.equal(opened[0].value.chartId, owned.chartId)
    assert.equal(studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...owned, route: { tab: 'style', axisId: 'yAxis' } }), true)
    assert.deepEqual(opened[1].route, { tab: 'style', section: 'verticalAxis', axisId: 'yAxis' })
    assert.throws(() => studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...owned, route: { tab: 'data', section: 'legend' } }), /数据标签页/)
    assert.throws(() => studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...owned, route: { tab: 'style', section: 'pieStyle' } }), /当前图表类型/)
    assert.equal(opened.length, 2)
    assert.equal(studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, owned), true); assert.equal(opened.length, 3); assert.equal(opened[2].route, undefined)
    assert.equal(studio.getWorkbookRevision(workbook.getId()), revision); assert.equal(history(studio, workbook.getId()), count); assert.deepEqual(json(workbook.save()), before)
  } finally { registration?.dispose(); studio.dispose() }
})

test('a late system copy after host disposal or workbook closure cannot publish its clipboard snapshot', async () => {
  for (const closeUnit of [false, true]) {
    const { studio, workbook, target } = await setup(), pending = deferred()
    let registration
    try {
      const owned = { ...target(), hostSessionId: 'late-copy' }
      registration = commands.registerChartObjectUIHost(studio, { sessionId: 'late-copy', target: () => owned,
        current: () => true, currentSheet: () => true, openPanel: () => true, exportImage: async () => true, writeClipboard: () => pending.promise })
      const copying = studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, owned)
      const rejected = assert.rejects(copying, /宿主|关闭|工作簿/)
      if (closeUnit) studio.disposeUnit(workbook.getId()); else registration.dispose()
      pending.resolve(true); await rejected; assert.equal(store(studio).clipboard, undefined)
    } finally { registration?.dispose(); studio.dispose() }
  }
})

test('copy/paste/delete retain image ownership through one author/history pipeline and JSON reopening', async () => {
  await withChartResourcePlatform(async () => {
    const compute = chartResourceGoldenTransport(), studio = new api.OfficeStudio({ computeClient: compute }), facade = api.FOffice.newAPI(studio)
    const workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet(), image = chartResourceFixture()
    let reopened
    try {
      sheet.getRange('A1:B2').setValues([['Word', 'Weight'], ['A', 1]])
      const original = await sheet.insertChart(sheet.newChart('wordCloud').setSource('A1:B2').setValueFields([1]).setSize(560, 400).setMaskImageResource(image).build())
      const target = { unitId: workbook.getId(), sheetId: sheet.getSheetId(), chartId: original.getChartId() }
      await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.copy, target)
      const count = history(studio, workbook.getId())
      await studio.executeCommand(commands.PASTE_COPIED_CHART_COMMAND, { unitId: workbook.getId(), sheetId: sheet.getSheetId(), newChartId: 'word-copy' })
      assert.equal(history(studio, workbook.getId()), count + 1)
      assert.equal(sheet.getChartById('word-copy').getOptions().wordCloud.maskImageId, image.id)
      const saved = json(workbook.save()); assert.deepEqual(saved.metadata.chartImageResources[image.id], image)
      reopened = new api.OfficeStudio({ workbook: saved, computeClient: chartResourceGoldenTransport() })
      assert.equal(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById('word-copy').getOptions().wordCloud.maskImageId, image.id)
      await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.delete, target)
      assert.deepEqual(workbook.save().metadata.chartImageResources[image.id], image)
      await studio.executeCommand(commands.CHART_OBJECT_COMMANDS.delete, { ...target, chartId: 'word-copy' })
      assert.equal(workbook.save().metadata.chartImageResources, undefined)
      workbook.undo(); assert.deepEqual(workbook.save().metadata.chartImageResources[image.id], image)
      workbook.redo(); assert.equal(workbook.save().metadata.chartImageResources, undefined)
    } finally { reopened?.dispose(); studio.dispose() }
  })
})
