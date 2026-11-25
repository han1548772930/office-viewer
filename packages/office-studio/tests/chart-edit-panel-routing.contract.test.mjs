import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P17 command/history drafts only; nothing in this file was executed this batch.
let server, api, commands
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, commands] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/charts/chart-object-commands.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const history = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
async function fixture() {
  const studio = new api.OfficeStudio(), facade = api.FOffice.newAPI(studio), workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C3').setValues([['Label', 'Value', 'Comparison'], ['North', 7, 3], ['South', 4, 9]])
  const chart = await sheet.insertChart(sheet.newChart('line').setSource('A1:C3').setValueFields([1, 2]).setTitle('Original').setSubtitle('Sub')
    .setOptions({ series: { 2: { rightAxis: true, name: 'Displayed comparison' } }, rightYAxis: { title: 'Secondary', scaleType: 'linear' } }).build())
  const target = () => {
    const model = studio.getWorkbookModel(workbook.getId()), value = model?.charts.find(item => item.id === chart.getChartId())
    return value && { unitId: model.id, sheetId: value.sheetId, chartId: value.id, expectedRevision: studio.getWorkbookRevision(model.id), snapshot: JSON.stringify(value) }
  }
  const opened = [], registrations = []
  function host(sessionId) {
    const registration = commands.registerChartObjectUIHost(studio, { sessionId, target: () => target() && { ...target(), hostSessionId: sessionId },
      current: value => { const active = target(); return !!active && value.unitId === active.unitId && value.sheetId === active.sheetId
        && value.chartId === active.chartId && value.expectedRevision === active.expectedRevision && value.snapshot === active.snapshot },
      currentSheet: (unit, subUnit) => unit === workbook.getId() && subUnit === sheet.getSheetId(),
      openPanel: (value, route) => { opened.push({ target: value, route }); return true }, exportImage: async () => true, writeClipboard: async () => true })
    registrations.push(registration); return registration
  }
  return { studio, workbook, sheet, chart, target, opened, host, dispose() { registrations.forEach(value => value.dispose()); studio.dispose() } }
}

test('a series route requires prepared current source identities and rejects a cold cache instead of selecting the first series', async () => {
  const f = await fixture()
  try {
    f.host('series-route'); f.sheet.getRange('A2').setValues([['North revised']])
    const before = json(f.workbook.save()), count = history(f.studio, f.workbook.getId()), revision = f.studio.getWorkbookRevision(f.workbook.getId())
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'style', seriesId: 'field:2' } }), /尚未就绪/)
    assert.deepEqual(f.opened, []); assert.deepEqual(json(f.workbook.save()), before)
    await f.chart.getComputedData()
    assert.equal(f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'style', seriesId: 'field:2' } }), true)
    assert.deepEqual(f.opened[0].route, { tab: 'style', section: 'series', seriesId: 'field:2' })
    assert.equal(f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'style', seriesId: 'name:"Comparison"' } }), true)
    assert.equal(f.opened[1].route.seriesId, 'name:"Comparison"')
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'style', seriesId: 'name:"Displayed comparison"' } }), /不存在/)
    assert.equal(history(f.studio, f.workbook.getId()), count); assert.equal(f.studio.getWorkbookRevision(f.workbook.getId()), revision)
  } finally { f.dispose() }
})

test('routed editing shares worksheet/source read permission checks and command veto before a host can open', async () => {
  const f = await fixture(), executed = []
  let veto
  try {
    f.host('permission-route'); f.studio.events.addEvent('CommandExecuted', event => { if (event.id === commands.CHART_OBJECT_COMMANDS.edit) executed.push(event) })
    const model = f.studio.getWorkbookModel(f.workbook.getId()), sheet = model.sheets.find(item => item.id === f.sheet.getSheetId())
    sheet.worksheetProtection = { id: 'route-guard', permissionId: 'route-permission', unitType: 'worksheet', ownerUserId: 'someone-else',
      allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'data' } }), /保护|编辑/)
    sheet.worksheetProtection.viewState = 'noOneElseCanView'
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'style', section: 'legend' } }), /查看|权限/)
    delete sheet.worksheetProtection
    const before = json(f.workbook.save()), revision = f.studio.getWorkbookRevision(f.workbook.getId()), count = history(f.studio, f.workbook.getId())
    veto = f.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commands.CHART_OBJECT_COMMANDS.edit) event.cancel = true })
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'style', titleKey: 'subtitle' } }), /canceled/)
    assert.deepEqual(f.opened, []); assert.deepEqual(executed, []); assert.deepEqual(json(f.workbook.save()), before)
    assert.equal(f.studio.getWorkbookRevision(f.workbook.getId()), revision); assert.equal(history(f.studio, f.workbook.getId()), count)
  } finally { veto?.dispose(); f.dispose() }
})

test('stale revision, author snapshot and disposed host-session routes cannot enter a replacement editor', async () => {
  const f = await fixture()
  try {
    const old = f.host('old-route'), captured = { ...f.target(), hostSessionId: 'old-route' }
    await f.chart.setTitle('External update')
    const before = json(f.workbook.save()), count = history(f.studio, f.workbook.getId()), route = { tab: 'style', titleKey: 'title' }
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...captured, route }), /内容已更新/)
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), snapshot: captured.snapshot, route }), /内容已更改/)
    old.dispose(); f.host('new-route')
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), hostSessionId: 'old-route', route }), /宿主已切换/)
    assert.deepEqual(f.opened, []); assert.deepEqual(json(f.workbook.save()), before); assert.equal(history(f.studio, f.workbook.getId()), count)
    assert.equal(f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), hostSessionId: 'new-route', route }), true)
    assert.equal(f.opened[0].target.hostSessionId, 'new-route')
  } finally { f.dispose() }
})

test('navigation stays outside content history while the real author update uses one undo/redo/JSON pipeline', async () => {
  const f = await fixture(), events = []
  let reopened
  try {
    f.host('history-route'); const id = f.workbook.getId(), before = json(f.workbook.save()), count = history(f.studio, id), revision = f.studio.getWorkbookRevision(id)
    f.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === commands.CHART_OBJECT_COMMANDS.edit) events.push('before') })
    f.studio.events.addEvent('CommandExecuted', event => { if (event.id === commands.CHART_OBJECT_COMMANDS.edit) events.push('executed') })
    for (const route of [{ tab: 'data' }, { tab: 'style', titleKey: 'rightYAxisTitle' }, { tab: 'style', section: 'lineAndArea' }]) {
      assert.equal(f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route }), true)
    }
    assert.deepEqual(events, ['before', 'executed', 'before', 'executed', 'before', 'executed'])
    assert.deepEqual(json(f.workbook.save()), before); assert.equal(history(f.studio, id), count); assert.equal(f.studio.getWorkbookRevision(id), revision)
    await f.chart.update({ subtitle: 'Authored subtitle', options: { rightYAxis: { title: 'Authored secondary', scaleType: 'linear' } } })
    const authored = f.chart.getInfo(); assert.equal(history(f.studio, id), count + 1)
    f.workbook.undo(); assert.equal(f.chart.getInfo().subtitle, 'Sub'); f.workbook.redo(); assert.deepEqual(f.chart.getInfo(), authored)
    const saved = json(f.workbook.save()); assert.equal(saved.charts.find(item => item.id === authored.id).route, undefined)
    reopened = new api.OfficeStudio({ workbook: saved })
    assert.deepEqual(api.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getChartById(authored.id).getInfo(), authored)
  } finally { reopened?.dispose(); f.dispose() }
})

test('closed workbooks and unavailable focused hosts fail before route navigation and leave the remaining workbook untouched', async () => {
  const f = await fixture()
  try {
    const registration = f.host('closed-route'), captured = { ...f.target(), hostSessionId: 'closed-route' }
    registration.dispose()
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...f.target(), route: { tab: 'data' } }), /宿主尚未绑定/)
    f.host('another-route'); const other = api.FOffice.newAPI(f.studio).createWorkbook({ id: 'surviving-route-unit' }), before = json(other.save())
    f.studio.disposeUnit(f.workbook.getId())
    assert.throws(() => f.studio.executeCommandSync(commands.CHART_OBJECT_COMMANDS.edit, { ...captured, route: { tab: 'style', section: 'legend' } }), /关闭/)
    assert.deepEqual(json(other.save()), before); assert.deepEqual(f.opened, [])
  } finally { f.dispose() }
})
