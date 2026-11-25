import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

// P18 Facade/command/history drafts. Validation remains deferred.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const history = (studio, id) => studio.editor.history.units[id]?.undo.length ?? 0
async function fixture(type = 'line', options = {}) {
  const studio = new api.OfficeStudio(), facade = api.FOffice.newAPI(studio), workbook = facade.getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues([['Label', 'Value', 'Comparison'], ['A', 9, 4], ['B', 2, 8], ['C', 8, 6]])
  const chart = await sheet.insertChart(sheet.newChart(type).setSource('A1:C4').setValueFields([1, 2])
    .setOptions({ autoGradientFill: false, allSeriesStyle: { point: { shape: 'none' } }, ...options }).build())
  return { studio, facade, workbook, sheet, chart }
}

test('FOffice enum and Builder shape/empty-value methods validate atomically and detach authored inputs', async () => {
  const f = await fixture()
  try {
    assert.deepEqual(f.facade.Enum.ChartAreaLineStyle, { Line: 'line', Smooth: 'smooth', Step: 'step' })
    assert.deepEqual(f.facade.Enum.ChartInvalidValueType, { Zero: 'zero', Break: 'break', Link: 'link' })
    const input = { area: { smooth: true, future: { retained: 4 } }, appearance: { background: 'navy' } }
    const builder = f.sheet.newChart('area').setSource('A1:C4').setValueFields([1]).setOptions(input)
    input.area.future.retained = 99
    assert.equal(builder.setAreaLineStyle(f.facade.Enum.ChartAreaLineStyle.Step), builder)
    assert.equal(builder.setInvalidValueStrategy(f.facade.Enum.ChartInvalidValueType.Link), builder)
    const valid = builder.build()
    assert.equal(valid.options.area.lineStyle, 'step'); assert.equal(valid.options.area.smooth, undefined)
    assert.equal(valid.options.area.future.retained, 4); assert.equal(valid.options.invalidValueStrategy, 'link')
    for (const value of [undefined, null, 1, {}, 'spline']) assert.throws(() => builder.setAreaLineStyle(value), TypeError)
    for (const value of [null, false, 'connect']) assert.throws(() => builder.setInvalidValueStrategy(value), TypeError)
    assert.deepEqual(builder.build(), valid)
    assert.equal(builder.clearAreaLineStyle(), builder); assert.equal(builder.resetInvalidValueStrategy(), builder)
    const reset = builder.build()
    assert.equal(reset.options.area.lineStyle, undefined); assert.equal(reset.options.area.smooth, undefined)
    assert.equal(reset.options.invalidValueStrategy, undefined); assert.equal(reset.options.area.future.retained, 4)
    assert.throws(() => f.sheet.newChart('scatter').setAreaLineStyle('smooth'), TypeError)
  } finally { f.studio.dispose() }
})

test('each FChart setter uses one real update command/history entry and leaves source values and prepared data intact', async () => {
  const f = await fixture(), executed = []
  let listener
  try {
    const id = f.workbook.getId(), count = history(f.studio, id), revision = f.studio.getWorkbookRevision(id), before = f.chart.getInfo()
    const cells = json(f.studio.getWorkbookModel(id).sheets.find(item => item.id === f.sheet.getSheetId()).cells), data = await f.chart.getComputedData()
    listener = f.studio.events.addEvent('CommandExecuted', event => { if (event.id === 'sheet.command.update-chart') executed.push(event) })
    assert.equal(await f.chart.setAreaLineStyle('smooth'), f.chart)
    assert.equal(f.chart.getAreaLineStyle(), 'smooth'); assert.equal(history(f.studio, id), count + 1)
    assert.equal(f.studio.getWorkbookRevision(id), revision + 1); assert.equal(executed.length, 1)
    const smooth = f.chart.getInfo(), scene = await f.chart.getComputedRenderCommands()
    assert.ok(scene.some(item => item.kind === 'line'))
    assert.deepEqual(await f.chart.getComputedData(), data); assert.deepEqual(smooth.source, before.source)
    f.workbook.undo(); assert.deepEqual(f.chart.getInfo(), before)
    f.workbook.redo(); assert.deepEqual(f.chart.getInfo(), smooth)
    const nextCount = history(f.studio, id), eventCount = executed.length
    assert.equal(await f.chart.setInvalidValueStrategy('link'), f.chart)
    assert.equal(f.chart.getInvalidValueStrategy(), 'link'); assert.equal(history(f.studio, id), nextCount + 1)
    assert.equal(executed.length, eventCount + 1)
    assert.deepEqual(f.studio.getWorkbookModel(id).sheets.find(item => item.id === f.sheet.getSheetId()).cells, cells)
  } finally { listener?.dispose(); f.studio.dispose() }
})

test('JSON save/reopen retains line shape and empty-value authors, and resets cannot resurrect old area smooth', async () => {
  const f = await fixture('area3D', { area: { smooth: true, future: 'retained' } })
  let reopened
  try {
    assert.equal(f.chart.getAreaLineStyle(), 'smooth'); assert.equal(f.chart.getInvalidValueStrategy(), 'break')
    await f.chart.setAreaLineStyle('step'); await f.chart.setInvalidValueStrategy('zero')
    const data = await f.chart.getComputedData(), scene = await f.chart.getComputedRenderCommands(), authored = f.chart.getInfo()
    reopened = new api.OfficeStudio({ workbook: json(f.workbook.save()) })
    const book = api.FOffice.newAPI(reopened).getActiveWorkbook(), chart = book.getActiveSheet().getChartById(f.chart.getChartId())
    assert.deepEqual(chart.getInfo(), authored); assert.equal(chart.getAreaLineStyle(), 'step'); assert.equal(chart.getInvalidValueStrategy(), 'zero')
    assert.deepEqual(await chart.getComputedData(), data); assert.deepEqual(await chart.getComputedRenderCommands(), scene)
    const count = history(reopened, book.getId())
    await chart.clearAreaLineStyle(); await chart.resetInvalidValueStrategy()
    assert.equal(history(reopened, book.getId()), count + 2)
    assert.equal(chart.getAreaLineStyle(), 'line'); assert.equal(chart.getInvalidValueStrategy(), 'break')
    assert.equal(chart.getOptions().area.smooth, undefined); assert.equal(chart.getOptions().area.future, 'retained')
    book.undo(); assert.equal(chart.getInvalidValueStrategy(), 'zero')
    book.undo(); assert.equal(chart.getAreaLineStyle(), 'step')
    book.redo(); book.redo()
    assert.equal(json(book.save()).charts[0].options.area.smooth, undefined)
    assert.equal(json(book.save()).charts[0].options.area.lineStyle, undefined)
  } finally { reopened?.dispose(); f.studio.dispose() }
})

test('permissions, BeforeCommandExecute veto and invalid enums prevent shape authorship without changing history', async () => {
  const f = await fixture(), executed = []
  let veto, listener
  try {
    const id = f.workbook.getId(), sheet = f.studio.getWorkbookModel(id).sheets.find(item => item.id === f.sheet.getSheetId())
    sheet.worksheetProtection = { id: 'shape-protection', permissionId: 'shape-permission', unitType: 'worksheet',
      ownerUserId: 'other-user', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
    const count = history(f.studio, id), before = json(f.workbook.save())
    await assert.rejects(() => f.chart.setAreaLineStyle('step'), /保护|编辑/)
    await assert.rejects(() => f.chart.setInvalidValueStrategy('zero'), /保护|编辑/)
    assert.deepEqual(json(f.workbook.save()), before); assert.equal(history(f.studio, id), count)
    sheet.worksheetProtection.viewState = 'noOneElseCanView'
    await assert.rejects(() => f.chart.getComputedData(), /查看|权限/)
    delete sheet.worksheetProtection
    const allowed = json(f.workbook.save()), revision = f.studio.getWorkbookRevision(id)
    listener = f.studio.events.addEvent('CommandExecuted', event => { if (event.id === 'sheet.command.update-chart') executed.push(event) })
    veto = f.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === 'sheet.command.update-chart') event.cancel = true })
    await assert.rejects(() => f.chart.setAreaLineStyle('smooth'), /canceled/)
    assert.deepEqual(json(f.workbook.save()), allowed); assert.deepEqual(executed, [])
    assert.equal(history(f.studio, id), count); assert.equal(f.studio.getWorkbookRevision(id), revision)
    veto.dispose(); veto = undefined
    for (const value of [undefined, null, false, 'curved', {}]) await assert.rejects(() => f.chart.setAreaLineStyle(value), TypeError)
    for (const value of [null, 'connect', 2]) await assert.rejects(() => f.chart.setInvalidValueStrategy(value), TypeError)
    assert.deepEqual(json(f.workbook.save()), allowed); assert.deepEqual(executed, [])
  } finally { veto?.dispose(); listener?.dispose(); f.studio.dispose() }
})

test('chart-type switches retain inactive configuration and return to the same shape; dedicated unsupported setters reject', async () => {
  const f = await fixture()
  try {
    await f.chart.setAreaLineStyle('step'); await f.chart.setInvalidValueStrategy('link')
    const source = f.chart.getDataSource(), data = await f.chart.getComputedData()
    await f.chart.setType('column')
    assert.equal(f.chart.getAreaLineStyle(), 'line'); assert.equal(f.chart.getOptions().area.lineStyle, 'step')
    const before = f.chart.getInfo(), count = history(f.studio, f.workbook.getId())
    await assert.rejects(() => f.chart.setAreaLineStyle('smooth'), TypeError)
    await assert.rejects(() => f.chart.clearAreaLineStyle(), TypeError)
    assert.deepEqual(f.chart.getInfo(), before); assert.equal(history(f.studio, f.workbook.getId()), count)
    await f.chart.setType('line')
    assert.equal(f.chart.getAreaLineStyle(), 'step'); assert.equal(f.chart.getInvalidValueStrategy(), 'link')
    assert.deepEqual(f.chart.getDataSource(), source); assert.deepEqual(await f.chart.getComputedData(), data)
    f.workbook.undo(); assert.equal(f.chart.getType(), 'column')
    f.workbook.redo(); assert.equal(f.chart.getAreaLineStyle(), 'step')
  } finally { f.studio.dispose() }
})

test('shape and empty-value histories remain isolated by workbook and a closed workbook cannot receive another update', async () => {
  const f = await fixture()
  try {
    const other = f.facade.createWorkbook({ id: 'independent-shape-book' }), sheet = other.getActiveSheet()
    sheet.getRange('A1:B4').setValues([['Label', 'Value'], ['A', 4], ['B', 8], ['C', 2]])
    const chart = await sheet.insertChart(sheet.newChart('area').setSource('A1:B4').setValueFields([1]).build())
    const firstCount = history(f.studio, f.workbook.getId()), otherCount = history(f.studio, other.getId())
    await f.chart.setAreaLineStyle('smooth'); await chart.setAreaLineStyle('step'); await chart.setInvalidValueStrategy('link')
    assert.equal(history(f.studio, f.workbook.getId()), firstCount + 1); assert.equal(history(f.studio, other.getId()), otherCount + 2)
    f.workbook.undo(); assert.equal(f.chart.getAreaLineStyle(), 'line'); assert.equal(chart.getAreaLineStyle(), 'step')
    other.undo(); assert.equal(chart.getInvalidValueStrategy(), 'break'); assert.equal(chart.getAreaLineStyle(), 'step')
    other.redo(); f.workbook.redo()
    const surviving = json(other.save())
    f.studio.disposeUnit(f.workbook.getId())
    await assert.rejects(() => f.chart.setAreaLineStyle('line'))
    assert.deepEqual(json(other.save()), surviving); assert.equal(chart.getInvalidValueStrategy(), 'link')
  } finally { f.studio.dispose() }
})
