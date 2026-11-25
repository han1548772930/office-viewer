import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { SUNBURST_TABLE } from './fixtures/chart-sunburst-fixtures.mjs'
import { chartHierarchyGoldenTransport as transport } from './fixtures/chart-hierarchy-facade-transport.mjs'

// P27 Facade/service drafts. No server, compiler, test, WASM or browser ran.
let server, api
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  api = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
function base(studio) {
  const workbook = api.FOffice.newAPI(studio).getActiveWorkbook(), sheet = workbook.getActiveSheet()
  sheet.getRange('A1:C8').setValues(SUNBURST_TABLE)
  return { workbook, sheet }
}
function newBuilder(sheet) { return sheet.newChart('sunburst').setSource('A1:C8').setHierarchyFields([0, 1]).setValueField(2).setAbsolutePosition(20, 30).setSize(560, 400) }
async function setup() {
  const compute = transport(), studio = new api.OfficeStudio({ computeClient: compute }), { workbook, sheet } = base(studio)
  const chart = await sheet.insertChart(newBuilder(sheet).build())
  return { compute, studio, workbook, sheet, chart }
}

test('all six Sunburst Builder methods chain, expose the instance Enum and detach nullable direct font/background inputs', () => {
  const studio = new api.OfficeStudio(), { sheet } = base(studio)
  try {
    assert.deepEqual(api.FOffice.newAPI(studio).Enum.ChartSunburstLabelPosition, { Inside: 'inside', Outside: 'outside' })
    const fields = [0, 1], background = { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }
    const builder = newBuilder(sheet)
    assert.equal(builder.setHierarchyFields(fields), builder); assert.equal(builder.setValueField(2), builder)
    assert.equal(builder.setSunburstLabel({ visible: true, contentType: 10, position: 'outside', backgroundColor: background }), builder)
    assert.equal(builder.setCenterLabel({ visible: true, family: 'Center Font', fontSize: 21, bold: true, format: '0.0' }), builder)
    const before = builder.build(); fields[0] = 2; background.stops[0].color = 'lime'; assert.deepEqual(builder.build(), before)
    assert.equal(builder.resetSunburstLabel(), builder); assert.equal(builder.build().options.sunburst.label, undefined)
    assert.equal(builder.build().options.sunburst.centerLabel.fontSize, 21)
    builder.setSunburstLabel({ position: null, fontSize: null, contentType: 0 })
    assert.equal(builder.build().options.sunburst.label.position, undefined); assert.equal(builder.build().options.sunburst.label.contentType, 0)
    assert.equal(builder.resetCenterLabel(), builder); assert.equal(builder.build().options.sunburst.centerLabel, undefined)
    assert.deepEqual(builder.build().options.mapping.hierarchyIndexes, [0, 1]); assert.deepEqual(builder.build().valueFields, [2])
    const current = builder.build()
    for (const operation of [() => builder.setHierarchyFields([]), () => builder.setHierarchyFields([0, 0]), () => builder.setHierarchyFields(Array(2)),
      () => builder.setValueField(-1), () => builder.setSunburstLabel({ position: 'left' }), () => builder.setSunburstLabel({ backgroundColor: { type: 'unknown' } }),
      () => builder.setCenterLabel({ font: { fontSize: 12 } }), () => builder.setCenterLabel({ fontSize: 0 })]) {
      assert.throws(operation); assert.deepEqual(builder.build(), current)
    }
    for (const method of ['setHierarchyFields', 'setValueField', 'setSunburstLabel', 'resetSunburstLabel', 'setCenterLabel', 'resetCenterLabel']) {
      const args = method === 'setHierarchyFields' ? [[0, 1]] : method === 'setValueField' ? [2] : method.startsWith('set') ? [{}] : []
      assert.throws(() => sheet.newChart('line')[method](...args), TypeError)
    }
  } finally { studio.dispose() }
})

test('live Sunburst options use one-command updates, signed layouts, independent resets, detached results and per-document internal reopen', async () => {
  const { compute, studio, workbook, chart } = await setup()
  let reopened
  try {
    const raw = await chart.getComputedData(), layout = await chart.getComputedLayout(), original = chart.getOptions()
    assert.deepEqual(raw.hierarchy.map(node => [node.name, node.value, node.selfValue, node.paletteIndex]), [['North', 13, -5, 1], ['South', 9, 3, 0]])
    assert.deepEqual(layout.sunburst.filter(node => node.depth === 0).map(node => node.value), [13, 9])
    assert.equal(layout.sunburst.filter(node => node.isSunburstResidual).length, 1)
    assert.equal((await chart.setHierarchyFields([0, 1])), chart); assert.equal((await chart.setValueField(2)), chart)
    const beforeApply = json(chart.getInfo()), beforeRevision = beforeApply.revision
    const beforeHistory = studio.editor.history.captureSavepoint(workbook.getId())
    await chart.update({ options: { sunburst: { label: { visible: false, position: 'outside' }, centerLabel: { visible: true, format: '0.00', fontSize: 22 } } } })
    assert.equal(chart.getInfo().revision, beforeRevision + 1)
    const afterApply = json(chart.getInfo()), afterHistory = studio.editor.history.captureSavepoint(workbook.getId())
    const commands = await chart.getComputedRenderCommands()
    assert.ok(commands.some(command => command.kind === 'text' && command.text === '22.00'))
    assert.equal(commands.some(command => command.kind === 'text' && command.text === 'North'), false)
    workbook.undo(); assert.deepEqual(json(chart.getInfo()), beforeApply)
    assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), beforeHistory)
    workbook.redo(); assert.deepEqual(json(chart.getInfo()), afterApply)
    assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), afterHistory)
    const gradient = { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] }
    await chart.setSunburstLabel({ visible: true, backgroundColor: gradient }); gradient.stops[0].color = 'lime'
    assert.equal(chart.getOptions().sunburst.label.backgroundColor.stops[0].color, 'red')
    const pattern = { type: 'pattern', preset: 'pct25', foregroundColor: 'black', backgroundColor: 'white' }
    await chart.setSunburstLabel({ backgroundColor: pattern }); pattern.foregroundColor = 'red'
    assert.equal(chart.getOptions().sunburst.label.backgroundColor.foregroundColor, 'black')
    assert.equal(Object.hasOwn(chart.getOptions().sunburst.label.backgroundColor, 'stops'), false)
    assert.equal((await chart.resetCenterLabel()), chart); assert.equal(chart.getOptions().sunburst.centerLabel, undefined)
    assert.equal(chart.getOptions().sunburst.label.visible, true)
    assert.equal((await chart.setCenterLabel({ visible: true, family: 'Direct Font', fontSize: 24 })), chart)
    const beforeReset = chart.getOptions(); assert.equal((await chart.resetSunburstLabel()), chart)
    assert.equal(chart.getOptions().sunburst.label, undefined); assert.equal(chart.getOptions().sunburst.centerLabel.fontSize, 24)
    workbook.undo(); assert.deepEqual(chart.getOptions(), beforeReset); workbook.redo(); assert.equal(chart.getOptions().sunburst.label, undefined)
    const saved = workbook.save(), detached = await chart.getComputedData()
    detached.hierarchy[0].selfValue = 100; assert.equal((await chart.getComputedData()).hierarchy[0].selfValue, -5)
    reopened = api.FOffice.newAPI(studio).createWorkbook({ ...saved, id: 'sunburst-reopened-draft' }, { makeCurrent: false })
    const copy = reopened.getActiveSheet().getCharts()[0]
    assert.deepEqual(copy.getOptions(), chart.getOptions()); assert.deepEqual(json(await copy.getComputedData()), json(raw))
    assert.ok((await copy.getComputedRenderCommands()).some(command => command.kind === 'text' && command.text === '22'))
    assert.equal(original.sunburst, undefined); assert.ok(compute.calls.length > 0)
    assert.ok(compute.calls.some(call => call.options.documentId !== compute.calls[0].options.documentId))
  } finally { reopened?.dispose(); studio.dispose() }
})

test('invalid live Sunburst inputs, intersecting or out-of-range mappings and inactive type APIs cannot publish content/history', async () => {
  const { studio, workbook, chart } = await setup()
  try {
    const before = JSON.stringify(workbook.save()), history = studio.editor.history.captureSavepoint(workbook.getId())
    for (const operation of [() => chart.setHierarchyFields([]), () => chart.setHierarchyFields([0, 0]), () => chart.setHierarchyFields([0, 2]),
      () => chart.setHierarchyFields([0, 3]), () => chart.setValueField(0), () => chart.setValueField(3),
      () => chart.setSunburstLabel({ position: 'left' }), () => chart.setSunburstLabel({ fontSize: 0 }),
      () => chart.setCenterLabel({ visible: 'yes' }), () => chart.setCenterLabel({ font: { bold: true } })]) {
      await assert.rejects(async () => { await operation() }); assert.equal(JSON.stringify(workbook.save()), before)
      assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), history)
    }
    await chart.update({ type: 'line', valueFields: [2], options: { mapping: undefined } })
    const inactive = JSON.stringify(workbook.save()), inactiveHistory = studio.editor.history.captureSavepoint(workbook.getId())
    for (const operation of [() => chart.setHierarchyFields([0, 1]), () => chart.setValueField(2), () => chart.setSunburstLabel({ visible: false }),
      () => chart.resetSunburstLabel(), () => chart.setCenterLabel({ visible: true }), () => chart.resetCenterLabel()]) await assert.rejects(async () => { await operation() }, TypeError)
    assert.equal(JSON.stringify(workbook.save()), inactive); assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), inactiveHistory)
  } finally { studio.dispose() }
})

test('worksheet edit protection blocks all six Sunburst operations and direct chart commands before Worker allocation', async () => {
  const { studio, workbook, sheet, chart } = await setup()
  let visitor
  try {
    studio.executeCommandSync('sheet.command.add-worksheet-protection', { unitId: workbook.getId(), sheetId: sheet.getSheetId(),
      rule: { id: 'sunburst-protection', permissionId: 'sunburst-permission', unitType: 'worksheet', ownerUserId: 'owner',
        editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [] } })
    const compute = transport(); visitor = new api.OfficeStudio({ workbook: json(workbook.save()), currentUserId: 'visitor', computeClient: compute })
    const targetWorkbook = api.FOffice.newAPI(visitor).getActiveWorkbook(), targetSheet = targetWorkbook.getActiveSheet(), target = targetSheet.getChartById(chart.getChartId())
    const contents = JSON.stringify(targetWorkbook.save()), history = visitor.editor.history.captureSavepoint(targetWorkbook.getId())
    for (const operation of [() => target.setHierarchyFields([1, 0]), () => target.setValueField(1), () => target.setSunburstLabel({ visible: false }),
      () => target.resetSunburstLabel(), () => target.setCenterLabel({ visible: true }), () => target.resetCenterLabel(),
      () => visitor.executeCommand('sheet.command.update-chart', { unitId: targetWorkbook.getId(), chartId: chart.getChartId(), patch: { title: 'Blocked Sunburst edit' } }),
      () => targetSheet.insertChart(newBuilder(targetSheet).build())]) {
      await assert.rejects(async () => { await operation() }, /编辑权限/)
      assert.equal(JSON.stringify(targetWorkbook.save()), contents); assert.deepEqual(visitor.editor.history.captureSavepoint(targetWorkbook.getId()), history)
    }
    assert.equal(compute.calls.length, 0)
  } finally { visitor?.dispose(); studio.dispose() }
})
