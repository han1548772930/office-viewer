import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Real Facade/command/history drafts. The harness runs only in unified validation.
let server, api, fixtures
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, fixtures] = await Promise.all([server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/tests/fixtures/sparkline-workbench-fixture.mjs')])
})
afterAll(async () => { await server?.close() })
function setup(legacy = false) {
  const workbook = fixtures.sparklineWorkbook()
  if (legacy) delete workbook.sheets[0].cells['0:6'].sparkline.groupId
  const studio = new api.OfficeStudio({ workbook }), office = api.FOffice.newAPI(studio), book = office.getActiveWorkbook(), sheet = book.getActiveSheet()
  return { studio, office, book, sheet, history: () => studio.editor.history.units[book.getId()]?.undo.length ?? 0, dispose: () => studio.dispose() }
}
const source = (row = 0) => ({ startRow: row, endRow: row, startColumn: 1, endColumn: 4 })
const target = (row = 0, col = 7) => ({ startRow: row, endRow: row, startColumn: col, endColumn: col })

test('public enum configuration keeps unspecified group fields and accepts value-only extrema patches', () => {
  const f = setup()
  try {
    const group = f.sheet.getSparklineGroupByCell(0, 6)
    assert.equal(group.setConfig({ type: f.office.Enum.SparklineTypeEnum.BAR_CHART, extremumMin: { value: -200 } }), group)
    const config = f.sheet.getAllSubSparkline().get('north').config
    assert.equal(config.type, api.SparklineTypeEnum.BAR_CHART); assert.equal(config.extremumMin.type, api.SparklineExtremumTypeEnum.GROUP_EXTREMUM)
    assert.equal(config.extremumMin.value, -200); assert.equal(config.points.highPoint.visible, true); assert.equal(config.seriesColor, '#246fce')
    f.book.undo(); assert.equal(f.sheet.getSparkline('G1').type, 'line')
  } finally { f.dispose() }
})
test('numeric-string and prototype-like enums reject without changing content or history', () => {
  const f = setup()
  try {
    const group = f.sheet.getSparklineGroupByCell(0, 6), before = JSON.stringify(f.book.save()), history = f.history()
    for (const config of [{ type: '1' }, { type: 'constructor' }, { type: '__proto__' }, { emptyShowAs: 'toString' },
      { nonNumShowAs: '__proto__' }, { extremumMin: { type: '2' } }, { extremumMax: { type: 'constructor' } },
      { points: { highPoint: { visible: null } } }, { axis: { reverse: 1 } }, { lineWidth: null }, { themeType: ['standard'] }]) assert.throws(() => group.setConfig(config))
    assert.equal(JSON.stringify(f.book.save()), before); assert.equal(f.history(), history)
  } finally { f.dispose() }
})
test('returned matrix mutations are detached and expose the complete structural API', () => {
  const f = setup()
  try {
    const matrix = f.sheet.getAllSubSparkline().get('north').sparklines, before = JSON.stringify(f.book.save())
    assert.ok(matrix instanceof api.SparklineRangeMatrix)
    matrix.moveRows(1, 1, 0); matrix.insertColumns(0, 2); matrix.removeColumns(0, 1)
    assert.equal(matrix.getValue(0, 7).startRow, 1); assert.equal(matrix.getSlice(0, 1, 7, 7).getSizeOf(), 2)
    assert.equal(JSON.stringify(f.book.save()), before)
  } finally { f.dispose() }
})
test('legacy single-item move rebinds to the committed group and supports a second move and deletion', () => {
  const f = setup(true)
  try {
    const handle = f.sheet.getSparklineByCell(0, 6), baseline = f.history()
    assert.ok(handle instanceof api.FSparkline); assert.equal(handle.changeDataSource(source(1), target()), handle)
    const groupId = f.sheet.getSparkline('H1').groupId
    assert.ok(groupId && !groupId.startsWith('ungrouped:')); assert.equal(f.sheet.getSparkline('G1'), null)
    assert.equal(handle.changeDataSource(source(2), target(2, 8)), handle); assert.equal(f.sheet.getSparkline('I3').groupId, groupId)
    handle.removeSparkline(); assert.equal(f.sheet.getSparkline('I3'), null); assert.equal(f.history(), baseline + 3)
    f.book.undo(); assert.equal(f.sheet.getSparkline('I3').sourceRange, 'B3:E3')
  } finally { f.dispose() }
})
test('legacy whole-group move keeps a valid group handle for configuration and removal', () => {
  const f = setup(true)
  try {
    const handle = f.sheet.getSparklineGroupByCell(0, 6)
    assert.ok(handle instanceof api.FSparklineGroup); assert.equal(handle.changeDataSource([source()], [target()]), handle)
    assert.equal(handle.setConfig({ seriesColor: '#123456' }), handle); assert.equal(f.sheet.getSparkline('H1').color, '#123456')
    handle.removeSparklineGroup(); assert.equal(f.sheet.getSparkline('H1'), null)
  } finally { f.dispose() }
})
test('cancelled changes leave the old handle at its original member', () => {
  const f = setup(true)
  try {
    const handle = f.sheet.getSparklineByCell(0, 6), before = JSON.stringify(f.book.save())
    const cancel = f.studio.events.addEvent('BeforeCommandExecute', event => { if (event.id === api.SET_SPARKLINES_COMMAND) event.cancel = true })
    assert.throws(() => handle.changeDataSource(source(1), target()), /canceled/); cancel.dispose()
    assert.equal(JSON.stringify(f.book.save()), before); handle.removeSparkline(); assert.equal(f.sheet.getSparkline('G1'), null)
  } finally { f.dispose() }
})
test('no-op moves return undefined and do not create a history entry', () => {
  const f = setup()
  try {
    const handle = f.sheet.getSparklineByCell(0, 6), baseline = f.history()
    assert.equal(handle.changeDataSource(source(), target(0, 6)), undefined); assert.equal(f.history(), baseline)
    handle.removeSparkline(); assert.equal(f.sheet.getSparkline('G1'), null)
  } finally { f.dispose() }
})
test('observer reentry cannot bind an old move handle to a later replacement', () => {
  const f = setup(true)
  try {
    const handle = f.sheet.getSparklineByCell(0, 6)
    const observation = f.studio.events.addEvent('SheetSparklineChanged', () => { observation.dispose(); f.sheet.clearSparklines('H1') })
    assert.equal(handle.changeDataSource(source(1), target()), undefined); assert.equal(f.sheet.getSparkline('H1'), null)
    assert.throws(() => handle.removeSparkline(), /已删除|重新组合/)
  } finally { f.dispose() }
})
test('owner generation changes reject a previously acquired handle', () => {
  const f = setup()
  try {
    const handle = f.sheet.getSparklineByCell(0, 6)
    f.studio.editor.history.clear(f.book.getId())
    assert.throws(() => handle.changeDataSource(source(1), target()), /被替换|关闭/)
  } finally { f.dispose() }
})
test('source edits and undo notify group data without changing the group spec', () => {
  const f = setup(), events = []
  try {
    f.studio.events.addEvent('SheetSparklineChanged', event => events.push(event))
    const before = JSON.stringify(f.sheet.getSparkline('G1'))
    f.sheet.getRange('B1').setValue(200); assert.equal(events.length, 1); assert.equal(events[0].sparklines.length, 1)
    assert.equal(JSON.stringify(f.sheet.getSparkline('G1')), before)
    f.book.undo(); assert.equal(events.length, 2)
    f.sheet.getRange('J50').setValue(100); assert.equal(events.length, 2)
  } finally { f.dispose() }
})
test('listener payloads are detached and listener errors do not undo a completed edit', () => {
  const f = setup(), seen = []
  try {
    f.studio.events.addEvent('SheetSparklineChanged', event => { event.sparklines[0].sparklines.reset(); throw new Error('subscriber failure') })
    f.studio.events.addEvent('SheetSparklineChanged', event => { seen.push(event.sparklines[0].sparklines.getSizeOf()) })
    f.sheet.getRange('B1').setValue(200)
    assert.deepEqual(seen, [2]); assert.equal(f.sheet.getRange('B1').getValue(), 200)
    assert.equal(f.studio.getEventListenerErrors().some(error => error.eventName === 'SheetSparklineChanged' && /subscriber failure/.test(String(error.error))), true)
  } finally { f.dispose() }
})
