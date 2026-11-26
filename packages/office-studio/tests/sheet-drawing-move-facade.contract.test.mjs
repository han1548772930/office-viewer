import assert from 'node:assert/strict'
import { afterAll, beforeAll, test, vi } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { sheetDrawingWorkbook, drawingRef as ref } from './fixtures/sheet-drawing-model-fixture.mjs'

// Deferred real command/Facade/content-history drafts. The SSR hooks have not run.
let server, api, host, broker
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  ;[api, host, broker] = await Promise.all([
    server.ssrLoadModule('/packages/office-studio/src/index.ts'),
    server.ssrLoadModule('/packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts'),
    server.ssrLoadModule('/packages/office-studio/src/api/sheet-drawing-selection.ts'),
  ])
})
afterAll(async () => { await server?.close() })
const json = value => JSON.parse(JSON.stringify(value))
const selection = [ref('image', 'a'), ref('chart', 'c')]
const count = f => f.studio.editor.history.units[f.unitId].undo.length
const scene = f => host.createSheetDrawingScene(f.studio.getWorkbookModel(f.unitId), 'data')
const matrix = (f, kind, id) => scene(f).objects.find(object => object.ref.kind === kind && object.ref.id === id).geometry.matrix
function setup(model = sheetDrawingWorkbook()) {
  const studio = new api.OfficeStudio({ workbook: model }), facade = api.FOffice.newAPI(studio), book = facade.getActiveWorkbook(), sheet = book.getActiveSheet()
  const unitId = book.getId(), commands = [], listener = studio.events.addEvent('CommandExecuted', value => commands.push(value.id))
  return { studio, facade, book, sheet, unitId, commands, dispose() { listener.dispose(); studio.dispose() } }
}

test('numeric Direction, mixed UI refs and Facade share one timed history with every actual movement retained', () => {
  const f = setup(), clock = vi.spyOn(Date, 'now').mockReturnValue(1000); let reopened
  try {
    const before = json(f.book.save()), oldScene = scene(f), initial = count(f)
    assert.equal(f.sheet.moveDrawings(selection, f.facade.Enum.Direction.RIGHT), f.sheet)
    clock.mockReturnValue(1100); f.sheet.moveDrawings([...selection].reverse(), f.facade.Enum.Direction.DOWN)
    clock.mockReturnValue(1200); f.sheet.moveDrawings(selection, f.facade.Enum.Direction.LEFT, 10)
    assert.equal(count(f), initial + 1)
    assert.equal(f.commands.filter(id => id === 'sheet.command.move-drawing').length, 3)
    for (const item of oldScene.objects.filter(object => selection.some(ref => ref.kind === object.ref.kind && ref.id === object.ref.id))) {
      const current = matrix(f, item.ref.kind, item.ref.id)
      assert.equal(current.e, item.geometry.matrix.e - 9); assert.equal(current.f, item.geometry.matrix.f + 1)
      for (const name of ['a', 'b', 'c', 'd']) assert.equal(current[name], item.geometry.matrix[name])
    }
    const after = json(f.book.save())
    assert.deepEqual(after.charts[0].size, before.charts[0].size); assert.deepEqual(after.charts[0].source, before.charts[0].source)
    assert.deepEqual(after.sheets[0].images[0].crop, before.sheets[0].images[0].crop)
    f.book.undo(); assert.deepEqual(json(f.book.save()), before)
    assert.deepEqual(broker.getSheetDrawingSelection(f.studio, f.unitId, 'data'), selection)
    f.book.redo(); assert.deepEqual(json(f.book.save()), after)
    reopened = new api.OfficeStudio({ workbook: after })
    assert.deepEqual(json(host.createSheetDrawingScene(reopened.getWorkbookModel(f.unitId), 'data')), json(scene(f)))
  } finally { clock.mockRestore(); reopened?.dispose(); f.dispose() }
})

test('the inclusive one-second window renews and a later burst creates another undo entry', () => {
  const f = setup(), clock = vi.spyOn(Date, 'now').mockReturnValue(1000)
  try {
    const initial = count(f); f.sheet.moveDrawings(selection, 1)
    clock.mockReturnValue(2000); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 1)
    clock.mockReturnValue(3001); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 2)
    f.book.undo(); assert.equal(matrix(f, 'chart', 'c').e, 52)
  } finally { clock.mockRestore(); f.dispose() }
})

test('saving between moves preserves an undo boundary and the real dirty savepoint', () => {
  const f = setup(), clock = vi.spyOn(Date, 'now').mockReturnValue(1000)
  try {
    const initial = count(f); f.sheet.moveDrawings(selection, 1)
    const saved = json(f.book.save()); f.studio.editor.history.markSaved(f.unitId)
    clock.mockReturnValue(1100); f.sheet.moveDrawings(selection, 1)
    assert.equal(count(f), initial + 2); assert.equal(f.studio.editor.history.isDirty(f.unitId), true)
    f.book.undo(); assert.deepEqual(json(f.book.save()), saved); assert.equal(f.studio.editor.history.isDirty(f.unitId), false)
  } finally { clock.mockRestore(); f.dispose() }
})

test('changing a shared selection away and back or replacing its owner separates move bursts', () => {
  const f = setup(), clock = vi.spyOn(Date, 'now').mockReturnValue(1000)
  try {
    const initial = count(f); f.sheet.moveDrawings(selection, 1)
    broker.setSheetDrawingSelection(f.studio, f.unitId, 'data', [ref('image', 'd')])
    broker.setSheetDrawingSelection(f.studio, f.unitId, 'data', selection)
    clock.mockReturnValue(1100); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 2)
    const binding = broker.bindSheetDrawingSelection(f.studio, f.unitId, () => ({ sheetId: 'data', refs: selection }))
    binding.dispose(); clock.mockReturnValue(1200); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 3)
  } finally { clock.mockRestore(); f.dispose() }
})

test('undo then redo and an intervening native-owner edit never amend a previous burst', () => {
  const f = setup(), clock = vi.spyOn(Date, 'now').mockReturnValue(1000)
  try {
    const initial = count(f); f.sheet.moveDrawings(selection, 1); f.book.undo(); f.book.redo()
    clock.mockReturnValue(1100); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 2)
    f.sheet.transformDrawings([ref('image', 'd')], { delta: { a: 1, b: 0, c: 0, d: 1, e: 1, f: 0 } })
    clock.mockReturnValue(1200); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 4)
  } finally { clock.mockRestore(); f.dispose() }
})

test('actual skeleton dimensions constrain an anchored chart while None images retain independent absolute movement', () => {
  const model = sheetDrawingWorkbook(); model.sheets[0].columnCount = 2; model.charts[0].position.x = 44 + 170
  model.sheets[0].images[0].absolutePosition.left = 0
  const f = setup(model)
  try {
    const initial = count(f), before = json(f.book.save())
    f.sheet.moveDrawings(selection, 1, 10)
    assert.equal(count(f), initial); assert.deepEqual(json(f.book.save()), before)
    f.sheet.moveDrawings([ref('image', 'a')], 3)
    assert.equal(matrix(f, 'image', 'a').e, -1); assert.equal(count(f), initial + 1)
  } finally { f.dispose() }
})

test('an inaccessible destination footprint rejects all mixed owners and leaves content/history intact', () => {
  const model = sheetDrawingWorkbook(), sheet = model.sheets[0]
  model.charts[0].position.x = 44 + 70; model.charts[0].size.width = 30
  sheet.rangeProtectionRules = [{ id: 'locked-column', permissionId: 'locked-column', unitType: 'range',
    ranges: [{ startRow: 0, endRow: 99, startCol: 1, endCol: 1 }], ownerUserId: 'peer', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }]
  const f = setup(model)
  try {
    const before = json(f.book.save()), initial = count(f)
    assert.throws(() => f.sheet.moveDrawings(selection, 1), /权限|permission/i)
    assert.deepEqual(json(f.book.save()), before); assert.equal(count(f), initial)
  } finally { f.dispose() }
})

test('invalid move parameters, stale selection and cancellation reject through the canonical registry', () => {
  const f = setup()
  try {
    const before = json(f.book.save()), initial = count(f), input = { unitId: f.unitId, sheetId: 'data', selection }
    assert.throws(() => f.studio.executeCommandSync('sheet.command.move-drawing', { ...input, direction: 'right' }), TypeError)
    assert.throws(() => f.studio.executeCommandSync('sheet.command.move-drawing', { ...input, direction: 1, step: 2 }), TypeError)
    assert.throws(() => f.studio.executeCommandSync('sheet.command.move-drawing', input), TypeError)
    broker.setSheetDrawingSelection(f.studio, f.unitId, 'data', [])
    assert.throws(() => f.studio.executeCommandSync('sheet.command.move-drawing', { ...input, direction: 1, expectedSelection: selection }), /选择/)
    const cancel = new AbortController(); cancel.abort()
    assert.throws(() => f.studio.executeCommandSync('sheet.command.move-drawing', { ...input, direction: 1, signal: cancel.signal }), /取消/)
    assert.deepEqual(json(f.book.save()), before); assert.equal(count(f), initial)
  } finally { f.dispose() }
})

test('a command observer that reenters authoring cannot attach a later nudge to stale history', () => {
  const f = setup(), clock = vi.spyOn(Date, 'now').mockReturnValue(1000); let entered = false
  const listener = f.studio.events.addEvent('CommandExecuted', event => {
    if (event.id !== 'sheet.command.move-drawing' || entered) return
    entered = true; f.sheet.transformDrawings(selection, { delta: { a: 1, b: 0, c: 0, d: 1, e: 5, f: 0 } })
  })
  try {
    const initial = count(f); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 2)
    clock.mockReturnValue(1100); f.sheet.moveDrawings(selection, 1); assert.equal(count(f), initial + 3)
    assert.equal(matrix(f, 'chart', 'c').e, 57)
  } finally { listener.dispose(); clock.mockRestore(); f.dispose() }
})
