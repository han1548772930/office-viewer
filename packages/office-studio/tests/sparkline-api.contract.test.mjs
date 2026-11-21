import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let apiModule

beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(), server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

afterAll(async () => { await server?.close() })

test('sparkline creation, editing, clearing, undo and workbook save use the shared command history', () => {
  const studio = new apiModule.OfficeStudio()
  const workbook = apiModule.FOffice.newAPI(studio).createWorkbook({
    sheets: [{ id: 'sparkline-sheet', name: 'Data', cellData: { 0: { 1: { v: 3 }, 2: { v: -1 }, 3: { v: 6 } } } }],
  })
  const sheet = workbook.getActiveSheet()
  const request = { targetRange: 'E1', sourceRange: 'B1:D1', type: 'column', color: '#187a55', negativeColor: '#c43b3b', emptyCells: 'zero', axis: 'individual' }
  sheet.setSparklines(request)
  assert.deepEqual(sheet.getSparkline('E1'), {
    type: request.type, sourceRange: request.sourceRange, color: request.color,
    negativeColor: request.negativeColor, emptyCells: request.emptyCells, axis: request.axis,
  })
  assert.equal(workbook.save().sheets[0].cells['0:4'].sparkline.sourceRange, 'B1:D1')

  sheet.setSparklines({ ...request, type: 'winloss', color: '#174ea6' })
  assert.equal(sheet.getSparkline('E1').type, 'winloss')
  workbook.undo()
  assert.equal(sheet.getSparkline('E1').type, 'column')
  sheet.clearSparklines('E1')
  assert.equal(sheet.getSparkline('E1'), null)
  workbook.undo()
  assert.equal(sheet.getSparkline('E1').type, 'column')

  const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
  assert.equal(apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getActiveSheet().getSparkline('E1').sourceRange, 'B1:D1')
  studio.dispose()
  reopened.dispose()
})

test('sparkline command validates target coordinates and respects protected target cells', () => {
  const owner = new apiModule.OfficeStudio({ currentUserId: 'owner' })
  const ownerApi = apiModule.FOffice.newAPI(owner)
  const workbook = ownerApi.createWorkbook({ sheets: [{ id: 'protected-sheet', name: 'Data' }] })
  workbook.getActiveSheet().setRangeProtection('E1', { ranges: [], editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [] })
  const studio = new apiModule.OfficeStudio({ workbook: workbook.save(), currentUserId: 'other-user' })
  const api = apiModule.FOffice.newAPI(studio)
  const protectedWorkbook = api.getActiveWorkbook()
  const sheet = protectedWorkbook.getActiveSheet()
  assert.throws(() => sheet.setSparklines({ targetRange: 'E1', sourceRange: 'A1:B1', type: 'line', color: '#167a55' }), /受保护/)
  assert.throws(() => studio.executeCommandSync('sheet.command.set-sparklines', { unitId: protectedWorkbook.id, sheetId: sheet.id, patches: { A0: null } }), /Invalid sparkline target/)
  owner.dispose()
  studio.dispose()
})

test('sparkline command rejects invalid styles and mismatched group dimensions without changing the workbook', () => {
  const studio = new apiModule.OfficeStudio()
  const workbook = apiModule.FOffice.newAPI(studio).createWorkbook({ sheets: [{ id: 'invalid-sparkline-sheet', name: 'Data' }] })
  const sheet = workbook.getActiveSheet()
  assert.throws(() => sheet.setSparklines({ targetRange: 'E1', sourceRange: 'A1:B1', type: 'line', color: 'green' }), /颜色格式无效/)
  assert.throws(() => sheet.setSparklines({ targetRange: 'E1:E2', sourceRange: 'A1:B1', type: 'line', color: '#167a55' }), /目标数必须与源数据系列数一致/)
  assert.equal(sheet.getSparkline('E1'), null)
  studio.dispose()
})
