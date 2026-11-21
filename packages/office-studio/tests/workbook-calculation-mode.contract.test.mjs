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

test('calculation mode persists and manual mode recalculates only on request', () => {
  const studio = new apiModule.OfficeStudio()
  const workbook = apiModule.FOffice.newAPI(studio).createWorkbook({
    sheets: [{ id: 'sheet-1', name: 'Sheet1', cellData: { 0: { 0: { v: 2 }, 1: { f: '=A1*2', v: 4 } } } }],
  })
  const sheet = workbook.getActiveSheet()
  assert.equal(workbook.getCalculationMode(), 'automatic')
  assert.equal(sheet.getRange('B1').getValue(), 4)

  workbook.setCalculationMode('manual')
  sheet.getRange('A1').setValue(3)
  assert.equal(sheet.getRange('B1').getValue(), 4)
  workbook.recalculate()
  assert.equal(sheet.getRange('B1').getValue(), 6)

  sheet.getRange('A1').setValue(5)
  assert.equal(sheet.getRange('B1').getValue(), 6)
  workbook.setCalculationMode('automatic')
  assert.equal(sheet.getRange('B1').getValue(), 10)

  const reopened = new apiModule.OfficeStudio({ workbook: workbook.save() })
  assert.equal(apiModule.FOffice.newAPI(reopened).getActiveWorkbook().getCalculationMode(), 'automatic')
  studio.dispose()
  reopened.dispose()
})

test('calculation mode command rejects invalid modes before modifying workbook state', () => {
  const studio = new apiModule.OfficeStudio()
  const id = studio.activeWorkbookId
  assert.throws(() => studio.executeCommandSync('sheet.command.set-calculation-mode', { unitId: id, mode: 'sometimes' }), /automatic or manual/)
  assert.equal(studio.getWorkbookModel(id).calculationMode, 'automatic')
  studio.dispose()
})
