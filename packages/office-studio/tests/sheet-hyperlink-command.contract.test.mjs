import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server
let apiModule

beforeAll(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'packages/office-studio/vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  apiModule = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

afterAll(async () => { await server?.close() })

test('cell hyperlink commands share model history, reject formulas, and retain links on reopen', () => {
  const studio = new apiModule.OfficeStudio()
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.createWorkbook({ sheets: [{ id: 'links', name: 'Links' }] })
    const params = { unitId: workbook.getId(), sheetId: 'links', row: 0, col: 0 }
    studio.executeCommandSync('sheet.command.set-cell-hyperlink', {
      ...params, draft: { target: 'https://example.com/report', display: 'Report' },
    })
    const cell = () => studio.getWorkbookModel(workbook.getId()).sheets[0].cells['0:0']
    assert.equal(cell().hyperlink, 'https://example.com/report')
    assert.equal(cell().value, 'Report')
    workbook.undo()
    assert.equal(cell(), undefined)
    workbook.redo()
    assert.equal(cell().hyperlink, 'https://example.com/report')

    const snapshot = studio.getWorkbookModel(workbook.getId())
    const reopened = api.createWorkbook({ ...snapshot, id: 'links-reopened' })
    assert.equal(studio.getWorkbookModel(reopened.getId()).sheets[0].cells['0:0'].hyperlink, 'https://example.com/report')
    studio.executeCommandSync('sheet.command.remove-cell-hyperlink', params)
    assert.equal(cell().hyperlink, null)
    assert.equal(cell().value, 'Report')
    workbook.undo()
    assert.equal(cell().hyperlink, 'https://example.com/report')

    studio.executeCommandSync('sheet.command.set-range-values', {
      unitId: workbook.getId(), sheetId: 'links',
      range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 }, values: [['=SUM(1,2)']],
    })
    assert.throws(() => studio.executeCommandSync('sheet.command.set-cell-hyperlink', {
      ...params, row: 1, draft: { target: 'https://example.com', display: 'Blocked' },
    }), /公式/)
    assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].cells['1:0'].formula, '=SUM(1,2)')
  } finally {
    studio.dispose()
  }
})
