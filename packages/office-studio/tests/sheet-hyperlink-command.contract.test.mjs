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

test('hyperlink label types share command history and survive reopening after replacement and removal', () => {
  const studio = new apiModule.OfficeStudio()
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.createWorkbook({ sheets: [{ id: 'typed-links', name: 'Typed links' }] })
    const params = { unitId: workbook.getId(), sheetId: 'typed-links', col: 0 }
    studio.executeCommandSync('sheet.command.set-range-values', {
      ...params, range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }, values: [[42.5], [false], [7]],
    })
    const cell = row => studio.getWorkbookModel(workbook.getId()).sheets[0].cells[`${row}:0`]
    const samples = [
      { display: '', expected: '42.5' },
      { display: '', expected: 'false' },
      { display: '=SUM(1,2)', expected: '=SUM(1,2)' },
    ]
    for (const [row, sample] of samples.entries()) {
      const original = structuredClone(cell(row))
      studio.executeCommandSync('sheet.command.set-cell-hyperlink', {
        ...params, row, draft: { target: 'https://example.com/report', display: sample.display },
      })
      assert.equal(cell(row).value, sample.expected)
      assert.equal(cell(row).dataType, 'text')
      assert.equal(cell(row).cachedDataType, null)
      assert.equal(cell(row).formula, null)
      workbook.undo()
      assert.deepEqual(cell(row), original)
      workbook.redo()
      assert.equal(cell(row).dataType, 'text')
      studio.executeCommandSync('sheet.command.remove-cell-hyperlink', { ...params, row })
      assert.equal(cell(row).hyperlink, null)
      assert.equal(cell(row).value, sample.expected)
      assert.equal(cell(row).dataType, 'text')
      workbook.undo()
      assert.equal(cell(row).hyperlink, 'https://example.com/report')
    }

    const snapshot = studio.getWorkbookModel(workbook.getId())
    const reopened = api.createWorkbook({ ...snapshot, id: 'typed-links-reopened' })
    for (const [row, sample] of samples.entries()) {
      const restored = studio.getWorkbookModel(reopened.getId()).sheets[0].cells[`${row}:0`]
      assert.equal(restored.value, sample.expected)
      assert.equal(restored.dataType, 'text')
      assert.equal(restored.cachedDataType, null)
      assert.equal(restored.formula, null)
      assert.equal(restored.hyperlink, 'https://example.com/report')
    }
  } finally {
    studio.dispose()
  }
})

test('workbook hyperlink navigation activates the target sheet and selects the linked range', () => {
  const studio = new apiModule.OfficeStudio()
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.createWorkbook({ sheets: [
      { id: 'source', name: 'Source' },
      { id: 'target', name: 'Quarter Report' },
    ] })
    workbook.navigateToSheetHyperlink('#gid=target&range=B2%3AC4')
    assert.equal(workbook.getActiveSheet()?.getSheetId(), 'target')
    assert.deepEqual(studio.getSelection(), {
      startRow: 1, startColumn: 1, endRow: 3, endColumn: 2,
      sheetId: 'target', unitId: workbook.getId(),
    })
    assert.throws(() => workbook.navigateToSheetHyperlink('#gid=missing&range=A1'), /Invalid internal worksheet hyperlink/)
    assert.throws(() => workbook.navigateToSheetHyperlink('https://example.com'), /Invalid internal worksheet hyperlink/)
  } finally {
    studio.dispose()
  }
})
