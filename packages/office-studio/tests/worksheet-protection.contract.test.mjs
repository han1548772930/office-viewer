import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
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

afterAll(async () => {
  await server?.close()
})

test('worksheet protection commands enforce permissions and participate in workbook undo, redo, and reopen', () => {
  const studio = new apiModule.OfficeStudio({ currentUserId: 'owner' })
  const api = apiModule.FOffice.newAPI(studio)
  try {
    const workbook = api.createWorkbook({ sheets: [{ id: 'protected-sheet', name: 'Forecast' }] })
    const sheet = workbook.getSheetBySheetId('protected-sheet')
    const rule = {
      id: 'worksheet-rule', permissionId: 'worksheet-permission', unitType: 'worksheet',
      ownerUserId: 'owner', description: 'Forecast', editState: 'designedUserCanEdit',
      viewState: 'othersCanView', allowedUserIds: ['editor'],
    }

    studio.executeCommandSync('sheet.command.add-worksheet-protection', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(), rule,
    })
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].worksheetProtection, rule)
    assert.throws(() => studio.executeCommandSync('sheet.command.add-worksheet-protection', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(), rule: { ...rule, editState: 'invalid' },
    }), /编辑权限无效/)
    const visitorStudio = new apiModule.OfficeStudio({ currentUserId: 'visitor' })
    try {
      const visitorApi = apiModule.FOffice.newAPI(visitorStudio)
      const visitorWorkbook = visitorApi.createWorkbook({ ...studio.getWorkbookModel(workbook.getId()), id: 'visitor-copy' })
      assert.throws(() => visitorStudio.executeCommandSync('sheet.command.set-range-values', {
        unitId: visitorWorkbook.getId(), sheetId: sheet.getSheetId(),
        range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }, values: [['blocked']],
      }), /工作表受保护/)
      assert.throws(() => visitorStudio.executeCommandSync('sheet.command.set-worksheet-name', {
        unitId: visitorWorkbook.getId(), sheetId: sheet.getSheetId(), name: 'Blocked rename',
      }), /工作表受保护/)
      assert.throws(() => visitorStudio.executeCommandSync('sheet.command.delete-worksheet-protection', {
        unitId: visitorWorkbook.getId(), sheetId: sheet.getSheetId(),
      }), /创建者/)
    } finally {
      visitorStudio.dispose()
    }

    const snapshot = studio.getWorkbookModel(workbook.getId())
    const reopened = api.createWorkbook({ ...snapshot, id: 'worksheet-protection-reopened' })
    assert.deepEqual(studio.getWorkbookModel(reopened.getId()).sheets[0].worksheetProtection, rule)

    workbook.undo()
    assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].worksheetProtection, undefined)
    workbook.redo()
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].worksheetProtection, rule)

    studio.executeCommandSync('sheet.command.delete-worksheet-protection', {
      unitId: workbook.getId(), sheetId: sheet.getSheetId(),
    })
    assert.equal(studio.getWorkbookModel(workbook.getId()).sheets[0].worksheetProtection, undefined)
    workbook.undo()
    assert.deepEqual(studio.getWorkbookModel(workbook.getId()).sheets[0].worksheetProtection, rule)
  } finally {
    studio.dispose()
  }
})

test('private worksheet visibility is enforced by selection and facade data access', () => {
  const ownerStudio = new apiModule.OfficeStudio({ currentUserId: 'owner' })
  const ownerApi = apiModule.FOffice.newAPI(ownerStudio)
  try {
    const source = ownerApi.createWorkbook({ sheets: [{
      id: 'private-sheet', name: 'Private', worksheetProtection: {
        id: 'private-rule', permissionId: 'private-permission', unitType: 'worksheet', ownerUserId: 'owner',
        editState: 'onlyMe', viewState: 'noOneElseCanView', allowedUserIds: [],
      },
    }] })
    const readerStudio = new apiModule.OfficeStudio({ currentUserId: 'reader' })
    const readerApi = apiModule.FOffice.newAPI(readerStudio)
    try {
      readerApi.createWorkbook({ ...ownerStudio.getWorkbookModel(source.getId()), id: 'reader-copy' })
      const workbook = readerApi.getWorkbook('reader-copy')
      const privateSheet = workbook.getSheetBySheetId('private-sheet')
      assert.throws(() => privateSheet.getSheetName(), /没有查看权限/)
      assert.throws(() => workbook.setActiveSheet('private-sheet'), /没有查看权限/)
      assert.throws(() => workbook.save(), /没有查看权限/)
    } finally {
      readerStudio.dispose()
    }
  } finally {
    ownerStudio.dispose()
  }
})
