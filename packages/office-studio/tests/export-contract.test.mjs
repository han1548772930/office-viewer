import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'

let server
let module

before(async () => {
  server = await createServer({
    configFile: resolve(process.cwd(), 'vite.config.ts'),
    root: process.cwd(),
    server: { middlewareMode: true, hmr: false },
    appType: 'custom',
    logLevel: 'error',
  })
  module = await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})

after(async () => {
  await server?.close()
})

test('export service rejects an already-aborted request before creating a worker', async () => {
  const controller = new AbortController()
  controller.abort()
  const studio = new module.OfficeStudio()
  const model = studio.getWorkbookModel(studio.activeWorkbookId)
  await assert.rejects(
    () => module.exportOfficeModel('workbook', model, { signal: controller.signal }),
    (error) => error instanceof module.OfficeExportError && error.code === 'CANCELLED',
  )
  studio.dispose()
})

test('export service keeps the public file name and type checks stable', async () => {
  const studio = new module.OfficeStudio()
  const workbook = studio.getWorkbookModel(studio.activeWorkbookId)
  assert.equal(module.officeDownloadName(workbook, 'workbook'), '未命名工作簿.xlsx')
  await assert.rejects(() => module.exportOfficeModel('document', workbook), /导出类型与当前模型不一致/)
  studio.dispose()
})
