import { expect, test } from '@playwright/test'

test('row grouping collapses, expands, undoes and redoes through the workbook model', async ({ page }) => {
  await page.goto('/')
  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()

  await page.evaluate(async () => {
    const runtimePath = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(runtimePath)
    const workbookId = studio.activeWorkbookId
    const workbook = studio.getWorkbookModel(workbookId)!
    studio.executeCommandSync('sheet.command.change-outline', {
      unitId: workbookId,
      sheetId: workbook.activeSheetId,
      action: { kind: 'group', axis: 'row', start: 1, end: 3 },
    })
  })

  const readOutline = () => page.evaluate(async () => {
    const runtimePath = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(runtimePath)
    const workbook = studio.getWorkbookModel(studio.activeWorkbookId)
    return workbook?.sheets.find(sheet => sheet.id === workbook.activeSheetId)?.outline?.groups ?? []
  })
  await expect.poll(async () => readOutline()).toMatchObject([{ axis: 'row', start: 1, end: 3, collapsed: false }])
  await grid.getByRole('button', { name: '行分组折叠' }).click()
  await expect.poll(async () => readOutline()).toMatchObject([{ axis: 'row', start: 1, end: 3, collapsed: true }])

  await page.keyboard.press('Control+z')
  await expect.poll(async () => readOutline()).toMatchObject([{ axis: 'row', start: 1, end: 3, collapsed: false }])
  await page.keyboard.press('Control+y')
  await expect.poll(async () => readOutline()).toMatchObject([{ axis: 'row', start: 1, end: 3, collapsed: true }])
})
