import { expect, test, type Locator, type Page } from '@playwright/test'

type RuntimeModule = typeof import('../../examples/workbench/src/runtime')
async function setup(page: Page): Promise<void> {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(path) as RuntimeModule
    const book = officeAPI.createWorkbook({ id: 'outline-dialog-fixture', name: '大纲弹窗验收', date1904: false }), sheet = book.getActiveSheet()!
    book.insertSheet('Other'); book.setActiveSheet(sheet)
    studio.setSelections(book.getId(), sheet.getSheetId(), [{ startRow: 2, endRow: 4, startColumn: 3, endColumn: 5 }])
  })
}
async function read(page: Page) {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(path) as RuntimeModule
    const book = officeAPI.getActiveWorkbook()!, workbook = studio.getWorkbookModel(book.getId())!
    return { workbook: JSON.parse(JSON.stringify(workbook)), revision: studio.getWorkbookRevision(book.getId()),
      history: studio.editor.history.units[book.getId()]?.undo.length ?? 0, snapshot: book.getSnapshot() }
  })
}
async function choose(page: Page, selector: Locator, label: string): Promise<void> {
  await selector.click()
  await page.locator('.ant-select-dropdown:visible').getByTitle(label, { exact: true }).click()
}
async function prompt(page: Page): Promise<Locator> {
  const grid = page.getByRole('grid', { name: '工作表画布' })
  await grid.focus(); await page.keyboard.press('Alt+Shift+ArrowRight')
  const dialog = page.getByRole('dialog', { name: '组合行列', exact: true })
  await expect(dialog).toBeVisible(); return dialog
}
async function settings(page: Page): Promise<Locator> {
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '数据', exact: true }).click()
  const direct = ribbon.getByRole('button', { name: '大纲设置', exact: true })
  if (await direct.count()) await direct.click()
  else {
    const button = ribbon.getByRole('button', { name: '组合', exact: true }), selector = ribbon.getByRole('combobox', { name: /组合/ })
    if (await button.count()) await button.click(); else await selector.click()
    const menuItem = page.getByRole('menuitem', { name: '大纲设置', exact: true })
    if (await menuItem.count()) await menuItem.click()
    else await page.locator('.ant-select-dropdown:visible').getByTitle('大纲设置', { exact: true }).click()
  }
  const dialog = page.getByRole('dialog', { name: '大纲汇总设置', exact: true })
  await expect(dialog).toBeVisible(); return dialog
}

// Deferred browser drafts: these exercise the actual workbench and Ant dialogs.
test('single-area shortcut opens an Ant axis prompt and publishes the chosen column group', async ({ page }) => {
  await setup(page)
  const before = await read(page), dialog = await prompt(page)
  await choose(page, dialog.getByRole('combobox', { name: '大纲分组方向' }), '组合 D–F 列')
  await dialog.getByRole('button', { name: '确定', exact: true }).click(); await expect(dialog).toBeHidden()
  await expect.poll(async () => (await read(page)).workbook.sheets[0].outline?.groups).toMatchObject([{ axis: 'column', start: 3, end: 5, collapsed: false }])
  const after = await read(page)
  expect(after.revision).toBe(before.revision + 1); expect(after.history).toBe(before.history + 1)
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeFocused()
})

test('axis prompt cancellation preserves workbook, history and Canvas focus', async ({ page }) => {
  await setup(page)
  const before = await read(page), dialog = await prompt(page)
  await dialog.getByRole('button', { name: '取消', exact: true }).click(); await expect(dialog).toBeHidden()
  const after = await read(page)
  expect(after.workbook).toEqual(before.workbook); expect(after.revision).toBe(before.revision); expect(after.history).toBe(before.history)
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeFocused()
})

test('a content edit makes the axis draft stale until explicit reread', async ({ page }) => {
  await setup(page)
  const dialog = await prompt(page)
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.addColumnOutline(7, 2)
  })
  await expect(dialog.getByRole('alert')).toContainText('更改')
  await expect(dialog.getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await dialog.getByRole('button', { name: '重新读取', exact: true }).click()
  await expect(dialog.getByRole('button', { name: '确定', exact: true })).toBeEnabled()
  await dialog.getByRole('button', { name: '确定', exact: true }).click()
  await expect.poll(async () => (await read(page)).workbook.sheets[0].outline?.groups.length).toBe(2)
  expect((await read(page)).workbook.sheets[0].outline.groups).toEqual(expect.arrayContaining([
    expect.objectContaining({ axis: 'row', start: 2, end: 4 }), expect.objectContaining({ axis: 'column', start: 7, end: 8 }),
  ]))
})

test('row and column summary settings publish atomically and survive undo, redo and public resource serialization', async ({ page }) => {
  await setup(page)
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.addRowOutline(2, 3).addColumnOutline(3, 3)
  })
  const before = await read(page), dialog = await settings(page)
  await choose(page, dialog.getByRole('combobox', { name: '大纲行汇总位置' }), '明细上方')
  await choose(page, dialog.getByRole('combobox', { name: '大纲列汇总位置' }), '明细左侧')
  await dialog.getByRole('button', { name: '确定', exact: true }).click(); await expect(dialog).toBeHidden()
  const after = await read(page), sheet = after.workbook.sheets[0]
  expect(sheet.outline).toMatchObject({ summaryAbove: true, summaryLeft: true })
  expect(after.revision).toBe(before.revision + 1); expect(after.history).toBe(before.history + 1)
  const resource = after.snapshot.resources!.find(item => item.name === 'SHEET_OUTLINE_PLUGIN')!
  expect(JSON.parse(resource.data)[sheet.id].officeStudio).toMatchObject({ schemaVersion: 1, summaryAbove: true, summaryLeft: true })
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule; officeAPI.getActiveWorkbook()!.undo()
  })
  expect((await read(page)).workbook.sheets[0].outline).toMatchObject({ summaryAbove: false, summaryLeft: false })
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule; officeAPI.getActiveWorkbook()!.redo()
  })
  expect((await read(page)).workbook.sheets[0].outline).toMatchObject({ summaryAbove: true, summaryLeft: true })
})

test('switching worksheets invalidates summary publication and cancellation retains the new host', async ({ page }) => {
  await setup(page)
  const dialog = await settings(page), before = await read(page)
  await choose(page, dialog.getByRole('combobox', { name: '大纲行汇总位置' }), '明细上方')
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const book = officeAPI.getActiveWorkbook()!; book.setActiveSheet(book.getSheetByName('Other')!)
  })
  await expect(dialog.getByRole('alert')).toContainText('过期')
  await expect(dialog.getByRole('button', { name: '确定', exact: true })).toBeDisabled()
  await dialog.getByRole('button', { name: '取消', exact: true }).click()
  const after = await read(page)
  expect(after.workbook.activeSheetId).toBe(after.workbook.sheets[1].id)
  expect(after.workbook.sheets[0].outline).toEqual(before.workbook.sheets[0].outline)
  expect(after.history).toBe(before.history)
})
