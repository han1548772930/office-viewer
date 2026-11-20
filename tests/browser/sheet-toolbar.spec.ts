import { expect, test, type Page, type Locator } from '@playwright/test'
import type { WorkbookModel } from '../../packages/office-studio/src/domain/model'
type RuntimeModule = typeof import('../../examples/workbench/src/runtime')

const ribbon = (page: Page): Locator => page.getByTestId('workbench-ribbon')
async function state(page: Page): Promise<WorkbookModel> {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(path) as RuntimeModule
    return JSON.parse(JSON.stringify(studio.getWorkbookModel(studio.activeWorkbookId)))
  })
}
async function cell(page: Page, key = '0:0') { const workbook = await state(page); return workbook.sheets.find(sheet => sheet.id === workbook.activeSheetId)?.cells[key] }
async function address(page: Page, text: string) {
  await page.getByRole('textbox', { name: '名称框', exact: true }).fill(text)
  await page.getByRole('textbox', { name: '名称框', exact: true }).press('Enter')
}
async function control(page: Page, name: string, role: 'button' | 'combobox' = 'button'): Promise<Locator> {
  let target = ribbon(page).getByRole(role, { name, exact: true })
  if (await target.count() && await target.isVisible()) return target
  const menu = page.getByRole('menu', { name: '更多功能', exact: true })
  if (!await menu.isVisible()) await ribbon(page).getByRole('button', { name: '更多功能', exact: true }).click()
  target = menu.getByRole(role, { name, exact: true })
  await expect(target).toBeVisible()
  return target
}
async function click(page: Page, name: string) { await (await control(page, name)).click() }
async function format(page: Page, name: string) {
  await (await control(page, '数字格式', 'combobox')).click()
  await page.locator('.ant-select-dropdown:visible').getByTitle(name, { exact: true }).click()
}
async function menu(page: Page, name: string, item: string) {
  await click(page, name)
  await page.getByRole('menu', { name, exact: true }).getByRole('button', { name: item, exact: true }).click()
}

test.beforeEach(async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeVisible()
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { univerAPI } = await import(path) as RuntimeModule
    const workbook = univerAPI.createWorkbook({ id: 'sheet-toolbar-workbook', name: '格式与顶部交互' })
    const sheet = workbook.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([[1234.5, 10], [.25, 20], [61.5, 30], ['1234.5', 40]])
  })
  await address(page, 'A1')
})

test('toolbar applies strike, font stepping, vertical alignment and range borders with undo/redo', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  await click(page, '删除线')
  await expect.poll(async () => (await cell(page))?.style.font.strike).toBe(true)
  await click(page, '增大字号')
  await expect.poll(async () => (await cell(page))?.style.font.size).toBe(12)
  await click(page, '减小字号')
  await expect.poll(async () => (await cell(page))?.style.font.size).toBe(11)
  await menu(page, '垂直对齐', '底端对齐')
  await expect.poll(async () => (await cell(page))?.style.vertical).toBe('bottom')
  await address(page, 'A1:B2')
  await menu(page, '边框', '外侧框线')
  await expect.poll(async () => Object.keys((await cell(page))?.style.border ?? {}).sort()).toEqual(['left', 'top'])
  await expect.poll(async () => Object.keys((await cell(page, '1:1'))?.style.border ?? {}).sort()).toEqual(['bottom', 'right'])
  await click(page, '撤销')
  await expect.poll(async () => Object.keys((await cell(page))?.style.border ?? {}).length).toBe(0)
  await click(page, '重做')
  await expect.poll(async () => Object.keys((await cell(page))?.style.border ?? {}).length).toBe(2)
  await click(page, '清除格式')
  await expect.poll(async () => (await cell(page))?.style.font.strike).toBe(false)
  expect((await cell(page))?.value).toBe(1234.5)
  expect(errors).toEqual([])
})

test('number-format controls change rendered domain display, decimal places and persisted raw values', async ({ page }) => {
  await format(page, '货币')
  await expect.poll(async () => (await cell(page))?.display).toBe('¥1,234.50')
  await click(page, '增加小数位')
  await expect.poll(async () => (await cell(page))?.display).toBe('¥1,234.500')
  await click(page, '减少小数位')
  await expect.poll(async () => (await cell(page))?.display).toBe('¥1,234.50')
  await format(page, '会计专用')
  await expect.poll(async () => (await cell(page))?.display).toBe('¥ 1,234.50')
  await address(page, 'A2')
  await format(page, '百分比')
  await expect.poll(async () => (await cell(page, '1:0'))?.display).toBe('25%')
  await address(page, 'A3')
  await format(page, '日期和时间')
  await expect.poll(async () => (await cell(page, '2:0'))?.display).toBe('1900-03-01 12:00:00')
  await address(page, 'A4')
  await format(page, '货币')
  await expect.poll(async () => (await cell(page, '3:0'))?.value).toBe('1234.5')
  await expect.poll(async () => (await cell(page, '3:0'))?.display).toBe('1234.5')
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { univerAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(univerAPI.getActiveWorkbook()!.save())) as WorkbookModel
    univerAPI.disposeUnit(saved.id!)
    univerAPI.createWorkbook(saved)
  })
  await expect.poll(async () => (await cell(page))?.style.numberFormat).toContain('¥')
  expect((await cell(page))?.value).toBe(1234.5)
})

test('narrow Ribbon keeps new format controls executable through overflow and Pro insert controls reachable', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 900 })
  await click(page, '删除线')
  await expect.poll(async () => (await cell(page))?.style.font.strike).toBe(true)
  await format(page, '货币')
  await expect.poll(async () => (await cell(page))?.display).toBe('¥1,234.50')
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await click(page, '切片器')
  await expect(page.getByRole('dialog', { name: '切片器与时间线', exact: true })).toBeVisible()
})

test('function picker inserts an editable formula and participates in content history', async ({ page }) => {
  await address(page, 'C1')
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await click(page, '插入函数')
  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(dialog.getByRole('alert')).toContainText('请输入函数参数')
  await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill('A1:A2')
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(dialog).toBeHidden()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=SUM(A1:A2)')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBe('=SUM(A1:A2)')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(1234.75)
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).not.toBe('=SUM(A1:A2)')
  await click(page, '重做')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(1234.75)
})
