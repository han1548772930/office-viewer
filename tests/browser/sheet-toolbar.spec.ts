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
    const { officeAPI } = await import(path) as RuntimeModule
    const workbook = officeAPI.createWorkbook({ id: 'sheet-toolbar-workbook', name: '格式与顶部交互' })
    const sheet = workbook.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([[1234.5, 10], [.25, 20], [61.5, 30], ['1234.5', 40]])
  })
  await address(page, 'A1')
})

test('View tab separates display and visibility tools and toggles real fullscreen state', async ({ page }) => {
  await ribbon(page).getByRole('tab', { name: '视图', exact: true }).click()
  await expect(ribbon(page).getByRole('group', { name: '显示', exact: true })).toBeVisible()
  await expect(ribbon(page).getByRole('group', { name: '可见性', exact: true })).toBeVisible()
  await expect(ribbon(page).getByRole('button', { name: '冻结窗格', exact: true })).toBeVisible()
  await expect(ribbon(page).getByRole('button', { name: '网格线', exact: true })).toBeVisible()
  await expect(ribbon(page).getByRole('combobox', { name: '缩放', exact: true })).toBeVisible()
  await click(page, '全屏')
  await expect.poll(() => page.evaluate(() => document.fullscreenElement?.classList.contains('excel-workbench') ?? false)).toBe(true)
  await click(page, '退出全屏')
  await expect.poll(() => page.evaluate(() => document.fullscreenElement === null)).toBe(true)
  await expect(ribbon(page).getByRole('button', { name: '全屏', exact: true })).toBeVisible()
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

test('format painter applies once to the completed drag range and participates in history', async ({ page }) => {
  await click(page, '鍔犵矖')
  await address(page, 'A1')
  await click(page, '鏍煎紡鍒?')

  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  const bounds = await grid.boundingBox()
  expect(bounds).not.toBeNull()
  if (!bounds) return
  const start = { x: bounds.x + 44 + 152 + 56, y: bounds.y + 26 + 26 + 13 }
  const end = { x: bounds.x + 44 + 152 + 112 + 56, y: bounds.y + 26 + 26 + 26 + 13 }
  await page.mouse.move(start.x, start.y)
  await page.mouse.down()
  await page.mouse.move(end.x, end.y, { steps: 4 })
  await page.mouse.up()

  for (const key of ['1:1', '2:1', '1:2', '2:2']) {
    await expect.poll(async () => (await cell(page, key))?.style.font.bold).toBe(true)
  }
  await click(page, '鎾ら攢')
  await expect.poll(async () => (await cell(page, '1:1'))?.style.font.bold).not.toBe(true)
  await click(page, '閲嶅仛')
  await expect.poll(async () => (await cell(page, '2:2'))?.style.font.bold).toBe(true)
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
    const { officeAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(officeAPI.getActiveWorkbook()!.save())) as WorkbookModel
    officeAPI.disposeUnit(saved.id!)
    officeAPI.createWorkbook(saved)
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

test('function category menu opens a picker filtered to supported formulas', async ({ page }) => {
  await address(page, 'C1')
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '插入函数选项', exact: true }).click()
  await page.getByRole('menu', { name: '插入函数', exact: true }).getByRole('button', { name: '逻辑', exact: true }).click()

  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await expect(dialog).toBeVisible()
  await expect(dialog.locator('p').first()).toContainText('逻辑')
  await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill('A1>0,"yes","no"')
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(dialog).toBeHidden()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=IF(A1>0,"yes","no")')
})

test('Data Ribbon applies STOP validation and blocks invalid cell entry without changing the model', async ({ page }) => {
  await address(page, 'A1:A2')
  await ribbon(page).getByRole('tab', { name: '数据', exact: true }).click()
  await click(page, '数据验证')

  const dialog = page.getByRole('dialog', { name: '数据验证', exact: true })
  await expect(dialog).toBeVisible()
  await dialog.locator('#sheet-validation-type').click()
  await page.locator('.ant-select-dropdown:visible').getByText('整数', { exact: true }).click()
  await dialog.locator('#sheet-validation-operator').click()
  await page.locator('.ant-select-dropdown:visible').getByText('大于', { exact: true }).click()
  await dialog.getByRole('textbox', { name: '数据验证规则值', exact: true }).fill('10')
  await dialog.locator('.ant-modal-footer .ant-btn-primary').click()
  await expect(dialog).toBeHidden()
  await expect.poll(async () => (await state(page)).sheets[0]?.dataValidations?.length).toBe(1)

  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  await address(page, 'A1')
  await grid.press('7')
  await page.locator('.canvas-cell-input').fill('12')
  await page.locator('.canvas-cell-input').press('Enter')
  await expect.poll(async () => (await cell(page, '0:0'))?.value).toBe(12)

  await address(page, 'A2')
  await grid.press('5')
  await page.locator('.canvas-cell-input').fill('5')
  await page.locator('.canvas-cell-input').press('Enter')
  await expect(page.getByRole('alert')).toContainText('数据验证未通过')
  await expect.poll(async () => (await cell(page, '1:0'))?.value).toBe(0.25)
})

test('data-validation list options render on the cell input bridge and persist after save/reopen', async ({ page }) => {
  await address(page, 'B1:B2')
  await ribbon(page).getByRole('tab', { name: '数据', exact: true }).click()
  await click(page, '数据验证')
  const dialog = page.getByRole('dialog', { name: '数据验证', exact: true })
  await dialog.locator('#sheet-validation-type').click()
  await page.locator('.ant-select-dropdown:visible').getByText('下拉列表', { exact: true }).click()
  await dialog.getByRole('textbox', { name: '数据验证列表选项', exact: true }).fill('Draft,Review,Approved')
  await dialog.locator('.ant-modal-footer .ant-btn-primary').click()
  await expect(dialog).toBeHidden()

  await address(page, 'B1')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  await grid.press('A')
  const input = page.locator('.canvas-cell-input')
  await expect(input).toHaveAttribute('list', /^sheet-validation-options-/)
  const listId = await input.getAttribute('list')
  await expect(page.locator(`#${listId} option`)).toHaveCount(3)

  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(officeAPI.getActiveWorkbook()!.save())) as WorkbookModel
    officeAPI.disposeUnit(saved.id!)
    officeAPI.createWorkbook(saved)
  })
  await address(page, 'B1')
  await grid.press('A')
  await expect(page.locator('.canvas-cell-input')).toHaveAttribute('list', /^sheet-validation-options-/)
})

test('text rotation, vertical text and shrink-to-fit redraw and participate in undo/redo', async ({ page }) => {
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('A1').setValues([['A long heading that should be compressed to fit this worksheet column']])
  })
  await address(page, 'A1')
  const canvas = page.locator('.canvas-surface canvas')
  await expect(canvas).toBeVisible()
  await page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))))
  const initialPixels = await canvas.screenshot()

  await menu(page, '文字方向', '左斜 45°')
  await expect.poll(async () => (await cell(page))?.style.rotation).toBe(-45)
  const rotatedPixels = await canvas.screenshot()
  expect(rotatedPixels.equals(initialPixels)).toBe(false)
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page))?.style.rotation ?? 0).toBe(0)
  expect((await canvas.screenshot()).equals(initialPixels)).toBe(true)
  await click(page, '重做')
  await expect.poll(async () => (await cell(page))?.style.rotation).toBe(-45)
  expect((await canvas.screenshot()).equals(rotatedPixels)).toBe(true)

  await menu(page, '文字方向', '竖排')
  await expect.poll(async () => (await cell(page))?.style.rotationVertical).toBe(true)
  const verticalPixels = await canvas.screenshot()
  expect(verticalPixels.equals(rotatedPixels)).toBe(false)
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page))?.style.rotationVertical ?? false).toBe(false)
  await click(page, '重做')
  await expect.poll(async () => (await cell(page))?.style.rotationVertical).toBe(true)

  await click(page, '自动换行')
  await expect.poll(async () => (await cell(page))?.style.wrap).toBe(true)
  const shrink = await control(page, '缩小字体填充')
  await expect(shrink).toBeDisabled()
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page))?.style.wrap ?? false).toBe(false)
  await click(page, '重做')
  await expect.poll(async () => (await cell(page))?.style.wrap).toBe(true)
  await click(page, '自动换行')
  await expect.poll(async () => (await cell(page))?.style.wrap).toBe(false)
  await menu(page, '文字方向', '水平')
  await expect.poll(async () => (await cell(page))?.style.rotationVertical ?? false).toBe(false)
  const horizontalPixels = await canvas.screenshot()
  await click(page, '缩小字体填充')
  await expect.poll(async () => (await cell(page))?.style.shrinkToFit).toBe(true)
  const fittedPixels = await canvas.screenshot()
  expect(fittedPixels.equals(horizontalPixels)).toBe(false)
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page))?.style.shrinkToFit ?? false).toBe(false)
  await click(page, '重做')
  await expect.poll(async () => (await cell(page))?.style.shrinkToFit).toBe(true)
})
