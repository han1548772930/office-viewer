import { expect, test, type Locator, type Page } from '@playwright/test'
import type { WorkbookModel } from '../../packages/office-studio/src/domain/model'
type Runtime = typeof import('../../examples/workbench/src/runtime')
type ViewAdapter = typeof import('../../packages/office-studio/src/services/excelViewAdapter')
type GridLayout = typeof import('../../packages/office-studio/src/rendering/spreadsheet/layout')

async function snapshot(page: Page): Promise<WorkbookModel> {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(path) as Runtime
    return JSON.parse(JSON.stringify(studio.getWorkbookModel(studio.activeWorkbookId))) as WorkbookModel
  })
}
async function historyState(page: Page): Promise<{ revision: number; undo: number; redo: number }> {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(path) as Runtime
    const history = studio.editor.history.units[studio.activeWorkbookId]
    if (!history) throw new Error('Active worksheet history is missing')
    return { revision: history.revision, undo: history.undo.length, redo: history.redo.length }
  })
}
async function visibleCellPoint(page: Page, row: number, column: number): Promise<{ x: number; y: number; hiddenHeights: number[] }> {
  return page.evaluate(async ({ row, column }) => {
    const resources = performance.getEntriesByType('resource')
    const loadedPath = (suffix: string): string => {
      const path = resources.filter(entry => new URL(entry.name).pathname.endsWith(suffix)).at(-1)?.name
      if (!path) throw new Error(`Expected a loaded editor module: ${suffix}`)
      return path
    }
    const { studio } = await import(loadedPath('/src/runtime.ts')) as Runtime
    const { toExcelWorkbookView } = await import(loadedPath('/services/excelViewAdapter.ts')) as ViewAdapter
    const { SpreadsheetGeometry } = await import(loadedPath('/rendering/spreadsheet/layout.ts')) as GridLayout
    const workbook = studio.getWorkbookModel(studio.activeWorkbookId)
    if (!workbook) throw new Error('Active workbook is missing')
    const view = toExcelWorkbookView(workbook), sheet = view.sheets.find(sheet => sheet.id === view.activeSheetId)
    const scroll = document.querySelector<HTMLElement>('.excel-workbench .canvas-scroll')
    if (!sheet || !scroll) throw new Error('Active Canvas viewport is missing')
    const geometry = new SpreadsheetGeometry(sheet)
    const rect = geometry.cellRect(row, column, { width: scroll.clientWidth, height: scroll.clientHeight, scrollLeft: scroll.scrollLeft, scrollTop: scroll.scrollTop, zoom: 1 })
    return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2, hiddenHeights: [geometry.rows.size(2), geometry.rows.size(4)] }
  }, { row, column })
}
async function seed(page: Page): Promise<void> {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeVisible()
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as Runtime
    const book = officeAPI.createWorkbook({ id: 'toolbar-data', name: '筛选与排序' })
    book.getActiveSheet()!.getRange('A1:C5').setValues([['地区', '销售', '货号'], ['东', 10, 'a'], ['西', 20, 'b'], ['东', 30, 'c'], ['西', 40, 'd']])
  })
  const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await name.fill('A1:C5'); await name.press('Enter')
}
async function open(page: Page, name: '筛选' | '自定义排序', choice: 'keep' | 'extend' = 'keep'): Promise<Locator> {
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '数据', exact: true }).click()
  if (name === '筛选') {
    const filter = ribbon.locator('[data-ribbon-item="sheet.data.filter"]')
    if (!(await snapshot(page)).sheets[0]!.autoFilter) await filter.getByRole('button', { name: '筛选', exact: true }).click()
    await filter.getByRole('button', { name: '筛选选项', exact: true }).click()
    await page.getByRole('menuitem', { name: '筛选条件', exact: true }).click()
  }
  if (name === '自定义排序') {
    await ribbon.locator('[data-ribbon-item="sheet.menu.sort"]').getByRole('button', { name: '排序', exact: true }).click()
    await page.getByRole('menuitem', { name, exact: true }).click()
    const reminder = page.getByRole('dialog', { name: '排序区域', exact: true })
    await expect(reminder).toBeVisible()
    await reminder.getByRole('radio', { name: choice === 'extend' ? '扩展选定区域' : '保持当前选区', exact: true }).check()
    await reminder.getByRole('button', { name: '确认区域', exact: true }).click()
  }
  const panel = page.getByRole('dialog', { name: '排序与筛选', exact: true })
  await expect(panel).toBeVisible()
  return panel
}
async function option(page: Page, control: Locator, label: string): Promise<void> {
  await control.click(); await page.locator('.ant-select-item-option').filter({ hasText: label }).click()
}

test('toolbar value filtering, clearing, undo and snapshot reopening use the live worksheet', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  await seed(page)
  const panel = await open(page, '筛选')
  await panel.getByRole('checkbox', { name: '西', exact: true }).uncheck()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([2, 4])
  await expect(panel.getByRole('checkbox', { name: '西', exact: true })).not.toBeChecked()
  await panel.locator('.ant-modal-close').click()
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  const point = await visibleCellPoint(page, 3, 0)
  expect(point.hiddenHeights).toEqual([0, 0])
  // The third painted row is logical row 4 after West rows 3 and 5 collapse.
  // A model-only hidden list would still hit A3 here, making this assertion fail.
  await grid.click({ position: { x: point.x, y: point.y } })
  const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await expect(name).toHaveValue('A4')
  await expect(grid.getByRole('gridcell')).toContainText('东')
  await grid.press('ArrowUp'); await expect(name).toHaveValue('A2')
  await grid.press('ArrowDown'); await expect(name).toHaveValue('A4')
  await grid.press('ArrowDown'); await expect(name).toHaveValue('A6')
  await open(page, '筛选')
  await expect(panel.getByRole('checkbox', { name: '西', exact: true })).not.toBeChecked()
  await panel.getByRole('button', { name: '全不选', exact: true }).click()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 2, 3, 4])
  await panel.getByRole('button', { name: '清除全部条件', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([])
  await panel.locator('.ant-modal-close').click()
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '开始', exact: true }).click()
  await ribbon.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 2, 3, 4])
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as Runtime
    const original = officeAPI.getActiveWorkbook()!, data = JSON.parse(JSON.stringify(original.save())) as WorkbookModel
    officeAPI.disposeUnit(original.getId()); officeAPI.createWorkbook(data)
  })
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 2, 3, 4])
  expect(errors).toEqual([])
})

test('filter heading controls open the clicked column and apply multiple colors through one history transaction', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  await seed(page)
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as Runtime
    const sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('B2').setBackground('#ff0000'); sheet.getRange('B4').setBackground('#00ff00')
    sheet.setAutoFilter('A1:C5')
  })
  const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await name.fill('C8'); await name.press('Enter')
  const columnButton = page.getByRole('button', { name: '筛选第 2 列', exact: true })
  await expect(columnButton).toBeVisible(); await columnButton.click()
  const panel = page.getByRole('dialog', { name: '排序与筛选', exact: true })
  await expect(panel.getByRole('combobox', { name: '筛选字段', exact: true })).toBeVisible()
  await option(page, panel.getByRole('combobox', { name: '筛选方式', exact: true }), '按颜色筛选')
  const before = await snapshot(page), history = await historyState(page)
  await panel.getByRole('checkbox', { name: '填充颜色 #ff0000', exact: true }).check()
  await panel.getByRole('checkbox', { name: '填充颜色 #00ff00', exact: true }).check()
  expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(history)
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([2, 4])
  expect((await snapshot(page)).sheets[0]!.autoFilter?.columns[1]?.kind).toBe('color')
  expect((await historyState(page)).undo).toBe(history.undo + 1)
  await panel.locator('.ant-modal-close').click()
  await expect(columnButton).toHaveAttribute('aria-pressed', 'true')
  await expect(name).toHaveValue('C8')
  await columnButton.click()
  await expect(panel.getByRole('checkbox', { name: '填充颜色 #ff0000', exact: true })).toBeChecked()
  await panel.getByRole('checkbox', { name: '文字颜色 #000000', exact: true }).check()
  await expect(panel.getByRole('checkbox', { name: '填充颜色 #ff0000', exact: true })).not.toBeChecked()
  await expect(panel.getByRole('checkbox', { name: '填充颜色 #00ff00', exact: true })).not.toBeChecked()
  await panel.getByRole('checkbox', { name: '文字颜色 #000000', exact: true }).uncheck()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.columns).toEqual({})
  expect(errors).toEqual([])
})

test('value search limits submission and select-all while draft choices leave history unchanged', async ({ page }) => {
  await seed(page)
  const panel = await open(page, '筛选')
  const search = panel.getByRole('textbox', { name: '搜索筛选值', exact: true })
  const before = await snapshot(page)
  const beforeHistory = await historyState(page)
  await search.fill('东')
  await expect(panel.getByRole('checkbox', { name: '东', exact: true })).toBeChecked()
  await expect(panel.getByRole('checkbox', { name: '西', exact: true })).toHaveCount(0)
  expect(await snapshot(page)).toEqual(before)
  expect(await historyState(page)).toEqual(beforeHistory)
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([2, 4])
  expect((await snapshot(page)).sheets[0]!.autoFilter?.columns[0]).toEqual({ kind: 'values', selectedItems: ['东'] })
  expect((await snapshot(page)).sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  expect((await historyState(page)).undo).toBe(beforeHistory.undo + 1)
  const filtered = await snapshot(page)
  const filteredHistory = await historyState(page)
  await search.fill('西')
  await panel.getByRole('button', { name: '全选', exact: true }).click()
  await search.fill('')
  await expect(panel.getByRole('checkbox', { name: '东', exact: true })).not.toBeChecked()
  await expect(panel.getByRole('checkbox', { name: '西', exact: true })).toBeChecked()
  await expect(panel.locator('.data-actions').first()).toContainText('1 / 2 个值')
  expect(await snapshot(page)).toEqual(filtered)
  expect(await historyState(page)).toEqual(filteredHistory)
  await search.fill('西')
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 3])
  expect((await snapshot(page)).sheets[0]!.autoFilter?.columns[0]).toEqual({ kind: 'values', selectedItems: ['西'] })
  expect((await historyState(page)).undo).toBe(filteredHistory.undo + 1)
  await panel.locator('.ant-modal-close').click()
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '开始', exact: true }).click()
  await ribbon.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([2, 4])
  await ribbon.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 3])
})

test('condition filtering rejects invalid numeric input without mutating content or history', async ({ page }) => {
  await seed(page)
  const panel = await open(page, '筛选')
  await option(page, panel.getByRole('combobox', { name: /筛选字段/ }), 'B · 销售')
  await option(page, panel.getByRole('combobox', { name: /筛选方式/ }), '按条件筛选')
  await option(page, panel.getByRole('combobox', { name: /筛选条件/ }), '大于或等于')
  await option(page, panel.getByRole('combobox', { name: /比较值类型/ }), '数值或日期序号')
  await panel.getByRole('textbox', { name: '筛选比较值', exact: true }).fill('invalid')
  const before = await snapshot(page)
  const beforeHistory = await historyState(page)
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect(panel.getByRole('alert')).toContainText('有效数值')
  expect(await snapshot(page)).toEqual(before)
  expect(await historyState(page)).toEqual(beforeHistory)
  await panel.getByRole('textbox', { name: '筛选比较值', exact: true }).fill('30')
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 2])
  await option(page, panel.getByRole('combobox', { name: /筛选字段/ }), 'A · 地区')
  await panel.getByRole('checkbox', { name: '西', exact: true }).uncheck()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 2, 4])
  expect(Object.keys((await snapshot(page)).sheets[0]!.autoFilter!.columns)).toEqual(['0', '1'])
  await option(page, panel.getByRole('combobox', { name: /筛选字段/ }), 'B · 销售')
  await expect(panel.getByRole('textbox', { name: '筛选比较值', exact: true })).toHaveValue('30')
  await panel.getByRole('button', { name: '清除此列条件', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([2, 4])
  expect(Object.keys((await snapshot(page)).sheets[0]!.autoFilter!.columns)).toEqual(['0'])
})

test('the custom two-condition form applies AND and OR atomically and preserves failed numeric drafts', async ({ page }) => {
  await seed(page)
  const panel = await open(page, '筛选')
  await option(page, panel.getByRole('combobox', { name: '筛选字段', exact: true }), 'B · 销售')
  await option(page, panel.getByRole('combobox', { name: '筛选方式', exact: true }), '自定义复合条件')
  await option(page, panel.getByRole('combobox', { name: '自定义条件 1', exact: true }), '大于')
  await option(page, panel.getByRole('combobox', { name: '自定义比较值 1 类型', exact: true }), '数值或日期序号')
  await panel.getByRole('textbox', { name: '自定义比较值 1', exact: true }).fill('15')
  await option(page, panel.getByRole('combobox', { name: '自定义条件 2', exact: true }), '小于')
  await option(page, panel.getByRole('combobox', { name: '自定义比较值 2 类型', exact: true }), '数值或日期序号')
  await panel.getByRole('textbox', { name: '自定义比较值 2', exact: true }).fill('invalid')
  const before = await snapshot(page), history = await historyState(page)
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect(panel.getByRole('alert')).toContainText('有效数值')
  await expect(panel.getByRole('textbox', { name: '自定义比较值 2', exact: true })).toHaveValue('invalid')
  expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(history)
  await panel.getByRole('textbox', { name: '自定义比较值 2', exact: true }).fill('35')
  await panel.getByRole('radio', { name: '并且', exact: true }).check()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([1, 4])
  expect((await historyState(page)).undo).toBe(history.undo + 1)
  await panel.getByRole('radio', { name: '或者', exact: true }).check()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([])
  await panel.locator('.ant-modal-close').click()
  await page.getByRole('button', { name: '筛选第 2 列', exact: true }).click()
  await expect(panel.getByRole('radio', { name: '或者', exact: true })).toBeChecked()
  await expect(panel.getByRole('textbox', { name: '自定义比较值 1', exact: true })).toHaveValue('15')
  await expect(panel.getByRole('textbox', { name: '自定义比较值 2', exact: true })).toHaveValue('35')
})

test('date tree group checks preserve both timestamps and cancellation leaves model and history unchanged', async ({ page }) => {
  await seed(page)
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as Runtime
    const sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('B2:B5').setValues([[45292], [45292.5], [45293], [45323]])
    sheet.getRange('B2:B5').setNumberFormat('yyyy-mm-dd hh:mm')
  })
  const panel = await open(page, '筛选')
  await option(page, panel.getByRole('combobox', { name: '筛选字段', exact: true }), 'B · 销售')
  const tree = panel.getByRole('tree', { name: '筛选值分组', exact: true })
  await expect(tree).toBeVisible()
  const before = await snapshot(page), history = await historyState(page)
  await panel.getByRole('button', { name: '全不选', exact: true }).click()
  await tree.locator('.ant-tree-treenode').filter({ has: page.locator('.ant-tree-node-content-wrapper').filter({ hasText: /^2024（4）$/u }) }).locator('.ant-tree-checkbox').click()
  expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(history)
  await panel.locator('.ant-modal-close').click()
  expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(history)
  await open(page, '筛选')
  await option(page, panel.getByRole('combobox', { name: '筛选字段', exact: true }), 'B · 销售')
  await panel.getByRole('textbox', { name: '搜索筛选值', exact: true }).fill('2024-01-01')
  await panel.getByRole('button', { name: '全选', exact: true }).click()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.columns[1]).toEqual({ kind: 'values', selectedItems: [45292, 45292.5] })
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([3, 4])
  expect((await historyState(page)).undo).toBe(history.undo + 1)
})

test('composition keeps the custom filter draft and disables changing its rule or submitting it until input ends', async ({ page }) => {
  await seed(page)
  const panel = await open(page, '筛选')
  await option(page, panel.getByRole('combobox', { name: '筛选方式', exact: true }), '自定义复合条件')
  await option(page, panel.getByRole('combobox', { name: '自定义条件 1', exact: true }), '包含')
  const input = panel.getByRole('textbox', { name: '自定义比较值 1', exact: true })
  await input.fill('东')
  const before = await snapshot(page), history = await historyState(page)
  await input.dispatchEvent('compositionstart', { data: '' })
  await expect(panel.getByRole('combobox', { name: '自定义条件 1', exact: true })).toBeDisabled()
  await expect(panel.getByRole('combobox', { name: '自定义条件 2', exact: true })).toBeDisabled()
  await expect(panel.getByRole('button', { name: '应用筛选', exact: true })).toBeDisabled()
  await input.press('Enter')
  await expect(input).toHaveValue('东')
  expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(history)
  await input.dispatchEvent('compositionend', { data: '东' })
  await expect(panel.getByRole('button', { name: '应用筛选', exact: true })).toBeEnabled()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.filteredRows).toEqual([2, 4])
})

test('custom multi-key sorting keeps headers and complete rows and is one reversible action', async ({ page }) => {
  await seed(page)
  const panel = await open(page, '自定义排序')
  await panel.getByRole('button', { name: '添加排序条件', exact: true }).click()
  await option(page, panel.getByRole('combobox', { name: /排序方向 2/ }), '降序')
  const before = await snapshot(page), beforeHistory = await historyState(page)
  await panel.getByRole('button', { name: '应用排序', exact: true }).click()
  await expect(panel.getByRole('status')).toContainText('已完成排序')
  await expect.poll(async () => { const model = await snapshot(page); return [model.sheets[0]!.cells['1:2']?.value, model.sheets[0]!.cells['2:2']?.value, model.sheets[0]!.cells['3:2']?.value, model.sheets[0]!.cells['4:2']?.value] }).toEqual(['c', 'a', 'd', 'b'])
  expect((await snapshot(page)).sheets[0]!.cells['0:0']?.value).toBe('地区')
  expect((await historyState(page)).undo).toBe(beforeHistory.undo + 1)
  const sorted = await snapshot(page)
  await panel.locator('.ant-modal-close').click()
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '开始', exact: true }).click()
  await ribbon.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.cells['1:2']?.value).toBe('a')
  expect((await snapshot(page)).sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  await ribbon.getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.cells['1:2']?.value).toBe('c')
  expect((await snapshot(page)).sheets[0]!.cells).toEqual(sorted.sheets[0]!.cells)
})

test('custom range cancellation leaves content/history/selection unchanged and extended sorting keeps the original active column', async ({ page }) => {
  await seed(page)
  const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await name.fill('B2'); await name.press('Enter')
  const before = await snapshot(page), depth = await historyState(page), ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '数据', exact: true }).click()
  await ribbon.locator('[data-ribbon-item="sheet.menu.sort"]').getByRole('button', { name: '排序', exact: true }).click()
  await page.getByRole('menuitem', { name: '自定义排序', exact: true }).click()
  const reminder = page.getByRole('dialog', { name: '排序区域', exact: true })
  await expect(reminder).toBeVisible(); await reminder.getByRole('radio', { name: '扩展选定区域', exact: true }).check()
  expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(depth)
  await reminder.getByRole('button', { name: '取消', exact: true }).click(); await expect(reminder).toBeHidden()
  await expect(name).toHaveValue('B2'); expect(await snapshot(page)).toEqual(before); expect(await historyState(page)).toEqual(depth)
  const panel = await open(page, '自定义排序', 'extend')
  await expect(panel.getByRole('textbox', { name: '排序筛选数据区域', exact: true })).toHaveValue('A1:C5')
  await expect(panel.locator('.data-sort-key').first()).toContainText('B · 销售')
  await option(page, panel.getByRole('combobox', { name: '排序方向 1', exact: true }), '降序')
  await panel.getByRole('button', { name: '应用排序', exact: true }).click()
  await expect(panel.getByRole('status')).toContainText('已完成排序')
  const sorted = await snapshot(page)
  expect([1, 2, 3, 4].map(row => sorted.sheets[0]!.cells[`${row}:2`]?.value)).toEqual(['d', 'c', 'b', 'a'])
  expect(sorted.sheets[0]!.cells['0:1']?.value).toBe('销售'); expect((await historyState(page)).undo).toBe(depth.undo + 1)
})

test('a single-row smart filter extends downward without absorbing adjacent columns', async ({ page }) => {
  await seed(page)
  const name = page.getByRole('textbox', { name: '名称框', exact: true }), before = await snapshot(page), depth = await historyState(page)
  await name.fill('A1:B1'); await name.press('Enter')
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '数据', exact: true }).click()
  await ribbon.locator('[data-ribbon-item="sheet.data.filter"]').getByRole('button', { name: '筛选', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).sheets[0]!.autoFilter?.range).toEqual({ startRow: 0, endRow: 4, startCol: 0, endCol: 1 })
  expect((await snapshot(page)).sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  expect((await historyState(page)).undo).toBe(depth.undo + 1)
})
