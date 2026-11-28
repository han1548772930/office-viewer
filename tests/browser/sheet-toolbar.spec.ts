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
  const menu = page.getByRole('dialog', { name: '更多功能', exact: true })
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
async function functionParameters(page: Page, name: string): Promise<Locator> {
  await click(page, '常用与统计')
  await page.getByRole('menu', { name: '常用与统计', exact: true }).getByRole('menuitem', { name: '插入函数', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await dialog.getByRole('combobox', { name: '函数名称', exact: true }).click()
  await dialog.getByRole('combobox', { name: '函数名称', exact: true }).fill(name)
  await page.locator('.ant-select-dropdown:visible .ant-select-item-option').filter({ hasText: `${name} · ` }).click()
  return dialog
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

test('history grid and continuous painter preserve data across two applications and cancel on Escape', async ({ page }) => {
  const painter = await control(page, '格式刷')
  const undo = ribbon(page).getByRole('button', { name: '撤销', exact: true })
  await expect(undo.locator('svg')).toHaveCSS('width', '18px')
  await expect(painter).toHaveText('')
  await click(page, '加粗')
  await painter.dblclick()
  await expect(painter).toHaveAttribute('aria-pressed', 'true')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  await grid.press('ArrowRight')
  await expect.poll(async () => (await cell(page, '0:1'))?.style.font.bold).toBe(true)
  await grid.press('ArrowDown')
  await expect.poll(async () => (await cell(page, '1:1'))?.style.font.bold).toBe(true)
  expect((await cell(page, '1:1'))?.value).toBe(20)
  await grid.press('Escape')
  await expect(painter).toHaveAttribute('aria-pressed', 'false')
  await undo.click()
  await expect.poll(async () => Boolean((await cell(page, '1:1'))?.style.font.bold)).toBe(false)
  expect((await cell(page, '0:1'))?.style.font.bold).toBe(true)
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
  await click(page, '边框选项')
  await page.getByRole('button', { name: '外侧框线', exact: true }).click()
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

test('the explicit function picker preserves parameter validation and saves the confirmed formula through history', async ({ page }) => {
  await address(page, 'C1')
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await click(page, '常用与统计')
  const commonMenu = page.getByRole('menu', { name: '常用与统计', exact: true })
  await expect(commonMenu.getByRole('menuitem')).toHaveText(['SUMIF', 'SUM', 'AVERAGE', 'IF', 'COUNT', 'SIN', 'MAX', '插入函数'])
  await commonMenu.getByRole('menuitem', { name: '插入函数', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await expect(dialog.getByText('SUM · 求和', { exact: true })).toBeVisible()
  await expect(dialog.locator('p').first()).toContainText('数学')
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(dialog.getByRole('alert')).toContainText('请输入函数参数')
  await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill('A1:A2')
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(dialog).toBeHidden()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=SUM(A1:A2)')
  expect((await cell(page, '0:2'))?.formula).not.toBe('=SUM(A1:A2)')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBe('=SUM(A1:A2)')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(1234.75)
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).not.toBe('=SUM(A1:A2)')
  await click(page, '重做')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(1234.75)
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(officeAPI.getActiveWorkbook()!.save())) as WorkbookModel
    officeAPI.disposeUnit(saved.id!)
    officeAPI.createWorkbook(saved)
  })
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBe('=SUM(A1:A2)')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(1234.75)
})

test('function categories insert batch totals in one reversible history action and retain the selected source cells', async ({ page }) => {
  await address(page, 'A1:B4')
  const before = await state(page)
  const depth = await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(path) as RuntimeModule
    return studio.editor.history.units[studio.activeWorkbookId]!.undo.length
  })
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await click(page, '常用与统计')
  await page.getByRole('menu', { name: '常用与统计', exact: true }).getByRole('menuitem', { name: 'SUM', exact: true }).click()
  await expect(page.getByRole('dialog', { name: '插入函数', exact: true })).toBeHidden()
  await expect.poll(async () => (await cell(page, '4:0'))?.formula).toBe('=SUM(A1:A4)')
  await expect.poll(async () => (await cell(page, '4:1'))?.formula).toBe('=SUM(B1:B4)')
  await expect.poll(async () => (await cell(page, '4:1'))?.value).toBe(100)
  expect((await state(page)).sheets[0]!.cells['3:0']?.value).toBe(before.sheets[0]!.cells['3:0']?.value)
  expect(await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(path) as RuntimeModule
    return studio.editor.history.units[studio.activeWorkbookId]!.undo.length
  })).toBe(depth + 1)
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page, '4:0'))?.formula).toBeUndefined()
  await click(page, '重做')
  await expect.poll(async () => (await cell(page, '4:1'))?.value).toBe(100)
})

test('Alt equals opens an uncommitted automatic sum in the formula input and commits through its existing input session', async ({ page }) => {
  await address(page, 'B5')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  await grid.focus(); await grid.press('Alt+=')
  const formula = page.getByRole('textbox', { name: '公式编辑', exact: true })
  await expect(formula).toHaveValue('=SUM(B1:B4')
  expect((await cell(page, '4:1'))?.formula).toBeUndefined()
  await formula.fill('=SUM(B1:B4)')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(async () => (await cell(page, '4:1'))?.value).toBe(100)
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page, '4:1'))?.formula).toBeUndefined()
})

test('function categories start automatic range editing while the explicit picker retains scalar and zero-argument parameters', async ({ page }) => {
  await address(page, 'C1')
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await click(page, '数学')
  await page.getByRole('menu', { name: '数学', exact: true }).getByRole('menuitem', { name: 'ABS', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await expect(dialog).toBeHidden()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=ABS(A1:B1')
  await page.getByRole('button', { name: '取消编辑', exact: true }).click()
  await expect(page.getByRole('button', { name: '确认编辑', exact: true })).toBeDisabled()
  expect((await cell(page, '0:2'))?.formula).toBeUndefined()

  await functionParameters(page, 'IF')
  await expect(dialog.getByText('IF · 条件判断', { exact: true })).toBeVisible()
  await expect(dialog.locator('p').first()).toContainText('逻辑')
  await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill('A1>0,"yes","no"')
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(dialog).toBeHidden()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=IF(A1>0,"yes","no")')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe('yes')
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBeUndefined()
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()

  for (const [category, name, label, argumentsText, expected] of [
    ['文本', 'LEFT', '截取左侧文本', '"Univer",3', 'Uni'],
    ['文本', 'MID', '按位置截取文本', '"Univer",2,3', 'niv'],
    ['文本', 'RIGHT', '截取右侧文本', '"Univer",3', 'ver'],
    ['文本', 'VALUE', '文本转数值', '"25%"', 0.25],
    ['逻辑', 'PERCENTOF', '子集占比', '2,5', 0.4],
    ['统计', 'AVEDEV', '平均绝对偏差', '2,4', 1],
    ['统计', 'DEVSQ', '偏差平方和', '2,4', 2],
    ['统计', 'GEOMEAN', '几何平均值', '4,9', 6],
    ['统计', 'HARMEAN', '调和平均值', '2,6', 3],
    ['统计', 'MODE.SNGL', '单一众数', '2,4,4', 4],
    ['日期与时间', 'WEEKDAY', '星期序号', '"2008-2-14",2', 4],
    ['数学', 'DEGREES', '弧度转角度', 'PI()', 180],
    ['数学', 'RADIANS', '角度转弧度', '270', 4.71238898038469],
    ['数学', 'ACOS', '反余弦', '0.5', 1.0471975511965979],
    ['数学', 'ACOSH', '反双曲余弦', '2', 1.3169578969248166],
    ['数学', 'ACOT', '反余切', '-2', 2.677945044588987],
    ['数学', 'ACOTH', '反双曲余切', '2', 0.5493061443340549],
    ['数学', 'ASIN', '反正弦', '0.5', 0.5235987755982989],
    ['数学', 'ASINH', '反双曲正弦', '10', 2.99822295029797],
    ['数学', 'ATAN', '反正切', '1', 0.7853981633974483],
    ['数学', 'ATANH', '反双曲正切', '0.5', 0.5493061443340548],
    ['数学', 'SIN', '正弦', '1', 0.8414709848078965],
    ['数学', 'COS', '余弦', '1', 0.5403023058681398],
    ['数学', 'TAN', '正切', '1', 1.5574077246549023],
    ['数学', 'COT', '余切', '1', 0.642092615934331],
    ['数学', 'SEC', '正割', '1', 1.8508157176809255],
    ['数学', 'CSC', '余割', '1', 1.188395105778121],
    ['数学', 'SINH', '双曲正弦', '1', 1.1752011936438014],
    ['数学', 'COSH', '双曲余弦', '1', 1.5430806348152437],
    ['数学', 'TANH', '双曲正切', '1', 0.7615941559557649],
    ['数学', 'COTH', '双曲余切', '1', 1.3130352854993315],
    ['数学', 'SECH', '双曲正割', '1', 0.6480542736638855],
    ['数学', 'CSCH', '双曲余割', '1', 0.8509181282393216],
    ['数学', 'EXP', '自然指数', '1', 2.718281828459045],
    ['数学', 'LN', '自然对数', 'EXP(3)', 3],
    ['数学', 'LOG', '指定底数对数', '8,2', 3],
    ['数学', 'LOG10', '常用对数', '100000', 5],
    ['数学', 'SIGN', '数值符号', '-10', -1],
    ['数学', 'ATAN2', '坐标反正切', '4,3', 0.6435011087932844],
    ['数学', 'EVEN', '向外取偶数', '1.5', 2],
    ['数学', 'ODD', '向外取奇数', '1.5', 3],
    ['数学', 'QUOTIENT', '整数商', '5,2', 2],
    ['数学', 'SQRTPI', 'π 倍数的平方根', '1', 1.7724538509055159],
    ['数学', 'FACT', '阶乘', '5', 120],
    ['数学', 'FACTDOUBLE', '双阶乘', '6', 48],
    ['数学', 'COMBIN', '组合数', '8,2', 28],
    ['数学', 'COMBINA', '可重复组合数', '8,2', 36],
    ['数学', 'ROUND', '四舍五入', '2.15,1.9', 2.2],
    ['数学', 'ROUNDBANK', '舍入到最近偶数', '2.345,2', 2.34],
    ['数学', 'ROUNDDOWN', '向零舍入', '3.2,0', 3],
    ['数学', 'ROUNDUP', '向外舍入', '3.2,0', 4],
    ['数学', 'TRUNC', '截去小数', '0.45,1', 0.4],
    ['数学', 'CEILING', '向上舍入到倍数', '2.5,1', 3],
    ['数学', 'CEILING.MATH', '按方向向上舍入', '-5.5,2,-1', -6],
    ['数学', 'CEILING.PRECISE', '向上舍入到绝对倍数', '4.3,2', 6],
    ['数学', 'FLOOR', '向下舍入到倍数', '3.7,2', 2],
    ['数学', 'FLOOR.MATH', '按方向向下舍入', '-5.5,2,-1', -4],
    ['数学', 'FLOOR.PRECISE', '向下舍入到绝对倍数', '-3.2,-1', -4],
    ['数学', 'ISO.CEILING', 'ISO 向上舍入', '-4.3,2', -4],
    ['数学', 'GCD', '最大公约数', '5,2', 1],
    ['数学', 'LCM', '最小公倍数', '5,2', 10],
    ['数学', 'MULTINOMIAL', '多项式系数', '5,2', 21],
    ['数学', 'MROUND', '舍入到最近倍数', '10,3', 9],
    ['数学', 'ABS', '绝对值', '"-5"', 5],
    ['数学', 'SQRT', '平方根', '"16"', 4],
    ['数学', 'INT', '向下取整', '5.999999999999999', 6],
    ['数学', 'MOD', '余数', '"5","2"', 1],
    ['数学', 'POWER', '乘方', '"5","2"', 25],
    ['数学', 'SUMSQ', '平方和', '3,4', 25],
    ['数学', 'SUMX2MY2', '平方差之和', '3,4', -7],
    ['数学', 'SUMX2PY2', '平方和之和', '3,4', 25],
    ['数学', 'SUMXMY2', '差平方之和', '3,4', 1],
    ['数学', 'SUMPRODUCT', '乘积求和', '3,4', 12],
    ['数学', 'SERIESSUM', '幂级数求和', '2,1,1,3', 6],
    ['数学', 'ARABIC', '罗马数字转数值', '"MCMXII"', 1912],
    ['数学', 'ROMAN', '数值转罗马数字', '499,0', 'CDXCIX'],
    ['数学', 'BASE', '数值转进制文本', '15,2,10', '0000001111'],
    ['数学', 'DECIMAL', '进制文本转数值', '"FF",16', 255],
    ['数学', 'RANDBETWEEN', '区间随机整数', '7,7', 7],
    ['数学', 'MDETERM', '矩阵行列式', '5', 5],
    ['信息', 'ERROR.TYPE', '错误类型编号', '"#NULL!"', 1],
    ['信息', 'ISBETWEEN', '检查区间', '7.9,1.2,12.45', true],
    ['信息', 'ISDATE', '检查日期和时间', '"2012-2-2"', true],
    ['信息', 'N', '值转数值', '7', 7],
    ['信息', 'TYPE', '值类型编号', 'A1', 1],
    ['信息', 'ISREF', '检查引用', 'A1', true],
    ['信息', 'ISEVEN', '检查偶数', '2.5', true],
    ['信息', 'ISODD', '检查奇数', '-3.9', true],
    ['信息', 'ISFORMULA', '检查公式', 'A1', false],
    ['信息', 'SHEET', '工作表序号', 'A1', 1],
    ['信息', 'ISEMAIL', '检查邮件地址', '"developer@univer.ai"', true],
    ['信息', 'ISURL', '检查网址', '"univer.ai"', true],
  ] as const) {
    await functionParameters(page, name)
    await expect(dialog.getByText(`${name} · ${label}`, { exact: true })).toBeVisible()
    await expect(dialog.locator('p').first()).toContainText(category)
    await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill(argumentsText)
    await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
    await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue(`=${name}(${argumentsText})`)
    expect((await cell(page, '0:2'))?.formula).toBeUndefined()
    await page.getByRole('button', { name: '确认编辑', exact: true }).click()
    if (typeof expected === 'number') await expect.poll(async () => (await cell(page, '0:2'))?.value).toBeCloseTo(expected, 12)
    if (typeof expected !== 'number') await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(expected)
    await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
    await click(page, '撤销')
    await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBeUndefined()
    await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  }

  for (const name of ['TRUE', 'FALSE', 'NOW', 'TODAY', 'NA', 'PI', 'RAND', 'SHEETS']) {
    await functionParameters(page, name)
    await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
    await expect(dialog).toBeHidden()
    await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue(`=${name}()`)
    if (name === 'PI' || name === 'SHEETS') {
      await page.getByRole('button', { name: '确认编辑', exact: true }).click()
      await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(name === 'PI' ? Math.PI : 1)
      await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
      await click(page, '撤销')
      await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBeUndefined()
      await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
      continue
    }
    if (name === 'RAND') {
      await page.getByRole('button', { name: '确认编辑', exact: true }).click()
      await expect.poll(async () => (await cell(page, '0:2'))?.value).toBeGreaterThanOrEqual(0)
      await expect.poll(async () => (await cell(page, '0:2'))?.value).toBeLessThan(1)
      await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
      await click(page, '撤销')
      await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBeUndefined()
      await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
      continue
    }
    await page.getByRole('button', { name: '取消编辑', exact: true }).click()
    expect((await cell(page, '0:2'))?.formula).toBeUndefined()
  }
})

test('information menu evaluates its selected predicate through editing, history and reopening', async ({ page }) => {
  await address(page, 'C1')
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await click(page, '信息')
  const functions = page.getByRole('menu', { name: '信息', exact: true })
  await expect(functions.getByRole('menuitem')).toHaveText([
    'ERROR.TYPE', 'ISBETWEEN', 'ISBLANK', 'ISDATE', 'ISEMAIL', 'ISERR', 'ISERROR', 'ISEVEN', 'ISFORMULA', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISODD', 'ISREF', 'ISTEXT', 'ISURL', 'N', 'NA', 'SHEET', 'SHEETS', 'TYPE',
  ])
  await functions.getByRole('menuitem', { name: 'ISERROR', exact: true }).click()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=ISERROR(A1:B1')
  await page.getByRole('button', { name: '取消编辑', exact: true }).click()
  await functionParameters(page, 'ISERROR')
  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await expect(dialog.getByText('ISERROR · 检查错误', { exact: true })).toBeVisible()
  await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill('1/0')
  await dialog.getByRole('button', { name: '插入到公式栏', exact: true }).click()
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('=ISERROR(1/0)')
  expect((await cell(page, '0:2'))?.formula).toBeUndefined()
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(true)
  await expect.poll(async () => (await cell(page, '0:2'))?.cachedDataType).toBe('boolean')
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await click(page, '撤销')
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBeUndefined()
  await click(page, '重做')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(true)
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(officeAPI.getActiveWorkbook()!.save())) as WorkbookModel
    officeAPI.disposeUnit(saved.id!)
    officeAPI.createWorkbook(saved)
  })
  await expect.poll(async () => (await cell(page, '0:2'))?.formula).toBe('=ISERROR(1/0)')
  await expect.poll(async () => (await cell(page, '0:2'))?.value).toBe(true)
})

test('formula menus and the open parameter form follow the whole selection permission', async ({ page }) => {
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(officeAPI.getActiveWorkbook()!.save())) as WorkbookModel
    saved.sheets[0]!.rangeProtectionRules = [{ id: 'formula-locked', permissionId: 'formula-locked-permission', ownerUserId: 'owner',
      editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [],
      ranges: [{ startRow: 0, endRow: 0, startCol: 3, endCol: 3 }] }]
    officeAPI.disposeUnit(saved.id!)
    officeAPI.createWorkbook(saved)
  })
  await address(page, 'C1')
  await ribbon(page).getByRole('tab', { name: '公式', exact: true }).click()
  await functionParameters(page, 'ABS')
  const dialog = page.getByRole('dialog', { name: '插入函数', exact: true })
  await dialog.getByRole('textbox', { name: '函数参数', exact: true }).fill('-5')
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('C1:D1').activate()
  })
  await expect(dialog.getByRole('textbox', { name: '函数参数', exact: true })).toBeDisabled()
  await expect(dialog.getByRole('button', { name: '插入到公式栏', exact: true })).toBeDisabled()
  await dialog.getByRole('button', { name: '取消', exact: true }).click()
  await expect(ribbon(page).getByRole('button', { name: '数学', exact: true })).toBeDisabled()
  await expect(ribbon(page).getByRole('button', { name: '编辑公式', exact: true })).toBeDisabled()
  expect((await cell(page, '0:2'))?.formula).toBeUndefined()
  await address(page, 'C1')
  await expect(ribbon(page).getByRole('button', { name: '数学', exact: true })).toBeEnabled()
})

test('Data Ribbon applies STOP validation and blocks invalid cell entry without changing the model', async ({ page }) => {
  await address(page, 'A1:A2')
  await ribbon(page).getByRole('tab', { name: '数据', exact: true }).click()
  await click(page, '数据验证')

  await page.getByRole('menuitem', { name: '管理验证规则', exact: true }).click()
  const panel = page.getByRole('complementary', { name: '数据验证', exact: true })
  await expect(panel).toBeVisible()
  expect((await state(page)).sheets[0]?.dataValidations ?? []).toHaveLength(0)
  await panel.getByRole('button', { name: '添加规则', exact: true }).click()
  await expect.poll(async () => (await state(page)).sheets[0]?.dataValidations?.[0]).toMatchObject({ type: 'decimal', operator: 'equal', formula1: '100', ranges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 0 }] })
  await panel.locator('#sheet-validation-type').click()
  await page.locator('.ant-select-dropdown:visible').getByText('整数', { exact: true }).click()
  await panel.locator('#sheet-validation-operator').click()
  await page.locator('.ant-select-dropdown:visible').getByText('大于', { exact: true }).click()
  await panel.getByRole('textbox', { name: '数据验证规则值', exact: true }).fill('10')
  await panel.getByRole('button', { name: /应\s*用/ }).click()
  await expect(panel.getByRole('button', { name: '添加规则', exact: true })).toBeVisible()
  await expect.poll(async () => (await state(page)).sheets[0]?.dataValidations?.length).toBe(1)

  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  const cellEditor = page.locator('.canvas-cell-input input')
  await address(page, 'A1')
  await grid.press('7')
  await cellEditor.fill('12')
  await cellEditor.press('Enter')
  await expect.poll(async () => (await cell(page, '0:0'))?.value).toBe(12)

  await address(page, 'A2')
  await grid.press('5')
  await cellEditor.fill('5')
  await cellEditor.press('Enter')
  await expect(page.getByRole('alert')).toContainText('数据验证未通过')
  await expect.poll(async () => (await cell(page, '1:0'))?.value).toBe(0.25)
})

test('data-validation list options render on the cell input bridge and persist after save/reopen', async ({ page }) => {
  await address(page, 'B1:B2')
  await ribbon(page).getByRole('tab', { name: '数据', exact: true }).click()
  await click(page, '数据验证')
  await page.getByRole('menuitem', { name: '添加规则', exact: true }).click()
  const panel = page.getByRole('complementary', { name: '数据验证', exact: true })
  await expect(panel).toBeVisible()
  await panel.locator('#sheet-validation-type').click()
  await page.locator('.ant-select-dropdown:visible').getByText('下拉列表', { exact: true }).click()
  await panel.getByRole('textbox', { name: '数据验证列表选项', exact: true }).fill('Draft,Review,Approved')
  await panel.getByRole('button', { name: /应\s*用/ }).click()
  await expect(panel.getByRole('button', { name: '添加规则', exact: true })).toBeVisible()

  await address(page, 'B1')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  await grid.press('A')
  const input = page.locator('.canvas-cell-input')
  await expect(input).toHaveCSS('border-top-left-radius', '0px')
  await expect(input).not.toHaveCSS('border-top-color', 'rgba(0, 0, 0, 0)')
  const cellEditor = input.locator('input')
  await cellEditor.click()
  await expect(page.getByRole('option', { name: 'Draft', exact: true })).toBeVisible()
  await page.getByRole('option', { name: 'Draft', exact: true }).click()
  await cellEditor.press('Enter')
  await expect.poll(async () => (await cell(page, '0:1'))?.value).toBe('Draft')

  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const saved = JSON.parse(JSON.stringify(officeAPI.getActiveWorkbook()!.save())) as WorkbookModel
    officeAPI.disposeUnit(saved.id!)
    officeAPI.createWorkbook(saved)
  })
  await address(page, 'B1')
  await grid.press('A')
  await expect(page.locator('.canvas-cell-input input')).toBeVisible()
  await page.locator('.canvas-cell-input input').press('ArrowDown')
  await expect(page.getByRole('option', { name: 'Draft', exact: true })).toBeVisible()
})

test('validation picks a sheet range into the draft before applying it to the existing rule', async ({ page }) => {
  await address(page, 'A1')
  await ribbon(page).getByRole('tab', { name: '数据', exact: true }).click()
  await click(page, '数据验证')
  await page.getByRole('menuitem', { name: '添加规则', exact: true }).click()
  const panel = page.getByRole('complementary', { name: '数据验证', exact: true })
  const picker = panel.locator('[data-pick-target="range"]')
  const firstRule = async () => (await state(page)).sheets[0]?.dataValidations?.[0]
  const original = await firstRule()
  expect(original?.id).toBeTruthy()
  await picker.click()
  await address(page, 'D4:E5')
  await expect(panel.getByRole('textbox', { name: '画布选区' })).toHaveValue('D4:E5')
  await panel.getByRole('button', { name: '使用选区' }).click()
  await expect(picker).toBeFocused()
  await expect(panel.locator('#sheet-validation-range')).toHaveValue('D4:E5')
  expect(await firstRule()).toEqual(original)
  await panel.getByRole('button', { name: /应\s*用/ }).click()
  await expect.poll(firstRule).toMatchObject({ id: original!.id, ranges: [{ startRow: 3, startCol: 3, endRow: 4, endCol: 4 }] })
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
