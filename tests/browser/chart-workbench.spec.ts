import { expect, test } from '@playwright/test'

test('inserts a radar chart from the worksheet and keeps its type across a saved snapshot', async ({ page }) => {
  await page.goto('/')
  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('A1:B4').setValues([
      ['Category', 'Units'], ['A', 6], ['B', 3], ['C', 9],
    ])
  })
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '图表', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入图表' })
  await dialog.getByPlaceholder('A1:C4').fill('A1:B4')
  await dialog.locator('.chart-insert-form .ant-select').first().click()
  await page.getByText('雷达图', { exact: true }).last().click()
  await dialog.locator('.ant-modal-footer button').last().click()
  await expect(page.locator('.chart-hitbox')).toHaveAttribute('data-rendered', 'true')

  const saved = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    const workbook = officeAPI.getActiveWorkbook()!
    const reopened = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false })
    const chart = reopened.getActiveSheet()!.getCharts()[0]!
    return { type: chart.getInfo().type, data: chart.getData() }
  })
  expect(saved).toMatchObject({
    type: 'radar',
    data: { categories: ['A', 'B', 'C'], series: [{ name: 'Units', values: [6, 3, 9] }] },
  })
})

test('inserts, moves, resizes and refreshes an editable chart on the worksheet canvas', async ({ page }) => {
  const pageErrors: string[] = []
  page.on('pageerror', error => pageErrors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')

  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)

  const nameBox = page.getByRole('textbox', { name: '名称框', exact: true })
  await nameBox.fill('A1:C4')
  await nameBox.press('Enter')
  await grid.press('Control+V')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))
    if (!resource) throw new Error('Workbench runtime module was not loaded')
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook().getActiveSheet().getRange('A1:C4').setValues([
      ['项目', '一月', '二月'], ['甲', 100, 150], ['乙', 200, 250], ['丙', 300, 350],
    ])
  })
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '图表', exact: true }).click()
  await page.getByPlaceholder('A1:C4').fill('A1:C4')
  await page.locator('.ant-modal-footer button').last().click()
  await expect(page.locator('.ant-modal-wrap')).toBeHidden()

  const chart = page.locator('.chart-hitbox')
  await expect(chart).toHaveCount(1)
  await expect(chart).toHaveRole('button')
  await expect(chart).toHaveAttribute('data-rendered', 'true')
  await expect(chart).toHaveAttribute('data-renderer', 'echarts-canvas')
  const bitmapSize = await chart.evaluate(element => ({ width: Number(element.getAttribute('data-bitmap-width')), height: Number(element.getAttribute('data-bitmap-height')) }))
  expect(bitmapSize.width).toBeGreaterThan(0)
  expect(bitmapSize.height).toBeGreaterThan(0)
  const initialSourceVersion = Number(await chart.getAttribute('data-source-version'))
  const before = await chart.evaluate(element => (element as HTMLElement).style.left)
  const bounds = await chart.boundingBox()
  expect(bounds).not.toBeNull()
  const hitTest = await page.evaluate(({ x, y }) => {
    const target = document.elementFromPoint(x, y)
    return { target: target?.outerHTML.slice(0, 180), chart: target?.closest('.chart-hitbox') !== null }
  }, { x: bounds!.x + 60, y: bounds!.y + 40 })
  expect(hitTest.chart, hitTest.target).toBe(true)
  await page.mouse.move(bounds!.x + 60, bounds!.y + 40)
  await page.mouse.down()
  await page.mouse.move(bounds!.x + 130, bounds!.y + 82, { steps: 4 })
  await page.mouse.up()
  await expect.poll(() => chart.evaluate(element => (element as HTMLElement).style.left)).not.toBe(before)
  await expect(chart).toHaveAttribute('aria-pressed', 'true')

  const handle = chart.getByLabel('调整图表大小')
  const resize = await handle.boundingBox()
  expect(resize).not.toBeNull()
  const widthBefore = await chart.evaluate((element) => (element as HTMLElement).style.width)
  await page.mouse.move(resize!.x + resize!.width / 2, resize!.y + resize!.height / 2)
  await page.mouse.down()
  await page.mouse.move(resize!.x + 80, resize!.y + 40, { steps: 4 })
  await page.mouse.up()
  await expect.poll(() => chart.evaluate((element) => (element as HTMLElement).style.width)).not.toBe(widthBefore)

  await nameBox.fill('B3')
  await nameBox.press('Enter')
  const formula = page.locator('.formula-input')
  await formula.fill('1000')
  await formula.press('Enter')
  await expect.poll(() => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))
    if (!resource) throw new Error('Workbench runtime module was not loaded')
    const { officeAPI } = await import(resource.name)
    return officeAPI.getActiveWorkbook().getActiveSheet().getCharts()[0].getData().series[0].values[1]
  })).toBe(1000)
  await expect.poll(async () => Number(await chart.getAttribute('data-source-version'))).toBeGreaterThan(initialSourceVersion)

  await chart.click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  await editor.getByRole('textbox', { name: '图表标题' }).fill('Quarterly revenue')
  await editor.getByRole('textbox', { name: '数据区域' }).fill('A1:B4')
  await editor.getByRole('textbox', { name: 'X 轴标题' }).fill('Product')
  await editor.getByRole('textbox', { name: 'Y 轴标题' }).fill('Revenue')
  await editor.getByRole('checkbox', { name: '显示图例' }).uncheck()
  await editor.locator('.chart-editor-actions button').last().click()
  await expect.poll(() => chart.getAttribute('aria-label')).toBe('Quarterly revenue')

  const chartSettings = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))
    if (!resource) throw new Error('Workbench runtime module was not loaded')
    const { officeAPI } = await import(resource.name)
    const workbook = officeAPI.getActiveWorkbook()!
    return workbook.getActiveSheet()!.getCharts()[0]!.getInfo()
  })
  expect(await chartSettings()).toMatchObject({
    title: 'Quarterly revenue',
    source: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
    legend: { visible: false },
    xAxis: { title: 'Product' },
    yAxis: { title: 'Revenue' },
  })
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await chartSettings()).title).toBe('')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.redo()
  })
  await expect.poll(async () => (await chartSettings()).title).toBe('Quarterly revenue')
  const reopenedChart = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    const workbook = officeAPI.getActiveWorkbook()!
    const reopened = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false })
    return reopened.getActiveSheet()!.getCharts()[0]!.getInfo()
  })
  expect(reopenedChart).toMatchObject({ title: 'Quarterly revenue', legend: { visible: false }, xAxis: { title: 'Product' }, yAxis: { title: 'Revenue' } })
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)
  expect(pageErrors).toEqual([])
})
