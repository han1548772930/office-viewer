import { expect, test } from '@playwright/test'

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
  await page.getByRole('button', { name: '插入', exact: true }).last().click()

  const chart = page.locator('.chart-hitbox')
  await expect(chart).toHaveCount(1)
  await expect(chart).toHaveRole('button')
  const before = await chart.evaluate(element => (element as HTMLElement).style.left)
  const bounds = await chart.boundingBox()
  expect(bounds).not.toBeNull()
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
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)
  expect(pageErrors).toEqual([])
})
