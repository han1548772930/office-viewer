import { expect, test } from '@playwright/test'

test('inserts, moves, resizes and refreshes an editable chart on the worksheet canvas', async ({ page }) => {
  const pageErrors: string[] = []
  page.on('pageerror', error => pageErrors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')

  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)

  await page.getByRole('button', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '图表', exact: true }).click()
  await page.getByPlaceholder('A1:C4').fill('A2:E5')
  await page.getByRole('button', { name: '插入', exact: true }).last().click()

  const chart = page.locator('.chart-hitbox')
  await expect(chart).toHaveCount(1)
  const before = await chart.evaluate(element => (element as HTMLElement).style.left)
  const bounds = await chart.boundingBox()
  expect(bounds).not.toBeNull()
  await page.mouse.move(bounds!.x + 60, bounds!.y + 40)
  await page.mouse.down()
  await page.mouse.move(bounds!.x + 130, bounds!.y + 82, { steps: 4 })
  await page.mouse.up()
  await expect.poll(() => chart.evaluate(element => (element as HTMLElement).style.left)).not.toBe(before)

  const movedBounds = await chart.boundingBox()
  expect(movedBounds).not.toBeNull()
  const imageBefore = await page.screenshot({ clip: movedBounds! })
  const gridBounds = await grid.boundingBox()
  expect(gridBounds).not.toBeNull()
  // B3 lies at x = row header + first column + half of the second column,
  // y = column header + two row heights + half a row height.
  await page.mouse.click(gridBounds!.x + 252, gridBounds!.y + 91)
  const formula = page.locator('.formula-input')
  await formula.fill('1000')
  await formula.press('Enter')
  await page.waitForTimeout(120)
  const imageAfter = await page.screenshot({ clip: movedBounds! })
  expect(imageAfter.equals(imageBefore)).toBe(false)
  expect(pageErrors).toEqual([])
})
