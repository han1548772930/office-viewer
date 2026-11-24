import { expect, test } from '@playwright/test'

// P16 workbench drafts. No dev server, browser, screenshots or pixels have run.
test('an explicit gradient editor keeps vector, alpha and stop order in one undoable apply and internal reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Category', 'Units'], ['A', 10], ['B', 20], ['C', 30]])
    await sheet.insertChart(sheet.newChart('column').setSource('A1:B4').setValueFields([1]).setAbsolutePosition(100, 100).setSize(520, 330)
      .setTitle('Explicit gradient draft').setAutoGradientFill(false).setAllSeriesStyle({ color: { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 1 },
        stops: [{ offset: 0, color: 'rgba(255, 0, 0, 0.8)' }, { offset: 1, color: 'blue', opacity: 0.5 }] } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Explicit gradient draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), color: chart.getOptions().allSeriesStyle?.color,
      revision: chart.getInfo().revision, paint: JSON.stringify(chart.getRenderCommands()), data: chart.getData() }
  })
  const before = await snapshot(), version = Number(await hitbox.getAttribute('data-source-version'))
  await hitbox.click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await editor.getByText('系列、标记与数据标签', { exact: true }).click()
  await editor.getByText('所有系列', { exact: true }).click()
  await expect(editor.getByLabel('系列颜色填充方式', { exact: true })).toBeVisible()
  const start = editor.getByRole('spinbutton', { name: '系列颜色渐变起点 X', exact: true })
  await start.fill('-0.25')
  await start.press('Tab')
  const opacity = editor.getByRole('spinbutton', { name: '系列颜色色标 1 不透明度', exact: true })
  await opacity.fill('0.4')
  await opacity.press('Tab')
  await editor.getByRole('button', { name: '添加系列颜色渐变色标', exact: true }).click()
  const third = editor.getByRole('spinbutton', { name: '系列颜色色标 3 位置', exact: true })
  await third.fill('0.75')
  await third.press('Tab')
  expect((await snapshot()).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  await expect.poll(async () => (await snapshot()).revision).toBe(before.revision + 1)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(version)
  const applied = await snapshot()
  expect(applied.color).toEqual({ type: 'linear-gradient', start: { x: -0.25, y: 0 }, end: { x: 1, y: 1 }, stops: [
    { offset: 0, color: 'rgba(255, 0, 0, 0.8)', opacity: 0.4 }, { offset: 1, color: 'blue', opacity: 0.5 }, { offset: 0.75, color: 'blue', opacity: 0.5 },
  ] })
  expect(applied.paint).not.toBe(before.paint)
  expect(applied.data).toEqual(before.data)
  const reopened = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const chart = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    return { color: chart.getOptions().allSeriesStyle?.color, described: chart.getDescription().series[1]?.color, paint: JSON.stringify(chart.getRenderCommands()) }
  })
  expect(reopened.color).toEqual(applied.color)
  expect(reopened.described).toEqual(applied.color)
  expect(reopened.paint).toBe(applied.paint)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).color).toEqual(before.color)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.redo()
  })
  await expect.poll(async () => (await snapshot()).color).toEqual(applied.color)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('a type switch retains a solid-only point gradient until the author resets it and cancel publishes no content', async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Category', 'Units'], ['A', 10], ['B', 20], ['C', 30]])
    await sheet.insertChart(sheet.newChart('column').setSource('A1:B4').setValueFields([1]).setTitle('Inactive gradient draft')
      .setAbsolutePosition(120, 100).setSize(480, 300).setSeries(1, { color: 'red', dataPoints: { 0: { color: { type: 'linear-gradient',
        start: { x: 0, y: 0 }, end: { x: 1, y: 0 }, stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] } } } }).build())
  })
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), type: chart.getType(), pointColor: chart.getOptions().series?.[1]?.dataPoints?.[0]?.color }
  })
  const before = await snapshot(), hitbox = page.getByRole('button', { name: 'Inactive gradient draft', exact: true })
  async function openLinePoint() {
    await hitbox.click()
    const editor = page.getByRole('dialog', { name: '图表设置' })
    await editor.getByLabel('图表类型', { exact: true }).click()
    await page.getByText('折线图', { exact: true }).last().click()
    await editor.getByText('系列、标记与数据标签', { exact: true }).click()
    await editor.getByText('单独设置系列', { exact: true }).click()
    await editor.getByText('单个数据项目', { exact: true }).click()
    await editor.getByLabel('选择数据项目', { exact: true }).click()
    await page.getByText('1 · A', { exact: true }).last().click()
    await expect(editor.getByRole('spinbutton', { name: '项目颜色渐变起点 X', exact: true })).toBeDisabled()
    await expect(editor.getByText('该图型的此颜色只接受纯色。已有渐变会保留，切换填充方式可继续编辑。', { exact: true })).toBeVisible()
    return editor
  }
  let editor = await openLinePoint()
  await editor.getByLabel('项目颜色填充方式', { exact: true }).click()
  await page.getByText('纯色', { exact: true }).last().click()
  await expect(editor.getByRole('spinbutton', { name: '项目颜色渐变起点 X', exact: true })).toHaveCount(0)
  expect((await snapshot()).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot()).contents).toBe(before.contents)
  editor = await openLinePoint()
  await editor.getByLabel('项目颜色填充方式', { exact: true }).click()
  await page.getByText('继承颜色', { exact: true }).last().click()
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect.poll(async () => (await snapshot()).type).toBe('line')
  await expect.poll(async () => (await snapshot()).pointColor).toBeUndefined()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).type).toBe('column')
  await expect.poll(async () => (await snapshot()).pointColor).toEqual(before.pointColor)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
