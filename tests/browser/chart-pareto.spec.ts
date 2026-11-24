import { expect, test } from '@playwright/test'

// P17 workbench drafts. No server, browser, screenshot or pixel check has run.
test('Pareto bar, cumulative-line and materialized-point settings publish one isolated apply with history and internal reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B6').setValues([['Cause', 'Count'], ['A', -5], ['B', 9], ['A', 8], ['zero', 0], ['negative', -1]])
    await sheet.insertChart(sheet.newChart('pareto').setSource('A1:B6').setCategoryField(0).setValueFields([1])
      .setAbsolutePosition(100, 100).setSize(520, 330).setTitle('Pareto style draft').setIncludeZeroValues(false)
      .setBarStyle({ opacity: 0.8, color: { type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 1 },
        stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] } })
      .setCumulativeLineStyle({ color: 'rgba(0, 128, 0, 0.8)', opacity: 0.5, width: 4, point: { shape: 'circle', size: 7 } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Pareto style draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), options: chart.getOptions(), revision: chart.getInfo().revision,
      categories: chart.getLayout().categories, raw: chart.getData(), paint: JSON.stringify(chart.getRenderCommands()) }
  })
  const before = await snapshot(), version = Number(await hitbox.getAttribute('data-source-version'))
  expect(before.categories).toEqual(['B', 'A'])
  await hitbox.click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await editor.getByText('图型专属设置', { exact: true }).click()
  await editor.getByText('帕累托柱形样式', { exact: true }).click()
  const opacity = editor.getByRole('spinbutton', { name: '柱形不透明度', exact: true })
  await opacity.fill('0.6')
  await opacity.press('Tab')
  const end = editor.getByRole('spinbutton', { name: '帕累托柱形颜色渐变终点 X', exact: true })
  await end.fill('1.25')
  await end.press('Tab')
  await editor.getByText('帕累托累计线样式', { exact: true }).click()
  const width = editor.getByRole('spinbutton', { name: '累计线宽', exact: true })
  await width.fill('0')
  await width.press('Tab')
  await expect(width).toHaveValue('4')
  await width.fill('6')
  await width.press('Tab')
  const lineOpacity = editor.getByRole('spinbutton', { name: '累计线不透明度', exact: true })
  await lineOpacity.fill('0.4')
  await lineOpacity.press('Tab')
  const size = editor.getByRole('spinbutton', { name: '累计线标记大小', exact: true })
  await size.fill('11')
  await size.press('Tab')
  await editor.getByRole('switch', { name: '保留零权重分类', exact: true }).click()
  await editor.getByLabel('选择帕累托累计线数据项目', { exact: true }).click()
  await page.getByText('2 · A', { exact: true }).last().click()
  const pointSize = editor.getByRole('spinbutton', { name: '累计线项目标记大小', exact: true })
  await pointSize.fill('13')
  await pointSize.press('Tab')
  await editor.getByLabel('累计线项目标记形状', { exact: true }).click()
  await page.getByText('菱形', { exact: true }).last().click()
  const pointBorder = editor.getByRole('spinbutton', { name: '累计线项目边框线宽', exact: true })
  await pointBorder.fill('0')
  await pointBorder.press('Tab')
  expect((await snapshot()).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  await expect.poll(async () => (await snapshot()).revision).toBe(before.revision + 1)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(version)
  const applied = await snapshot()
  expect(applied.categories).toEqual(['B', 'A', 'zero'])
  expect(applied.raw).toEqual(before.raw)
  expect(applied.options.pareto!.barStyle!.opacity).toBe(0.6)
  expect(applied.options.pareto!.barStyle!.color).toEqual({ type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1.25, y: 1 },
    stops: [{ offset: 0, color: 'red' }, { offset: 1, color: 'blue' }] })
  expect(applied.options.pareto!.cumulativeLineStyle!.width).toBe(6)
  expect(applied.options.pareto!.cumulativeLineStyle!.opacity).toBe(0.4)
  expect(applied.options.pareto!.cumulativeLineStyle!.point!.size).toBe(11)
  expect(applied.options.pareto!.cumulativeLineStyle!.dataPoints![1]).toEqual({ size: 13, shape: 'diamond', border: { width: 0 } })
  expect(applied.paint).not.toBe(before.paint)
  expect(JSON.parse(applied.paint).some((command: { kind: string; width?: number }) => command.kind === 'line' && command.width === 6)).toBe(true)
  const restored = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const chart = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    return { options: chart.getOptions(), categories: chart.getLayout().categories, paint: JSON.stringify(chart.getRenderCommands()) }
  })
  expect(restored.options).toEqual(applied.options)
  expect(restored.categories).toEqual(applied.categories)
  expect(restored.paint).toBe(applied.paint)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).options).toEqual(before.options)
  await expect.poll(async () => (await snapshot()).categories).toEqual(before.categories)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.redo()
  })
  await expect.poll(async () => (await snapshot()).options).toEqual(applied.options)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('Pareto reset buttons clear complete dedicated styles while cancel and undo preserve authored values', async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Cause', 'Count'], ['A', 20], ['B', 10], ['zero', 0]])
    await sheet.insertChart(sheet.newChart('pareto').setSource('A1:B4').setCategoryField(0).setValueFields([1])
      .setAbsolutePosition(120, 100).setSize(480, 300).setTitle('Pareto reset draft').setPalette(['purple', 'red', 'green', 'orange'])
      .setIncludeZeroValues(true).setBarStyle({ color: 'purple', opacity: 0.6 })
      .setCumulativeLineStyle({ color: 'red', width: 8, point: { color: 'blue', size: 14 }, label: { visible: true } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Pareto reset draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), options: chart.getOptions(), layout: chart.getLayout(), paint: chart.getRenderCommands() }
  })
  async function open() {
    await hitbox.click()
    const editor = page.getByRole('dialog', { name: '图表设置' })
    await editor.getByText('图型专属设置', { exact: true }).click()
    await editor.getByText('帕累托柱形样式', { exact: true }).click()
    await editor.getByText('帕累托累计线样式', { exact: true }).click()
    return editor
  }
  const before = await snapshot()
  let editor = await open()
  await editor.getByRole('button', { name: '帕累托柱形样式恢复默认', exact: true }).click()
  await editor.getByRole('button', { name: '帕累托累计线样式恢复默认', exact: true }).click()
  expect((await snapshot()).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot()).contents).toBe(before.contents)
  editor = await open()
  await editor.getByRole('button', { name: '帕累托累计线样式恢复默认', exact: true }).click()
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  await expect.poll(async () => (await snapshot()).options.pareto!.cumulativeLineStyle).toBeUndefined()
  const reset = await snapshot()
  expect(reset.options.pareto!.barStyle).toEqual(before.options.pareto!.barStyle)
  expect(reset.options.pareto!.includeZeroValues).toBe(true)
  expect(reset.layout.series.at(-1)!.color).toBe('orange')
  expect(reset.layout.series.at(-1)!.points).toHaveLength(3)
  expect(reset.paint.some(command => command.kind === 'line' && command.width === 2)).toBe(true)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).options).toEqual(before.options)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.redo()
  })
  await expect.poll(async () => (await snapshot()).options).toEqual(reset.options)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
