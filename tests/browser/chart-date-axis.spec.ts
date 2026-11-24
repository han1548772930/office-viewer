import { expect, test, type Locator, type Page } from '@playwright/test'

// P22 workbench drafts only. No server, browser, pixels, theme/IME or Office sample has run.
async function applyAndClose(editor: Locator) {
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  const close = editor.getByRole('button', { name: '关闭', exact: true }).last()
  await expect(close).toBeEnabled(); await expect(editor).toBeVisible()
  await close.click(); await expect(editor).toBeHidden()
}
async function insertDateChart(page: Page, title: string, mode: 'ordinary' | 'resets' | 'bubble' = 'ordinary') {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async ({ name, mode }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const table = [['Date', 'Units', 'Size'], [45382, 30, 3], [45322, 10, 1], [45351, 20, 2], [45322, 11, 4]]
    sheet.getRange('A1:C5').setValues(table)
    const builder = sheet.newChart(mode === 'bubble' ? 'bubble' : 'line').setSource('A1:C5').setCategoryField(0)
      .setValueFields(mode === 'bubble' ? [1, 2] : [1]).setAbsolutePosition(110, 90).setSize(560, 420).setTitle(name)
      .setLegend({ visible: false }).setUseDateAxis(true)
      .setXAxis({ scaleType: 'linear', baseTimeUnit: 'days', majorTimeUnit: 'days', majorUnit: 15, line: { width: 2, color: '#2563eb' } })
      .setYAxis({ scaleType: 'linear', min: 0, max: 100, majorUnit: 25 })
    if (mode === 'resets') builder.setXAxis({ min: 45292, max: 45412, reverse: true, format: 'yyyy-mm-dd',
      baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: 2 })
    if (mode === 'bubble') builder.setBubbleMapping({ xIndex: 0, yIndex: 1, sizeIndex: 2 })
      .setXAxis({ min: 45290, max: 45420, baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: 2 })
    await sheet.insertChart(builder.build())
  }, { name: title, mode })
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}

async function snapshot(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    const info = chart.getInfo(), commands = await chart.getComputedRenderCommands()
    return { contents: JSON.stringify(workbook.save()), revision: info.revision, axes: { x: info.xAxis, y: info.yAxis },
      options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(), layout: await chart.getComputedLayout(),
      commands: JSON.stringify(commands), dates: commands.filter((command: { kind: string; text?: string }) => command.kind === 'text' && /^2024-\d{2}-\d{2}$/.test(command.text ?? '')).map((command: { text: string }) => command.text) }
  })
}

async function openAxes(page: Page, title: string, both = false) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  await editor.getByText('坐标轴', { exact: true }).click()
  await editor.getByText('X 轴 / 分类轴', { exact: true }).click()
  const x = editor.locator('.ant-collapse-item').filter({ has: page.getByText('X 轴 / 分类轴', { exact: true }) }).last()
  if (both) await editor.getByText('Y 轴 / 数值轴', { exact: true }).click()
  const y = editor.locator('.ant-collapse-item').filter({ has: page.getByText('Y 轴 / 数值轴', { exact: true }) }).last()
  return { editor, x, y }
}

async function choose(page: Page, panel: Locator, name: string, value: string) {
  await panel.getByRole('combobox', { name, exact: true }).click()
  await page.getByRole('option', { name: value, exact: true }).click()
}
function selected(page: Page, panel: Locator, name: string) {
  return panel.locator('.ant-form-item').filter({ has: page.getByRole('combobox', { name, exact: true }) }).last().locator('.ant-select-selection-item')
}
async function history(page: Page, action: 'undo' | 'redo') {
  await page.evaluate(async operation => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    if (operation === 'undo') workbook.undo()
    else workbook.redo()
  }, action)
}
async function reopen(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const chart = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    const info = chart.getInfo()
    return { axes: { x: info.xAxis, y: info.yAxis }, options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(), commands: JSON.stringify(await chart.getComputedRenderCommands()) }
  })
}

test('antd date basic and major units alter calendar geometry once, preserve source rows and survive numeric interval edits, history and reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Date calendar units draft', hitbox = await insertDateChart(page, title), before = await snapshot(page)
  const { editor, x, y } = await openAxes(page, title, true)
  await choose(page, x, '日期基本单位', '月')
  await choose(page, x, '日期主刻度单位', '月')
  await x.getByRole('spinbutton', { name: '日期主单位间隔', exact: true }).fill('2')
  await x.getByRole('spinbutton', { name: '日期主单位间隔', exact: true }).press('Tab')
  await y.getByRole('spinbutton', { name: '主刻度间隔', exact: true }).fill('20')
  await y.getByRole('spinbutton', { name: '主刻度间隔', exact: true }).press('Tab')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await applyAndClose(editor)
  const applied = await snapshot(page)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.axes.x).toMatchObject({ baseTimeUnit: 'months', majorTimeUnit: 'months', majorUnit: 2 })
  expect(applied.axes.y.majorUnit).toBe(20)
  expect(applied.dates).toEqual(['2024-01-01', '2024-03-01'])
  expect(applied.layout.dateOrder).toEqual([1, 3, 2, 0])
  const points = applied.layout.series[0].points
  expect(points.map((point: { categoryIndex: number; xValue: number }) => [point.categoryIndex, point.xValue])).toEqual([[1, 45322], [3, 45322], [2, 45351], [0, 45382]])
  expect(points[0].x).toBe(points[1].x)
  expect(points[2].x - points[0].x).toBeCloseTo(points[3].x - points[2].x)
  expect(applied.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ axes: applied.axes, options: applied.options, data: applied.data, commands: applied.commands })
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).axes).toEqual(before.axes)
  await history(page, 'redo'); await expect.poll(async () => (await snapshot(page)).axes).toEqual(applied.axes)
  await history(page, 'undo'); await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('date fields restore independent inheritance in detached drafts and remain cleared after close, apply, undo and internal reopen', async ({ page }) => {
  const title = 'Date field resets draft', hitbox = await insertDateChart(page, title, 'resets'), before = await snapshot(page)
  async function resetFields() {
    const { editor, x } = await openAxes(page, title)
    await x.getByRole('button', { name: '日期主刻度单位恢复默认', exact: true }).click()
    await expect(selected(page, x, '日期主刻度单位')).toHaveText('月')
    await expect(selected(page, x, '日期基本单位')).toHaveText('月')
    await expect(x.getByRole('spinbutton', { name: '日期主单位间隔', exact: true })).toHaveValue('2')
    await x.getByRole('button', { name: '日期主单位间隔恢复默认', exact: true }).click()
    await expect(x.getByRole('spinbutton', { name: '日期主单位间隔', exact: true })).toHaveValue('')
    await x.getByRole('button', { name: '日期基本单位恢复默认', exact: true }).click()
    await expect(selected(page, x, '日期基本单位')).toHaveText('日')
    await expect(selected(page, x, '日期主刻度单位')).toHaveText('日')
    for (const name of ['日期主刻度单位', '日期主单位间隔', '日期基本单位']) await expect(x.getByRole('button', { name: `${name}恢复默认`, exact: true })).toBeDisabled()
    await expect(x.getByRole('switch', { name: '日期轴', exact: true })).toHaveAttribute('aria-checked', 'true')
    await expect(x.getByRole('switch', { name: '反向轴', exact: true })).toHaveAttribute('aria-checked', 'true')
    expect((await snapshot(page)).contents).toBe(before.contents)
    return editor
  }
  let editor = await resetFields()
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(before.contents)
  editor = await resetFields()
  await applyAndClose(editor)
  const applied = await snapshot(page)
  for (const field of ['baseTimeUnit', 'majorTimeUnit', 'majorUnit'] as const) expect(applied.axes.x[field]).toBeUndefined()
  expect(applied.axes.x).toMatchObject({ min: 45292, max: 45412, reverse: true, format: 'yyyy-mm-dd', line: before.axes.x.line })
  expect(applied.axes.y).toEqual(before.axes.y); expect(applied.options.mapping.useDateAxis).toBe(true)
  expect(applied.layout.dateAxis).toBe(true); expect(applied.revision).toBe(before.revision + 1)
  expect(applied.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ axes: applied.axes, options: applied.options, data: applied.data, commands: applied.commands })
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).axes).toEqual(before.axes)
  await history(page, 'redo'); await expect.poll(async () => (await snapshot(page)).axes).toEqual(applied.axes)
  await history(page, 'undo'); await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('turning dates off preserves inactive calendar fields and accepts a fractional numeric interval, while toggle reset persists', async ({ page }) => {
  const title = 'Date off numeric interval draft', hitbox = await insertDateChart(page, title, 'bubble'), before = await snapshot(page)
  async function turnOff() {
    const { editor, x } = await openAxes(page, title)
    await x.getByRole('switch', { name: '日期轴', exact: true }).click()
    await expect(x.getByRole('combobox', { name: '日期基本单位', exact: true })).toHaveCount(0)
    await expect(x.getByRole('combobox', { name: '日期主刻度单位', exact: true })).toHaveCount(0)
    await x.getByRole('spinbutton', { name: '主刻度间隔', exact: true }).fill('0.25')
    await x.getByRole('spinbutton', { name: '主刻度间隔', exact: true }).press('Tab')
    expect((await snapshot(page)).contents).toBe(before.contents)
    return editor
  }
  let editor = await turnOff()
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(before.contents)
  editor = await turnOff()
  await applyAndClose(editor)
  const off = await snapshot(page)
  expect(off.axes.x).toMatchObject({ baseTimeUnit: 'months', majorTimeUnit: 'years', majorUnit: 0.25, min: 45290, max: 45420 })
  expect(off.options.mapping.useDateAxis).toBe(false); expect(off.layout.dateAxis).toBeUndefined()
  expect(off.revision).toBe(before.revision + 1)
  expect(await reopen(page)).toEqual({ axes: off.axes, options: off.options, data: off.data, commands: off.commands })
  const reopened = await openAxes(page, title)
  await reopened.x.getByRole('button', { name: '日期轴恢复默认', exact: true }).click()
  await expect(reopened.x.getByRole('button', { name: '日期轴恢复默认', exact: true })).toBeDisabled()
  await expect(reopened.x.getByRole('switch', { name: '日期轴', exact: true })).toHaveAttribute('aria-checked', 'false')
  expect((await snapshot(page)).contents).toBe(off.contents)
  await applyAndClose(reopened.editor)
  const reset = await snapshot(page)
  expect(reset.options.mapping.useDateAxis).toBeUndefined(); expect(reset.axes).toEqual(off.axes)
  expect(reset.revision).toBe(off.revision + 1)
  expect(await reopen(page)).toEqual({ axes: reset.axes, options: reset.options, data: reset.data, commands: reset.commands })
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(off.options)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).axes).toEqual(before.axes)
  await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('fractional calendar counts and excessive authored date tick budgets keep the settings open without publishing history', async ({ page }) => {
  const title = 'Date interval rejection draft', hitbox = await insertDateChart(page, title), before = await snapshot(page), { editor, x } = await openAxes(page, title)
  await choose(page, x, '日期基本单位', '月'); await choose(page, x, '日期主刻度单位', '月')
  const interval = x.getByRole('spinbutton', { name: '日期主单位间隔', exact: true })
  await interval.fill('0.5'); await interval.press('Tab')
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeVisible()
  await expect(editor.getByText('日期轴主单位须为正数，月或年步进须为安全整数。', { exact: true }).first()).toBeVisible()
  expect((await snapshot(page)).contents).toBe(before.contents)
  await choose(page, x, '日期基本单位', '日'); await choose(page, x, '日期主刻度单位', '日')
  await interval.fill('0.001'); await interval.press('Tab')
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeVisible(); await expect(editor.getByText(/日期轴主刻度超过10000个/).first()).toBeVisible()
  expect((await snapshot(page)).contents).toBe(before.contents)
  await interval.fill('15'); await interval.press('Tab')
  await applyAndClose(editor)
  expect((await snapshot(page)).revision).toBe(before.revision + 1)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).axes).toEqual(before.axes)
  await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
