import { expect, test, type Page } from '@playwright/test'

// P20 workbench drafts. No server, browser, screenshot or pixel check has run.
async function insertSankey(page: Page, title = 'Sankey interaction draft') {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async name => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:C6').setValues([['Source', 'Target', 'Weight'], ['R', 'A', 10], ['A', 'B', 4], ['A', 'C', 6], ['B', 'D', 4], ['C', 'E', 6]])
    await sheet.insertChart(sheet.newChart('sankey').setSource('A1:C6').setCategoryField(0).setValueFields([1, 2])
      .setMapping({ edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } }).setAbsolutePosition(100, 100).setSize(540, 350)
      .setLegend({ visible: false }).setTitle(name).build())
  }, title)
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}

async function snapshot(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    const layout = await chart.getComputedLayout(), commands = await chart.getComputedRenderCommands()
    return { contents: JSON.stringify(workbook.save()), info: chart.getInfo(), options: chart.getOptions(), data: chart.getData(), layout,
      commands: JSON.stringify(commands), labels: commands.filter((command: { kind: string }) => command.kind === 'text').map((command: { text: string }) => command.text) }
  })
}

async function nodePoint(page: Page, title: string, name = 'B') {
  const state = await snapshot(page), box = await page.getByRole('button', { name: title, exact: true }).boundingBox()
  if (!box) throw new Error('Sankey hitbox is unavailable')
  const node = state.layout.sankey.nodes.find((node: { name: string }) => node.name === name)
  if (!node) throw new Error('Sankey node is unavailable')
  return { x: box.x + (node.x + node.width / 2 - state.layout.frame.x) * box.width / state.info.size.width,
    y: box.y + (node.y + node.height / 2 - state.layout.frame.y) * box.height / state.info.size.height }
}

test('Sankey trajectory hover and node dragging publish Canvas versions without content history or whole-chart movement', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  const title = 'Sankey interaction draft', hitbox = await insertSankey(page, title), before = await snapshot(page)
  const point = await nodePoint(page, title), originalBox = await hitbox.boundingBox(), initialVersion = Number(await hitbox.getAttribute('data-source-version'))
  await page.mouse.move(point.x, point.y)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(initialVersion)
  expect((await snapshot(page)).contents).toBe(before.contents)
  const hoverVersion = Number(await hitbox.getAttribute('data-source-version'))
  await page.mouse.down()
  await page.mouse.move(point.x + 36, point.y + 18, { steps: 10 })
  await page.mouse.up()
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(hoverVersion)
  const after = await snapshot(page)
  expect(after.contents).toBe(before.contents)
  expect(after.info).toEqual(before.info)
  expect(after.data).toEqual(before.data)
  expect(after.layout).toEqual(before.layout)
  expect(await hitbox.boundingBox()).toEqual(originalBox)
  await expect(page.getByRole('dialog', { name: '图表设置' })).toBeHidden()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('Escape and pointer cancellation discard node drags, and source changes cancel stale previews before pointerup', async ({ page }) => {
  const title = 'Sankey cancellation draft', hitbox = await insertSankey(page, title), before = await snapshot(page)
  for (const reason of ['escape', 'pointercancel'] as const) {
    const point = await nodePoint(page, title)
    await page.mouse.move(point.x, point.y); await page.mouse.down()
    await page.mouse.move(point.x + 30, point.y + 16, { steps: 5 })
    if (reason === 'escape') await page.keyboard.press('Escape')
    else await page.evaluate(() => window.dispatchEvent(new PointerEvent('pointercancel', { pointerId: 1, bubbles: true })))
    await page.mouse.up()
    expect((await snapshot(page)).contents).toBe(before.contents)
    expect((await snapshot(page)).layout).toEqual(before.layout)
    await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  }
  const point = await nodePoint(page, title)
  await page.mouse.move(point.x, point.y); await page.mouse.down()
  await page.mouse.move(point.x + 35, point.y + 15, { steps: 5 })
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('C3').setValue(8)
  })
  await page.mouse.up()
  await expect.poll(async () => (await snapshot(page)).data.graph.nodes.find((node: { name: string }) => node.name === 'A')!.value).toBe(14)
  const changed = await snapshot(page)
  expect(changed.info).toEqual(before.info)
  await expect(page.getByRole('dialog', { name: '图表设置' })).toBeHidden()
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('Sankey labels and node opacity use isolated antdv-next drafts, one apply, history and internal reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  const title = 'Sankey styles draft', hitbox = await insertSankey(page, title), before = await snapshot(page)
  expect(before.labels).toContain('B')
  await hitbox.click({ position: { x: 30, y: 20 } })
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await editor.getByText('系列、标记与数据标签', { exact: true }).click()
  await editor.getByText('所有系列', { exact: true }).click()
  const all = editor.locator('.ant-collapse-item').filter({ has: page.getByText('所有系列', { exact: true }) }).last()
  await expect(all.getByRole('switch', { name: '显示数据标签', exact: true })).toHaveAttribute('aria-checked', 'true')
  await expect(all.getByLabel('标记形状', { exact: true })).toHaveCount(0)
  await expect(all.getByRole('spinbutton', { name: '标记大小', exact: true })).toHaveCount(0)
  const content = all.getByLabel('标签内容', { exact: true })
  await content.click()
  const contentOptions = page.locator('.ant-select-dropdown:visible')
  await contentOptions.getByText('系列', { exact: true }).last().click()
  await contentOptions.getByText('数值', { exact: true }).last().click()
  await content.press('Escape')
  const position = all.getByLabel('标签位置', { exact: true })
  await position.click(); await page.locator('.ant-select-dropdown:visible').getByText('右侧', { exact: true }).last().click()
  const format = all.getByRole('textbox', { name: '标签数字格式', exact: true })
  await format.fill('0.00'); await format.press('Tab')
  const opacity = all.getByRole('spinbutton', { name: '填充不透明度', exact: true })
  await opacity.fill('0.5'); await opacity.press('Tab')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  await expect.poll(async () => (await snapshot(page)).info.revision).toBe(before.info.revision + 1)
  const applied = await snapshot(page)
  expect(applied.options.allSeriesStyle.label.contentType).toBe(14)
  expect(applied.options.allSeriesStyle.label.position).toBe('right')
  expect(applied.options.allSeriesStyle.fillOpacity).toBe(0.5)
  expect(applied.labels).toContain('Weight,B,4.00')
  expect(applied.data).toEqual(before.data)
  const reopened = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const chart = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    return { options: chart.getOptions(), commands: JSON.stringify(await chart.getComputedRenderCommands()) }
  })
  expect(reopened.options).toEqual(applied.options); expect(reopened.commands).toBe(applied.commands)
  for (const operation of ['undo', 'redo'] as const) {
    await page.evaluate(async action => {
      const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
      const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
      if (action === 'undo') workbook.undo()
      else workbook.redo()
    }, operation)
    await expect.poll(async () => (await snapshot(page)).options).toEqual(operation === 'undo' ? before.options : applied.options)
  }
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})
