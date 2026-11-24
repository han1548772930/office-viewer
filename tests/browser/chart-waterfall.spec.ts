import { expect, test, type Locator, type Page } from '@playwright/test'

// P18 workbench drafts. No server, browser, screenshot or pixel check has run.
async function snapshot(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    const [layout, commands] = await Promise.all([chart.getComputedLayout(), chart.getComputedRenderCommands()])
    return { contents: JSON.stringify(workbook.save()), options: JSON.parse(JSON.stringify(chart.getOptions())),
      revision: chart.getInfo().revision, raw: chart.getData(), layout, paint: JSON.stringify(commands) }
  })
}

async function reopen(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const chart = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    const [layout, commands] = await Promise.all([chart.getComputedLayout(), chart.getComputedRenderCommands()])
    return { options: JSON.parse(JSON.stringify(chart.getOptions())), raw: chart.getData(), layout, paint: JSON.stringify(commands) }
  })
}

async function history(page: Page, action: 'undo' | 'redo') {
  await page.evaluate(async operation => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()![operation]()
  }, action)
}

async function openWaterfall(page: Page, title: string): Promise<Locator> {
  await page.getByRole('button', { name: title, exact: true }).click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  if (!await editor.getByRole('switch', { name: '显示自动合计', exact: true }).isVisible()) {
    await editor.getByText('图型专属设置', { exact: true }).click()
  }
  return editor
}

async function number(editor: Locator, label: string, value: string) {
  const field = editor.getByRole('spinbutton', { name: label, exact: true })
  await field.fill(value)
  await field.press('Tab')
}

async function choose(page: Page, editor: Locator, label: string, value: string) {
  await editor.getByLabel(label, { exact: true }).click()
  await page.getByText(value, { exact: true }).last().click()
}

test('waterfall role, source-total and connector edits publish one apply with shared cumulative geometry, history and internal reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:C3').setValues([['Label', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]])
    await sheet.insertChart(sheet.newChart('waterfall').setSource('A1:C3').setCategoryField(0).setValueFields([1, 2])
      .setAbsolutePosition(100, 100).setSize(520, 330).setTitle('Waterfall source draft')
      .setPalette(['red', 'blue', 'green', 'orange']).setPositiveStyle(1, { color: 'red' })
      .setNegativeStyle(2, { color: 'blue' }).setSubtotalStyle(2, { color: 'purple' })
      .setStackType('arrangement').setUseSubtotal(true).setOptions({ waterfall: { pointRoles: [], connector: false } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Waterfall source draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const before = await snapshot(page), version = Number(await hitbox.getAttribute('data-source-version'))
  expect(before.layout.waterfall.points.filter((point: { synthetic: boolean }) => point.synthetic)
    .map((point: { value: number }) => point.value)).toEqual([9, 9])
  const editor = await openWaterfall(page, 'Waterfall source draft')
  await editor.getByRole('textbox', { name: '角色显示名称', exact: true }).fill('Gain')
  await number(editor, '角色填充不透明度', '0.55')
  await editor.getByText('角色内的单个数据项目', { exact: true }).click()
  await choose(page, editor, '选择瀑布图角色数据项目', '1 · A')
  await number(editor, '角色项目不透明度', '0.35')
  await editor.getByRole('switch', { name: '显示累计连接线', exact: true }).click()
  await number(editor, '连接线宽', '3')
  await number(editor, '连接线不透明度', '0.65')
  await choose(page, editor, '连接线型', '实线')
  await editor.getByText('源项目的变化与合计', { exact: true }).click()
  await choose(page, editor, '瀑布图项目含义来源系列', 'Second')
  await choose(page, editor, '瀑布图项目含义来源项目', '1 · A')
  await choose(page, editor, '源项目含义', '合计')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  await expect.poll(async () => (await snapshot(page)).revision).toBe(before.revision + 1)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(version)
  const applied = await snapshot(page)
  expect(applied.raw).toEqual(before.raw)
  expect(applied.options.waterfall.pointRoles).toEqual([{ seriesIndex: 2, dataPointIndex: 0, role: 'total' }])
  expect(applied.options.waterfall.styles.positive).toMatchObject({ name: 'Gain', fillOpacity: 0.55, dataPoints: { 0: { fillOpacity: 0.35 } } })
  expect(applied.options.waterfall.connector).toMatchObject({ width: 3, opacity: 0.65, dashType: 'solid' })
  expect(applied.layout.waterfall.points.filter((point: { synthetic: boolean }) => point.synthetic)
    .map((point: { value: number }) => point.value)).toEqual([2, 2])
  expect(applied.layout.series[1].points[0]).toMatchObject({ dataPointIndex: 0, categoryIndex: 3, baseValue: 0, value: 4, waterfallRole: 'subtotal' })
  expect(applied.layout.waterfall.connectors).toHaveLength(3)
  expect(applied.paint).not.toBe(before.paint)
  expect(JSON.parse(applied.paint).some((command: { kind: string; width?: number }) => command.kind === 'line' && command.width === 3)).toBe(true)
  const restored = await reopen(page)
  expect(restored).toEqual({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  expect((await snapshot(page)).layout).toEqual(before.layout)
  await history(page, 'redo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(applied.options)
  expect((await snapshot(page)).paint).toBe(applied.paint)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('stacked waterfall clears source roles, paints one independently styled global subtotal and preserves authorship when hiding it', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:C3').setValues([['Label', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]])
    await sheet.insertChart(sheet.newChart('waterfall').setSource('A1:C3').setCategoryField(0).setValueFields([1, 2])
      .setAbsolutePosition(100, 100).setSize(520, 330).setTitle('Waterfall global draft')
      .setSubtotalStyle(2, { name: 'Source only', color: 'purple', fillOpacity: 0.8 })
      .setOptions({ waterfall: { pointRoles: [{ seriesIndex: 2, dataPointIndex: 0, role: 'total' }], connector: true,
        styles: { subtotal: { name: 'Grand', color: 'orange', fillOpacity: 0.4, label: { visible: true, contentType: 'series' } } } } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Waterfall global draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const before = await snapshot(page)
  let editor = await openWaterfall(page, 'Waterfall global draft')
  await choose(page, editor, '瀑布图排列方式', '按分类堆积')
  await expect(editor.getByText('源项目的变化与合计', { exact: true })).toHaveCount(0)
  await choose(page, editor, '瀑布图角色', '小计 / 合计')
  await expect(editor.getByText('堆积瀑布图的自动合计使用全局共享样式。', { exact: true })).toBeVisible()
  await expect(editor.getByText('全局合计共享样式', { exact: true })).toBeVisible()
  await editor.getByRole('textbox', { name: '角色显示名称', exact: true }).fill('Grand total')
  await number(editor, '角色填充不透明度', '0.25')
  await editor.getByText('角色内的单个数据项目', { exact: true }).click()
  await choose(page, editor, '选择瀑布图角色数据项目', '自动全局合计')
  await number(editor, '角色项目不透明度', '0.6')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const stacked = await snapshot(page)
  expect(stacked.revision).toBe(before.revision + 1)
  expect(stacked.options.waterfall.stackType).toBe('stacked')
  expect(stacked.options.waterfall.pointRoles).toBeUndefined()
  expect(stacked.options.waterfall.styles.subtotal).toMatchObject({ name: 'Grand total', fillOpacity: 0.25, dataPoints: { 2: { fillOpacity: 0.6 } } })
  expect(stacked.options.series[2].waterfallStyles.subtotal).toEqual(before.options.series[2].waterfallStyles.subtotal)
  expect(stacked.layout.categories).toEqual(['A', 'B', 'Grand total'])
  expect(stacked.layout.series.map((series: { points: unknown[] }) => series.points.length)).toEqual([2, 2])
  expect(stacked.layout.series[0].points[1]).toMatchObject({ baseValue: 12, value: -3, dataPointIndex: 1, categoryIndex: 1 })
  expect(stacked.layout.series[1].points[1]).toMatchObject({ baseValue: 14, value: -2, dataPointIndex: 1, categoryIndex: 1 })
  expect(stacked.layout.waterfall.points.at(-1)).toMatchObject({ seriesPosition: 0xffff_ffff, synthetic: true, baseValue: 0, value: 9, dataPointIndex: 2 })
  expect(stacked.layout.waterfall.connectors).toHaveLength(1)
  expect(JSON.parse(stacked.paint).some((command: { kind: string; text?: string }) => command.kind === 'text' && command.text === 'Grand total')).toBe(true)
  editor = await openWaterfall(page, 'Waterfall global draft')
  await editor.getByRole('switch', { name: '显示自动合计', exact: true }).click()
  expect((await snapshot(page)).contents).toBe(stacked.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const hidden = await snapshot(page)
  expect(hidden.revision).toBe(stacked.revision + 1)
  expect(hidden.layout.categories).toEqual(['A', 'B'])
  expect(hidden.layout.waterfall.points).toHaveLength(4)
  expect(hidden.layout.waterfall.points.some((point: { synthetic: boolean }) => point.synthetic)).toBe(false)
  expect(hidden.options.waterfall.styles).toEqual(stacked.options.waterfall.styles)
  expect(hidden.raw).toEqual(before.raw)
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).layout.categories).toEqual(stacked.layout.categories)
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'redo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(stacked.options)
  await history(page, 'redo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(hidden.options)
  expect(await reopen(page)).toEqual({ options: hidden.options, raw: hidden.raw, layout: hidden.layout, paint: hidden.paint })
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('waterfall role, point, connector and source-total resets stay isolated on cancel and survive apply without reinitializing defaults', async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Label', 'Value'], ['A', 6], ['B', -2], ['C', 4]])
    await sheet.insertChart(sheet.newChart('waterfall').setSource('A1:B4').setCategoryField(0).setValueFields([1])
      .setAbsolutePosition(100, 100).setSize(480, 300).setTitle('Waterfall reset draft').setPalette(['purple', 'red'])
      .setUseSubtotal(false).setOptions({ waterfall: { pointRoles: [{ seriesIndex: 1, dataPointIndex: 2, role: 'total' }],
        connector: { color: 'black', width: 4, opacity: 0.5, dashType: 'solid' },
        styles: { positive: { color: 'lime', fillOpacity: 0.6, dataPoints: { 0: { color: 'blue', fillOpacity: 0.3 } } } } } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Waterfall reset draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const before = await snapshot(page)
  async function resetDraft() {
    const editor = await openWaterfall(page, 'Waterfall reset draft')
    await editor.getByText('角色内的单个数据项目', { exact: true }).click()
    await choose(page, editor, '选择瀑布图角色数据项目', '1 · A')
    await editor.getByRole('button', { name: '恢复角色项目继承样式', exact: true }).click()
    await editor.getByRole('button', { name: '恢复该角色继承样式', exact: true }).click()
    await editor.getByRole('button', { name: '显示累计连接线恢复默认', exact: true }).click()
    await expect(editor.getByRole('switch', { name: '显示累计连接线', exact: true })).not.toBeChecked()
    await editor.getByText('源项目的变化与合计', { exact: true }).click()
    await choose(page, editor, '瀑布图项目含义来源项目', '3 · C')
    await editor.getByRole('button', { name: '源项目含义恢复默认', exact: true }).click()
    return editor
  }
  let editor = await resetDraft()
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(before.contents)
  editor = await resetDraft()
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const reset = await snapshot(page)
  expect(reset.revision).toBe(before.revision + 1)
  expect(reset.options.waterfall.styles?.positive).toBeUndefined()
  expect(reset.options.waterfall.connector).toBeUndefined()
  expect(reset.options.waterfall.pointRoles).toBeUndefined()
  expect(reset.options.waterfall.useSubtotal).toBe(false)
  expect(reset.raw).toEqual(before.raw)
  expect(before.layout.series[0].points[2]).toMatchObject({ baseValue: 0, value: 4, waterfallRole: 'subtotal' })
  expect(reset.layout.series[0].points[2]).toMatchObject({ baseValue: 4, value: 4, waterfallRole: 'positive' })
  expect(reset.layout.waterfall.connectors).toHaveLength(2)
  expect(reset.paint).not.toBe(before.paint)
  expect(await reopen(page)).toEqual({ options: reset.options, raw: reset.raw, layout: reset.layout, paint: reset.paint })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'redo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(reset.options)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
