import { expect, test, type Locator, type Page } from '@playwright/test'

// P21 workbench drafts. No server, browser, screenshot or pixel check has run.
async function insertHeatmap(page: Page, title: string, configured = false) {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async ({ name, configured }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:C4').setValues([['Region', 'Units', 'Revenue'], ['North', 1, 2], ['South', 3, 4], ['West', 5, 6]])
    sheet.getRange('C2:C4').setNumberFormat('0.00')
    const builder = sheet.newChart('heatmap').setSource('A1:C4').setCategoryField(0).setValueFields([1, 2])
      .setAbsolutePosition(100, 100).setSize(540, 360).setTitle(name).setPalette(['#0000ff', '#ff0000'])
      .setLegend({ visible: true, position: 'bottom' }).setValueRange(0, 8).setValueUnit(' kg ')
    if (configured) builder.setVisualMapType('piecewise').setCellLabel({ visible: true, position: 'top', fontSize: 14, format: '0.000' })
    await sheet.insertChart(builder.build())
  }, { name: title, configured })
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}

async function snapshot(page: Page) {
  return page.evaluate(async () => {
    const resources = performance.getEntriesByType('resource')
    const runtime = resources.find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const adapter = resources.find(entry => /\/features\/charts\/echarts-option\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(runtime.name), { chartEChartsOption } = await import(adapter.name)
    const workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    const info = chart.getInfo(), layout = await chart.getComputedLayout(), commands = await chart.getComputedRenderCommands()
    const option = chartEChartsOption(workbook.save(), info), legend = option.officeChartPointer.geometry.heatmap.legend
    return { contents: JSON.stringify(workbook.save()), info, options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(), layout,
      commands: JSON.stringify(commands), labels: commands.filter((command: { kind: string }) => command.kind === 'text').map((command: { text: string }) => command.text),
      // Adapter geometry uses local CSS pixels; public getComputedLayout uses worksheet coordinates.
      legend: { bar: legend.bar, pieces: legend.pieces, visible: legend.visible } }
  })
}

async function chartPoint(page: Page, title: string, target: 'cell' | 'colorbar') {
  const state = await snapshot(page), box = await page.getByRole('button', { name: title, exact: true }).boundingBox()
  if (!box) throw new Error('Heatmap hitbox is unavailable')
  const cell = state.layout.series[1].points.find((point: { categoryIndex: number }) => point.categoryIndex === 1)
  const bar = state.legend.bar
  if (target === 'colorbar' && !bar) throw new Error('Continuous Heatmap colorbar is unavailable')
  const x = target === 'cell' ? cell.x - state.layout.frame.x : bar.x + bar.width / 2
  const y = target === 'cell' ? cell.y - state.layout.frame.y : bar.y + bar.height / 2
  return { x: box.x + x * box.width / state.info.size.width, y: box.y + y * box.height / state.info.size.height }
}

async function openSettings(page: Page, title: string) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  return editor
}

async function choose(page: Page, editor: Locator, label: string, value: string) {
  await editor.getByLabel(label, { exact: true }).click()
  await page.locator('.ant-select-dropdown:visible').getByText(value, { exact: true }).last().click()
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
    return { options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(), commands: JSON.stringify(await chart.getComputedRenderCommands()) }
  })
}

test('cell and colorbar hover update the Canvas version without source, author or content-history changes', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  const title = 'Heatmap hover draft', hitbox = await insertHeatmap(page, title), before = await snapshot(page)
  let version = Number(await hitbox.getAttribute('data-source-version'))
  for (const target of ['cell', 'colorbar'] as const) {
    const point = await chartPoint(page, title, target)
    await page.mouse.move(point.x, point.y)
    await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(version)
    version = Number(await hitbox.getAttribute('data-source-version'))
    const after = await snapshot(page)
    expect(after.contents).toBe(before.contents)
    expect(after.data).toEqual(before.data)
    expect(after.layout).toEqual(before.layout)
    expect(after.commands).toBe(before.commands)
    await page.mouse.move(point.x, point.y)
    await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBe(version)
  }
  await expect(page.getByRole('dialog', { name: '图表设置' })).toBeHidden()
  await history(page, 'undo')
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('pointer leave and cancellation clear hover, and a source edit replaces stale geometry while preserving authored range', async ({ page }) => {
  const title = 'Heatmap source draft', hitbox = await insertHeatmap(page, title), before = await snapshot(page)
  for (const reason of ['pointerleave', 'pointercancel'] as const) {
    const point = await chartPoint(page, title, 'colorbar')
    const baseline = Number(await hitbox.getAttribute('data-source-version'))
    await page.mouse.move(point.x, point.y)
    await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(baseline)
    const hoverVersion = Number(await hitbox.getAttribute('data-source-version'))
    if (reason === 'pointerleave') await hitbox.dispatchEvent('pointerleave')
    else await page.evaluate(() => window.dispatchEvent(new PointerEvent('pointercancel', { pointerId: 1, bubbles: true })))
    await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(hoverVersion)
    expect((await snapshot(page)).contents).toBe(before.contents)
    const box = await hitbox.boundingBox()
    if (!box) throw new Error('Heatmap hitbox is unavailable')
    await page.mouse.move(box.x + 20, box.y + 18)
  }
  const point = await chartPoint(page, title, 'cell')
  await page.mouse.move(point.x, point.y)
  const hoveredVersion = Number(await hitbox.getAttribute('data-source-version'))
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('C3').setValue(12)
  })
  await expect.poll(async () => (await snapshot(page)).data.series[1].values[1]).toBe(12)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(hoveredVersion)
  const changed = await snapshot(page)
  expect(changed.data.series[1].formattedValues[1]).toBe('12.00')
  expect(changed.info).toEqual(before.info)
  expect(changed.options.heatmap).toEqual(before.options.heatmap)
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).data).toEqual(before.data)
  await history(page, 'undo')
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('antd Heatmap range, unit, value labels and colorbar type stay isolated until one apply, then undo and JSON reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  const title = 'Heatmap options draft', hitbox = await insertHeatmap(page, title), before = await snapshot(page)
  expect(before.labels).toContain('0 kg'); expect(before.labels).toContain('8 kg')
  expect(before.data.series[1].formattedValues).toEqual(['2.00', '4.00', '6.00'])
  const editor = await openSettings(page, title)
  await choose(page, editor, '色阶类型', '分段')
  await editor.getByText('图型专属设置', { exact: true }).click()
  await expect(editor.getByLabel('单元格标签位置', { exact: true })).toHaveCount(0)
  for (const [name, value] of [['色阶最小值', '1'], ['色阶最大值', '6']] as const) {
    const input = editor.getByRole('spinbutton', { name, exact: true })
    await input.fill(value); await input.press('Tab')
  }
  const unit = editor.getByRole('textbox', { name: '数值单位', exact: true })
  await unit.fill(' t '); await unit.press('Tab')
  await editor.getByRole('switch', { name: '显示单元格标签', exact: true }).click()
  await choose(page, editor, '单元格标签位置', '下方')
  const format = editor.getByRole('textbox', { name: '单元格标签数字格式', exact: true })
  await format.fill('0.000'); await format.press('Tab')
  await editor.getByText('系列、标记与数据标签', { exact: true }).click()
  await editor.getByText('所有系列', { exact: true }).click()
  const all = editor.locator('.ant-collapse-item').filter({ has: page.getByText('所有系列', { exact: true }) }).last()
  await expect(all.getByLabel('标签内容', { exact: true })).toHaveCount(0)
  await expect(all.getByLabel('标记形状', { exact: true })).toHaveCount(0)
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  await expect.poll(async () => (await snapshot(page)).info.revision).toBe(before.info.revision + 1)
  const applied = await snapshot(page)
  expect(applied.options.heatmap).toMatchObject({ visualMapType: 'piecewise', min: 1, max: 6, unit: ' t ', label: { visible: true, position: 'bottom', format: '0.000' } })
  expect(applied.labels).toContain('4.000')
  expect(applied.labels).not.toContain('4.00')
  expect(applied.labels.some((text: string) => text.endsWith(' t'))).toBe(false)
  expect(applied.legend.pieces).toHaveLength(8)
  expect(applied.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ options: applied.options, data: applied.data, commands: applied.commands })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'redo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(applied.options)
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('individual range, unit, type and whole-label resets preserve content on cancel and restore defaults in one apply', async ({ page }) => {
  const title = 'Heatmap reset draft', hitbox = await insertHeatmap(page, title, true), before = await snapshot(page)
  async function resetDraft() {
    const editor = await openSettings(page, title)
    await editor.getByRole('button', { name: '色阶类型恢复默认', exact: true }).click()
    await editor.getByText('图型专属设置', { exact: true }).click()
    for (const name of ['色阶最小值', '色阶最大值', '数值单位', '显示单元格标签']) {
      await editor.getByRole('button', { name: `${name}恢复默认`, exact: true }).click()
    }
    await expect(editor.getByRole('switch', { name: '显示单元格标签', exact: true })).toHaveAttribute('aria-checked', 'false')
    await expect(editor.getByLabel('单元格标签位置', { exact: true })).toHaveCount(0)
    expect((await snapshot(page)).contents).toBe(before.contents)
    return editor
  }
  let editor = await resetDraft()
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(before.contents)
  editor = await resetDraft()
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const reset = await snapshot(page)
  expect(reset.info.revision).toBe(before.info.revision + 1)
  for (const key of ['min', 'max', 'visualMapType', 'unit', 'label']) expect(reset.options.heatmap?.[key]).toBeUndefined()
  expect(reset.layout.min).toBe(1); expect(reset.layout.max).toBe(6)
  expect(reset.legend.bar).toBeDefined(); expect(reset.legend.pieces).toHaveLength(0)
  expect(reset.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ options: reset.options, data: reset.data, commands: reset.commands })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'undo')
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
