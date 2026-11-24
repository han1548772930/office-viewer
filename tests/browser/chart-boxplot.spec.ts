import { expect, test, type Locator, type Page } from '@playwright/test'

// P23 workbench drafts only. No server, browser, pixels, theme/IME, Office or benchmark has run.
// Tests that draw independent outliers explicitly select the local inclusive/IQR extension.
async function applyAndClose(editor: Locator) {
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  const close = editor.getByRole('button', { name: '关闭', exact: true }).last()
  await expect(close).toBeEnabled(); await expect(editor).toBeVisible()
  await close.click(); await expect(editor).toBeHidden()
}
async function insertBoxplot(page: Page, title: string, mode: 'ordinary' | 'large' | 'legacy' = 'ordinary') {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async ({ name, mode }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const samples = mode === 'large' ? Array.from({ length: 17_000 }, (_value, index) => index === 16_999 ? 100 : 1)
      : [100, null, -100, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 100, -100]
    const table = [['Sample', 'Measurements'], ...samples.map((value, index) => [`Sample ${index + 1}`, value])]
    const range = `A1:B${table.length}`
    sheet.getRange(range).setValues(table)
    const builder = sheet.newChart('boxplot').setSource(range).setCategoryField(0).setValueFields([1])
      .setAbsolutePosition(110, 90).setSize(560, 420).setTitle(name).setLegend({ visible: true })
      .setYAxis({ scaleType: 'linear', min: -120, max: 240 })
      .setOptions({ boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr',
        ...(mode === 'legacy' ? { whiskerMultiplier: 2, showOutliers: false } : {}) } })
      .setSeries(1, { color: 'green', point: { shape: 'circle', size: 4 }, dataPoints: {
        0: { color: 'red', shape: 'rect', size: 10, fillOpacity: 0.25, border: { color: 'navy', width: 2, opacity: 0.8 } },
        3: { color: 'gold', shape: 'circle', size: 18 },
        13: { color: 'blue', shape: 'square', size: 14, fillOpacity: 0.75 },
      } })
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
    const [layout, commands] = await Promise.all([chart.getComputedLayout(), chart.getComputedRenderCommands()])
    return { contents: JSON.stringify(workbook.save()), revision: chart.getInfo().revision, options: JSON.parse(JSON.stringify(chart.getOptions())),
      raw: chart.getData(), layout, paint: JSON.stringify(commands) }
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
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    if (operation === 'undo') workbook.undo()
    else workbook.redo()
  }, action)
}
async function openSamples(page: Page, title: string) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  await editor.getByText('系列、标记与数据标签', { exact: true }).click()
  await editor.getByText('单独设置系列', { exact: true }).click()
  await editor.getByText('单个原始样本', { exact: true }).click()
  const samples = editor.locator('.ant-collapse-item').filter({ has: page.getByText('单个原始样本', { exact: true }) }).last()
  return { editor, samples }
}
async function number(panel: Locator, label: string, value: string) {
  const field = panel.getByRole('spinbutton', { name: label, exact: true })
  await field.fill(value); await field.press('Tab')
}
async function choose(page: Page, panel: Locator, label: string, value: string) {
  await panel.getByRole('combobox', { name: label, exact: true }).click()
  await page.getByRole('option', { name: value, exact: true }).click()
}
function selected(page: Page, panel: Locator, label: string) {
  return panel.locator('.ant-form-item').filter({ has: page.getByRole('combobox', { name: label, exact: true }) }).last().locator('.ant-select-selection-item')
}
const drawCommands = (paint: string): Array<{ kind: string; fill?: string; color?: string; width?: number; x1?: number; x2?: number; y1?: number; y2?: number; points?: Array<{ x: number; y: number }> }> => JSON.parse(paint)

test('antd raw-sample edits keep duplicate outliers independent and publish one chart update with history and internal reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Boxplot independent samples draft', hitbox = await insertBoxplot(page, title), before = await snapshot(page)
  const { editor, samples } = await openSamples(page, title)
  await number(samples, '原始样本序号', '1')
  await expect(selected(page, samples, '选择数据项目')).toHaveText('1 · Sample 1 · 100')
  await number(samples, '项目标记大小', '12'); await number(samples, '项目填充不透明度', '0.4')
  await choose(page, samples, '项目标记形状', '菱形')
  await number(samples, '原始样本序号', '14')
  await expect(selected(page, samples, '选择数据项目')).toHaveText('14 · Sample 14 · 100')
  await number(samples, '项目标记大小', '16'); await choose(page, samples, '项目标记形状', '方形')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await applyAndClose(editor)
  const applied = await snapshot(page)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.series[1].dataPoints[0]).toMatchObject({ color: 'red', shape: 'diamond', size: 12, fillOpacity: 0.4 })
  expect(applied.options.series[1].dataPoints[13]).toMatchObject({ color: 'blue', shape: 'rect', size: 16, fillOpacity: 0.75 })
  expect(applied.raw).toEqual(before.raw); expect(applied.layout).toEqual(before.layout)
  expect(drawCommands(applied.paint).some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 0, 0.4)')).toBe(true)
  expect(drawCommands(applied.paint).some(command => command.kind === 'rect' && command.width === 16 && command.fill === 'rgba(0, 0, 255, 0.75)')).toBe(true)
  expect(await reopen(page)).toEqual({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'redo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(applied.options)
  await history(page, 'undo'); await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('null samples report a Form error and an inlier preset activates at its original row after source editing without changing author scopes', async ({ page }) => {
  const title = 'Boxplot source activation draft', hitbox = await insertBoxplot(page, title), before = await snapshot(page), { editor, samples } = await openSamples(page, title)
  await number(samples, '原始样本序号', '2')
  await expect(samples.getByText('请输入包含有限数值的原始样本序号。', { exact: true })).toBeVisible()
  await expect(samples.getByRole('spinbutton', { name: '项目标记大小', exact: true })).toHaveCount(0)
  expect((await snapshot(page)).contents).toBe(before.contents)
  await number(samples, '原始样本序号', '4')
  await expect(samples.getByText('请输入包含有限数值的原始样本序号。', { exact: true })).toHaveCount(0)
  await choose(page, samples, '项目标记形状', '菱形'); await number(samples, '项目标记大小', '20'); await number(samples, '项目填充不透明度', '0.6')
  await applyAndClose(editor)
  const preset = await snapshot(page)
  expect(preset.options.series[1].dataPoints[3]).toMatchObject({ color: 'gold', shape: 'diamond', size: 20, fillOpacity: 0.6 })
  expect(drawCommands(preset.paint).some(command => command.fill === 'rgba(255, 215, 0, 0.6)')).toBe(false)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('B2:B5').setValues([[1], [null], [-100], [200]])
  })
  const refreshed = await snapshot(page)
  expect(refreshed.options).toEqual(preset.options)
  expect(refreshed.raw.series[0].values[0]).toBe(1); expect(refreshed.raw.series[0].values[3]).toBe(200)
  expect(drawCommands(refreshed.paint).some(command => command.kind === 'polygon' && command.fill === 'rgba(255, 215, 0, 0.6)')).toBe(true)
  expect(drawCommands(refreshed.paint).some(command => command.fill === 'rgba(255, 0, 0, 0.25)')).toBe(false)
  expect(refreshed.options.series[1].dataPoints[0]).toEqual(before.options.series[1].dataPoints[0])
  expect(await reopen(page)).toEqual({ options: refreshed.options, raw: refreshed.raw, layout: refreshed.layout, paint: refreshed.paint })
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).raw).toEqual(preset.raw)
  expect((await snapshot(page)).paint).toBe(preset.paint)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('a source sample beyond the bounded selector remains directly editable and resettable through close, apply, history and reopen', async ({ page }) => {
  const title = 'Boxplot long source draft', hitbox = await insertBoxplot(page, title, 'large'), before = await snapshot(page)
  let { editor, samples } = await openSamples(page, title)
  await number(samples, '原始样本序号', '17000')
  await expect(selected(page, samples, '选择数据项目')).toHaveText('17000 · Sample 17000 · 100')
  await number(samples, '项目标记大小', '18'); await number(samples, '项目填充不透明度', '0.35'); await choose(page, samples, '项目标记形状', '菱形')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await applyAndClose(editor)
  const applied = await snapshot(page)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.series[1].dataPoints[16999]).toMatchObject({ shape: 'diamond', size: 18, fillOpacity: 0.35 })
  expect(applied.raw).toEqual(before.raw)
  expect(drawCommands(applied.paint).some(command => command.kind === 'polygon' && command.fill === 'rgba(0, 128, 0, 0.35)')).toBe(true)
  expect(await reopen(page)).toEqual({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  async function resetDraft() {
    const opened = await openSamples(page, title)
    await number(opened.samples, '原始样本序号', '17000')
    await opened.samples.getByRole('button', { name: '恢复项目默认样式', exact: true }).click()
    await expect(opened.samples.getByRole('spinbutton', { name: '项目标记大小', exact: true })).toHaveValue('')
    expect((await snapshot(page)).contents).toBe(applied.contents)
    return opened
  }
  ;({ editor, samples } = await resetDraft())
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(applied.contents)
  ;({ editor, samples } = await resetDraft())
  await applyAndClose(editor)
  const reset = await snapshot(page)
  expect(reset.options.series[1].dataPoints[16999]).toBeUndefined()
  expect(reset.options.series[1].dataPoints[0]).toEqual(applied.options.series[1].dataPoints[0])
  expect(reset.revision).toBe(applied.revision + 1); expect(reset.raw).toEqual(before.raw)
  expect(reset.paint).toBe(before.paint)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(applied.options)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('Boxplot controls restore fixed half-median/range defaults from a legacy IQR draft while preserving inactive sample author styles', async ({ page }) => {
  const title = 'Boxplot statistical defaults draft', hitbox = await insertBoxplot(page, title, 'legacy'), before = await snapshot(page)
  async function resetDraft() {
    await hitbox.click({ position: { x: 24, y: 18 } })
    const editor = page.getByRole('dialog', { name: '图表设置' })
    await expect(editor).toBeVisible(); await editor.getByText('图型专属设置', { exact: true }).click()
    const panel = editor.locator('.ant-collapse-item').filter({ has: page.getByText('图型专属设置', { exact: true }) }).last()
    await expect(selected(page, panel, '四分位方法')).toHaveText('包含中位数')
    await expect(selected(page, panel, '须线范围')).toHaveText('IQR 范围')
    await expect(panel.getByRole('spinbutton', { name: '须线 IQR 倍数', exact: true })).toHaveValue('2')
    await panel.getByRole('button', { name: '须线范围恢复默认', exact: true }).click()
    await expect(selected(page, panel, '须线范围')).toHaveText('IQR 范围')
    await expect(panel.getByRole('button', { name: '须线范围恢复默认', exact: true })).toBeDisabled()
    await panel.getByRole('button', { name: '四分位方法恢复默认', exact: true }).click()
    await expect(selected(page, panel, '四分位方法')).toHaveText('上下半组中位数')
    await expect(selected(page, panel, '须线范围')).toHaveText('IQR 范围')
    await panel.getByRole('button', { name: '须线 IQR 倍数恢复默认', exact: true }).click()
    await expect(selected(page, panel, '须线范围')).toHaveText('全样本最小/最大')
    await expect(panel.getByRole('spinbutton', { name: '须线 IQR 倍数', exact: true })).toHaveCount(0)
    await panel.getByRole('button', { name: '显示异常点恢复默认', exact: true }).click()
    await expect(panel.getByRole('switch', { name: '显示异常点', exact: true })).toHaveAttribute('aria-checked', 'true')
    for (const label of ['四分位方法', '须线范围', '显示异常点']) await expect(panel.getByRole('button', { name: `${label}恢复默认`, exact: true })).toBeDisabled()
    expect((await snapshot(page)).contents).toBe(before.contents)
    return editor
  }
  let editor = await resetDraft()
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(before.contents)
  editor = await resetDraft()
  await applyAndClose(editor)
  const reset = await snapshot(page)
  expect(reset.options.boxplot).toEqual({}); expect(reset.options.series).toEqual(before.options.series)
  expect(reset.revision).toBe(before.revision + 1); expect(reset.raw).toEqual(before.raw)
  expect(drawCommands(reset.paint).some(command => command.fill === 'rgba(255, 0, 0, 0.25)')).toBe(false)
  const plot = reset.layout.plot, maxY = plot.y + (1 - (100 - reset.layout.min) / (reset.layout.max - reset.layout.min)) * plot.height
  expect(drawCommands(reset.paint).some(command => command.kind === 'line' && command.color === 'rgba(0, 128, 0, 1)'
    && Math.abs((command.y1 ?? Infinity) - maxY) < 1e-6 && Math.abs((command.y2 ?? Infinity) - maxY) < 1e-6)).toBe(true)
  expect(await reopen(page)).toEqual({ options: reset.options, raw: reset.raw, layout: reset.layout, paint: reset.paint })
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'redo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(reset.options)
  await history(page, 'undo'); await history(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
