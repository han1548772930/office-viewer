import { expect, test, type Locator, type Page } from '@playwright/test'
import { applyStatisticalSettings, bootStatisticalWorkbench, changeChartNumber, chartSettingsPanel, chooseChartValue,
  openStatisticalSettings, openStatisticalTypePanel, selectedChartValue, statisticalChartHistory,
  statisticalChartSnapshot, statisticalPaint } from './helpers/chart-statistical-settings'

// P25 workbench drafts only. Source/Worker/Canvas assertions are unexecuted expectations.
async function insertHistogram(page: Page, title: string) {
  await bootStatisticalWorkbench(page)
  await page.evaluate(async name => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const table = [['Row', 'Value', 'Other'], ['A', 0, 40], ['B', 1, 30], ['C', 2, 20], ['D', 3, 10], ['E', 4, 0], ['Missing', null, 70]]
    sheet.getRange('A1:C7').setValues(table)
    const builder = sheet.newChart('histogram').setSource('A1:C7').setCategoryField(0).setHistogramField(1)
      .setAbsolutePosition(110, 90).setSize(560, 420).setTitle(name).setYAxis({ scaleType: 'linear', min: 0, max: 4 })
      .setBinCount(4).setUnderflowThreshold(1).setOverflowThreshold(3).setBinGap(0.25)
      .setOptions({ invalidValueStrategy: 'zero' })
      .setSeries(1, { color: 'blue', dataPoints: { 1: { color: 'red', fillOpacity: 0.4 } } })
    await sheet.insertChart(builder.build())
  }, title)
  await expect(page.getByRole('button', { name: title, exact: true })).toHaveAttribute('data-rendered', 'true')
}
async function openHistogramPoints(page: Page, editor: Locator) {
  await editor.getByText('系列、标记与数据标签', { exact: true }).click()
  await editor.getByText('单独设置系列', { exact: true }).click()
  await editor.getByText('单个数据项目', { exact: true }).click()
  return chartSettingsPanel(page, editor, '单个数据项目')
}
function blueRedBars(paint: string) {
  return statisticalPaint(paint).filter(command => command.kind === 'rect'
    && (command.fill?.startsWith('rgba(0, 0, 255,') || command.fill?.startsWith('rgba(255, 0, 0,')))
}

test('Histogram data controls precede paint and materialized bin point edits apply one update, history and internal reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Histogram materialized bins draft'
  await insertHistogram(page, title)
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), data = chartSettingsPanel(page, editor, '数据源与字段')
  expect(before.legend.visible).toBe(false)
  expect(before.layout.categories).toEqual(['≤ 1', '[2, 2.5]', '(2.5, 3]', '> 3'])
  expect(before.layout.series[0].points.map((point: { value: number }) => point.value)).toEqual([2, 1, 1, 1])
  expect(before.raw.series[0].values).toEqual([0, 1, 2, 3, 4, null])
  await expect(selectedChartValue(page, data, '样本字段')).toHaveText('Value')
  await expect(data.locator('.ant-form-item').filter({ has: page.getByRole('combobox', { name: '样本字段', exact: true }) }).last().locator('.ant-select'))
    .not.toHaveClass(/ant-select-multiple/)
  await expect(selectedChartValue(page, data, '分箱方式')).toHaveText('按箱数')
  await expect(data.getByRole('spinbutton', { name: '箱数', exact: true })).toHaveValue('4')
  const type = await openStatisticalTypePanel(page, editor)
  await expect(type.getByRole('combobox', { name: '分箱方式', exact: true })).toHaveCount(0)
  await expect(type.getByRole('spinbutton', { name: '下溢箱阈值（≤）', exact: true })).toHaveCount(0)
  await changeChartNumber(type, '箱间隙比例', '0.5')
  const points = await openHistogramPoints(page, editor)
  await changeChartNumber(points, '数据项目序号', '2')
  await expect(selectedChartValue(page, points, '选择数据项目')).toHaveText('2 · [2, 2.5]')
  await changeChartNumber(points, '项目填充不透明度', '0.6')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const applied = await statisticalChartSnapshot(page), bars = blueRedBars(applied.paint)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.histogram.binGap).toBe(0.5); expect(applied.options.series[1].dataPoints[1]).toMatchObject({ color: 'red', fillOpacity: 0.6 })
  expect(applied.raw).toEqual(before.raw); expect(applied.layout).toEqual(before.layout)
  expect(bars).toHaveLength(4)
  for (const bar of bars) expect(bar.width! / (applied.layout.plot.width / 4)).toBeCloseTo(0.5, 10)
  expect(bars[1]!.fill).toBe('rgba(255, 0, 0, 0.6)')
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(applied.options)
  const resetEditor = await openStatisticalSettings(page, title), resetPoints = await openHistogramPoints(page, resetEditor)
  await changeChartNumber(resetPoints, '数据项目序号', '2')
  await resetPoints.getByRole('button', { name: '恢复项目默认样式', exact: true }).click()
  await applyStatisticalSettings(resetEditor)
  const restored = await statisticalChartSnapshot(page)
  expect(restored.options.series[1].dataPoints[1]).toBeUndefined()
  expect(blueRedBars(restored.paint).every(bar => bar.fill === 'rgba(0, 0, 255, 1)')).toBe(true)
  expect(restored.options.histogram).toEqual(applied.options.histogram)
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('Width uses threshold origin and truncates its last bin, while Count/threshold/gap resets restore Scott Auto through the real settings command', async ({ page }) => {
  const title = 'Histogram binning restore draft'
  await insertHistogram(page, title)
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), data = chartSettingsPanel(page, editor, '数据源与字段')
  await chooseChartValue(page, data, '分箱方式', '按箱宽')
  await expect(data.getByRole('spinbutton', { name: '箱数', exact: true })).toHaveCount(0)
  await changeChartNumber(data, '箱宽', '0.75')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const width = await statisticalChartSnapshot(page)
  expect(width.options.histogram.binning).toEqual({ mode: 'width', width: 0.75 })
  expect(width.layout.categories).toEqual(['≤ 1', '(1, 1.75]', '(1.75, 2.5]', '(2.5, 3]', '> 3'])
  expect(width.layout.series[0].points.map((point: { value: number }) => point.value)).toEqual([2, 0, 1, 1, 1])
  const defaultsEditor = await openStatisticalSettings(page, title), defaultsData = chartSettingsPanel(page, defaultsEditor, '数据源与字段')
  await chooseChartValue(page, defaultsData, '分箱方式', '按箱数'); await changeChartNumber(defaultsData, '箱数', '4')
  await defaultsData.getByRole('button', { name: '箱数恢复默认', exact: true }).click()
  await expect(selectedChartValue(page, defaultsData, '分箱方式')).toHaveText('Scott 自动')
  await expect(defaultsData.getByRole('spinbutton', { name: '箱数', exact: true })).toHaveCount(0)
  await defaultsData.getByRole('button', { name: '下溢箱阈值（≤）恢复默认', exact: true }).click()
  await defaultsData.getByRole('button', { name: '上溢箱阈值（>）恢复默认', exact: true }).click()
  const defaultsType = await openStatisticalTypePanel(page, defaultsEditor)
  await defaultsType.getByRole('button', { name: '箱间隙比例恢复默认', exact: true }).click()
  await expect(defaultsType.getByRole('spinbutton', { name: '箱间隙比例', exact: true })).toHaveValue('0.007')
  expect((await statisticalChartSnapshot(page)).contents).toBe(width.contents)
  await applyStatisticalSettings(defaultsEditor)
  const defaults = await statisticalChartSnapshot(page)
  expect(defaults.options.histogram).toEqual({})
  expect(defaults.layout.categories).toEqual(['[0, 2]', '(2, 4]'])
  expect(defaults.layout.series[0].points.map((point: { value: number }) => point.value)).toEqual([3, 2])
  for (const bar of blueRedBars(defaults.paint)) expect(bar.width! / (defaults.layout.plot.width / 2)).toBeCloseTo(0.993, 10)
  expect(defaults.raw).toEqual(before.raw)
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: defaults.options, raw: defaults.raw, layout: defaults.layout, paint: defaults.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(width.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(defaults.options)
})
