import { expect, test, type Page } from '@playwright/test'
import { applyStatisticalSettings, bootStatisticalWorkbench, changeChartNumber, chartSettingsPanel, chooseChartValue,
  openStatisticalSettings, openStatisticalTypePanel, selectedChartValue, statisticalChartHistory,
  statisticalChartSnapshot, statisticalPaint } from './helpers/chart-statistical-settings'

// P25 workbench drafts only. No test, server, browser, theme/IME, Office or benchmark ran.
async function insertCandlestick(page: Page, title: string, mode: 'ohlc' | 'hlc' | 'empty-hlc' = 'ohlc', legacy = false) {
  await bootStatisticalWorkbench(page)
  await page.evaluate(async ({ name, mode, legacy }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const rows = mode === 'empty-hlc' ? [['First', null, null, null, null], ['Second', null, null, null, null]]
      : [['Rising', 10, 15, 8, 12], ['Falling', 12, 14, 7, 9], ['Doji', 11, 13, 10, 11], ['Gap', null, 13, 9, 10]]
    const table = [['Date', 'Open', 'High', 'Low', 'Close'], ...rows], range = 'A1:E' + table.length
    sheet.getRange(range).setValues(table)
    const builder = sheet.newChart('candlestick').setSource(range).setCategoryField(0).setValueFields([1, 2, 3, 4])
      .setCandlestickFields({ ...(mode === 'ohlc' ? { openIndex: 1 } : {}), highIndex: 2, lowIndex: 3, closeIndex: 4 })
      .setAbsolutePosition(110, 90).setSize(560, 420).setTitle(name).setYAxis({ scaleType: 'linear', min: 0, max: 20 })
      .setCandleWidth(40)
      .setRisingStyle({ color: 'red', fillOpacity: 0.3, border: { color: 'navy', width: 3, opacity: 0.8 } })
      .setFallingStyle({ color: 'blue', hollow: true, border: { color: 'green', width: 2, opacity: 0.6 } })
      .setDojiStyle({ border: { color: 'black', width: 2, opacity: 0.4 } })
    if (legacy) builder.setOptions({ candlestick: { risingColor: 'purple', fallingColor: 'orange', hollow: true } })
    await sheet.insertChart(builder.build())
  }, { name: title, mode, legacy })
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}

test('antd state inspection stays detached and independent Rising/Falling/Doji edits publish one update with history and internal reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Candlestick independent states draft'
  await insertCandlestick(page, title)
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), panel = await openStatisticalTypePanel(page, editor)
  expect(before.legend.visible).toBe(false)
  await expect(selectedChartValue(page, panel, 'K 线状态')).toHaveText('上涨')
  await chooseChartValue(page, panel, 'K 线状态', '十字线')
  await expect(panel.getByRole('spinbutton', { name: '蜡烛填充不透明度', exact: true })).toHaveCount(0)
  await expect(panel.getByRole('switch', { name: '空心蜡烛', exact: true })).toHaveCount(0)
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await changeChartNumber(panel, '蜡烛边框宽度', '5'); await changeChartNumber(panel, '蜡烛边框不透明度', '0.35')
  await chooseChartValue(page, panel, 'K 线状态', '下跌')
  await changeChartNumber(panel, '蜡烛填充不透明度', '0.65')
  await panel.getByRole('switch', { name: '空心蜡烛', exact: true }).click()
  await changeChartNumber(panel, '蜡烛宽度', '48')
  await chooseChartValue(page, panel, 'K 线状态', '上涨')
  await changeChartNumber(panel, '蜡烛填充不透明度', '0.45')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const applied = await statisticalChartSnapshot(page), paint = statisticalPaint(applied.paint)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.candlestick).toMatchObject({ candleWidth: 48, rising: { fillOpacity: 0.45 },
    falling: { fillOpacity: 0.65, hollow: false }, doji: { border: { width: 5, opacity: 0.35 } } })
  expect(applied.raw.candleFields).toEqual({ openIndex: 1, highIndex: 2, lowIndex: 3, closeIndex: 4 })
  expect(applied.raw.candles[3]).toBeNull()
  expect(paint.some(command => command.kind === 'rect' && command.width === 48 && command.fill === 'rgba(255, 0, 0, 0.45)')).toBe(true)
  expect(paint.some(command => command.kind === 'rect' && command.width === 48 && command.fill === 'rgba(0, 0, 255, 0.65)')).toBe(true)
  expect(paint.some(command => command.kind === 'line' && command.width === 5 && command.color === 'rgba(0, 0, 0, 0.35)')).toBe(true)
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(applied.options)
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('explicit HLC keeps an empty Open selection with four legacy value fields and applies conditional high/low and close styles', async ({ page }) => {
  const title = 'Candlestick HLC overlays draft'
  await insertCandlestick(page, title, 'hlc')
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title)
  const data = chartSettingsPanel(page, editor, '数据源与字段'), panel = await openStatisticalTypePanel(page, editor)
  await expect(selectedChartValue(page, data, '开盘字段（可选）')).toHaveCount(0)
  await expect(selectedChartValue(page, data, '收盘字段')).toHaveText('Close')
  await expect(panel.getByRole('combobox', { name: 'K 线状态', exact: true })).toHaveCount(0)
  await expect(panel.getByRole('spinbutton', { name: '蜡烛宽度', exact: true })).toHaveCount(0)
  await expect(panel.getByRole('switch', { name: '显示高低线', exact: true })).toBeChecked()
  await expect(panel.getByRole('switch', { name: '显示收盘标记', exact: true })).toBeChecked()
  await panel.getByRole('switch', { name: '显示高低线', exact: true }).click()
  await expect(panel.getByRole('spinbutton', { name: '高低线宽度', exact: true })).toHaveCount(0)
  await panel.getByRole('switch', { name: '显示高低线', exact: true }).click()
  await changeChartNumber(panel, '高低线宽度', '3'); await changeChartNumber(panel, '高低线不透明度', '0.5')
  await chooseChartValue(page, panel, '高低线线型', '虚线')
  await chooseChartValue(page, panel, '收盘标记形状', '短划线'); await changeChartNumber(panel, '收盘标记大小', '16')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const applied = await statisticalChartSnapshot(page)
  expect(applied.options.mapping.candlestick).toEqual({ highIndex: 2, lowIndex: 3, closeIndex: 4 })
  expect(applied.valueFields).toEqual([1, 2, 3, 4])
  expect(applied.options.candlestick).toMatchObject({ highLowLine: { visible: true, border: { width: 3, opacity: 0.5, dashType: 'dashed' } },
    closeMarker: { point: { shape: 'dash', size: 16 } } })
  expect(statisticalPaint(applied.paint).filter(command => command.kind === 'line' && command.color === 'rgba(78, 85, 101, 1)' && command.width === 1
    && command.x1 !== command.x2 && command.y1 === command.y2).map(command => command.x2! - command.x1!)).toEqual([16, 16, 16, 16])
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('B2').setValues([[99]])
  })
  const refreshed = await statisticalChartSnapshot(page)
  expect(refreshed.raw.candleFields).toEqual({ highIndex: 2, lowIndex: 3, closeIndex: 4 })
  expect(refreshed.raw.candles.every((candle: { open?: number }) => candle.open === undefined)).toBe(true)
  expect(refreshed.paint).toBe(applied.paint)
  const resetEditor = await openStatisticalSettings(page, title), resetPanel = await openStatisticalTypePanel(page, resetEditor)
  await chooseChartValue(page, resetPanel, '收盘标记形状', '无标记')
  await applyStatisticalSettings(resetEditor)
  const none = await statisticalChartSnapshot(page)
  expect(statisticalPaint(none.paint).filter(command => command.kind === 'line' && command.color === 'rgba(78, 85, 101, 1)' && command.width === 1
    && command.x1 !== command.x2 && command.y1 === command.y2)).toHaveLength(0)
  const defaultsEditor = await openStatisticalSettings(page, title), defaults = await openStatisticalTypePanel(page, defaultsEditor)
  await defaults.getByRole('button', { name: '恢复 K 线高低线默认值', exact: true }).click()
  await defaults.getByRole('button', { name: '恢复 K 线收盘标记默认值', exact: true }).click()
  await expect(defaults.getByRole('switch', { name: '显示高低线', exact: true })).toBeChecked()
  await expect(defaults.getByRole('switch', { name: '显示收盘标记', exact: true })).toBeChecked()
  await expect(defaults.getByRole('spinbutton', { name: '高低线宽度', exact: true })).toHaveValue('1')
  await expect(selectedChartValue(page, defaults, '收盘标记形状')).toHaveText('点')
  await applyStatisticalSettings(defaultsEditor)
  const restored = await statisticalChartSnapshot(page)
  expect(restored.options.candlestick.highLowLine).toBeUndefined(); expect(restored.options.candlestick.closeMarker).toBeUndefined()
  expect(statisticalPaint(restored.paint).filter(command => command.kind === 'line' && command.color === 'rgba(78, 85, 101, 1)' && command.width === 1
    && command.x1 !== command.x2 && command.y1 === command.y2).map(command => command.x2! - command.x1!)).toEqual([4, 4, 4, 4])
})

test('state restore retains legacy inheritance while all-state restore removes it, and an empty HLC source never displays OHLC body controls', async ({ page }) => {
  const title = 'Candlestick legacy restore draft'
  await insertCandlestick(page, title, 'ohlc', true)
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), panel = await openStatisticalTypePanel(page, editor)
  await panel.getByRole('button', { name: '恢复所选 K 线状态继承样式', exact: true }).click()
  await panel.getByRole('button', { name: '蜡烛宽度恢复默认', exact: true }).click()
  await expect(panel.getByRole('switch', { name: '空心蜡烛', exact: true })).toBeChecked()
  await applyStatisticalSettings(editor)
  const inherited = await statisticalChartSnapshot(page)
  expect(inherited.options.candlestick.rising).toBeUndefined(); expect(inherited.options.candlestick.candleWidth).toBeUndefined()
  expect(inherited.options.candlestick).toMatchObject({ risingColor: 'purple', fallingColor: 'orange', hollow: true })
  expect(statisticalPaint(inherited.paint).some(command => command.kind === 'rect' && command.stroke === 'rgba(128, 0, 128, 1)' && command.fill === 'transparent')).toBe(true)
  const allEditor = await openStatisticalSettings(page, title), allPanel = await openStatisticalTypePanel(page, allEditor)
  await allPanel.getByRole('button', { name: '恢复全部 K 线默认设置', exact: true }).click()
  expect((await statisticalChartSnapshot(page)).contents).toBe(inherited.contents)
  await applyStatisticalSettings(allEditor)
  const defaults = await statisticalChartSnapshot(page)
  expect(defaults.options.candlestick).toBeUndefined(); expect(defaults.options.mapping).toEqual(before.options.mapping)
  expect(defaults.revision).toBe(before.revision + 2)
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(inherited.options)
  await insertCandlestick(page, 'Candlestick empty HLC draft', 'empty-hlc')
  const emptyEditor = await openStatisticalSettings(page, 'Candlestick empty HLC draft'), emptyPanel = await openStatisticalTypePanel(page, emptyEditor)
  await expect(emptyPanel.getByRole('combobox', { name: 'K 线状态', exact: true })).toHaveCount(0)
  await expect(emptyPanel.getByRole('spinbutton', { name: '蜡烛宽度', exact: true })).toHaveCount(0)
  await emptyEditor.getByRole('button', { name: '关闭', exact: true }).last().click()
})
