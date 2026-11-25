import assert from 'node:assert/strict'
import { resolve } from 'node:path'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import { createServer } from 'vite'

// P21 deferred drafts: real ExcelWorkbench, Ant controls and Canvas2D hits.
// No Vite/dev process, browser, test, check, build, WASM or benchmark was started.
let server, browser, context, page, baseURL
const errors = []
beforeAll(async () => {
  server = await createServer({ configFile: resolve(process.cwd(), 'vite.config.ts'), root: process.cwd(),
    server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('No HTTP port was allocated')
  baseURL = `http://127.0.0.1:${address.port}/packages/office-studio/tests/fixtures/chart-object-actions.html`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)
beforeEach(async () => {
  errors.length = 0
  context = await browser.newContext({ viewport: { width: 1440, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] })
  page = await context.newPage(); page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await page.waitForFunction(() => window.chartObjectFixture?.state().charts.length === 2)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}, 60_000)
afterEach(async () => { try { await page?.evaluate(() => window.chartObjectFixture?.dispose()); assert.deepEqual(errors, []) } finally { await context?.close() } })
afterAll(async () => { await browser?.close(); await server?.close() })
const state = () => page.evaluate(() => window.chartObjectFixture.state())
const chart = value => value.charts.find(chart => chart.id === 'native-object')
const hitbox = () => page.locator('.chart-hitbox[data-chart-id="native-object"]')
const menu = () => page.locator('.chart-element-float-menu')
const field = key => menu().locator(`[data-chart-option-key=${JSON.stringify(key)}]`)
const drawer = () => page.locator('.ant-drawer').filter({ has: page.getByText('图表设置', { exact: true }) })
const drawerField = key => drawer().locator(`[data-chart-option-key=${JSON.stringify(key)}]`)
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const rows = [['Category', 'Original'], ['A', 8], ['B', 5], ['C', 3], ['D', 2]]
const compoundOptions = () => ({ series: { 1: { dataPoints: { 2: { color: '#334455' } } } }, pie: {
  label: { visible: true, position: 'outside', fontSize: 12 }, aggregateLabel: 'Others', aggregateStyle: { color: '#445566' },
  sliceStyles: { 0: { color: '#112233' }, 2: { color: '#abcdef' } }, secondaryPlot: { split: { type: 'position', count: 2 } },
} })
async function mount(options) {
  await page.evaluate(options => window.chartObjectFixture.mount(options), options)
  await expect(hitbox()).toHaveAttribute('data-rendered', 'true')
}
async function frames() { await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))) }
async function closeMenu() {
  if (await menu().count()) { await menu().getByRole('button', { name: '关闭图表元素工具栏', exact: true }).click(); await expect(menu()).toHaveCount(0) }
}
async function closeDrawer() {
  await drawer().getByRole('button', { name: '关闭', exact: true }).click(); await expect(drawer()).not.toBeVisible()
}
async function select(selector, { shift = false } = {}) {
  await closeMenu(); const measured = await page.evaluate(selector => window.chartObjectFixture.elementPoint(selector), selector)
  if (shift) await page.keyboard.down('Shift')
  try { await page.mouse.click(measured.viewport.x, measured.viewport.y) } finally { if (shift) await page.keyboard.up('Shift') }
  await expect(menu()).toBeVisible(); await expect(menu()).toHaveAttribute('data-chart-element-kind', shift ? 'series' : selector.kind)
  if (!shift && selector.markKind) await expect(menu()).toHaveAttribute('data-chart-mark-kind', selector.markKind)
  return measured
}
async function number(key, value) { const input = field(key).getByRole('spinbutton'); await input.fill(String(value)); await input.blur() }
async function choose(key, label) {
  await field(key).getByRole('combobox').click(); await menu().locator('.ant-select-dropdown:visible').getByText(label, { exact: true }).click()
}
async function settledAuthor() {
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 0
    && document.querySelector('.chart-hitbox[data-chart-id="native-object"]')?.getAttribute('aria-busy') === 'false')
  await frames()
}
async function apply() {
  const button = menu().getByRole('button', { name: '应用图表元素修改', exact: true })
  await expect(button).toBeEnabled(); await button.click(); await settledAuthor()
}
async function more() {
  await menu().getByRole('button', { name: '图表元素更多设置', exact: true }).click()
  await expect(menu()).toHaveCount(0); await expect(drawer()).toBeVisible()
}
async function promote() {
  await menu().getByRole('button', { name: '选择整个图表系列', exact: true }).click()
  await expect(menu()).toHaveAttribute('data-chart-element-kind', 'mark'); await expect(menu()).toHaveAttribute('data-chart-mark-kind', 'seriesStyle')
}

test('point promotion edits a dedicated numeric seriesStyle through one author command, preserves point maps and follows undo/save/reopen', async () => {
  await mount({ nativeType: 'line', multiSeries: true, rows: [['Category', 'Original', 'Comparison'], ['A', 8, 2], ['B', 5, 3], ['C', 3, 7]],
    nativeOptions: { series: { 1: { name: 'Display name', point: { size: 8 }, dataPoints: { 1: { size: 14 } } },
      2: { color: '#223344', rightAxis: true, name: 'Other display' } } } })
  const before = await state(); await select({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 1 }); await promote()
  await expect(menu()).toHaveAttribute('data-chart-element-key', 'mark:seriesStyle:field:1')
  await expect(field('series.1.name')).toHaveCount(0); await number('series.1.fillOpacity', 0.45); await number('series.1.border.width', 3)
  assert.deepEqual((await state()).charts, before.charts); await apply()
  let next = await state(); assert.equal(chart(next).options.series[1].fillOpacity, 0.45); assert.equal(chart(next).options.series[1].border.width, 3)
  assert.equal(chart(next).options.series[1].name, 'Display name'); assert.deepEqual(chart(next).options.series[1].dataPoints, chart(before).options.series[1].dataPoints)
  assert.deepEqual(chart(next).options.series[2], chart(before).options.series[2]); assert.equal(next.history.undo, before.history.undo + 1)
  await expect(menu().getByRole('button', { name: '应用图表元素修改', exact: true })).toBeDisabled()
  await menu().getByRole('button', { name: '恢复图表元素默认格式', exact: true }).click(); await apply()
  next = await state(); assert.equal(chart(next).options.series[1].fillOpacity, undefined); assert.deepEqual(chart(next).options.series[1].dataPoints, chart(before).options.series[1].dataPoints)
  assert.equal(next.history.undo, before.history.undo + 2); assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
  await page.evaluate(() => window.chartObjectFixture.undo()); assert.equal(chart(await state()).options.series[1].fillOpacity, 0.45)
  await page.evaluate(() => window.chartObjectFixture.redo()); await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
  await select({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 1 }, { shift: true }); await expect(menu()).toHaveAttribute('data-chart-element-key', 'series:field:1')
}, 60_000)

test('an exact-name seriesStyle uses the visible author layer and More opens its existing full series form after a scoped reset', async () => {
  await mount({ nativeType: 'line', rows, nativeOptions: { series: { 1: { color: '#112233' } },
    namedSeries: [{ name: 'Original', style: { name: 'Named display', color: '#abcdef', dataPoints: { 2: { size: 16 } } } }] } })
  const before = await state(); await select({ kind: 'dataPoint', seriesId: 'name:"Original"', dataIndex: 1 }); await promote()
  await number('namedSeries."Original".fillOpacity', 0.5); await apply()
  let next = await state(); assert.equal(chart(next).options.namedSeries[0].style.fillOpacity, 0.5); assert.deepEqual(chart(next).options.series, chart(before).options.series)
  await menu().getByRole('button', { name: '恢复图表元素默认格式', exact: true }).click(); await apply(); next = await state()
  assert.equal(chart(next).options.namedSeries[0].style.color, undefined); assert.equal(chart(next).options.namedSeries[0].style.name, 'Named display')
  assert.deepEqual(chart(next).options.namedSeries[0].style.dataPoints, chart(before).options.namedSeries[0].style.dataPoints)
  const history = next.history.undo; await more()
  await expect(drawer().locator('[data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', 'name:"Original"')
  await expect(drawer().getByRole('combobox', { name: '选择图表系列', exact: true })).toBeFocused()
  assert.equal((await state()).history.undo, history)
  assert.deepEqual((await state()).commands.find(command => command.id === 'sheet.command.chart.open-focused-panel').params.route,
    { tab: 'style', section: 'series', seriesId: 'name:"Original"' })
}, 60_000)

test('the measured global Pie label mark edits every label, resets its own layer and consumes Delete through canonical history', async () => {
  await mount({ nativeType: 'pie', rows, nativeOptions: { series: { 1: { label: { visible: true, fontSize: 11 } } },
    pie: { label: { visible: true, position: 'inside', fontSize: 14 }, sliceStyles: { 2: { color: '#abcdef' } } } } })
  const before = await state(); await select({ kind: 'mark', markKind: 'pieStyle' }); await expect(menu()).toContainText('全部扇区的数据标签')
  await number('pie.label.fontSize', 19); await apply(); let next = await state()
  assert.equal(chart(next).options.pie.label.fontSize, 19); assert.deepEqual(chart(next).options.pie.sliceStyles, chart(before).options.pie.sliceStyles)
  const labels = (await page.evaluate(() => window.chartObjectFixture.scene())).filter(command => command.kind === 'text' && command.chartCanvasElement?.directEdit?.kind === 'pieStyle')
  assert.equal(labels.length, 4); assert.ok(labels.every(command => command.font.includes('19px')))
  await menu().getByRole('button', { name: '恢复图表元素默认格式', exact: true }).click(); await apply()
  next = await state(); assert.equal(chart(next).options.pie.label, undefined); assert.deepEqual(chart(next).options.series, chart(before).options.series)
  await select({ kind: 'mark', markKind: 'pieStyle' }); await hitbox().focus(); await hitbox().press('Delete'); await expect(menu()).toHaveCount(0)
  await settledAuthor(); next = await state(); assert.equal(chart(next).options.pie.label.visible, false)
  assert.equal(next.charts.length, 2); assert.equal(next.history.undo, before.history.undo + 3)
  assert.ok((await page.evaluate(() => window.chartObjectFixture.elements())).every(target => target.element.directEdit?.kind !== 'pieStyle'))
  await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
}, 60_000)

test('composite source-slice edits preserve source points and generated aggregate and follow the real plot after repartitioning', async () => {
  for (const nativeType of ['pieOfPie', 'barOfPie']) {
    await mount({ nativeType, rows, nativeOptions: compoundOptions() }); const before = await state(),
      selected = await select({ kind: 'mark', markKind: 'pieSliceStyle', sourcePointIndex: 2, plot: 'secondary' })
    assert.equal(selected.target.element.directEdit.sourceField, 1); assert.equal(selected.target.element.directEdit.sourcePointIndex, 2)
    await number('pie.sliceStyles.2.fillOpacity', 0.4)
    if (nativeType === 'pieOfPie') await number('pie.sliceStyles.2.explosion', 20)
    else await expect(field('pie.sliceStyles.2.explosion')).toHaveCount(0)
    await apply(); let next = await state()
    assert.equal(chart(next).options.pie.sliceStyles[2].fillOpacity, 0.4); assert.deepEqual(chart(next).options.series, chart(before).options.series)
    assert.deepEqual(chart(next).options.pie.aggregateStyle, chart(before).options.pie.aggregateStyle)
    assert.deepEqual(chart(next).options.pie.sliceStyles[0], chart(before).options.pie.sliceStyles[0]); assert.equal(next.history.undo, before.history.undo + 1)
    await page.evaluate(() => window.chartObjectFixture.updateChart({ options: { pie: { secondaryPlot: { split: { type: 'position', count: 1 } } } } }))
    await select({ kind: 'mark', markKind: 'pieSliceStyle', sourcePointIndex: 2, plot: 'primary' })
    await expect(field('pie.sliceStyles.2.explosion')).toHaveCount(1); await number('pie.sliceStyles.2.explosion', 15); await apply()
    next = await state(); assert.equal(chart(next).options.pie.sliceStyles[2].explosion, 15)
    assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
    await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
    const reopened = await page.evaluate(() => window.chartObjectFixture.elementPoint({ kind: 'mark', markKind: 'pieSliceStyle', sourcePointIndex: 2, plot: 'primary' }))
    assert.equal(reopened.target.element.directEdit.sourcePointIndex, 2)
  }
}, 60_000)

test('dirty Waterfall More focuses the exact numeric, named or global role field and its real full form resets remain scoped', async () => {
  const cases = [
    { nativeType: 'waterfall', multiSeries: true, rows: [['Category', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]],
      nativeOptions: { waterfall: { pointRoles: [], connector: true }, series: { 1: { waterfallStyles: { negative: { dataPoints: { 1: { color: '#abcdef' } } } } } } },
      seriesId: 'field:1', role: 'negative', prefix: 'series.1' },
    { nativeType: 'waterfall', rows: [['Category', 'Original'], ['A', 6], ['B', -2], ['C', 4]],
      nativeOptions: { waterfall: { pointRoles: [], connector: true }, namedSeries: [{ name: 'Original', style: { color: '#abcdef' } }] },
      seriesId: 'name:"Original"', role: 'positive', prefix: 'namedSeries."Original"' },
    { nativeType: 'waterfall', multiSeries: true, rows: [['Category', 'First', 'Second'], ['A', 10, 4], ['B', -3, -2]],
      nativeOptions: { waterfall: { stackType: 'stacked', useSubtotal: true, connector: true } }, role: 'subtotal', prefix: 'waterfall.styles' },
  ]
  for (const item of cases) {
    await mount(item); const before = await state()
    await select({ kind: 'mark', markKind: 'waterfallStyle', markPart: item.role, ...(item.seriesId ? { seriesId: item.seriesId } : { globalMark: true }) })
    const key = item.seriesId ? `${item.prefix}.waterfallStyles.${item.role}` : `${item.prefix}.${item.role}`
    await number(`${key}.fillOpacity`, 0.5); await more()
    await expect(drawer().locator('.chart-waterfall-settings')).toHaveAttribute('data-chart-waterfall-role', item.role)
    await expect(drawer().locator('.chart-waterfall-settings [data-chart-route-series]')).toHaveAttribute('data-chart-selected-series', item.seriesId ?? 'all')
    await expect(drawerField(`${key}.color`).getByRole('combobox')).toBeFocused()
    const next = await state(), command = next.commands.find(command => command.id === 'sheet.command.chart.open-focused-panel')
    assert.deepEqual(command.params.route.focus, { kind: 'waterfallRole', role: item.role }); assert.equal(command.params.route.seriesId, item.seriesId)
    assert.equal(next.history.undo, before.history.undo + 1); assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
    assert.equal(chart(next).options.series?.[4294967295], undefined)
    if (item.seriesId === 'field:1') {
      await drawer().getByText('角色内的单个数据项目', { exact: true }).click()
      const point = drawer().getByRole('combobox', { name: '选择瀑布图角色数据项目', exact: true })
      await point.click(); await point.press('ArrowDown'); await point.press('Enter')
      await drawer().getByRole('button', { name: '恢复角色项目继承样式', exact: true }).click()
      await drawer().getByRole('button', { name: '应用', exact: true }).click(); await settledAuthor()
      const reset = await state(); assert.equal(chart(reset).options.series[1].waterfallStyles.negative.dataPoints[1], undefined)
      assert.equal(chart(reset).options.series[1].waterfallStyles.negative.fillOpacity, 0.5)
      assert.deepEqual(reset.model.sheets[0].cells, before.model.sheets[0].cells)
    } else if (!item.seriesId) {
      await drawer().getByRole('button', { name: '恢复该角色继承样式', exact: true }).click()
      await drawer().getByRole('button', { name: '应用', exact: true }).click(); await settledAuthor()
      assert.equal(chart(await state()).options.waterfall.styles.subtotal, undefined)
    }
    await closeDrawer()
  }
}, 60_000)

test('Pie More focuses aggregate, connector, global label and source-slice controls and the full slice reset survives save/reopen', async () => {
  await mount({ nativeType: 'pieOfPie', rows, nativeOptions: compoundOptions() })
  const cases = [
    { markKind: 'pieAggregateStyle', focus: { kind: 'pieAggregate' }, dirty: 'pie.aggregateStyle.fillOpacity', value: 0.4, key: 'pie.aggregateStyle.color', role: 'combobox' },
    { markKind: 'pieConnectorStyle', focus: { kind: 'pieConnector' }, dirty: 'pie.secondaryPlot.connector.width', value: 4, key: 'pie.secondaryPlot.connector.visible', role: 'switch' },
    { markKind: 'pieStyle', focus: { kind: 'pieLabel' }, dirty: 'pie.label.fontSize', value: 19, key: 'pie.label.visible', role: 'switch' },
    { markKind: 'pieSliceStyle', sourcePointIndex: 2, plot: 'secondary', focus: { kind: 'pieSlice', sourcePointIndex: 2 },
      dirty: 'pie.sliceStyles.2.fillOpacity', value: 0.5, key: 'pie.sliceStyles.2.color', role: 'combobox' },
  ]
  for (const item of cases) {
    const before = await state(); await select({ kind: 'mark', markKind: item.markKind,
      ...(item.sourcePointIndex === undefined ? {} : { sourcePointIndex: item.sourcePointIndex, plot: item.plot }) })
    await number(item.dirty, item.value); await more(); await expect(drawerField(item.key).getByRole(item.role)).toBeFocused()
    const next = await state(), command = next.commands.filter(command => command.id === 'sheet.command.chart.open-focused-panel').at(-1)
    assert.deepEqual(command.params.route.focus, item.focus); assert.equal(next.history.undo, before.history.undo + 1)
    if (item.sourcePointIndex === undefined) await closeDrawer()
  }
  await expect(drawer().locator('.chart-pie-slice-settings')).toHaveAttribute('data-chart-pie-source-point', '2')
  const before = await state(); await drawer().getByRole('button', { name: '恢复所选扇区继承样式', exact: true }).click()
  assert.deepEqual((await state()).charts, before.charts)
  await drawer().getByRole('button', { name: '应用', exact: true }).click(); await settledAuthor()
  const next = await state(); assert.equal(chart(next).options.pie.sliceStyles[2].color, undefined); assert.equal(chart(next).options.pie.sliceStyles[2].fillOpacity, undefined)
  assert.deepEqual(chart(next).options.series, chart(before).options.series); assert.deepEqual(chart(next).options.pie.aggregateStyle, chart(before).options.pie.aggregateStyle)
  assert.deepEqual(chart(next).options.pie.sliceStyles[0], chart(before).options.pie.sliceStyles[0]); assert.equal(next.history.undo, before.history.undo + 1)
  await closeDrawer(); await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
}, 60_000)

test('real native and shared Boxplot summary/outlier edits preserve statistics and sample indices through canonical author history', async () => {
  const samples = [100, null, -100, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 100, -100], sourceRows = [['Category', 'Samples'], ...samples.map((value, index) => [`Sample ${index + 1}`, value])]
  for (const shared of [false, true]) {
    await mount({ nativeType: 'boxplot', rows: sourceRows, nativeConfig: { yAxis: { scaleType: 'linear', min: -120, max: 120 } }, nativeOptions: {
      boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr', showOutliers: true }, series: { 1: { color: '#112233',
        ...(shared ? { fillPattern: { preset: 'ltUpDiag', foregroundColor: 'black', backgroundColor: 'white' } } : {}),
        dataPoints: { 0: { size: 18, shape: 'rect' }, 2: { size: 16, shape: 'rect' }, 14: { size: 8, shape: 'diamond' }, 13: { size: 6, shape: 'circle', fillPattern: null } },
      } } } })
    const before = await state(); await select({ kind: 'mark', markKind: 'seriesStyle', seriesId: 'field:1' })
    await number('series.1.fillOpacity', 0.6); await apply()
    const measured = await select({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 13 })
    assert.equal(measured.target.element.dataIndex, 13); await number('point.size', 20); await choose('point.shape', '菱形'); await apply()
    const next = await state(); assert.equal(chart(next).options.series[1].fillOpacity, 0.6)
    assert.equal(chart(next).options.series[1].dataPoints[13].size, 20); assert.equal(chart(next).options.series[1].dataPoints[13].shape, 'diamond')
    assert.deepEqual(chart(next).options.series[1].dataPoints[0], chart(before).options.series[1].dataPoints[0]); assert.equal(next.history.undo, before.history.undo + 2)
    const after = await page.evaluate(() => window.chartObjectFixture.elementPoint({ kind: 'dataPoint', seriesId: 'field:1', dataIndex: 13 }))
    assert.ok(after.target.bounds.width > measured.target.bounds.width); assert.equal(after.target.element.dataIndex, 13)
    const computed = await page.evaluate(() => window.chartObjectFixture.data())
    assert.deepEqual(computed.categories, sourceRows.slice(1).map(row => row[0])); assert.deepEqual(computed.series[0].values, samples)
    assert.deepEqual(next.model.sheets[0].cells, before.model.sheets[0].cells)
    await hitbox().focus(); await hitbox().press('Delete'); assert.equal((await state()).history.undo, next.history.undo); assert.equal((await state()).charts.length, 2)
    await page.evaluate(() => window.chartObjectFixture.reopen()); assert.deepEqual(chart(await state()), chart(next))
  }
}, 60_000)

test('new label/slice mark drafts honor read-only and IME and cannot revive after a pending completion changes chart type and active sheet', async () => {
  await mount({ readonly: true, nativeType: 'pieOfPie', rows, nativeOptions: compoundOptions() }); const guarded = await state()
  await select({ kind: 'mark', markKind: 'pieSliceStyle', sourcePointIndex: 2, plot: 'secondary' })
  await expect(field('pie.sliceStyles.2.fillOpacity').getByRole('spinbutton')).toBeDisabled()
  await expect(menu().getByRole('button', { name: '图表元素更多设置', exact: true })).toBeDisabled()
  await hitbox().focus(); await hitbox().press('Delete'); assert.deepEqual((await state()).charts, guarded.charts)
  await mount({ nativeType: 'pieOfPie', rows, nativeOptions: compoundOptions() }); const before = await state()
  await select({ kind: 'mark', markKind: 'pieStyle' }); const input = field('pie.label.format').getByRole('textbox')
  await input.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })); node.value = '0.000'
    node.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '0', isComposing: true }))
    for (const key of ['Enter', 'Escape']) node.dispatchEvent(new KeyboardEvent('keydown', { key, isComposing: true, bubbles: true, cancelable: true })) })
  await expect(menu()).toBeVisible(); assert.deepEqual((await state()).charts, before.charts)
  await input.evaluate(node => { node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '0.000' })); node.dispatchEvent(new InputEvent('input', { bubbles: true })) })
  await input.press('Escape'); await expect(menu()).toHaveCount(0); await expect(hitbox()).toBeFocused(); assert.equal((await state()).history.undo, before.history.undo)
  await select({ kind: 'mark', markKind: 'pieSliceStyle', sourcePointIndex: 2, plot: 'secondary' }); await number('pie.sliceStyles.2.fillOpacity', 0.4)
  await page.evaluate(() => window.chartObjectFixture.holdAuthor()); await menu().getByRole('button', { name: '应用图表元素修改', exact: true }).click()
  await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 1)
  await page.evaluate(() => { void window.chartObjectFixture.updateChart({ type: 'line' }) }); await page.waitForFunction(() => window.chartObjectFixture.state().pendingAuthors === 2)
  await expect(menu()).toHaveCount(0); await page.evaluate(() => window.chartObjectFixture.switchSheet('destination')); await grid().focus()
  await page.evaluate(() => window.chartObjectFixture.finishAuthors()); await frames()
  await expect(menu()).toHaveCount(0); await expect(drawer()).toHaveCount(0); await expect(grid()).toBeFocused()
  assert.equal(chart(await state()).options.pie.sliceStyles[2].fillOpacity, 0.4); assert.deepEqual((await state()).model.sheets[0].cells, before.model.sheets[0].cells)
}, 60_000)
