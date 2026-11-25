import { expect, test } from '@playwright/test'
import { insertHierarchyChart } from './helpers/chart-hierarchy-settings'
import { applyStatisticalSettings, changeChartNumber, chooseChartValue, openStatisticalSettings, openStatisticalTypePanel,
  selectedChartValue, statisticalChartHistory, statisticalChartSnapshot } from './helpers/chart-statistical-settings'

// P27 workbench drafts. None of these source/Worker/Canvas/history assertions ran.
interface HierarchyNode { name: string; value: number; selfValue: number; paletteIndex?: number; children?: HierarchyNode[] }
function semantic(nodes: HierarchyNode[]): unknown[] {
  return nodes.map(node => ({ name: node.name, value: node.value, self: node.selfValue, children: semantic(node.children ?? []) }))
}
function words(paint: string): string[] {
  return (JSON.parse(paint) as Array<{ kind: string; text?: string }>).filter(command => command.kind === 'text').map(command => command.text ?? '')
}

test('signed hierarchy, hidden default center and conditional node/center controls publish one isolated update', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Sunburst signed center draft'; await insertHierarchyChart(page, title, 'sunburst')
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), type = await openStatisticalTypePanel(page, editor)
  expect(semantic(before.raw.hierarchy)).toEqual([
    { name: 'North', value: 13, self: -5, children: [{ name: 'Leaf B', value: 12, self: 12, children: [] }, { name: 'Leaf C', value: 6, self: 6, children: [] }] },
    { name: 'South', value: 9, self: 3, children: [{ name: 'Leaf A', value: 6, self: 6, children: [] }] },
  ])
  expect(before.layout.sunburst.filter((node: { depth: number }) => node.depth === 0).map((node: { value: number }) => node.value)).toEqual([13, 9])
  await expect(editor.getByRole('combobox', { name: '图例位置', exact: true })).toHaveCount(0)
  await expect(type.getByRole('switch', { name: '显示中心标签', exact: true })).toHaveAttribute('aria-checked', 'false')
  await expect(type.getByRole('textbox', { name: '中心标签数字格式', exact: true })).toHaveCount(0)
  await expect(type.getByRole('switch', { name: '显示节点标签', exact: true })).toHaveAttribute('aria-checked', 'true')
  await expect(selectedChartValue(page, type, '节点标签位置')).toHaveText('自动')
  await type.getByRole('switch', { name: '显示中心标签', exact: true }).click()
  await type.getByRole('textbox', { name: '中心标签数字格式', exact: true }).fill('0.00')
  await changeChartNumber(type, '中心标签字号', '22')
  await type.getByRole('switch', { name: '显示节点标签', exact: true }).click()
  await expect(type.getByRole('combobox', { name: '节点标签内容', exact: true })).toHaveCount(0)
  await expect(type.getByRole('spinbutton', { name: '中心标签字号', exact: true })).toBeVisible()
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const applied = await statisticalChartSnapshot(page)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.sunburst).toMatchObject({ label: { visible: false }, centerLabel: { visible: true, format: '0.00', fontSize: 22 } })
  expect(words(applied.paint)).toContain('22.00'); expect(words(applied.paint)).not.toContain('North'); expect(applied.raw).toEqual(before.raw)
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(applied.options)
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('category/value, placement, format and gradient/pattern backgrounds reset independently of center settings', async ({ page }) => {
  const title = 'Sunburst node backgrounds draft'; await insertHierarchyChart(page, title, 'sunburst')
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), type = await openStatisticalTypePanel(page, editor)
  await type.getByRole('switch', { name: '显示中心标签', exact: true }).click()
  await changeChartNumber(type, '中心标签字号', '20')
  await chooseChartValue(page, type, '节点标签内容', '数值')
  await type.getByRole('combobox', { name: '节点标签内容', exact: true }).press('Escape')
  await type.getByRole('textbox', { name: '节点标签数字格式', exact: true }).fill('0.000')
  await chooseChartValue(page, type, '节点标签位置', '内部')
  await chooseChartValue(page, type, '节点标签背景填充方式', '线性渐变')
  await changeChartNumber(type, '节点标签背景渐变终点 X', '0.75')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const gradient = await statisticalChartSnapshot(page)
  expect(gradient.options.sunburst.label).toMatchObject({ contentType: 10, position: 'inside', format: '0.000', backgroundColor: { type: 'linear-gradient', end: { x: 0.75 } } })
  expect(words(gradient.paint)).toContain('13.000'); expect(words(gradient.paint)).toContain('22')
  const patternEditor = await openStatisticalSettings(page, title), patternType = await openStatisticalTypePanel(page, patternEditor)
  await chooseChartValue(page, patternType, '节点标签背景类型', '图案填充')
  await expect(patternType.getByRole('spinbutton', { name: '节点标签背景渐变终点 X', exact: true })).toHaveCount(0)
  await expect(patternType.getByRole('combobox', { name: '节点标签背景图案', exact: true })).toBeVisible()
  await applyStatisticalSettings(patternEditor)
  const pattern = await statisticalChartSnapshot(page)
  expect(pattern.options.sunburst.label.backgroundColor).toMatchObject({ type: 'pattern', preset: 'pct50' })
  expect(pattern.options.sunburst.label.backgroundColor.stops).toBeUndefined(); expect(pattern.options.sunburst.centerLabel.fontSize).toBe(20)
  const resetEditor = await openStatisticalSettings(page, title), resetType = await openStatisticalTypePanel(page, resetEditor)
  await resetType.getByRole('button', { name: '恢复旭日图节点标签默认值', exact: true }).click()
  await expect(resetType.getByRole('textbox', { name: '节点标签数字格式', exact: true })).toHaveCount(0)
  await expect(selectedChartValue(page, resetType, '节点标签位置')).toHaveText('自动')
  await applyStatisticalSettings(resetEditor)
  const defaults = await statisticalChartSnapshot(page)
  expect(defaults.options.sunburst).toEqual({ centerLabel: { visible: true, fontSize: 20 } }); expect(defaults.raw).toEqual(before.raw)
  expect(words(defaults.paint)).toContain('North'); expect(words(defaults.paint)).toContain('22')
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(pattern.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(defaults.options)
})

test('live six-method resets and signed source updates refresh the semantic center through history and internal reopen', async ({ page }) => {
  const title = 'Sunburst live source draft'; await insertHierarchyChart(page, title, 'sunburst')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), chart = officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts()[0]!
    await chart.setHierarchyFields([0, 1]); await chart.setValueField(2)
    await chart.setCenterLabel({ visible: true, fontSize: 21, format: '0.00' }); await chart.setSunburstLabel({ visible: false })
    await chart.resetSunburstLabel()
  })
  const independent = await statisticalChartSnapshot(page)
  expect(independent.options.sunburst.label).toBeUndefined(); expect(independent.options.sunburst.centerLabel.fontSize).toBe(21)
  expect(words(independent.paint)).toContain('22.00')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name); officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('C3').setValue(2)
  })
  const changed = await statisticalChartSnapshot(page)
  expect(semantic(changed.raw.hierarchy)).toEqual([
    { name: 'South', value: 9, self: 3, children: [{ name: 'Leaf A', value: 6, self: 6, children: [] }] },
    { name: 'North', value: 3, self: -5, children: [{ name: 'Leaf C', value: 6, self: 6, children: [] }, { name: 'Leaf B', value: 2, self: 2, children: [] }] },
  ])
  expect(words(changed.paint)).toContain('12.00')
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: changed.options, raw: changed.raw, layout: changed.layout, paint: changed.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => semantic((await statisticalChartSnapshot(page)).raw.hierarchy)).toEqual(semantic(independent.raw.hierarchy))
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => semantic((await statisticalChartSnapshot(page)).raw.hierarchy)).toEqual(semantic(changed.raw.hierarchy))
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name); await officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts()[0]!.resetCenterLabel()
  })
  expect((await statisticalChartSnapshot(page)).options.sunburst.centerLabel).toBeUndefined()
})

test('actual residual bitmap hover, leave and reentry change only source versions', async ({ page }) => {
  const title = 'Sunburst residual hover draft', hitbox = await insertHierarchyChart(page, title, 'sunburst'), before = await statisticalChartSnapshot(page)
  const box = await hitbox.boundingBox()
  if (!box) throw new Error('Sunburst hitbox is unavailable')
  const node = before.layout.sunburst.find((entry: { isSunburstResidual?: boolean }) => entry.isSunburstResidual)
  const angle = (node.start + node.end) / 2, radius = (node.inner + node.outer) / 2, plot = before.layout.plot
  const point = { x: box.x + (plot.x + plot.width / 2 + Math.cos(angle) * radius) / before.layout.frame.width * box.width,
    y: box.y + (plot.y + plot.height / 2 + Math.sin(angle) * radius) / before.layout.frame.height * box.height }
  const version = () => hitbox.getAttribute('data-source-version').then(value => Number(value))
  for (let attempt = 0; attempt < 2; attempt++) {
    const previous = await version(); await page.mouse.move(point.x, point.y); await expect.poll(version).toBeGreaterThan(previous)
    const active = await version(); await page.mouse.move(point.x, point.y); await expect.poll(version).toBe(active)
    expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
    await page.mouse.move(box.x - 10, box.y - 10); await expect.poll(version).toBeGreaterThan(active)
  }
  expect(await statisticalChartSnapshot(page)).toMatchObject({ contents: before.contents, raw: before.raw, layout: before.layout, paint: before.paint })
  await statisticalChartHistory(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
