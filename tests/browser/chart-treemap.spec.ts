import { expect, test, type Page } from '@playwright/test'
import { applyStatisticalSettings, bootStatisticalWorkbench, changeChartNumber, chooseChartValue,
  openStatisticalSettings, openStatisticalTypePanel, selectedChartValue, statisticalChartHistory,
  statisticalChartSnapshot } from './helpers/chart-statistical-settings'

// P26 workbench drafts. None of these source/Worker/Canvas/history assertions ran.
async function insertTreemap(page: Page, title: string) {
  await bootStatisticalWorkbench(page)
  await page.evaluate(async name => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:C8').setValues([
      ['Region', 'Category', 'Weight'], ['South', 'Leaf A', 8], ['North', 'Leaf B', 12], [null, 'Leaf C', 6],
      ['South', null, 3], ['South', 'Leaf A', -2], ['South', 'Dropped', -4], ['North', null, -5],
    ])
    sheet.getRange('C2:C8').setNumberFormat('0.00')
    const builder = sheet.newChart('treemap').setSource('A1:C8').setHierarchyFields([0, 1]).setValueField(2)
      .setAbsolutePosition(110, 90).setSize(560, 420).setTitle(name).setPalette(['#112233', '#ddeeff'])
      .setParentLabelLayout('overlapping').resetParentLabelLayout()
      .setTreemapLabel({ visible: false }).resetTreemapLabel()
    await sheet.insertChart(builder.build())
  }, title)
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}

interface HierarchyNode { name: string; value: number; selfValue: number; paletteIndex?: number; children?: HierarchyNode[] }
function semantic(nodes: HierarchyNode[]): unknown[] {
  return nodes.map(node => ({ name: node.name, value: node.value, self: node.selfValue,
    children: semantic(node.children ?? []) }))
}
interface TextCommand { kind: string; text?: string }
function words(paint: string): string[] {
  return (JSON.parse(paint) as TextCommand[]).filter(command => command.kind === 'text').map(command => command.text ?? '')
}

test('signed source carries scoped names and materializes residuals, while parent and leaf label settings publish one isolated update', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Treemap signed hierarchy draft'
  await insertTreemap(page, title)
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title)
  expect(semantic(before.raw.hierarchy)).toEqual([
    { name: 'North', value: 13, self: -5, children: [
      { name: 'Leaf B', value: 12, self: 12, children: [] }, { name: 'Leaf C', value: 6, self: 6, children: [] },
    ] },
    { name: 'South', value: 9, self: 3, children: [{ name: 'Leaf A', value: 6, self: 6, children: [] }] },
  ])
  expect(before.raw.hierarchy.map((node: HierarchyNode) => node.paletteIndex)).toEqual([1, 0])
  expect(before.layout.treemap.filter((rect: { isTreemapResidual?: boolean }) => rect.isTreemapResidual)
    .map((rect: { value: number; node: HierarchyNode }) => ({ value: rect.value, owner: rect.node.name }))).toEqual([{ value: 3, owner: 'South' }])
  await expect(editor.getByRole('combobox', { name: '图例位置', exact: true })).toHaveCount(0)
  const type = await openStatisticalTypePanel(page, editor)
  await expect(selectedChartValue(page, type, '父级标签布局')).toHaveText('独立标题带')
  await expect(type.getByRole('switch', { name: '显示叶级标签', exact: true })).toHaveAttribute('aria-checked', 'true')
  await expect(type.getByRole('textbox', { name: '叶级标签数字格式', exact: true })).toHaveCount(0)
  await type.getByRole('switch', { name: '显示叶级标签', exact: true }).click()
  await expect(type.getByRole('combobox', { name: '叶级标签内容', exact: true })).toHaveCount(0)
  await expect(type.getByRole('combobox', { name: '父级标签布局', exact: true })).toBeVisible()
  await chooseChartValue(page, type, '父级标签布局', '叠放')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const applied = await statisticalChartSnapshot(page), labels = words(applied.paint)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.treemap).toMatchObject({ parentLabelLayout: 'overlapping', label: { visible: false } })
  expect(labels).toContain('North'); expect(labels).toContain('South')
  expect(labels).not.toContain('Leaf A'); expect(applied.raw).toEqual(before.raw)
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(applied.options)
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('leaf category/value, numeric format, gradient and pattern backgrounds reset independently of the parent layout', async ({ page }) => {
  const title = 'Treemap labels background draft'
  await insertTreemap(page, title)
  const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), type = await openStatisticalTypePanel(page, editor)
  await chooseChartValue(page, type, '父级标签布局', '隐藏')
  await chooseChartValue(page, type, '叶级标签内容', '数值')
  await type.getByRole('combobox', { name: '叶级标签内容', exact: true }).press('Escape')
  await type.getByRole('textbox', { name: '叶级标签数字格式', exact: true }).fill('0.000')
  await chooseChartValue(page, type, '叶级标签背景填充方式', '线性渐变')
  await changeChartNumber(type, '叶级标签背景渐变终点 X', '0.75')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const gradient = await statisticalChartSnapshot(page)
  expect(gradient.options.treemap.label).toMatchObject({ contentType: 10, format: '0.000', backgroundColor: { type: 'linear-gradient', end: { x: 0.75 } } })
  expect(words(gradient.paint)).toContain('Leaf A'); expect(words(gradient.paint)).toContain('6.000')
  expect(words(gradient.paint)).not.toContain('North'); expect(words(gradient.paint)).not.toContain('South')
  const patternEditor = await openStatisticalSettings(page, title), patternType = await openStatisticalTypePanel(page, patternEditor)
  await chooseChartValue(page, patternType, '叶级标签背景类型', '图案填充')
  await expect(patternType.getByRole('spinbutton', { name: '叶级标签背景渐变终点 X', exact: true })).toHaveCount(0)
  await expect(patternType.getByRole('combobox', { name: '标签背景图案', exact: true })).toBeVisible()
  await applyStatisticalSettings(patternEditor)
  const pattern = await statisticalChartSnapshot(page)
  expect(pattern.options.treemap.label.backgroundColor).toMatchObject({ type: 'pattern', preset: 'pct50' })
  expect(pattern.options.treemap.label.backgroundColor.stops).toBeUndefined()
  expect(pattern.options.treemap.parentLabelLayout).toBe('none')
  const resetEditor = await openStatisticalSettings(page, title), resetType = await openStatisticalTypePanel(page, resetEditor)
  await resetType.getByRole('button', { name: '恢复树图叶级标签默认值', exact: true }).click()
  await expect(resetType.getByRole('textbox', { name: '叶级标签数字格式', exact: true })).toHaveCount(0)
  await applyStatisticalSettings(resetEditor)
  const defaults = await statisticalChartSnapshot(page)
  expect(defaults.options.treemap).toEqual({ parentLabelLayout: 'none' }); expect(defaults.raw).toEqual(before.raw)
  expect(words(defaults.paint)).toContain('Leaf A'); expect(words(defaults.paint)).not.toContain('North')
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: defaults.options, raw: defaults.raw, layout: defaults.layout, paint: defaults.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(pattern.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(defaults.options)
})

test('live Facade resets preserve independent options and a signed source edit invalidates the tree through history and internal reopen', async ({ page }) => {
  const title = 'Treemap live source draft'
  await insertTreemap(page, title)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), chart = officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts()[0]!
    await chart.setHierarchyFields([0, 1]); await chart.setValueField(2)
    await chart.setTreemapLabel({ visible: false, format: '0.00' }); await chart.setParentLabelLayout('none')
    await chart.resetParentLabelLayout()
  })
  const independent = await statisticalChartSnapshot(page)
  expect(independent.options.treemap.label).toEqual({ visible: false, format: '0.00' })
  expect(independent.options.treemap.parentLabelLayout).toBeUndefined(); expect(words(independent.paint)).toContain('North')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    await sheet.getCharts()[0]!.resetTreemapLabel(); sheet.getRange('C3').setValue(2)
  })
  const changed = await statisticalChartSnapshot(page)
  expect(changed.options.treemap).toEqual({})
  expect(semantic(changed.raw.hierarchy)).toEqual([
    { name: 'South', value: 9, self: 3, children: [{ name: 'Leaf A', value: 6, self: 6, children: [] }] },
    { name: 'North', value: 3, self: -5, children: [
      { name: 'Leaf C', value: 6, self: 6, children: [] }, { name: 'Leaf B', value: 2, self: 2, children: [] },
    ] },
  ])
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: changed.options, raw: changed.raw, layout: changed.layout, paint: changed.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => semantic((await statisticalChartSnapshot(page)).raw.hierarchy)).toEqual(semantic(independent.raw.hierarchy))
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => semantic((await statisticalChartSnapshot(page)).raw.hierarchy)).toEqual(semantic(changed.raw.hierarchy))
})

test('actual bitmap hover, leave and reentry change source versions without changing the hierarchy or content history', async ({ page }) => {
  const title = 'Treemap bitmap hover draft', hitbox = await insertTreemap(page, title), before = await statisticalChartSnapshot(page)
  const box = await hitbox.boundingBox()
  if (!box) throw new Error('Treemap hitbox is unavailable')
  const point = { x: box.x + box.width * 0.52, y: box.y + box.height * 0.58 }
  const version = () => hitbox.getAttribute('data-source-version').then(value => Number(value))
  for (let attempt = 0; attempt < 2; attempt++) {
    const previous = await version()
    await page.mouse.move(point.x, point.y); await expect.poll(version).toBeGreaterThan(previous)
    const active = await version()
    await page.mouse.move(point.x, point.y); await expect.poll(version).toBe(active)
    expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
    await page.mouse.move(box.x - 10, box.y - 10); await expect.poll(version).toBeGreaterThan(active)
  }
  expect(await statisticalChartSnapshot(page)).toMatchObject({ contents: before.contents, raw: before.raw, layout: before.layout, paint: before.paint })
  await statisticalChartHistory(page, 'undo'); await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
