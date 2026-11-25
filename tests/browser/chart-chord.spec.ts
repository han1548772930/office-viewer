import { expect, test } from '@playwright/test'
import { chordHistorySavepoint, chordScreenPoint, chordSourceVersion, chordTooltip, hoverChord, insertChordChart,
  openChordSettings, setChordAppearance } from './helpers/chart-chord-settings'
import { applyStatisticalSettings, changeChartNumber, chooseChartValue, openStatisticalTypePanel,
  selectedChartValue, statisticalChartHistory, statisticalChartSnapshot } from './helpers/chart-statistical-settings'

// P28 drafts. Real workbench/Worker/Canvas, history, reopen and input assertions
// are declared for unified validation; none ran during implementation.
const labelDetails = ['节点标签位置', '节点标签颜色', '节点标签字号', '节点标签粗体', '节点标签斜体',
  '节点标签字体', '节点标签旋转角', '节点标签距离']
const words = (paint: string) => (JSON.parse(paint) as Array<{ kind: string; text?: string }>).
  filter(command => command.kind === 'text').map(command => command.text)

test('28 settings, conditional labels and repairable radii commit once through Apply/history/internal reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Chord settings draft'; await insertChordChart(page, title)
  const before = await statisticalChartSnapshot(page), history = await chordHistorySavepoint(page)
  expect(before.raw.chord.nodes.map((node: { value: number }) => node.value)).toEqual([5, 5, 5, 5])
  const editor = await openChordSettings(page, title), type = await openStatisticalTypePanel(page, editor)
  const settings = type.locator('.chart-chord-settings')
  await expect(editor.locator('.chart-chord-mapping .ant-form-item')).toHaveCount(3)
  await expect(settings.locator('.ant-form-item')).toHaveCount(25)
  await expect(type.getByRole('spinbutton', { name: '内半径比例', exact: true })).toHaveValue('0.58')
  await expect(type.getByRole('spinbutton', { name: '外半径比例', exact: true })).toHaveValue('0.78')
  await expect(type.getByRole('spinbutton', { name: '节点间隔角', exact: true })).toHaveValue('2')
  await expect(selectedChartValue(page, type, '节点标签位置')).toHaveText('外部')
  await expect(editor.getByRole('combobox', { name: '图例位置', exact: true })).toHaveCount(0)
  const innerBox = await type.getByRole('spinbutton', { name: '内半径比例', exact: true }).boundingBox()
  const outerBox = await type.getByRole('spinbutton', { name: '外半径比例', exact: true }).boundingBox()
  expect(innerBox).not.toBeNull(); expect(outerBox).not.toBeNull()
  expect(Math.abs(innerBox!.y - outerBox!.y)).toBeLessThan(2); expect(outerBox!.x).toBeGreaterThan(innerBox!.x)
  await type.getByRole('switch', { name: '显示节点标签', exact: true }).click()
  await expect(settings.locator('.ant-form-item')).toHaveCount(17)
  for (const label of labelDetails) await expect(settings.locator(`[aria-label="${label}"]`)).toHaveCount(0)
  await changeChartNumber(type, '内半径比例', '0.9')
  await expect(type.getByText(/内半径须小于外半径/)).toBeVisible()
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  expect(await chordHistorySavepoint(page)).toBe(history)
  await changeChartNumber(type, '外半径比例', '0.95')
  await expect(type.getByText(/内半径须小于外半径/)).toHaveCount(0)
  await changeChartNumber(type, '节点不透明度', '0.45')
  await changeChartNumber(type, '弦带不透明度', '0.35')
  await chooseChartValue(page, type, '悬停聚焦', '自身')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const applied = await statisticalChartSnapshot(page)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.chord).toMatchObject({ layout: { innerRadiusRatio: 0.9, outerRadiusRatio: 0.95 },
    node: { opacity: 0.45 }, ribbon: { opacity: 0.35 }, label: { visible: false }, emphasis: { focus: 'self' } })
  expect(applied.raw.chord).toEqual(before.raw.chord)
  for (const name of ['B', 'A', 'C', 'D']) expect(words(applied.paint)).not.toContain(name)
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(applied.options)
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeFocused()
  await expect(page.locator('.chart-render-error')).toHaveCount(0); expect(errors).toEqual([])
})

test('mapping excludes endpoints from weight choices, allows count and diagnoses a temporary role swap before Apply', async ({ page }) => {
  const title = 'Chord mapping draft'; await insertChordChart(page, title, true)
  const before = await statisticalChartSnapshot(page), editor = await openChordSettings(page, title)
  await editor.getByRole('combobox', { name: '和弦权重字段', exact: true }).click()
  await expect(page.getByRole('option', { name: '计数', exact: true })).toBeVisible()
  await expect(page.getByRole('option', { name: 'Weight', exact: true })).toBeVisible()
  await expect(page.getByRole('option', { name: 'Alternate', exact: true })).toBeVisible()
  await expect(page.getByRole('option', { name: 'Source', exact: true })).toHaveCount(0)
  await expect(page.getByRole('option', { name: 'Target', exact: true })).toHaveCount(0)
  await editor.getByRole('combobox', { name: '和弦权重字段', exact: true }).press('Escape')
  await chooseChartValue(page, editor, '和弦权重字段', '计数')
  await chooseChartValue(page, editor, '和弦来源字段', 'Target')
  await expect(editor.getByText('来源与目标须使用不同字段。', { exact: true })).toBeVisible()
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await chooseChartValue(page, editor, '和弦目标字段', 'Source')
  await expect(editor.getByText('来源与目标须使用不同字段。', { exact: true })).toHaveCount(0)
  await applyStatisticalSettings(editor)
  const counted = await statisticalChartSnapshot(page)
  expect(counted.options.mapping.chord).toEqual({ sourceIndex: 1, targetIndex: 0 })
  expect(counted.options.mapping.edge).toBeUndefined(); expect(counted.valueFields).toEqual([0])
  expect(counted.raw.chord.weightMode).toBe('count')
  expect(counted.raw.chord.links.map((link: { value: number }) => link.value)).toEqual([2, 2])
  expect(counted.raw.chord.links[0]).toMatchObject({ source: 'A', target: 'B', formattedValue: '2' })
  expect(counted.revision).toBe(before.revision + 1)
  expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: counted.options, raw: counted.raw, layout: counted.layout })
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).raw.chord.weightMode).toBe('value')
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).raw.chord.weightMode).toBe('count')
})

test('font/color and group resets remain independent and preserve authored theme tokens', async ({ page }) => {
  const title = 'Chord independent resets draft'; await insertChordChart(page, title, false, true)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), chart = officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts()[0]!
    await chart.setChordLayout({ padAngle: 7, clockwise: false })
    await chart.setChordLabel({ color: 'gray.700', fontSize: 18, bold: true, italic: true, family: 'Georgia', rotation: 27, distance: 15 })
    await chart.setChordRibbon({ color: 'target', opacity: 0.3, width: 2, type: 'dotted' }); await chart.setChordEmphasis({ focus: 'self' })
  })
  const before = await statisticalChartSnapshot(page), editor = await openChordSettings(page, title), type = await openStatisticalTypePanel(page, editor)
  await type.getByRole('button', { name: '恢复和弦标签字体', exact: true }).click()
  await type.getByRole('button', { name: '恢复和弦标签颜色', exact: true }).click()
  await type.getByRole('button', { name: '恢复和弦布局默认值', exact: true }).click()
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
  await applyStatisticalSettings(editor)
  const reset = await statisticalChartSnapshot(page)
  expect(reset.options.chord.node.color).toBe('blue.500'); expect(reset.options.chord.layout).toBeUndefined()
  expect(reset.options.chord.label).toEqual({ rotation: 27, distance: 15 })
  expect(reset.options.chord.ribbon).toEqual(before.options.chord.ribbon); expect(reset.options.chord.emphasis).toEqual(before.options.chord.emphasis)
  expect(reset.revision).toBe(before.revision + 1)
  await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
  await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(reset.options)
  const allEditor = await openChordSettings(page, title), allType = await openStatisticalTypePanel(page, allEditor)
  await allType.getByRole('button', { name: '恢复全部和弦默认设置', exact: true }).click()
  await applyStatisticalSettings(allEditor)
  const defaults = await statisticalChartSnapshot(page)
  expect(defaults.options.chord).toBeUndefined(); expect(defaults.options.mapping).toEqual(reset.options.mapping)
  expect(defaults.raw.chord).toEqual(reset.raw.chord); expect(words(defaults.paint)).toContain('B')
})

test('node/ribbon aggregate tooltips keep keyboard focus and history, cache hover and refresh source values', async ({ page }) => {
  const title = 'Chord content tooltip draft', hitbox = await insertChordChart(page, title)
  const before = await statisticalChartSnapshot(page), history = await chordHistorySavepoint(page)
  const grid = page.getByRole('grid', { name: '工作表画布' }); await grid.focus()
  const { point } = await hoverChord(page, hitbox)
  await expect(chordTooltip(page)).toContainText('B'); await expect(chordTooltip(page)).toContainText('Weight：5.00')
  await expect(grid).toBeFocused(); await expect(chordTooltip(page)).toHaveCSS('pointer-events', 'none')
  const version = await chordSourceVersion(hitbox)
  await page.mouse.move(point.x + 0.1, point.y); await expect.poll(() => chordSourceVersion(hitbox)).toBe(version)
  await hoverChord(page, hitbox, 'link')
  await expect(chordTooltip(page)).toContainText('B — A'); await expect(chordTooltip(page)).toContainText('Weight：5.00')
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents); expect(await chordHistorySavepoint(page)).toBe(history)
  await page.keyboard.press('Escape'); await expect(chordTooltip(page)).toBeHidden()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name); officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('C2:C5').setValues([[4], [6], [8], [2]])
  })
  await expect(hitbox).toHaveAttribute('aria-busy', 'false')
  await hoverChord(page, hitbox); await expect(chordTooltip(page)).toContainText('Weight：10.00')
  const updated = await statisticalChartSnapshot(page)
  expect(updated.revision).toBe(before.revision); expect(updated.raw.chord.links.map((link: { value: number }) => link.value)).toEqual([10, 10])
  await statisticalChartHistory(page, 'undo'); await hoverChord(page, hitbox); await expect(chordTooltip(page)).toContainText('Weight：5.00')
})

test('viewable protected Chord keeps tooltip interaction while edit commands fail without changing history', async ({ page }) => {
  const title = 'Chord read-only tooltip draft', hitbox = await insertChordChart(page, title)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, sheet = workbook.getActiveSheet()!
    const snapshot = workbook.save(), protectedSheet = snapshot.sheets.find(candidate => candidate.id === sheet.getSheetId())
    if (!protectedSheet) throw new Error('Protected Chord sheet is missing')
    // Loading another owner's viewable workbook is a valid entry path. The
    // protection command correctly refuses creating a rule for another user.
    protectedSheet.worksheetProtection = { id: 'chord-view-rule', permissionId: 'chord-view-permission', unitType: 'worksheet',
      ownerUserId: 'other-user', editState: 'onlyMe', viewState: 'othersCanView', allowedUserIds: [] }
    officeAPI.createWorkbook({ ...snapshot, id: crypto.randomUUID() })
    officeAPI.disposeUnit(workbook.getId())
  })
  await expect(hitbox).toHaveClass(/chart-view-only/)
  const before = await statisticalChartSnapshot(page), history = await chordHistorySavepoint(page)
  await hoverChord(page, hitbox); await expect(chordTooltip(page)).toContainText('Weight：5.00')
  const message = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), chart = officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts()[0]!
    try { await chart.setChordNode({ opacity: 0.2 }); return 'Unexpected success' }
    catch (error) { return error instanceof Error ? error.message : String(error) }
  })
  expect(message).toMatch(/编辑权限/)
  expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents); expect(await chordHistorySavepoint(page)).toBe(history)
  const box = await hitbox.boundingBox(); if (!box) throw new Error('Chord hitbox is unavailable')
  await page.mouse.move(box.x - 8, box.y - 8); await expect(chordTooltip(page)).toBeHidden()
})

test('theme, zoom, scroll, blur and workbench disposal clear tooltips and preserve the content model', async ({ page }) => {
  const title = 'Chord lifecycle draft', hitbox = await insertChordChart(page, title, false, true)
  const before = await statisticalChartSnapshot(page), history = await chordHistorySavepoint(page)
  await hoverChord(page, hitbox)
  const light = await chordTooltip(page).evaluate(element => getComputedStyle(element).backgroundColor)
  const version = await chordSourceVersion(hitbox)
  await setChordAppearance(page, '深色'); await expect(chordTooltip(page)).toBeHidden()
  await expect.poll(() => chordSourceVersion(hitbox)).toBeGreaterThan(version)
  await hoverChord(page, hitbox)
  const dark = await chordTooltip(page).evaluate(element => getComputedStyle(element).backgroundColor)
  expect(dark).not.toBe(light); expect((await statisticalChartSnapshot(page)).options.chord.node.color).toBe('blue.500')
  const oldWidth = Number(await hitbox.getAttribute('data-bitmap-width'))
  await page.getByRole('button', { name: '放大工作表', exact: true }).click()
  await expect(chordTooltip(page)).toBeHidden()
  await expect.poll(async () => Number(await hitbox.getAttribute('data-bitmap-width'))).toBe(Math.round(oldWidth * 1.1))
  await hoverChord(page, hitbox)
  await page.locator('.canvas-scroll').evaluate(element => { element.scrollTop = 60; element.scrollLeft = 30 })
  await expect(chordTooltip(page)).toBeHidden(); await hoverChord(page, hitbox)
  await page.evaluate(() => window.dispatchEvent(new Event('blur')))
  await expect(chordTooltip(page)).toBeHidden(); await hoverChord(page, hitbox)
  await page.keyboard.press('Escape'); await expect(chordTooltip(page)).toBeHidden()
  await hoverChord(page, hitbox); await page.getByRole('tab', { name: '文档', exact: true }).click()
  await expect(hitbox).toBeHidden(); await expect(chordTooltip(page)).toBeHidden()
  await page.getByRole('tab', { name: '表格', exact: true }).click()
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const snapshot = await statisticalChartSnapshot(page), point = await chordScreenPoint(hitbox, snapshot)
  await page.mouse.move(point.x, point.y); await expect(chordTooltip(page)).toContainText('Weight：5.00')
  expect(snapshot.contents).toBe(before.contents); expect(await chordHistorySavepoint(page)).toBe(history)
  await statisticalChartHistory(page, 'undo'); await expect(hitbox).toHaveCount(0); await expect(chordTooltip(page)).toBeHidden()
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
