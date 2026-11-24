import { expect, test } from '@playwright/test'

// P15: authored controls are tested through the workbench only during future unified validation.
test('gradient reset and manual plot controls remain drafts until one undoable apply and survive reopen', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Category', 'Units'], ['A', 10], ['B', 20], ['C', 30]])
    await sheet.insertChart(sheet.newChart('area').setSource('A1:B4').setValueFields([1]).setAbsolutePosition(100, 100)
      .setSize(500, 320).setTitle('Gradient plot draft').build())
  })
  const hitbox = page.getByRole('button', { name: 'Gradient plot draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), authored: chart.getOptions().autoGradientFill, gradient: chart.getDescription().autoGradientFill,
      plotArea: chart.getOptions().plotArea, plot: chart.getLayout().plot, revision: chart.getInfo().revision }
  })
  await hitbox.click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  const before = await snapshot(), version = Number(await hitbox.getAttribute('data-source-version'))
  await editor.getByText('图型专属设置', { exact: true }).click()
  const gradient = editor.getByRole('switch', { name: '自动渐变填充', exact: true }), reset = editor.getByRole('button', { name: '自动渐变填充恢复默认', exact: true })
  await expect(gradient).toHaveAttribute('aria-checked', 'true')
  await expect(reset).toBeDisabled()
  await gradient.click()
  await expect(gradient).toHaveAttribute('aria-checked', 'false')
  await expect(reset).toBeEnabled()
  await reset.click()
  await expect(gradient).toHaveAttribute('aria-checked', 'true')
  await expect(reset).toBeDisabled()
  expect((await snapshot()).contents).toBe(before.contents)
  await gradient.click()
  await editor.getByText('绘图区', { exact: true }).click()
  await editor.getByRole('switch', { name: '手动设置绘图区', exact: true }).click()
  for (const [name, value] of [['左侧位置 (%)', '15'], ['顶部位置 (%)', '25'], ['宽度 (%)', '60'], ['高度 (%)', '50']] as const) {
    const input = editor.getByRole('spinbutton', { name, exact: true })
    await input.fill(value)
    await input.press('Tab')
  }
  expect((await snapshot()).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect.poll(async () => (await snapshot()).plotArea).toEqual({ left: 0.15, top: 0.25, width: 0.6, height: 0.5 })
  await expect.poll(async () => (await snapshot()).gradient).toBe(false)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(version)
  const applied = await snapshot()
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.plot.width).toBe(300)
  expect(applied.plot.height).toBe(160)
  const reopened = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const saved = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    return { options: saved.getOptions(), gradient: saved.getDescription().autoGradientFill, plot: saved.getLayout().plot }
  })
  expect(reopened.options.plotArea).toEqual(applied.plotArea)
  expect(reopened.gradient).toBe(false)
  expect(reopened.plot).toEqual(applied.plot)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).plotArea).toBeUndefined()
  await expect.poll(async () => (await snapshot()).gradient).toBe(true)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.redo()
  })
  await expect.poll(async () => (await snapshot()).plotArea).toEqual(applied.plotArea)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('an invalid manual rectangle reports its constraint and publishes no chart content or history', async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Category', 'Units'], ['A', 10], ['B', 20], ['C', 30]])
    await sheet.insertChart(sheet.newChart('column').setSource('A1:B4').setValueFields([1]).setTitle('Invalid plot draft').build())
  })
  const hitbox = page.getByRole('button', { name: 'Invalid plot draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  await hitbox.click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), revision: chart.getInfo().revision, plotArea: chart.getOptions().plotArea }
  })
  const before = await snapshot()
  await editor.getByText('绘图区', { exact: true }).click()
  await editor.getByRole('switch', { name: '手动设置绘图区', exact: true }).click()
  for (const [name, value] of [['左侧位置 (%)', '95'], ['宽度 (%)', '90']] as const) {
    const input = editor.getByRole('spinbutton', { name, exact: true })
    await input.fill(value)
    await input.press('Tab')
  }
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor.getByText('绘图区须完整位于图表范围内。', { exact: true }).first()).toBeVisible()
  expect(await snapshot()).toEqual(before)
  for (const [name, value] of [['左侧位置 (%)', '20'], ['顶部位置 (%)', '20'], ['宽度 (%)', '60'], ['高度 (%)', '60']] as const) {
    const input = editor.getByRole('spinbutton', { name, exact: true })
    await input.fill(value)
    await input.press('Tab')
  }
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect.poll(async () => (await snapshot()).plotArea).toEqual({ left: 0.2, top: 0.2, width: 0.6, height: 0.6 })
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).plotArea).toBeUndefined()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect(hitbox).toHaveCount(0)
})

// P12/P14/P15 drafts for unified validation. These new cases have not been run in the implementation phase.
test('hover at worksheet zoom/scroll changes only the raster and pointer cancellation discards a drag draft', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Date', 'Units'], [59, 10], [60, 20], [63, 30]])
    await sheet.insertChart(sheet.newChart('line').setSource('A1:B4').setValueFields([1]).setUseDateAxis(true)
      .setAbsolutePosition(120, 100).setSize(420, 280).setTitle('Pointer draft')
      .setAxisPointer({ indicatorLineColor: 'red', indicatorLineType: 'dashed', indicatorLabelColor: 'navy', indicatorLabelTextColor: 'white' }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Pointer draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const slider = page.getByRole('slider', { name: '工作表缩放' })
  await slider.press('Home')
  for (let index = 0; index < 14; index++) await slider.press('ArrowRight')
  await expect(slider).toHaveAttribute('aria-valuenow', '150')
  await page.locator('.canvas-scroll').evaluate(element => { element.scrollLeft = 60; element.scrollTop = 40 })
  await expect.poll(() => hitbox.evaluate(element => Math.round(element.getBoundingClientRect().width))).toBe(630)
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    const info = chart.getInfo(), plot = chart.getLayout().plot
    return { contents: JSON.stringify(workbook.save()), revision: info.revision, position: info.position,
      plot: { x: plot.x - info.position.x, y: plot.y - info.position.y, width: plot.width, height: plot.height } }
  })
  const before = await snapshot(), bounds = (await hitbox.boundingBox())!, version = Number(await hitbox.getAttribute('data-source-version'))
  const pointer = { x: bounds.x + (before.plot.x + before.plot.width / 3) * 1.5,
    y: bounds.y + (before.plot.y + before.plot.height / 2) * 1.5 }
  await page.mouse.move(pointer.x, pointer.y)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(version)
  expect((await snapshot()).contents).toBe(before.contents)
  expect((await snapshot()).revision).toBe(before.revision)
  const hoverVersion = Number(await hitbox.getAttribute('data-source-version'))
  await page.mouse.move(10, 10)
  await expect.poll(async () => Number(await hitbox.getAttribute('data-source-version'))).toBeGreaterThan(hoverVersion)
  await page.mouse.move(pointer.x, pointer.y)
  await page.mouse.down()
  await page.mouse.move(pointer.x + 60, pointer.y + 30)
  await page.evaluate(() => window.dispatchEvent(new PointerEvent('pointercancel')))
  await page.mouse.up()
  await expect.poll(async () => (await snapshot()).position).toEqual(before.position)
  expect((await snapshot()).contents).toBe(before.contents)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})

test('pattern controls edit a detached draft, apply a real solid override, and preserve history and internal reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Category', 'Units'], ['A', 10], ['B', 20], ['C', 30]])
    await sheet.insertChart(sheet.newChart('column').setSource('A1:B4').setValueFields([1]).setAbsolutePosition(100, 100).setSize(420, 280)
      .setTitle('Pattern draft').setAllSeriesStyle({ fillOpacity: 1, fillPattern: { preset: 'horz', foregroundColor: '#ff0000', backgroundColor: '#ffff00' } }).build())
  })
  const hitbox = page.getByRole('button', { name: 'Pattern draft', exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  const snapshot = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    return { contents: JSON.stringify(workbook.save()), pattern: chart.getOptions().allSeriesStyle?.fillPattern,
      red: chart.getRenderCommands().filter(command => command.kind === 'polygon' && command.fill === 'rgba(255, 0, 0, 1)').length }
  })
  const before = await snapshot()
  expect(before.red).toBeGreaterThan(0)
  const openPatternPanel = async () => {
    await hitbox.click()
    const editor = page.getByRole('dialog', { name: '图表设置' })
    await expect(editor).toBeVisible()
    await editor.getByText('系列、标记与数据标签', { exact: true }).click()
    await editor.getByText('所有系列', { exact: true }).click()
    return editor
  }
  let editor = await openPatternPanel()
  await editor.getByLabel('填充纹理', { exact: true }).click()
  await page.getByText('向上对角线', { exact: true }).last().click()
  await expect(editor.getByLabel('纹理前景色', { exact: true })).toBeVisible()
  expect((await snapshot()).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot()).contents).toBe(before.contents)
  editor = await openPatternPanel()
  await editor.getByLabel('填充纹理', { exact: true }).click()
  await page.getByText('无纹理', { exact: true }).last().click()
  await expect(editor.getByLabel('纹理前景色', { exact: true })).toHaveCount(0)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect.poll(async () => (await snapshot()).pattern).toBeNull()
  await expect.poll(async () => (await snapshot()).red).toBe(0)
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await snapshot()).pattern).toEqual(before.pattern)
  await expect.poll(async () => (await snapshot()).red).toBeGreaterThan(0)
  const reopened = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const saved = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false })
    return saved.getActiveSheet()!.getCharts()[0]!.getOptions().allSeriesStyle?.fillPattern
  })
  expect(reopened).toEqual(before.pattern)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('inserts a radar chart from the worksheet and keeps its type across a saved snapshot', async ({ page }) => {
  await page.goto('/')
  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.getActiveSheet()!.getRange('A1:B4').setValues([
      ['Category', 'Units'], ['A', 6], ['B', 3], ['C', 9],
    ])
  })
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '图表', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入图表' })
  await dialog.getByPlaceholder('A1:C4').fill('A1:B4')
  await dialog.locator('.chart-insert-form .ant-select').first().click()
  await page.getByText('雷达图', { exact: true }).last().click()
  await dialog.locator('.ant-modal-footer button').last().click()
  await expect(page.locator('.chart-hitbox')).toHaveAttribute('data-rendered', 'true')

  const saved = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    const workbook = officeAPI.getActiveWorkbook()!
    const reopened = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false })
    const chart = reopened.getActiveSheet()!.getCharts()[0]!
    return { type: chart.getInfo().type, data: chart.getData() }
  })
  expect(saved).toMatchObject({
    type: 'radar',
    data: { categories: ['A', 'B', 'C'], series: [{ name: 'Units', values: [6, 3, 9] }] },
  })
})

test('inserts, moves, resizes and refreshes an editable chart on the worksheet canvas', async ({ page }) => {
  const pageErrors: string[] = []
  page.on('pageerror', error => pageErrors.push(error.message))
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')

  const grid = page.getByRole('grid', { name: '工作表画布' })
  await expect(grid).toBeVisible()
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)

  const nameBox = page.getByRole('textbox', { name: '名称框', exact: true })
  await nameBox.fill('A1:C4')
  await nameBox.press('Enter')
  await grid.press('Control+V')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))
    if (!resource) throw new Error('Workbench runtime module was not loaded')
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook().getActiveSheet().getRange('A1:C4').setValues([
      ['项目', '一月', '二月'], ['甲', 100, 150], ['乙', 200, 250], ['丙', 300, 350],
    ])
  })
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '图表', exact: true }).click()
  await page.getByPlaceholder('A1:C4').fill('A1:C4')
  await page.locator('.ant-modal-footer button').last().click()
  await expect(page.locator('.ant-modal-wrap')).toBeHidden()

  const chart = page.locator('.chart-hitbox')
  await expect(chart).toHaveCount(1)
  await expect(chart).toHaveRole('button')
  await expect(chart).toHaveAttribute('data-rendered', 'true')
  await expect(chart).toHaveAttribute('data-renderer', 'echarts-canvas')
  const bitmapSize = await chart.evaluate(element => ({ width: Number(element.getAttribute('data-bitmap-width')), height: Number(element.getAttribute('data-bitmap-height')) }))
  expect(bitmapSize.width).toBeGreaterThan(0)
  expect(bitmapSize.height).toBeGreaterThan(0)
  const initialSourceVersion = Number(await chart.getAttribute('data-source-version'))
  const before = await chart.evaluate(element => (element as HTMLElement).style.left)
  const bounds = await chart.boundingBox()
  expect(bounds).not.toBeNull()
  const hitTest = await page.evaluate(({ x, y }) => {
    const target = document.elementFromPoint(x, y)
    return { target: target?.outerHTML.slice(0, 180), chart: target?.closest('.chart-hitbox') !== null }
  }, { x: bounds!.x + 60, y: bounds!.y + 40 })
  expect(hitTest.chart, hitTest.target).toBe(true)
  await page.mouse.move(bounds!.x + 60, bounds!.y + 40)
  await page.mouse.down()
  await page.mouse.move(bounds!.x + 130, bounds!.y + 82, { steps: 4 })
  await page.mouse.up()
  await expect.poll(() => chart.evaluate(element => (element as HTMLElement).style.left)).not.toBe(before)
  await expect(chart).toHaveAttribute('aria-pressed', 'true')

  const handle = chart.getByLabel('调整图表大小')
  const resize = await handle.boundingBox()
  expect(resize).not.toBeNull()
  const widthBefore = await chart.evaluate((element) => (element as HTMLElement).style.width)
  await page.mouse.move(resize!.x + resize!.width / 2, resize!.y + resize!.height / 2)
  await page.mouse.down()
  await page.mouse.move(resize!.x + 80, resize!.y + 40, { steps: 4 })
  await page.mouse.up()
  await expect.poll(() => chart.evaluate((element) => (element as HTMLElement).style.width)).not.toBe(widthBefore)

  await nameBox.fill('B3')
  await nameBox.press('Enter')
  const formula = page.locator('.formula-input')
  await formula.fill('1000')
  await formula.press('Enter')
  await expect.poll(() => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))
    if (!resource) throw new Error('Workbench runtime module was not loaded')
    const { officeAPI } = await import(resource.name)
    return officeAPI.getActiveWorkbook().getActiveSheet().getCharts()[0].getData().series[0].values[1]
  })).toBe(1000)
  await expect.poll(async () => Number(await chart.getAttribute('data-source-version'))).toBeGreaterThan(initialSourceVersion)

  await chart.click()
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  await editor.getByRole('textbox', { name: '图表标题' }).fill('Quarterly revenue')
  await editor.getByRole('textbox', { name: '图表数据区域' }).fill('A1:B4')
  await editor.getByText('坐标轴', { exact: true }).click()
  await editor.getByText('X 轴 / 分类轴', { exact: true }).click()
  const xPanel = editor.locator('.ant-collapse-item').filter({ has: editor.getByText('X 轴 / 分类轴', { exact: true }) }).last()
  await xPanel.getByRole('textbox', { name: '轴标题', exact: true }).fill('Product')
  await editor.getByText('Y 轴 / 数值轴', { exact: true }).click()
  const yPanel = editor.locator('.ant-collapse-item').filter({ has: editor.getByText('Y 轴 / 数值轴', { exact: true }) }).last()
  await yPanel.getByRole('textbox', { name: '轴标题', exact: true }).fill('Revenue')
  await editor.getByRole('switch', { name: '显示图例' }).click()
  await editor.locator('.chart-editor-actions button').last().click()
  await expect.poll(() => chart.getAttribute('aria-label')).toBe('Quarterly revenue')

  const chartSettings = async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))
    if (!resource) throw new Error('Workbench runtime module was not loaded')
    const { officeAPI } = await import(resource.name)
    const workbook = officeAPI.getActiveWorkbook()!
    return workbook.getActiveSheet()!.getCharts()[0]!.getInfo()
  })
  expect(await chartSettings()).toMatchObject({
    title: 'Quarterly revenue',
    source: { range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
    legend: { visible: false },
    xAxis: { title: 'Product' },
    yAxis: { title: 'Revenue' },
  })
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.undo()
  })
  await expect.poll(async () => (await chartSettings()).title).toBe('')
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    officeAPI.getActiveWorkbook()!.redo()
  })
  await expect.poll(async () => (await chartSettings()).title).toBe('Quarterly revenue')
  const reopenedChart = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find((entry) => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    const workbook = officeAPI.getActiveWorkbook()!
    const reopened = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false })
    return reopened.getActiveSheet()!.getCharts()[0]!.getInfo()
  })
  expect(reopenedChart).toMatchObject({ title: 'Quarterly revenue', legend: { visible: false }, xAxis: { title: 'Product' }, yAxis: { title: 'Revenue' } })
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)
  expect(pageErrors).toEqual([])
})
