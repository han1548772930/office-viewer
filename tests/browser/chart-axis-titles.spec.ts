import { expect, test, type Page } from '@playwright/test'

// P12 workbench drafts only. No server, browser, pixel or Office check has run.
async function insertChart(page: Page, title: string, configured = false) {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async ({ name, configured }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B4').setValues([['Month', 'Units'], ['January', 10], ['February', 20], ['March', 40]])
    const builder = sheet.newChart('line').setSource('A1:B4').setCategoryField(0).setValueFields([1])
      .setAbsolutePosition(100, 100).setSize(540, 400).setTitle(name).setLegend({ visible: false })
      .setOption('appearance', { font: { family: 'Georgia', fontSize: 18, color: '#334155', bold: true, italic: false } })
      .setXAxis({ title: 'Physical months', titleAlignment: 'start', reverse: false,
        line: { color: '#2563eb', width: 2 }, tick: { length: 6 } })
      .setYAxis({ title: 'Physical units', titleAlignment: 'start', reverse: false, min: 0, max: 100 })
    if (configured) builder.setXAxis({ titleAlignment: 'end', reverse: true,
      titleStyle: { family: 'Aptos', fontSize: 24, color: '#be123c', bold: false, italic: true } })
    await sheet.insertChart(builder.build())
  }, { name: title, configured })
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}

async function snapshot(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, chart = workbook.getActiveSheet()!.getCharts()[0]!
    const info = chart.getInfo(), commands = await chart.getComputedRenderCommands()
    return { contents: JSON.stringify(workbook.save()), revision: info.revision,
      axes: JSON.parse(JSON.stringify({ x: info.xAxis, y: info.yAxis })), options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(),
      layout: await chart.getComputedLayout(), commands: JSON.stringify(commands),
      titles: commands.filter((command: { kind: string; text?: string }) => command.kind === 'text'
        && ['Physical months', 'Physical units'].includes(command.text ?? '')) }
  })
}

async function openAxes(page: Page, title: string, both = false) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  await editor.getByText('坐标轴', { exact: true }).click()
  await editor.getByText('X 轴 / 分类轴', { exact: true }).click()
  const x = editor.locator('.ant-collapse-item').filter({ has: page.getByText('X 轴 / 分类轴', { exact: true }) }).last()
  if (both) await editor.getByText('Y 轴 / 数值轴', { exact: true }).click()
  const y = editor.locator('.ant-collapse-item').filter({ has: page.getByText('Y 轴 / 数值轴', { exact: true }) }).last()
  return { editor, x, y }
}

async function history(page: Page, operation: 'undo' | 'redo') {
  await page.evaluate(async action => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    if (action === 'undo') workbook.undo()
    else workbook.redo()
  }, operation)
}

async function reopen(page: Page) {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    const chart = officeAPI.createWorkbook({ ...workbook.save(), id: crypto.randomUUID() }, { makeCurrent: false }).getActiveSheet()!.getCharts()[0]!
    const info = chart.getInfo()
    return { axes: JSON.parse(JSON.stringify({ x: info.xAxis, y: info.yAxis })), options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(), commands: JSON.stringify(await chart.getComputedRenderCommands()) }
  })
}

test('antd axis title alignment, inherited font and reverse edits apply once against physical edges and survive history and reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  const title = 'Physical axis titles draft', hitbox = await insertChart(page, title), before = await snapshot(page)
  expect(before.titles).toHaveLength(2)
  expect(before.titles.every((command: { font: string }) => /18px.*Georgia/.test(command.font))).toBe(true)
  const { editor, x, y } = await openAxes(page, title, true)
  // Each axis places the title font before the separate axis-label font controls.
  const titleSize = x.getByRole('spinbutton', { name: '字号', exact: true }).first()
  await expect(titleSize).toHaveValue('18')
  await expect(x.getByRole('button', { name: '字号恢复默认', exact: true })).toBeDisabled()
  await titleSize.fill('22'); await titleSize.press('Tab')
  await x.getByRole('combobox', { name: '轴标题对齐', exact: true }).click()
  await page.getByRole('option', { name: '终点', exact: true }).click()
  await x.getByRole('switch', { name: '反向轴', exact: true }).click()
  await y.getByRole('switch', { name: '反向轴', exact: true }).click()
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const applied = await snapshot(page), horizontal = applied.titles.find((command: { text: string }) => command.text === 'Physical months')!,
    vertical = applied.titles.find((command: { text: string }) => command.text === 'Physical units')!
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.axes.x).toMatchObject({ titleAlignment: 'end', reverse: true, titleStyle: { fontSize: 22 } })
  expect(applied.axes.y.reverse).toBe(true)
  // Computed layout/commands both use worksheet coordinates, including frame position.
  expect(horizontal.x).toBe(applied.layout.plot.x + applied.layout.plot.width)
  expect(horizontal.align).toBe('right')
  expect(horizontal.font).toMatch(/22px.*Georgia/)
  expect(vertical.y).toBe(applied.layout.plot.y)
  expect(vertical.align).toBe('right')
  expect(vertical.x).toBe(before.titles.find((command: { text: string }) => command.text === 'Physical units')!.x)
  expect(applied.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ axes: applied.axes, options: applied.options, data: applied.data, commands: applied.commands })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).axes).toEqual(before.axes)
  await history(page, 'redo')
  await expect.poll(async () => (await snapshot(page)).axes).toEqual(applied.axes)
  await history(page, 'undo')
  await history(page, 'undo')
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
  expect(errors).toEqual([])
})

test('axis title font and alignment reset independently, and whole-title reset preserves unrelated fields through cancel, apply and reopen', async ({ page }) => {
  const title = 'Axis title resets draft', hitbox = await insertChart(page, title, true), before = await snapshot(page)
  async function resetFields() {
    const { editor, x } = await openAxes(page, title)
    for (const name of ['字体', '字号', '文字颜色', '粗体', '斜体', '轴标题对齐']) {
      await x.getByRole('button', { name: `${name}恢复默认`, exact: true }).click()
    }
    await expect(x.getByRole('textbox', { name: '轴标题', exact: true })).toHaveValue('Physical months')
    await expect(x.getByRole('textbox', { name: '字体', exact: true }).first()).toHaveValue('Georgia')
    await expect(x.getByRole('spinbutton', { name: '字号', exact: true }).first()).toHaveValue('18')
    await expect(x.getByRole('switch', { name: '粗体', exact: true }).first()).toHaveAttribute('aria-checked', 'true')
    await expect(x.getByRole('switch', { name: '斜体', exact: true }).first()).toHaveAttribute('aria-checked', 'false')
    await expect(x.getByRole('switch', { name: '反向轴', exact: true })).toHaveAttribute('aria-checked', 'true')
    expect((await snapshot(page)).contents).toBe(before.contents)
    return editor
  }
  let editor = await resetFields()
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(before.contents)
  editor = await resetFields()
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const fieldsReset = await snapshot(page)
  expect(fieldsReset.revision).toBe(before.revision + 1)
  expect(fieldsReset.axes.x.title).toBe('Physical months')
  expect(fieldsReset.axes.x.titleAlignment).toBeUndefined()
  for (const field of ['family', 'fontSize', 'color', 'bold', 'italic'] as const) expect(fieldsReset.axes.x.titleStyle?.[field]).toBeUndefined()
  expect(fieldsReset.axes.x).toMatchObject({ reverse: true, line: before.axes.x.line, tick: before.axes.x.tick })
  expect(fieldsReset.axes.y).toEqual(before.axes.y)
  expect(await reopen(page)).toEqual({ axes: fieldsReset.axes, options: fieldsReset.options, data: fieldsReset.data, commands: fieldsReset.commands })
  async function resetGroup() {
    const { editor, x } = await openAxes(page, title)
    await x.getByRole('button', { name: '轴标题恢复默认', exact: true }).click()
    await expect(x.getByRole('textbox', { name: '轴标题', exact: true })).toHaveValue('')
    await expect(x.getByRole('button', { name: '轴标题恢复默认', exact: true })).toBeDisabled()
    await expect(x.getByRole('switch', { name: '反向轴', exact: true })).toHaveAttribute('aria-checked', 'true')
    expect((await snapshot(page)).contents).toBe(fieldsReset.contents)
    return editor
  }
  editor = await resetGroup()
  await editor.getByRole('button', { name: '关闭', exact: true }).last().click()
  await expect(editor).toBeHidden()
  expect((await snapshot(page)).contents).toBe(fieldsReset.contents)
  editor = await resetGroup()
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const groupReset = await snapshot(page)
  expect(groupReset.revision).toBe(fieldsReset.revision + 1)
  expect(groupReset.axes.x.title).toBeUndefined()
  expect(groupReset.axes.x.titleStyle).toBeUndefined()
  expect(groupReset.axes.x.titleAlignment).toBeUndefined()
  expect(groupReset.axes.x).toMatchObject({ reverse: true, line: before.axes.x.line, tick: before.axes.x.tick })
  expect(groupReset.axes.y).toEqual(before.axes.y)
  expect(groupReset.titles.map((command: { text: string }) => command.text)).toEqual(['Physical units'])
  expect(groupReset.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ axes: groupReset.axes, options: groupReset.options, data: groupReset.data, commands: groupReset.commands })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).axes).toEqual(fieldsReset.axes)
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).axes).toEqual(before.axes)
  await history(page, 'undo')
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
