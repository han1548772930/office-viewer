import { expect, test, type Page } from '@playwright/test'

// P11 workbench drafts. No server, browser, screenshot or pixel check has run.
async function insertRadar(page: Page, title: string, configured = false) {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async ({ name, configured }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    sheet.getRange('A1:B5').setValues([['Direction', 'Score'], ['North', 0], ['East', 25], ['South', 50], ['West', 100]])
    const builder = sheet.newChart('radar').setSource('A1:B5').setCategoryField(0).setValueFields([1])
      .setAbsolutePosition(100, 100).setSize(500, 340).setTitle(name).setLegend({ visible: false }).setYAxis({ min: 0, max: 100 })
    if (configured) builder.setShape('circle').setFill(true).setRadarAxisLabel({ visible: false, format: '0.000', fontSize: 18, family: 'Georgia' })
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
    const commands = await chart.getComputedRenderCommands()
    return { contents: JSON.stringify(workbook.save()), revision: chart.getInfo().revision, options: JSON.parse(JSON.stringify(chart.getOptions())), data: chart.getData(),
      layout: await chart.getComputedLayout(), commands: JSON.stringify(commands),
      labels: commands.filter((command: { kind: string }) => command.kind === 'text').map((command: { text: string; font: string }) => ({ text: command.text, font: command.font })) }
  })
}

async function openSettings(page: Page, title: string) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const editor = page.getByRole('dialog', { name: '图表设置' })
  await expect(editor).toBeVisible()
  await editor.getByText('图型专属设置', { exact: true }).click()
  const section = editor.locator('.ant-collapse-item').filter({ has: page.getByText('图型专属设置', { exact: true }) }).last()
  return { editor, section }
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

test('Radar defaults expose five labeled intervals and antd label edits preserve hidden drafts until one apply and reopen', async ({ page }) => {
  const errors: string[] = []
  page.on('pageerror', error => errors.push(error.message))
  const title = 'Radar default draft', hitbox = await insertRadar(page, title), before = await snapshot(page)
  const numeric = before.labels.filter((label: { text: string }) => ['20', '40', '60', '80', '100'].includes(label.text))
  expect(numeric).toHaveLength(20)
  expect(numeric.every((label: { font: string }) => /12px/.test(label.font))).toBe(true)
  const { editor, section } = await openSettings(page, title)
  const toggle = section.getByRole('switch', { name: '显示数值刻度标签', exact: true })
  await expect(toggle).toHaveAttribute('aria-checked', 'true')
  await expect(section.getByRole('button', { name: '显示数值刻度标签恢复默认', exact: true })).toBeDisabled()
  const font = section.getByRole('spinbutton', { name: '字号', exact: true }), format = section.getByRole('textbox', { name: '刻度数字格式', exact: true })
  await expect(font).toHaveValue('12')
  await font.fill('18'); await font.press('Tab')
  await format.fill('0.0'); await format.press('Tab')
  await toggle.click()
  await expect(section.getByRole('textbox', { name: '刻度数字格式', exact: true })).toHaveCount(0)
  await toggle.click()
  await expect(section.getByRole('textbox', { name: '刻度数字格式', exact: true })).toHaveValue('0.0')
  expect((await snapshot(page)).contents).toBe(before.contents)
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  await expect(editor).toBeHidden()
  const applied = await snapshot(page)
  expect(applied.revision).toBe(before.revision + 1)
  expect(applied.options.radar.axisLabel).toMatchObject({ visible: true, fontSize: 18, format: '0.0' })
  expect(applied.labels.some((label: { text: string; font: string }) => label.text === '20.0' && /18px/.test(label.font))).toBe(true)
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

test('Radar shape, fill and whole-label resets cancel without history and restore visible defaults in one apply', async ({ page }) => {
  const title = 'Radar reset draft', hitbox = await insertRadar(page, title, true), before = await snapshot(page)
  async function resetDraft() {
    const { editor, section } = await openSettings(page, title)
    await expect(section.getByRole('switch', { name: '显示数值刻度标签', exact: true })).toHaveAttribute('aria-checked', 'false')
    for (const name of ['雷达网格', '填充雷达区域', '显示数值刻度标签']) {
      await section.getByRole('button', { name: `${name}恢复默认`, exact: true }).click()
    }
    await expect(section.getByRole('switch', { name: '显示数值刻度标签', exact: true })).toHaveAttribute('aria-checked', 'true')
    await expect(section.getByRole('switch', { name: '填充雷达区域', exact: true })).toHaveAttribute('aria-checked', 'false')
    await expect(section.getByRole('spinbutton', { name: '字号', exact: true })).toHaveValue('12')
    await expect(section.getByRole('textbox', { name: '刻度数字格式', exact: true })).toHaveValue('')
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
  expect(reset.revision).toBe(before.revision + 1)
  expect(reset.options.radar?.shape).toBeUndefined()
  expect(reset.options.radar?.fill).toBeUndefined()
  expect(reset.options.radar?.axisLabel).toBeUndefined()
  expect(reset.labels.some((label: { text: string }) => label.text === '20')).toBe(true)
  expect(reset.data).toEqual(before.data)
  expect(await reopen(page)).toEqual({ options: reset.options, data: reset.data, commands: reset.commands })
  await history(page, 'undo')
  await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'undo')
  await expect(hitbox).toHaveCount(0)
  await expect(page.locator('.chart-render-error')).toHaveCount(0)
})
