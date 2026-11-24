import { expect, type Locator, type Page } from '@playwright/test'

// Shared navigation for statistical-type drafts. Importing this helper runs nothing.
export async function bootStatisticalWorkbench(page: Page) {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
}

export function chartSettingsPanel(page: Page, editor: Locator, header: string) {
  return editor.locator('.ant-collapse-item').filter({ has: page.getByText(header, { exact: true }) }).last()
}
export async function openStatisticalSettings(page: Page, title: string) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const editor = page.getByRole('dialog', { name: '图表设置', exact: true })
  await expect(editor).toBeVisible()
  await expect(editor.getByText('正在读取图表数据源…', { exact: true })).toHaveCount(0)
  return editor
}
export async function openStatisticalTypePanel(page: Page, editor: Locator) {
  await editor.getByText('图型专属设置', { exact: true }).click()
  return chartSettingsPanel(page, editor, '图型专属设置')
}
export async function changeChartNumber(panel: Locator, label: string, value: string) {
  const field = panel.getByRole('spinbutton', { name: label, exact: true })
  await field.fill(value); await field.press('Tab')
}
export async function chooseChartValue(page: Page, panel: Locator, label: string, value: string) {
  await panel.getByRole('combobox', { name: label, exact: true }).click()
  await page.getByRole('option', { name: value, exact: true }).click()
}
export function selectedChartValue(page: Page, panel: Locator, label: string) {
  return panel.locator('.ant-form-item').filter({ has: page.getByRole('combobox', { name: label, exact: true }) })
    .last().locator('.ant-select-selection-item')
}
export async function applyStatisticalSettings(editor: Locator) {
  await editor.getByRole('button', { name: '应用', exact: true }).click()
  const close = editor.getByRole('button', { name: '关闭', exact: true }).last()
  await expect(close).toBeEnabled(); await expect(editor).toBeVisible()
  await close.click(); await expect(editor).toBeHidden()
}

export async function statisticalChartSnapshot(page: Page, reopen = false) {
  return page.evaluate(async createCopy => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), current = officeAPI.getActiveWorkbook()!
    const workbook = createCopy ? officeAPI.createWorkbook({ ...current.save(), id: crypto.randomUUID() }, { makeCurrent: false }) : current
    const chart = workbook.getActiveSheet()!.getCharts()[0]!, info = chart.getInfo()
    const [raw, layout, commands] = await Promise.all([chart.getComputedData(), chart.getComputedLayout(), chart.getComputedRenderCommands()])
    return { contents: JSON.stringify(workbook.save()), revision: info.revision, legend: info.legend, valueFields: info.valueFields,
      options: JSON.parse(JSON.stringify(chart.getOptions())), raw: JSON.parse(JSON.stringify(raw)), layout, paint: JSON.stringify(commands) }
  }, reopen)
}
export async function statisticalChartHistory(page: Page, action: 'undo' | 'redo') {
  await page.evaluate(async operation => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    if (operation === 'undo') workbook.undo(); else workbook.redo()
  }, action)
}

export interface StatisticalDrawCommand {
  kind: string
  fill?: string
  stroke?: string
  color?: string
  width?: number
  height?: number
  x1?: number
  x2?: number
  y1?: number
  y2?: number
}
export const statisticalPaint = (paint: string): StatisticalDrawCommand[] => JSON.parse(paint)
