import { expect, test, type Page } from '@playwright/test'
import type { WorkbookModel } from '../../packages/office-studio/src/domain/model'

type RuntimeModule = typeof import('../../examples/workbench/src/runtime')

async function state(page: Page) {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(path) as RuntimeModule
    const unitId = studio.activeWorkbookId!
    const model = JSON.parse(JSON.stringify(studio.getWorkbookModel(unitId))) as WorkbookModel
    const sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    return {
      model,
      undo: studio.editor.history.units[unitId].undo.length,
      firstOptions: sheet.getCellValidationOptions(0, 0),
      lastOptions: sheet.getCellValidationOptions(3, 0),
    }
  })
}

async function reopen(page: Page, model: WorkbookModel) {
  await page.evaluate(async snapshot => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    officeAPI.createWorkbook({ ...snapshot, id: 'quick-validation-reopened' })
  }, model)
}

test.beforeEach(async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeVisible()
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(path) as RuntimeModule
    studio.regionalSettings.setLocale('zh-CN')
    const book = officeAPI.createWorkbook({ id: 'quick-validation-book', name: '快捷验证', calculationMode: 'manual' })
    const sheet = book.getActiveSheet()!
    studio.setSelections(book.getId(), sheet.getSheetId(), [
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      { startRow: 3, endRow: 3, startColumn: 0, endColumn: 0 },
    ])
    studio.editor.history.clear(book.getId())
  })
})

for (const values of [['是', '否'], ['未开始', '进行中', '已完成'], ['A', 'B', 'C']]) {
  test(`quick dropdown preset ${values.join('/')} writes one rule to all selected rectangles`, async ({ page }) => {
    const ribbon = page.getByTestId('workbench-ribbon')
    await ribbon.getByRole('tab', { name: '插入', exact: true }).click()
    await ribbon.getByRole('button', { name: '下拉列表', exact: true }).click()
    await page.getByRole('dialog', { name: '下拉列表', exact: true }).getByRole('button', { name: values.join(' / '), exact: true }).click()
    await expect.poll(async () => (await state(page)).firstOptions).toEqual(values)
    const saved = await state(page)
    expect(saved.lastOptions).toEqual(values)
    expect(saved.undo).toBe(1)
    expect(saved.model.sheets[0].dataValidations).toHaveLength(1)
    expect(saved.model.sheets[0].dataValidations![0].ranges).toEqual([
      { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      { startRow: 3, endRow: 3, startCol: 0, endCol: 0 },
    ])
    await ribbon.getByRole('button', { name: '撤销', exact: true }).click()
    await expect.poll(async () => (await state(page)).firstOptions).toEqual([])
    await ribbon.getByRole('button', { name: '重做', exact: true }).click()
    await reopen(page, (await state(page)).model)
    expect((await state(page)).firstOptions).toEqual(values)
    expect((await state(page)).lastOptions).toEqual(values)
  })
}

test('editing a quick list keeps commas and quotes inside each option and commits one history entry', async ({ page }) => {
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(path) as RuntimeModule
    const book = officeAPI.getActiveWorkbook()!
    book.getActiveSheet()!.setDataValidation('A1', {
      type: 'list', formula1: '"Red,"Green, dark","It""s blue""', renderMode: 'custom',
    })
    studio.editor.history.clear(book.getId())
  })
  const ribbon = page.getByTestId('workbench-ribbon')
  await ribbon.getByRole('tab', { name: '数据', exact: true }).click()
  await ribbon.getByRole('button', { name: '数据验证', exact: true }).click()
  await page.getByRole('menuitem', { name: '管理验证规则', exact: true }).click()
  const panel = page.getByRole('complementary', { name: '数据验证', exact: true })
  await panel.locator('.validation-rule-main').first().click()
  const input = panel.getByRole('textbox', { name: '数据验证列表选项', exact: true })
  await expect(input).toHaveValue('Red,"Green, dark","It""s blue"')
  await input.fill('Red,"Green, dark","It""s blue",Approved')
  await panel.getByRole('button', { name: /应\s*用/ }).click()
  const expected = ['Red', 'Green, dark', 'It"s blue', 'Approved']
  await expect.poll(async () => (await state(page)).firstOptions).toEqual(expected)
  expect((await state(page)).undo).toBe(1)
  await ribbon.getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(async () => (await state(page)).firstOptions).toEqual(expected.slice(0, 3))
  await ribbon.getByRole('button', { name: '重做', exact: true }).click()
  await reopen(page, (await state(page)).model)
  expect((await state(page)).firstOptions).toEqual(expected)
})

for (const [label, showTime] of [['日期', false], ['日期和时间', true]] as const) {
  test(`quick ${label} creates the unbounded date rule over all selected rectangles`, async ({ page }) => {
    const ribbon = page.getByTestId('workbench-ribbon')
    await ribbon.getByRole('tab', { name: '插入', exact: true }).click()
    await ribbon.getByRole('button', { name: '日期选项', exact: true }).click()
    await page.getByRole('menu', { name: '日期', exact: true }).getByRole('menuitem', { name: label, exact: true }).click()
    await expect.poll(async () => (await state(page)).model.sheets[0].dataValidations?.length).toBe(1)
    const saved = await state(page)
    const rule = saved.model.sheets[0].dataValidations![0]
    expect(rule.type).toBe('date')
    expect(rule.operator).toBeUndefined()
    expect(rule.formula1).toBeUndefined()
    expect(rule.ranges).toHaveLength(2)
    expect(rule.metadata?.showTime === true).toBe(showTime)
    expect(saved.undo).toBe(1)
    await ribbon.getByRole('button', { name: '撤销', exact: true }).click()
    await expect.poll(async () => (await state(page)).model.sheets[0].dataValidations?.length ?? 0).toBe(0)
    await ribbon.getByRole('button', { name: '重做', exact: true }).click()
    await reopen(page, (await state(page)).model)
    expect((await state(page)).model.sheets[0].dataValidations![0]).toEqual(rule)
  })
}
