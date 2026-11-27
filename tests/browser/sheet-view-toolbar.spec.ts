import { expect, test, type Page } from '@playwright/test'
import type { WorkbookModel } from '../../packages/office-studio/src/domain/model'

type Runtime = typeof import('../../examples/workbench/src/runtime')

async function capture(page: Page) {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(path) as Runtime
    const unitId = studio.activeWorkbookId
    const history = studio.editor.history.units[unitId]!
    return {
      model: JSON.parse(JSON.stringify(studio.getWorkbookModel(unitId))) as WorkbookModel,
      revision: studio.getWorkbookRevision(unitId), undo: history.undo.length, redo: history.redo.length,
      dirty: studio.editor.history.isDirty(unitId),
    }
  })
}

test.beforeEach(async ({ page }) => {
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeVisible()
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as Runtime
    const book = officeAPI.createWorkbook({ id: 'sheet-view-operations', name: '视图操作' })
    book.getActiveSheet()!.getRange('A1:C3').setValues([['标题', '金额', '备注'], ['东', 20, '保留'], ['西', 30, '原文']])
  })
  const name = page.getByRole('textbox', { name: '名称框', exact: true })
  await name.fill('B2')
  await name.press('Enter')
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '视图', exact: true }).click()
})

test('crosshair toggle and color selection preserve workbook content, history and focus', async ({ page }) => {
  const ribbon = page.getByTestId('workbench-ribbon')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  const toggle = ribbon.getByRole('button', { name: '十字高亮', exact: true })
  const options = ribbon.getByRole('button', { name: '十字高亮选项', exact: true })
  const before = await capture(page)
  await expect(toggle).toHaveAttribute('aria-pressed', 'false')
  await toggle.click()
  await expect(toggle).toHaveAttribute('aria-pressed', 'true')
  await expect(grid).toBeFocused()
  expect(await capture(page)).toEqual(before)
  await toggle.click()
  await expect(toggle).toHaveAttribute('aria-pressed', 'false')
  await options.click()
  await page.getByRole('menu', { name: '十字高亮', exact: true }).getByRole('menuitem', { name: '紫色', exact: true }).click()
  await expect(toggle).toHaveAttribute('aria-pressed', 'true')
  await expect(grid).toBeFocused()
  expect(await capture(page)).toEqual(before)
  await options.click()
  await expect(page.getByRole('menu', { name: '十字高亮', exact: true }).getByRole('menuitem', { name: '紫色', exact: true })).toHaveClass(/ant-menu-item-selected/)
})

test('crosshair operations keep the cell input draft and return focus without committing it', async ({ page }) => {
  const ribbon = page.getByTestId('workbench-ribbon')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  const before = await capture(page)
  await grid.press('F2')
  const editor = page.locator('.canvas-cell-input input')
  await expect(editor).toBeVisible()
  await editor.fill('未提交的编辑')
  await ribbon.getByRole('button', { name: '十字高亮', exact: true }).click()
  await expect(editor).toHaveValue('未提交的编辑')
  await expect(editor).toBeFocused()
  expect(await capture(page)).toEqual(before)
  await ribbon.getByRole('button', { name: '十字高亮选项', exact: true }).click()
  await page.getByRole('menu', { name: '十字高亮', exact: true }).getByRole('menuitem', { name: '绿色', exact: true }).click()
  await expect(editor).toHaveValue('未提交的编辑')
  await expect(editor).toBeFocused()
  expect(await capture(page)).toEqual(before)
  await editor.press('Escape')
  await expect(editor).not.toBeVisible()
  expect(await capture(page)).toEqual(before)
})
