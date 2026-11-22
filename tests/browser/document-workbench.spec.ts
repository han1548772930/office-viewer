import { expect, test } from '@playwright/test'

type WorkbenchSnapshot = {
  text: string
  fontFamilies: string[]
  blocks: string[]
  controls: Array<{ type: string; startOffset: number; endOffset: number }>
}

async function snapshot(page: import('@playwright/test').Page): Promise<WorkbenchSnapshot> {
  return page.evaluate(() => {
    const root = document.querySelector<HTMLElement>('.document-workbench')
    let component = root?.__vueParentComponent
    while (component && component.type?.__name !== 'DocumentWorkbench') component = component.parent
    const state = component?.setupState as { model?: { blocks?: Array<{ type: string; runs?: Array<{ text: string; fontFamily: string }> }>; contentControls?: Array<{ type: string; startOffset: number; endOffset: number }> } } | undefined
    const model = state?.model
    const runs = model?.blocks?.flatMap((block) => block.runs ?? []) ?? []
    return {
      text: runs.map((run) => run.text).join(''),
      fontFamilies: runs.map((run) => run.fontFamily),
      blocks: model?.blocks?.map((block) => block.type) ?? [],
      controls: model?.contentControls?.map(({ type, startOffset, endOffset }) => ({ type, startOffset, endOffset })) ?? [],
    }
  })
}

test.beforeEach(async ({ page }) => {
  await page.goto('/')
  await page.getByRole('tab', { name: '文档', exact: true }).click()
  await expect(page.locator('.document-workbench')).toBeVisible()
  await page.getByRole('textbox', { name: '文档画布编辑区' }).waitFor()
})

test('an antdv-next font selection retains the Canvas text range', async ({ page }) => {
  const bridge = page.getByLabel('文档文字输入')
  await bridge.focus()
  await bridge.pressSequentially('选区保留')
  await bridge.press('Control+A')
  const font = page.locator('.document-workbench').getByRole('combobox', { name: '字体', exact: true })
  await font.click()
  await page.locator('.ant-select-dropdown:visible .ant-select-item').filter({ hasText: 'Georgia' }).click()
  const result = await snapshot(page)
  expect(result.text).toBe('选区保留')
  expect(result.fontFamilies).toContain('Georgia')
})

test('Canvas input bridge commits Chinese composition as one undoable edit', async ({ page }) => {
  const bridge = page.getByLabel('文档文字输入')
  await bridge.focus()
  await bridge.evaluate((element) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '中文输入' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文输入' }))
  })
  await expect.poll(async () => (await snapshot(page)).text).toBe('中文输入')
  await bridge.focus()
  await bridge.press('Control+Z')
  await expect.poll(async () => (await snapshot(page)).text).toBe('')
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '重做', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).text).toBe('中文输入')
})

test('table insertion modal applies dimensions and Ctrl+Z removes the inserted table', async ({ page }) => {
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '表格', exact: true }).click()
  await page.getByRole('menuitem', { name: '插入表格', exact: true }).click()
  const dialog = page.locator('.ant-modal')
  await expect(dialog).toBeVisible()
  await dialog.getByRole('spinbutton').nth(0).fill('2')
  await dialog.getByRole('spinbutton').nth(1).fill('3')
  await dialog.getByRole('button', { name: /插\s*入/ }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(async () => (await snapshot(page)).blocks).toContain('table')
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '开始', exact: true }).click()
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).blocks).not.toContain('table')
})

test('document context actions use the antdv-next menu and preserve paragraph/table groups', async ({ page }) => {
  const canvas = page.getByRole('textbox', { name: '文档画布编辑区' })
  const bounds = await canvas.boundingBox()
  if (!bounds) throw new Error('Document canvas is not visible')
  await page.mouse.click(bounds.x + 240, bounds.y + 100, { button: 'right' })
  const menu = page.locator('.document-context-menu .ant-menu')
  await expect(menu).toBeVisible()
  await expect(menu).toHaveAttribute('role', 'menu')
  await expect(menu.getByRole('menuitem')).toContainText(['复制', '剪切', '粘贴', '选择性粘贴', '删除', '段落设置', '分节设置'])
  await expect(menu.getByRole('separator')).toHaveCount(2)
  await page.keyboard.press('Escape')
  await expect(page.locator('.document-context-menu')).toBeHidden()
  await expect(page.getByLabel('文档文字输入')).toBeFocused()

  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '表格', exact: true }).click()
  await page.getByRole('menuitem', { name: '插入表格', exact: true }).click()
  const dialog = page.locator('.ant-modal')
  await expect(dialog).toBeVisible()
  await dialog.getByRole('spinbutton').nth(0).fill('2')
  await dialog.getByRole('spinbutton').nth(1).fill('2')
  await dialog.getByRole('button', { name: /插\s*入/ }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(async () => (await snapshot(page)).blocks).toContain('table')
  const table = page.getByRole('textbox', { name: '文档画布编辑区' })
  const tableBounds = await table.boundingBox()
  if (!tableBounds) throw new Error('Document canvas is not visible')
  const tablePoint = await page.evaluate(() => {
    const root = document.querySelector<HTMLElement>('.document-workbench')
    let component = root?.__vueParentComponent
    while (component && component.type?.__name !== 'DocumentWorkbench') component = component.parent
    const state = component?.setupState as { canvasRef?: { getLayout?: () => { width: number; cells: Array<{ x: number; y: number; width: number; height: number; target: { blockId: string } }> } } }
    const canvas = state.canvasRef
    const layout = canvas?.getLayout?.()
    const table = document.querySelector<HTMLElement>('.document-scroll')
    if (!layout || !table) return undefined
    const viewport = table.getBoundingClientRect()
    const firstCell = layout.cells[0]
    if (!firstCell) return undefined
    const zoom = 1
    const centeredOffset = Math.max(0, (table.clientWidth - layout.width * zoom) / 2)
    const point = {
      x: viewport.left + centeredOffset + (firstCell.x + firstCell.width / 2) * zoom - table.scrollLeft,
      y: viewport.top + firstCell.y + firstCell.height / 2 - table.scrollTop,
    }
    return point
  })
  if (!tablePoint) throw new Error('Inserted table cell is not present in the document layout')
  await page.mouse.click(tablePoint.x, tablePoint.y, { button: 'right' })
  const groupedMenu = page.locator('.document-context-menu .ant-menu')
  await expect(groupedMenu).toBeVisible()
  await expect(groupedMenu.getByRole('separator')).toHaveCount(3)
  await expect(groupedMenu.getByRole('menuitem')).toContainText(['复制', '段落设置', '在上方插入行', '删除表格'])
})
