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
    const state = root?.__vueParentComponent?.setupState as { model?: { blocks?: Array<{ type: string; runs?: Array<{ text: string; fontFamily: string }> }>; contentControls?: Array<{ type: string; startOffset: number; endOffset: number }> } } | undefined
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
  const dialog = page.locator('.ant-modal')
  await expect(dialog).toBeVisible()
  await dialog.getByRole('spinbutton').nth(0).fill('2')
  await dialog.getByRole('spinbutton').nth(1).fill('3')
  await dialog.locator('.ant-modal-footer button').filter({ hasText: '确定' }).click()
  await expect(dialog).toBeHidden()
  await expect.poll(async () => (await snapshot(page)).blocks).toContain('table')
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '开始', exact: true }).click()
  await page.getByTestId('workbench-ribbon').getByRole('button', { name: '撤销', exact: true }).click()
  await expect.poll(async () => (await snapshot(page)).blocks).not.toContain('table')
})
