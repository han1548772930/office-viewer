import { expect, test, type Page } from '@playwright/test'
import type { DocumentModel } from '../../packages/office-studio/src/domain/model'
import type { DocumentTable } from '../../packages/office-studio/src/features/document/model/document-model'

type RuntimeModule = typeof import('../../examples/workbench/src/runtime')
const ribbon = (page: Page) => page.getByTestId('workbench-ribbon')
const bridge = (page: Page) => page.getByLabel('文档文字输入')

async function state(page: Page): Promise<{ model: DocumentModel; text: string }> {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').find((entry) => new URL(entry.name).pathname === '/src/runtime.ts')?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(path) as RuntimeModule
    return JSON.parse(JSON.stringify({ model: studio.getDocumentModel(studio.activeDocumentId), text: officeAPI.getActiveDocument()?.getText() ?? '' }))
  })
}
async function table(page: Page): Promise<DocumentTable | undefined> {
  return (await state(page)).model.blocks.find((block) => block.type === 'table')?.metadata.table as DocumentTable | undefined
}
async function tableAction(page: Page, label: string): Promise<void> {
  await ribbon(page).getByRole('tab', { name: '表格工具', exact: true }).click()
  await ribbon(page).getByRole('button', { name: label, exact: true }).click()
}
async function history(page: Page, label: '撤销' | '重做'): Promise<void> {
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await ribbon(page).getByRole('button', { name: label, exact: true }).click()
}
async function insertTable(page: Page): Promise<void> {
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '表格', exact: true }).click()
  await page.getByRole('menu', { name: '表格', exact: true }).getByRole('button', { name: '插入表格', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入表格', exact: true })
  await dialog.getByRole('spinbutton').nth(0).fill('2')
  await dialog.getByRole('spinbutton').nth(1).fill('2')
  await dialog.getByRole('button', { name: /^确\s*定$/ }).click()
  await expect(dialog).toBeHidden()
  await expect(ribbon(page).getByRole('tab', { name: '表格工具', exact: true })).toHaveAttribute('aria-selected', 'true')
}
test.beforeEach(async ({ page }) => {
  await page.goto('/')
  await page.getByRole('tab', { name: '文档', exact: true }).click()
  await expect(page.getByRole('textbox', { name: '文档画布编辑区' })).toBeVisible()
})

test('header session preserves the body caret while inserting a field and restores the previous tab', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 })
  await bridge(page).focus(); await bridge(page).pressSequentially('body')
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '页眉和页脚', exact: true }).click()
  const story = page.getByRole('textbox', { name: '编辑页眉', exact: true })
  await expect(story).toBeFocused()
  await story.fill('report')
  await story.evaluate((el: HTMLTextAreaElement) => el.setSelectionRange(2, 2))
  await ribbon(page).getByRole('button', { name: '页码', exact: true }).click()
  await page.getByRole('menuitem', { name: '当前页码' }).click()
  await expect.poll(async () => (await state(page)).model.sectionHeaderFooters?.first?.pageFields?.map(f => [f.field, f.offset])).toEqual([['PAGE', 2]])
  await expect(story).toBeFocused()
  await story.press('Escape')
  await expect(story).toHaveCount(0)
  await expect(ribbon(page).getByRole('tab', { name: '插入', exact: true })).toHaveAttribute('aria-selected', 'true')
  await expect(bridge(page)).toBeFocused()
  await bridge(page).pressSequentially('!')
  await expect.poll(async () => (await state(page)).text.trim()).toBe('body!')
})

test('compact position control edits the active section without closing the story session', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 })
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '页眉和页脚', exact: true }).click()
  await page.setViewportSize({ width: 820, height: 900 })
  await ribbon(page).getByRole('button', { name: '位置', exact: true }).click()
  const margin = page.getByRole('spinbutton', { name: '页眉距顶端' }).filter({ visible: true })
  await margin.fill('25.5'); await margin.press('Enter')
  await expect.poll(async () => (await state(page)).model.sectionSetup?.marginHeader).toBe(25.5)
  await expect(page.getByRole('textbox', { name: '编辑页眉' })).toBeVisible()
})

test('table context edits real rows and columns, restores a valid caret, and records structural undo', async ({ page }) => {
  await insertTable(page)
  await tableAction(page, '在下方插入行')
  await expect.poll(async () => (await table(page))?.rows.length).toBe(3)
  await bridge(page).focus(); await bridge(page).pressSequentially('X')
  await expect.poll(async () => (await table(page))?.rows[1]?.cells[0]?.blocks[0]?.runs.map((run) => run.text).join('')).toBe('X')
  await tableAction(page, '在右侧插入列')
  await expect.poll(async () => (await table(page))?.rows[0]?.cells.length).toBe(3)
  await history(page, '撤销')
  await expect.poll(async () => (await table(page))?.rows[0]?.cells.length).toBe(2)
  await history(page, '重做')
  await expect.poll(async () => (await table(page))?.rows[0]?.cells.length).toBe(3)
  await tableAction(page, '删除所选列')
  await expect.poll(async () => (await table(page))?.rows[0]?.cells.length).toBe(2)
  await tableAction(page, '删除所选行')
  await expect.poll(async () => (await table(page))?.rows.length).toBe(2)
  await tableAction(page, '删除表格')
  await expect.poll(async () => (await table(page)) === undefined).toBe(true)
  await expect(ribbon(page).getByRole('tab', { name: '表格工具', exact: true })).toHaveCount(0)
  await bridge(page).focus(); await bridge(page).pressSequentially('Z')
  await expect.poll(async () => (await state(page)).text.trim()).toBe('Z')
  await history(page, '撤销')
  await history(page, '撤销')
  await expect.poll(async () => (await table(page))?.rows.length).toBe(2)
})

test('a locked cell control prevents deleting its table and the failed command creates no undo step', async ({ page }) => {
  await insertTable(page)
  await bridge(page).focus(); await bridge(page).pressSequentially('保护')
  await page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').find((entry) => new URL(entry.name).pathname === '/src/runtime.ts')?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(path) as RuntimeModule
    const document = officeAPI.getActiveDocument()!
    const table = document.getModel().blocks.find((block) => block.type === 'table')!.metadata.table as DocumentTable
    document.addContentControl({ id: 'protected-cell', type: 'text', blockId: table.rows[0]!.cells[0]!.blocks[0]!.id, startOffset: 0, endOffset: 2, lock: 'sdtLocked' })
  })
  await tableAction(page, '删除表格')
  await expect(page.locator('.document-notice')).toContainText('wrapper is locked')
  expect((await table(page))?.rows.length).toBe(2)
  expect((await state(page)).model.contentControls?.map((control) => control.id)).toContain('protected-cell')
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).model.contentControls?.length ?? 0).toBe(0)
  expect((await table(page))?.rows.length).toBe(2)
})

test('symbol and compound emoji pickers replace the preserved Canvas selection and undo as one action', async ({ page }) => {
  await bridge(page).focus(); await bridge(page).pressSequentially('replace')
  await bridge(page).press('Control+A')
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '符号', exact: true }).click()
  await page.getByRole('dialog', { name: '插入符号', exact: true }).getByRole('button', { name: '插入 ©', exact: true }).click()
  await expect.poll(async () => (await state(page)).text).toBe('©')
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).text).toBe('replace')
  await bridge(page).focus(); await bridge(page).press('Control+A')
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '表情', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '插入表情', exact: true })
  await dialog.getByRole('textbox', { name: '自定义字符', exact: true }).fill('👨‍👩‍👧‍👦')
  await dialog.getByRole('button', { name: '插入字符', exact: true }).click()
  await expect.poll(async () => (await state(page)).text).toBe('👨‍👩‍👧‍👦')
  await bridge(page).focus(); await bridge(page).press('Backspace')
  await expect.poll(async () => (await state(page)).text).toBe('')
})

test('the content-control split menu starts the chosen real control type', async ({ page }) => {
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '内容控件选项', exact: true }).click()
  await page.getByRole('button', { name: '复选框控件', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '内容控件', exact: true })
  await expect(dialog.getByRole('region', { name: '新增内容控件', exact: true })).toBeVisible()
  await dialog.getByRole('button', { name: '插入到当前范围', exact: true }).click()
  await expect.poll(async () => (await state(page)).model.contentControls?.[0]?.type).toBe('checkbox')
})

test('toolbar cut writes the real browser clipboard and removes content only as an undoable model command', async ({ page }) => {
  await bridge(page).focus(); await bridge(page).pressSequentially('Clipboard')
  await bridge(page).press('Control+A')
  await bridge(page).press('Control+B')
  await ribbon(page).getByRole('button', { name: '剪切', exact: true }).click()
  await expect.poll(async () => (await state(page)).text).toBe('')
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('Clipboard')
  const html = await page.evaluate(async () => {
    const item = (await navigator.clipboard.read()).find(candidate => candidate.types.includes('text/html'))
    return item ? (await item.getType('text/html')).text() : ''
  })
  expect(html).toMatch(/<strong>Clipboard<\/strong>/)
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).text).toBe('Clipboard')
  expect((await state(page)).model.blocks.flatMap(block => block.runs).some(run => run.text === 'Clipboard' && run.bold)).toBe(true)
  await history(page, '重做')
  await expect.poll(async () => (await state(page)).text).toBe('')
})

test('toolbar paste retains HTML runs and uses the same command as the native paste input event', async ({ page }) => {
  await bridge(page).focus(); await bridge(page).pressSequentially('old')
  await bridge(page).press('Control+A')
  await page.evaluate(async () => {
    await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob(['Source'], { type: 'text/plain' }), 'text/html': new Blob(['<p><strong>Source</strong></p>'], { type: 'text/html' }) })])
  })
  await pasteMode(page, '保留源格式')
  await expect.poll(async () => (await state(page)).text).toBe('Source')
  expect((await state(page)).model.blocks.flatMap((block) => block.runs).some((run) => run.text === 'Source' && run.bold)).toBe(true)
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).text).toBe('old')
  await bridge(page).focus()
  await bridge(page).press('Control+A')
  await bridge(page).evaluate((element) => {
    const data = new DataTransfer()
    data.setData('text/plain', 'Native')
    data.setData('text/html', '<p><i>Native</i></p>')
    element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }))
  })
  await expect.poll(async () => (await state(page)).text).toBe('Native')
  expect((await state(page)).model.blocks.flatMap((block) => block.runs).some((run) => run.text === 'Native' && run.italic)).toBe(true)
})

test('plain-text paste ignores source HTML marks and preserves destination formatting', async ({ page }) => {
  await page.evaluate(async () => {
    await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob(['Plain'], { type: 'text/plain' }), 'text/html': new Blob(['<p><strong>Plain</strong></p>'], { type: 'text/html' }) })])
  })
  await pasteMode(page, '仅保留文本')
  await expect.poll(async () => (await state(page)).text).toBe('Plain')
  expect((await state(page)).model.blocks.flatMap((block) => block.runs).every((run) => !run.bold)).toBe(true)
})

async function pasteMode(page: Page, label: '保留源格式' | '匹配目标格式' | '仅保留文本') {
  await ribbon(page).getByRole('button', { name: '粘贴', exact: true }).click()
  await page.getByRole('menuitem', { name: label, exact: true }).click()
}

test('native Ctrl+Shift+V shares plain paste semantics and the following ordinary paste retains its source', async ({ page }) => {
  await bridge(page).focus()
  for (const [text, plain] of [['Plain', true], ['Rich', false]] as const) {
    await bridge(page).evaluate((element, { text, plain }) => {
      if (plain) element.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, ctrlKey: true, shiftKey: true, key: 'V' }))
      const data = new DataTransfer()
      data.setData('text/plain', text)
      data.setData('text/html', `<p><strong>${text}</strong></p>`)
      element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }))
    }, { text, plain })
    await expect.poll(async () => (await state(page)).text).toBe(plain ? 'Plain' : 'PlainRich')
  }
  const runs = (await state(page)).model.blocks.flatMap(block => block.runs)
  expect(runs.find(run => run.text === 'Plain')?.bold).toBe(false)
  expect(runs.find(run => run.text === 'Rich')?.bold).toBe(true)
})

test('match destination keeps the source emphasis in the destination character formatting', async ({ page }) => {
  await bridge(page).focus()
  await bridge(page).pressSequentially('Base')
  const before = (await state(page)).model.blocks.flatMap(block => block.runs).find(run => run.text.includes('Base'))
  await page.evaluate(async () => navigator.clipboard.write([new ClipboardItem({
    'text/plain': new Blob(['Mixed'], { type: 'text/plain' }),
    'text/html': new Blob(['<p><strong style="font-family:Georgia;font-size:32pt">Mixed</strong></p>'], { type: 'text/html' }),
  })]))
  await pasteMode(page, '匹配目标格式')
  await expect.poll(async () => (await state(page)).text).toBe('BaseMixed')
  const pasted = (await state(page)).model.blocks.flatMap(block => block.runs).find(run => run.text === 'Mixed')
  expect(pasted).toMatchObject({ bold: true, fontFamily: before?.fontFamily, fontSize: before?.fontSize })
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).text).toBe('Base')
})

test('a pending clipboard read cannot overwrite newer Canvas input', async ({ page }) => {
  await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
      readText: () => new Promise<string>((resolve) => { (window as Window & { resolveTestClipboard?: (value: string) => void }).resolveTestClipboard = resolve }),
    } })
  })
  await pasteMode(page, '仅保留文本')
  await expect.poll(() => page.evaluate(() => typeof (window as Window & { resolveTestClipboard?: unknown }).resolveTestClipboard)).toBe('function')
  await bridge(page).focus(); await bridge(page).pressSequentially('N')
  await page.evaluate(() => (window as Window & { resolveTestClipboard?: (value: string) => void }).resolveTestClipboard?.('stale'))
  await expect(page.locator('.document-notice')).toContainText('文档或选区已变化')
  expect((await state(page)).text).toBe('N')
})
