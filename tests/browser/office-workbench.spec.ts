import { expect, test, type Locator, type Page } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { inflateRawSync } from 'node:zlib'
import type { DocumentModel, WorkbookModel } from '../../packages/office-studio/src/domain/model'

type RuntimeModule = typeof import('../../examples/workbench/src/runtime')
const runtimeErrors = new WeakMap<Page, string[]>()

/** Reading a model must not use FWorkbook/FDocument.save(), which creates a savepoint. */
async function models(page: Page): Promise<{ workbook: WorkbookModel; document: DocumentModel; canUndo: boolean; canRedo: boolean }> {
  return page.evaluate(async () => {
    // Reuse Vite's loaded module URL, including its HMR query, instead of
    // creating a second OfficeStudio instance by importing a bare URL.
    const runtimePath = performance.getEntriesByType('resource').filter((entry) => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio } = await import(runtimePath) as RuntimeModule
    const workbook = studio.getWorkbookModel(studio.activeWorkbookId)
    const document = studio.getDocumentModel(studio.activeDocumentId)
    if (!workbook || !document) throw new Error('The active workspace models are missing')
    return JSON.parse(JSON.stringify({ workbook, document, canUndo: studio.editor.canUndo, canRedo: studio.editor.canRedo }))
  })
}

async function cellValue(page: Page, address: string): Promise<unknown> {
  const { workbook } = await models(page)
  const parts = /^([A-Z]+)(\d+)$/.exec(address)
  if (!parts) throw new Error(`Invalid test cell address: ${address}`)
  const column = [...parts[1]!].reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1
  const key = `${Number(parts[2]) - 1}:${column}`
  return workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId)?.cells[key]?.value
}

function ribbon(page: Page): Locator { return page.getByTestId('workbench-ribbon') }
function visibleRibbonButton(page: Page, name: string): Locator { return ribbon(page).getByRole('button', { name, exact: true }) }

/** Inspect the downloaded Office ZIP independently of the application's writer. */
function readOfficePart(bytes: Buffer, expectedPath: string): string {
  let end = bytes.length - 22
  const earliestEnd = Math.max(0, bytes.length - 65_557)
  while (end >= earliestEnd && bytes.readUInt32LE(end) !== 0x06054b50) end -= 1
  if (end < earliestEnd) throw new Error('The downloaded file has no ZIP central directory')
  const entries = bytes.readUInt16LE(end + 10)
  let cursor = bytes.readUInt32LE(end + 16)
  for (let index = 0; index < entries; index += 1) {
    if (bytes.readUInt32LE(cursor) !== 0x02014b50) throw new Error('The Office central directory is invalid')
    const method = bytes.readUInt16LE(cursor + 10)
    const compressedSize = bytes.readUInt32LE(cursor + 20)
    const nameLength = bytes.readUInt16LE(cursor + 28)
    const extraLength = bytes.readUInt16LE(cursor + 30)
    const commentLength = bytes.readUInt16LE(cursor + 32)
    const path = bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8')
    const local = bytes.readUInt32LE(cursor + 42)
    if (path === expectedPath) {
      if (bytes.readUInt32LE(local) !== 0x04034b50) throw new Error('The Office part has no local ZIP header')
      const dataStart = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28)
      const compressed = bytes.subarray(dataStart, dataStart + compressedSize)
      if (method === 0) return compressed.toString('utf8')
      if (method === 8) return inflateRawSync(compressed).toString('utf8')
      throw new Error(`Unsupported ZIP compression: ${method}`)
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  throw new Error(`The downloaded Office file is missing ${expectedPath}`)
}

async function gotoCell(page: Page, address: string): Promise<void> {
  const nameBox = page.getByRole('textbox', { name: '名称框', exact: true })
  await nameBox.fill(address)
  await nameBox.press('Enter')
  await expect(nameBox).toHaveValue(address)
}

async function switchWorkspace(page: Page, name: '表格' | '文档'): Promise<void> {
  await page.getByRole('navigation', { name: '工作台', exact: true }).getByRole('tab', { name, exact: true }).click()
  await expect(page.locator(name === '表格' ? '.excel-workbench' : '.document-workbench')).toBeVisible()
}

async function expectHeight(locator: Locator, height: number): Promise<void> {
  await expect.poll(async () => (await locator.boundingBox())?.height).toBe(height)
}

test.beforeEach(async ({ page }) => {
  const errors: string[] = []
  runtimeErrors.set(page, errors)
  page.on('pageerror', (error) => errors.push(error.message))
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeVisible()
})

test.afterEach(async ({ page }) => { expect(runtimeErrors.get(page) ?? []).toEqual([]) })

async function chooseLayout(page: Page, name: '经典工具栏' | '双行工具栏' | '折叠分类工具栏' | '单行合并工具栏' = '经典工具栏') {
  await page.getByRole('button', { name: '设置', exact: true }).click()
  await page.getByRole('menuitem', { name: /^工具栏布局/ }).click()
  await page.getByRole('menuitem', { name, exact: true }).click()
}

test('collapsed chooses a category in the toolbar; simple merges categories without changing selection', async ({ page }) => {
  await chooseLayout(page, '折叠分类工具栏')
  await expect(ribbon(page).getByRole('tablist')).toHaveCount(0)
  await ribbon(page).getByRole('button', { name: '开始', exact: true }).press('ArrowDown')
  await page.getByRole('menuitem', { name: '插入', exact: true }).click()
  await expect(ribbon(page).getByRole('toolbar', { name: '插入', exact: true })).toBeVisible()
  await chooseLayout(page, '单行合并工具栏')
  await expect(ribbon(page).getByRole('tablist')).toHaveCount(0)
  await expect(ribbon(page).locator('.ribbon-category-toggle')).toHaveCount(0)
  await expect(ribbon(page).getByRole('toolbar', { name: '开始', exact: true })).toBeVisible()
  await chooseLayout(page, '双行工具栏')
  await expect(ribbon(page).getByRole('tab', { name: '插入', exact: true })).toHaveAttribute('aria-selected', 'true')
})

test('product links honor the source header order, route and browser history', async ({ page }) => {
  await page.goto('/#docs')
  const header = page.locator('.workspace-header')
  await expectHeight(header, 36)
  const docs = header.getByRole('tab', { name: '文档', exact: true }), sheets = header.getByRole('tab', { name: '表格', exact: true })
  await expect(docs).toHaveAttribute('aria-current', 'page')
  await expect(sheets).toHaveAttribute('href', '#sheets')
  const left = await header.locator('.workspace-brand').boundingBox(), tab = await sheets.boundingBox(), next = await docs.boundingBox()
  expect(left && tab && next && left.x + left.width <= tab.x && tab.x + tab.width <= next.x).toBeTruthy()
  await sheets.click()
  await expect(page).toHaveURL(/#sheets$/)
  await page.goBack(); await expect(docs).toHaveAttribute('aria-selected', 'true')
  await page.goForward(); await expect(sheets).toHaveAttribute('aria-selected', 'true')
  await page.goto('/#unknown'); await expect(page).toHaveURL(/#sheets$/)
})

test('workspace defaults to the source example grid and remembers layout across products and reload', async ({ page }) => {
  const assertGrid = async () => {
    await expectHeight(ribbon(page).locator('.workbench-ribbon-tabs'), 32)
    await expectHeight(page.getByTestId('ribbon-grid-toolbar'), 69)
    await expect(ribbon(page).locator('.ribbon-grid-cells').first()).toHaveCSS('grid-template-rows', '24px 24px')
  }
  await assertGrid()
  await switchWorkspace(page, '文档')
  await assertGrid()
  await chooseLayout(page)
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
  await switchWorkspace(page, '表格')
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
  await page.reload()
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
  await chooseLayout(page, '双行工具栏')
  await assertGrid()
})

test('Sheets uses classic Ribbon, a formula header, Canvas and one product footer', async ({ page }) => {
  await chooseLayout(page)
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(error.message))
  await expect(ribbon(page)).toHaveCount(1)
  await expectHeight(ribbon(page).locator('.workbench-ribbon-tabs'), 36)
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
  const formula = page.getByTestId('sheet-formula-bar')
  await expectHeight(formula, 28)
  const nameBox = formula.getByRole('textbox', { name: '名称框', exact: true })
  await expect.poll(async () => (await nameBox.boundingBox())?.width).toBe(100)
  await expect(nameBox).toHaveCSS('border-top-left-radius', '0px')
  await expect(formula.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveCSS('border-top-left-radius', '0px')
  const fontSelector = ribbon(page).locator('[data-ribbon-item="sheet.command.set-font-family"] .ribbon-select').filter({ visible: true })
  await expect(fontSelector).toHaveCSS('border-top-left-radius', '4px')
  await expect(fontSelector).toHaveCSS('border-top-color', 'rgba(0, 0, 0, 0)')
  const grid = page.getByRole('grid', { name: '工作表画布', exact: true })
  await grid.press('A')
  const cellEditor = page.locator('.canvas-cell-input')
  await expect(cellEditor).toHaveCSS('border-top-left-radius', '0px')
  await expect(cellEditor).not.toHaveCSS('border-top-color', 'rgba(0, 0, 0, 0)')
  await cellEditor.press('Escape')
  await expectHeight(page.getByTestId('sheet-footer'), 36)
  await expect(page.getByTestId('sheet-footer')).toBeInViewport()
  await expect(page.locator('.excel-workbench footer')).toHaveCount(1)
  await expect(page.locator('.app-footer')).toHaveCount(0)
  await expect(page.locator('canvas[data-render-surface="true"]')).toHaveCount(1)
  await expect(page.locator('.excel-workbench [contenteditable="true"]')).toHaveCount(0)
  const positions = await Promise.all([ribbon(page), formula, page.getByRole('grid', { name: '工作表画布' }), page.getByTestId('sheet-footer')].map((part) => part.boundingBox()))
  expect(positions.every(Boolean)).toBe(true)
  for (let index = 1; index < positions.length; index += 1) expect(positions[index]!.y).toBeGreaterThanOrEqual(positions[index - 1]!.y + positions[index - 1]!.height)
  expect(positions.at(-1)!.y + positions.at(-1)!.height).toBeLessThanOrEqual(page.viewportSize()!.height)
  expect(errors).toEqual([])
})

test('Ribbon tabs implement roving keyboard focus and keep their classic row dimensions', async ({ page }) => {
  await chooseLayout(page)
  const tabs = ribbon(page).getByRole('tab')
  const names = await tabs.allTextContents()
  expect(names).toContain('开始')
  expect(names).toContain('插入')
  expect(names).toContain('视图')
  await tabs.first().focus()
  await tabs.first().press('End')
  await expect(tabs.last()).toBeFocused()
  await expect(tabs.last()).toHaveAttribute('aria-selected', 'true')
  await expect(tabs.last()).toHaveAttribute('tabindex', '0')
  await tabs.last().press('Home')
  await expect(tabs.first()).toBeFocused()
  await tabs.first().press('ArrowRight')
  await expect(tabs.nth(1)).toBeFocused()
  await expect(tabs.nth(1)).toHaveAttribute('aria-selected', 'true')
  await expect(tabs.first()).toHaveAttribute('tabindex', '-1')
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
})

test('the formula session cancels, commits and participates in the document history', async ({ page }) => {
  await gotoCell(page, 'B3')
  const input = page.getByRole('textbox', { name: '公式编辑', exact: true })
  await input.fill('取消的值')
  await page.getByRole('button', { name: '取消编辑', exact: true }).click()
  await expect.poll(() => cellValue(page, 'B3')).toBeUndefined()
  await input.fill('314')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(() => cellValue(page, 'B3')).toBe(314)
  await visibleRibbonButton(page, '撤销').click()
  await expect.poll(() => cellValue(page, 'B3')).toBeUndefined()
  await visibleRibbonButton(page, '重做').click()
  await expect.poll(() => cellValue(page, 'B3')).toBe(314)
  await page.getByRole('button', { name: '展开公式栏', exact: true }).click()
  await expectHeight(page.getByTestId('sheet-formula-bar'), 80)
  await page.getByRole('button', { name: '折叠公式栏', exact: true }).click()
  await expectHeight(page.getByTestId('sheet-formula-bar'), 28)
})

test('switching worksheets and products preserves content and isolated undo history', async ({ page }) => {
  await gotoCell(page, 'C4')
  await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill('工作簿内容')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(() => cellValue(page, 'C4')).toBe('工作簿内容')
  const originalSheet = (await models(page)).workbook.activeSheetId
  await page.getByRole('button', { name: '添加工作表', exact: true }).click()
  await expect.poll(async () => (await models(page)).workbook.sheets.length).toBe(2)
  const firstSheetName = (await models(page)).workbook.sheets.find((sheet) => sheet.id === originalSheet)!.name
  await page.getByRole('tablist', { name: '工作表', exact: true }).getByRole('tab', { name: firstSheetName, exact: true }).click()
  await gotoCell(page, 'C4')
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('工作簿内容')
  await switchWorkspace(page, '文档')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.evaluate((element) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '独立文档内容' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '独立文档内容' }))
  })
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('独立文档内容')
  await bridge.press('Control+Z')
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('')
  await switchWorkspace(page, '表格')
  await gotoCell(page, 'C4')
  await expect(page.getByRole('textbox', { name: '公式编辑', exact: true })).toHaveValue('工作簿内容')
  await visibleRibbonButton(page, '撤销').click()
  await expect.poll(async () => (await models(page)).workbook.sheets.length).toBe(1)
  await expect.poll(() => cellValue(page, 'C4')).toBe('工作簿内容')
  await switchWorkspace(page, '文档')
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('')
  await visibleRibbonButton(page, '重做').click()
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('独立文档内容')
})

test('narrow Sheets keeps overflow tools operable and the editor inside the viewport', async ({ page }) => {
  await chooseLayout(page)
  await page.setViewportSize({ width: 620, height: 800 })
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(error.message))
  await gotoCell(page, 'B2')
  await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill('窄屏工具')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  const trigger = ribbon(page).getByRole('button', { name: '更多功能', exact: true })
  await expect(trigger).toBeVisible()
  await trigger.click()
  const menu = page.getByRole('dialog', { name: '更多功能', exact: true })
  await expect(menu).toBeVisible()
  await menu.getByRole('button', { name: '清除内容', exact: true }).click()
  await expect.poll(() => cellValue(page, 'B2')).toBe('')
  await expect(menu).toBeVisible()
  await trigger.focus()
  await trigger.press('ArrowDown')
  await expect(menu).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(menu).toBeHidden()
  await expect(trigger).toBeFocused()
  const extent = await page.evaluate(() => ({ width: innerWidth, scroll: document.documentElement.scrollWidth }))
  expect(extent.scroll).toBeLessThanOrEqual(extent.width)
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
  expect(errors).toEqual([])
})

test('Docs keeps its logical text range while formatting through classic Ribbon', async ({ page }) => {
  await chooseLayout(page)
  await switchWorkspace(page, '文档')
  await expect(ribbon(page)).toHaveCount(1)
  await expectHeight(ribbon(page).locator('.workbench-ribbon-tabs'), 36)
  await expectHeight(ribbon(page).locator('.workbench-ribbon-toolbar'), 40)
  await expect(page.locator('.document-workbench footer')).toHaveCount(1)
  await expect(page.locator('.document-workbench [contenteditable="true"]')).toHaveCount(0)
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.pressSequentially('选区格式保留')
  await bridge.press('Control+A')
  await visibleRibbonButton(page, '加粗').click()
  await expect.poll(async () => (await models(page)).document.blocks[0]!.runs.every((run) => run.bold)).toBe(true)
  await expect(visibleRibbonButton(page, '加粗')).toHaveAttribute('aria-pressed', 'true')
  await visibleRibbonButton(page, '撤销').click()
  await expect.poll(async () => (await models(page)).document.blocks[0]!.runs.every((run) => run.bold)).toBe(false)
  await visibleRibbonButton(page, '重做').click()
  await expect.poll(async () => (await models(page)).document.blocks[0]!.runs.map((run) => run.text).join('')).toBe('选区格式保留')
  await expect.poll(async () => (await models(page)).document.blocks[0]!.runs.every((run) => run.bold)).toBe(true)
  await expect(page.locator('.document-workbench canvas[data-render-surface="true"]')).toHaveCount(1)
})

test('Docs traditional outline is configurable and navigates stable paragraph identities', async ({ page }) => {
  await page.evaluate(async () => {
    const runtimePath = performance.getEntriesByType('resource').filter((entry) => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { officeAPI } = await import(runtimePath) as RuntimeModule
    const paragraphStyle = { align: 'left' as const, indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 }
    const run = (text: string) => ({ id: `run-${text}`, text, bold: false, italic: false, underline: false, strike: false, color: '#1f2937', fontFamily: 'Arial', fontSize: 11 })
    officeAPI.createDocument({
      id: 'outline-browser-fixture', name: '大纲导航测试', kind: 'document', modified: false, wordCount: 0,
      page: { width: 794, height: 1123, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 },
      blocks: [
        { id: 'heading-a', type: 'heading', level: 1, runs: [run('第一章')], style: paragraphStyle, metadata: {} },
        { id: 'paragraph-a', type: 'paragraph', level: 0, runs: [run('第一章正文')], style: paragraphStyle, metadata: {} },
        { id: 'heading-b', type: 'heading', level: 2, runs: [run('第一节')], style: paragraphStyle, metadata: {} },
        { id: 'paragraph-b', type: 'paragraph', level: 0, runs: [run('第一节正文')], style: paragraphStyle, metadata: {} },
      ],
    })
  })
  await expect(page.locator('.document-workbench')).toBeVisible()
  await expect(page.getByRole('navigation', { name: '文档大纲', exact: true })).toHaveCount(0)
  await ribbon(page).getByRole('tab', { name: '视图', exact: true }).click()
  await visibleRibbonButton(page, '文档大纲').click()
  const outline = page.getByRole('navigation', { name: '文档大纲', exact: true })
  await expect(outline).toBeVisible()
  const firstHeading = outline.locator('[data-outline-item="heading-a"]')
  const nextHeading = outline.locator('[data-outline-item="heading-b"]')
  await nextHeading.click()
  await expect(nextHeading).toHaveAttribute('aria-current', 'location')
  await firstHeading.focus()
  await firstHeading.press('ArrowDown')
  await expect(nextHeading).toBeFocused()
  await nextHeading.press('ArrowUp')
  await expect(firstHeading).toBeFocused()
  await firstHeading.press('Enter')
  await expect(firstHeading).toHaveAttribute('aria-current', 'location')
  await visibleRibbonButton(page, '文档大纲').click()
  await expect(outline).toHaveCount(0)
  await expect.poll(async () => (await models(page)).document.blocks.map((block) => block.id)).toEqual(['heading-a', 'paragraph-a', 'heading-b', 'paragraph-b'])
})

test('Chinese Canvas composition is a single reversible document edit', async ({ page }) => {
  await switchWorkspace(page, '文档')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.evaluate((element) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '中文输入' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文输入' }))
  })
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('中文输入')
  await bridge.press('Control+Z')
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('')
  await bridge.press('Control+Y')
  await expect.poll(async () => (await models(page)).document.blocks.flatMap((block) => block.runs).map((run) => run.text).join('')).toBe('中文输入')
})

test('Docs font selector applies to the retained Canvas text selection', async ({ page }) => {
  await switchWorkspace(page, '文档')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.pressSequentially('字体选区')
  await bridge.press('Control+A')
  const font = ribbon(page).getByRole('combobox', { name: '字体', exact: true })
  await font.fill('Georgia')
  await font.press('Enter')
  await expect.poll(async () => (await models(page)).document.blocks[0]!.runs.every((run) => run.fontFamily === 'Georgia')).toBe(true)
  await expect.poll(async () => (await models(page)).document.blocks[0]!.runs.map((run) => run.text).join('')).toBe('字体选区')
})

test('switching products restores the worksheet viewport and a usable Canvas surface', async ({ page }) => {
  await gotoCell(page, 'T45')
  const scroller = page.locator('.excel-workbench .canvas-scroll')
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(300)
  await expect.poll(() => scroller.evaluate((element) => element.scrollLeft)).toBeGreaterThan(300)
  const position = await scroller.evaluate((element) => ({ top: element.scrollTop, left: element.scrollLeft }))
  await switchWorkspace(page, '文档')
  await expect(page.locator('.document-workbench canvas[data-render-surface="true"]')).toHaveCount(1)
  await switchWorkspace(page, '表格')
  await expect(page.getByRole('textbox', { name: '名称框', exact: true })).toHaveValue('T45')
  await expect.poll(() => scroller.evaluate((element) => ({ top: element.scrollTop, left: element.scrollLeft }))).toEqual(position)
  const surface = page.locator('.excel-workbench canvas[data-render-surface="true"]')
  await expect.poll(() => surface.evaluate((canvas: HTMLCanvasElement) => canvas.width > 0 && canvas.height > 0)).toBe(true)
  await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill('恢复后编辑')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(() => cellValue(page, 'T45')).toBe('恢复后编辑')
})

test('saving an edited document downloads a real DOCX with the edited text', async ({ page }) => {
  await switchWorkspace(page, '文档')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.pressSequentially('原生文件写出验证')
  await expect.poll(async () => (await models(page)).document.modified).toBe(true)
  const downloadPromise = page.waitForEvent('download')
  await page.getByRole('button', { name: '保存 Office 文件', exact: true }).click()
  const download = await downloadPromise
  expect(download.suggestedFilename()).toMatch(/\.docx$/i)
  const path = await download.path()
  expect(path).not.toBeNull()
  const bytes = await readFile(path!)
  expect(bytes.subarray(0, 4).toString('hex')).toBe('504b0304')
  const documentXml = readOfficePart(bytes, 'word/document.xml')
  expect(documentXml.replace(/<[^>]+>/g, '')).toContain('原生文件写出验证')
  expect(readOfficePart(bytes, '[Content_Types].xml')).toContain('wordprocessingml.document.main+xml')
  await expect.poll(async () => (await models(page)).document.modified).toBe(false)
})

test('the application file menu creates documents and the shared rename command is undoable', async ({ page }) => {
  await page.getByRole('button', { name: '文件菜单', exact: true }).click()
  await page.getByRole('menuitem', { name: '新建文档', exact: true }).click()
  await expect(page.locator('.document-workbench')).toBeVisible()
  await page.getByRole('button', { name: '重命名当前文件', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '重命名文件', exact: true })
  const name = dialog.getByRole('textbox', { name: '文件名称', exact: true })
  await name.fill('无效/名称')
  await dialog.getByRole('button', { name: /^确\s*定$/ }).click()
  await expect(dialog.getByRole('alert')).toBeVisible()
  await name.fill('布局闭环文档.docx')
  await name.press('Enter')
  await expect(dialog).toBeHidden()
  await expect(page.getByRole('button', { name: '重命名当前文件', exact: true })).toHaveText('布局闭环文档.docx')
  await expect.poll(async () => (await models(page)).document.name).toBe('布局闭环文档.docx')
  await visibleRibbonButton(page, '撤销').click()
  await expect.poll(async () => (await models(page)).document.name).toBe('未命名文档')
  await visibleRibbonButton(page, '重做').click()
  await expect.poll(async () => (await models(page)).document.name).toBe('布局闭环文档.docx')
  await expect(page).toHaveTitle('布局闭环文档.docx · Office Studio')
})

test('header save commits a pending formula session to a genuine XLSX workbook', async ({ page }) => {
  await gotoCell(page, 'B3')
  await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill('17')
  const downloadPromise = page.waitForEvent('download')
  await page.getByRole('button', { name: '保存 Office 文件', exact: true }).click()
  const download = await downloadPromise
  expect(download.suggestedFilename()).toMatch(/\.xlsx$/i)
  const path = await download.path()
  expect(path).not.toBeNull()
  const bytes = await readFile(path!)
  const worksheet = readOfficePart(bytes, 'xl/worksheets/sheet1.xml')
  expect(worksheet).toMatch(/<c\b[^>]*r="B3"[^>]*>[\s\S]*?<v>17<\/v>/)
  expect(readOfficePart(bytes, '[Content_Types].xml')).toContain('spreadsheetml.sheet.main+xml')
  await expect.poll(() => cellValue(page, 'B3')).toBe(17)
  await expect.poll(async () => (await models(page)).workbook.modified).toBe(false)
})

test('a failed Office import reports its error and retains the editable workbook', async ({ page }) => {
  await gotoCell(page, 'D2')
  await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill('导入失败仍保留')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  const workbookId = (await models(page)).workbook.id
  await page.getByLabel('选择 Office 文件', { exact: true }).setInputFiles({ name: 'empty.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.alloc(0) })
  await expect(page.locator('.parser-error')).toBeVisible()
  await expect(page.locator('.parser-error')).toContainText(/空|为空|解析|读取|损坏/)
  await expect.poll(async () => (await models(page)).workbook.id).toBe(workbookId)
  await expect.poll(() => cellValue(page, 'D2')).toBe('导入失败仍保留')
  await page.getByRole('button', { name: '关闭错误提示', exact: true }).click()
  await expect(page.locator('.parser-error')).toHaveCount(0)
  await gotoCell(page, 'E2')
  await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill('仍可继续编辑')
  await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  await expect.poll(() => cellValue(page, 'E2')).toBe('仍可继续编辑')
})

test('the name box navigates actual ranges and the footer statistics follow the selection', async ({ page }) => {
  for (const [address, value] of [['B2', '10'], ['B3', '20']]) {
    await gotoCell(page, address!)
    await page.getByRole('textbox', { name: '公式编辑', exact: true }).fill(value!)
    await page.getByRole('button', { name: '确认编辑', exact: true }).click()
  }
  await gotoCell(page, 'B2:B3')
  const statistic = page.getByRole('button', { name: '选区统计', exact: true })
  await expect(statistic).toHaveAttribute('title', '求和: 30')
  await statistic.click()
  await page.getByRole('menuitem', { name: '计数: 2', exact: true }).click()
  await expect(statistic).toHaveAttribute('title', '计数: 2')
  const nameBox = page.getByRole('textbox', { name: '名称框', exact: true })
  await nameBox.fill('A0')
  await nameBox.press('Enter')
  await expect(nameBox).toHaveAttribute('aria-invalid', 'true')
  await nameBox.press('Escape')
  await expect(nameBox).toHaveValue('B2:B3')
  await gotoCell(page, 'D4')
  await expect(statistic).toHaveCount(0)
})

test('sheet names validate duplicates and hidden sheets can be restored through the footer menu', async ({ page }) => {
  const originalName = (await models(page)).workbook.sheets[0]!.name
  await page.getByRole('button', { name: '添加工作表', exact: true }).click()
  await page.getByRole('button', { name: '工作表操作', exact: true }).click()
  await page.getByRole('menuitem', { name: '重命名', exact: true }).click()
  const dialog = page.getByRole('dialog', { name: '重命名工作表', exact: true })
  const field = dialog.getByRole('textbox', { name: '工作表名称', exact: true })
  await field.fill(originalName)
  await field.press('Enter')
  await expect(dialog.getByRole('alert')).toContainText(/已存在|重复/)
  await field.fill('恢复测试表')
  await field.press('Enter')
  await expect(dialog).toBeHidden()
  const tab = page.getByRole('tablist', { name: '工作表', exact: true }).getByRole('tab', { name: '恢复测试表', exact: true })
  await expect(tab).toHaveAttribute('aria-selected', 'true')
  const id = (await models(page)).workbook.activeSheetId
  await page.getByRole('button', { name: '工作表操作', exact: true }).click()
  await page.getByRole('menuitem', { name: '隐藏工作表', exact: true }).click()
  await expect(tab).toHaveCount(0)
  await expect.poll(async () => (await models(page)).workbook.sheets.find((sheet) => sheet.id === id)?.visible).toBe(false)
  await page.getByRole('button', { name: '所有工作表', exact: true }).click()
  await page.getByRole('menuitem', { name: '恢复测试表（隐藏，点击显示）', exact: true }).click()
  await expect(tab).toHaveAttribute('aria-selected', 'true')
  await expect.poll(async () => (await models(page)).workbook.sheets.find((sheet) => sheet.id === id)?.visible).toBe(true)
  await visibleRibbonButton(page, '撤销').click()
  await expect(tab).toHaveCount(0)
  await visibleRibbonButton(page, '重做').click()
  await expect(tab).toHaveAttribute('aria-selected', 'true')
})

test('editing a heading through Canvas updates its outline label and undo restores it', async ({ page }) => {
  await switchWorkspace(page, '文档')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.pressSequentially('原标题')
  await bridge.press('Control+A')
  await ribbon(page).getByRole('combobox', { name: '段落样式', exact: true }).click()
  await page.getByRole('option', { name: '标题 1', exact: true }).click()
  await expect.poll(async () => (await models(page)).document.blocks[0]!.type).toBe('heading')
  const blockId = (await models(page)).document.blocks[0]!.id
  await ribbon(page).getByRole('tab', { name: '视图', exact: true }).click()
  await visibleRibbonButton(page, '文档大纲').click()
  const heading = page.getByRole('navigation', { name: '文档大纲', exact: true }).locator(`[data-outline-item="${blockId}"]`)
  await expect(heading).toHaveText('原标题')
  await heading.click()
  await bridge.press('Control+A')
  await bridge.evaluate((element) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
    element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '更新后的标题' }))
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '更新后的标题' }))
  })
  await expect(heading).toHaveText('更新后的标题')
  await bridge.press('Control+Z')
  await expect(heading).toHaveText('原标题')
  await bridge.press('Control+Y')
  await expect(heading).toHaveText('更新后的标题')
})

interface WorkspaceHistoryDraftProbe {
  unitId: string
  cancel: boolean
  navigate: boolean
  before: number
  completed: number
  dispose: () => void
}
type WorkspaceHistoryDraftWindow = Window & { workspaceHistoryDraft?: WorkspaceHistoryDraftProbe }

// Production App/Ribbon/shortcut drafts, 2026-11-22 Asia/Shanghai.
// These cases have not been run; all alignment validation remains deferred.
async function installWorkspaceHistoryDraft(page: Page, kind: 'workbook' | 'document', direction: 'undo' | 'redo') {
  await page.evaluate(async ({ kind, direction }) => {
    const runtimePath = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(runtimePath) as RuntimeModule
    const book = officeAPI.getActiveWorkbook(), doc = officeAPI.getActiveDocument(), range = book?.getActiveSheet()?.getRange('A1')
    if (!book || !doc || !range) throw new Error('History draft could not find both active models')
    range.setInputText('sheet saved')
    doc.insertText(0, 'doc saved')
    // Establish synthetic fixture savepoints without invoking a file writer.
    studio.markSaved('workbook', book.id)
    studio.markSaved('document', doc.id)
    range.setInputText('sheet changed')
    doc.insertText(3, ' changed')
    studio.setActiveKind(kind)
    const unitId = kind === 'workbook' ? book.id : doc.id
    if (direction === 'redo') studio.undo(kind, unitId)
    if (kind === 'document') {
      const blockId = doc.getModel().blocks[0]?.id
      if (!blockId) throw new Error('History draft has no text paragraph')
      const selection = { anchor: { blockId, offset: 1 }, focus: { blockId, offset: 1 } }
      doc.setInputSelection(selection)
      doc.executeCommand({ id: 'doc.command.format-text', patch: { fontSize: 27, color: undefined } }, selection)
    }
    const probe: WorkspaceHistoryDraftProbe = { unitId, cancel: true, navigate: false, before: 0, completed: 0, dispose: () => {} }
    const before = officeAPI.addEvent(direction === 'undo' ? officeAPI.Event.BeforeUndo : officeAPI.Event.BeforeRedo, event => {
      if (event.unitId !== unitId) return
      probe.before++
      if (probe.cancel) event.cancel = true
      else if (probe.navigate) studio.setActiveKind(kind === 'workbook' ? 'document' : 'workbook')
    })
    const completed = officeAPI.addEvent(direction === 'undo' ? officeAPI.Event.Undo : officeAPI.Event.Redo, event => {
      if (event.unitId === unitId) probe.completed++
    })
    probe.dispose = () => { before.dispose(); completed.dispose() }
    ;(window as WorkspaceHistoryDraftWindow).workspaceHistoryDraft = probe
  }, { kind, direction })
  await expect(page.locator(kind === 'workbook' ? '.excel-workbench' : '.document-workbench')).toBeVisible()
}

async function workspaceHistoryDraftSnapshot(page: Page) {
  return page.evaluate(async () => {
    const runtimePath = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, officeAPI } = await import(runtimePath) as RuntimeModule
    const workbook = studio.getWorkbookModel(studio.activeWorkbookId), documentModel = studio.getDocumentModel(studio.activeDocumentId)
    const probe = (window as WorkspaceHistoryDraftWindow).workspaceHistoryDraft
    if (!workbook?.id || !documentModel?.id || !probe) throw new Error('History draft lost its models or observations')
    const input = studio.getDocumentInputState(documentModel.id)
    const histories = [workbook.id, documentModel.id].map(unitId => {
      const unit = studio.editor.history.units[unitId]
      if (!unit) throw new Error('History draft lost a unit stack')
      return { unitId, undo: unit.undo.length, redo: unit.redo.length, revision: unit.revision,
        generation: unit.generation, savedRevision: unit.savedRevision, dirty: studio.editor.history.isDirty(unitId) }
    })
    const state: { workbook: WorkbookModel; document: DocumentModel; histories: typeof histories;
      input: { state?: typeof input; clearedColor: boolean } } = JSON.parse(JSON.stringify({ workbook, document: documentModel, histories,
        input: { state: input, clearedColor: Boolean(input && Object.hasOwn(input.marks, 'color') && input.marks.color === undefined) } }))
    return {
      state,
      activeKind: studio.editor.state.activeKind, before: probe.before, completed: probe.completed,
      workbookText: officeAPI.getActiveWorkbook()?.getActiveSheet()?.getRange('A1').getValue(),
      documentText: officeAPI.getActiveDocument()?.getText(),
    }
  })
}

async function invokeWorkspaceHistoryDraft(page: Page, direction: 'undo' | 'redo', entry: 'ribbon' | 'global') {
  if (entry === 'ribbon') {
    const button = visibleRibbonButton(page, direction === 'undo' ? '撤销' : '重做')
    await expect(button).toBeEnabled()
    await button.click()
  } else {
    // Force the production shell shortcut path, outside native text editors.
    await page.evaluate(() => { document.body.tabIndex = -1; document.body.focus() })
    await page.keyboard.press(direction === 'undo' ? 'Control+z' : 'Control+y')
  }
}

for (const kind of ['workbook', 'document'] as const) for (const direction of ['undo', 'redo'] as const) for (const entry of ['ribbon', 'global'] as const) {
  test(`production ${kind} ${entry} ${direction} retains canceled content and a Before listener navigation`, async ({ page }) => {
    await installWorkspaceHistoryDraft(page, kind, direction)
    const before = await workspaceHistoryDraftSnapshot(page)
    await invokeWorkspaceHistoryDraft(page, direction, entry)
    await expect.poll(async () => (await workspaceHistoryDraftSnapshot(page)).before).toBe(1)
    const canceled = await workspaceHistoryDraftSnapshot(page)
    expect(canceled.state).toEqual(before.state)
    expect(canceled.activeKind).toBe(kind)
    expect(canceled.completed).toBe(0)

    await page.evaluate(() => {
      const probe = (window as WorkspaceHistoryDraftWindow).workspaceHistoryDraft
      if (!probe) throw new Error('History draft observations are unavailable')
      probe.cancel = false; probe.navigate = true
    })
    await invokeWorkspaceHistoryDraft(page, direction, entry)
    await expect.poll(async () => (await workspaceHistoryDraftSnapshot(page)).before).toBe(2)
    const navigated = await workspaceHistoryDraftSnapshot(page)
    expect(navigated.activeKind).toBe(kind === 'workbook' ? 'document' : 'workbook')
    expect({ ...navigated.state, input: undefined }).toEqual({ ...before.state, input: undefined })
    expect(navigated.completed).toBe(0)

    await page.evaluate(async kind => {
      const runtimePath = performance.getEntriesByType('resource').filter(entry => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
      const { studio } = await import(runtimePath) as RuntimeModule
      const probe = (window as WorkspaceHistoryDraftWindow).workspaceHistoryDraft
      if (!probe) throw new Error('History draft observations are unavailable')
      probe.navigate = false
      studio.setActiveKind(kind)
    }, kind)
    await expect(page.locator(kind === 'workbook' ? '.excel-workbench' : '.document-workbench')).toBeVisible()
    await invokeWorkspaceHistoryDraft(page, direction, entry)
    await expect.poll(async () => (await workspaceHistoryDraftSnapshot(page)).completed).toBe(1)
    const completed = await workspaceHistoryDraftSnapshot(page)
    expect(completed.before).toBe(3)
    expect(completed.activeKind).toBe(kind)
    expect(completed.workbookText).toBe(kind === 'workbook' && direction === 'undo' ? 'sheet saved' : 'sheet changed')
    expect(completed.documentText).toBe(kind === 'document' && direction === 'undo' ? 'doc saved' : 'doc changed saved')
    const targetIndex = kind === 'workbook' ? 0 : 1, neighborIndex = 1 - targetIndex
    expect(completed.state.histories[targetIndex].dirty).toBe(direction === 'redo')
    expect(completed.state.histories[neighborIndex]).toEqual(before.state.histories[neighborIndex])
    await page.evaluate(() => (window as WorkspaceHistoryDraftWindow).workspaceHistoryDraft?.dispose())
  })
}
