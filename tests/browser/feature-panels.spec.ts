import { expect, test, type Locator, type Page } from '@playwright/test'
import type { CellScalar, DocumentModel, WorkbookModel } from '../../packages/office-studio/src/domain/model'

type RuntimeModule = typeof import('../../examples/workbench/src/runtime')
const runtimeErrors = new WeakMap<Page, string[]>()

function ribbon(page: Page): Locator { return page.getByTestId('workbench-ribbon') }

/** antdv-next exposes virtual accessibility options separately from the visible popup rows. */
async function selectOption(page: Page, combobox: Locator, label: string): Promise<void> {
  await combobox.click()
  await page.locator('.ant-select-dropdown:visible').getByTitle(label, { exact: true }).click()
}

/** Observe the live model without accidentally creating a savepoint. */
async function state(page: Page): Promise<{ workbook: WorkbookModel; document: DocumentModel; pivotData: CellScalar[][]; documentText: string }> {
  return page.evaluate(async () => {
    const path = performance.getEntriesByType('resource').filter((entry) => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { studio, univerAPI } = await import(path) as RuntimeModule
    return JSON.parse(JSON.stringify({
      workbook: studio.getWorkbookModel(studio.activeWorkbookId),
      document: studio.getDocumentModel(studio.activeDocumentId),
      pivotData: univerAPI.getActiveWorkbook()?.getPivotTables()[0]?.getData() ?? [],
      documentText: univerAPI.getActiveDocument()?.getText() ?? '',
    }))
  })
}

async function seedPivot(page: Page, mode: 'region' | 'date'): Promise<void> {
  await page.evaluate(async (dataMode) => {
    const path = performance.getEntriesByType('resource').filter((entry) => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { univerAPI } = await import(path) as RuntimeModule
    const workbook = univerAPI.createWorkbook({ id: 'feature-panels-workbook', name: '面板交互验收', date1904: false })
    const source = workbook.getActiveSheet()!
    const values: CellScalar[][] = dataMode === 'region'
      ? [['地区', '销售额'], ['华东', 100], ['华南', 200], ['华北', 300], ['华东', 50]]
      : [['日期', '销售额'], [1, 100], [59, 200], [60, 300], [61, 400]]
    source.getRange('A1:B5').setValues(values)
    const target = workbook.insertSheet('数据分析')
    const pivot = await workbook.addPivotTable({
      unitId: workbook.getId(), subUnitId: source.getSheetId(), sheetName: source.getSheetName(),
      range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 },
    }, 'existing', { unitId: workbook.getId(), subUnitId: target.getSheetId(), row: 0, col: 0 })
    await pivot.addField('0', 'row')
    await pivot.addField('1', 'value')
    workbook.setActiveSheet(target)
  }, mode)
  await expect(page.getByRole('tab', { name: '数据分析', exact: true })).toHaveAttribute('aria-selected', 'true')
}

async function openSlicerPanel(page: Page, kind: '切片器' | '时间线'): Promise<Locator> {
  await ribbon(page).getByRole('tab', { name: '数据', exact: true }).click()
  await ribbon(page).getByRole('button', { name: kind, exact: true }).click()
  const panel = page.getByRole('dialog', { name: '切片器与时间线', exact: true })
  await expect(panel).toBeVisible()
  return panel
}

async function openControlPanel(page: Page): Promise<Locator> {
  await ribbon(page).getByRole('tab', { name: '插入', exact: true }).click()
  await ribbon(page).getByRole('button', { name: '内容控件', exact: true }).click()
  const panel = page.getByRole('dialog', { name: '内容控件', exact: true })
  await expect(panel).toBeVisible()
  return panel
}

async function closePanel(panel: Locator): Promise<void> {
  await panel.locator('.ant-modal-close').click()
  await expect(panel).toBeHidden()
}

async function history(page: Page, action: '撤销' | '重做'): Promise<void> {
  await ribbon(page).getByRole('tab', { name: '开始', exact: true }).click()
  await ribbon(page).getByRole('button', { name: action, exact: true }).click()
}

/** Explicit internal snapshot persistence; this does not stand in for Office export. */
async function reopenSnapshot(page: Page, kind: 'workbook' | 'document'): Promise<void> {
  await page.evaluate(async (unitKind) => {
    const path = performance.getEntriesByType('resource').filter((entry) => new URL(entry.name).pathname === '/src/runtime.ts').at(-1)?.name ?? '/src/runtime.ts'
    const { univerAPI } = await import(path) as RuntimeModule
    if (unitKind === 'workbook') {
      const original = univerAPI.getActiveWorkbook()!
      const snapshot = JSON.parse(JSON.stringify(original.save())) as WorkbookModel
      univerAPI.disposeUnit(original.getId())
      univerAPI.createWorkbook(snapshot)
    } else {
      const original = univerAPI.getActiveDocument()!
      const snapshot = JSON.parse(JSON.stringify(original.save())) as DocumentModel
      univerAPI.disposeUnit(original.getId())
      univerAPI.createDocument(snapshot)
    }
  }, kind)
}

async function addTextControl(page: Page, text: string): Promise<Locator> {
  await page.getByRole('navigation', { name: '工作台', exact: true }).getByRole('tab', { name: '文档', exact: true }).click()
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.pressSequentially(text)
  await bridge.press('Control+A')
  const panel = await openControlPanel(page)
  await panel.getByRole('button', { name: '添加控件', exact: true }).click()
  await panel.getByRole('textbox', { name: '控件标题', exact: true }).fill('客户字段')
  await panel.getByRole('button', { name: '插入到当前范围', exact: true }).click()
  await expect.poll(async () => (await state(page)).document.contentControls?.length).toBe(1)
  await expect(panel.getByRole('textbox', { name: '控件文本', exact: true })).toHaveValue(text)
  return panel
}

test.beforeEach(async ({ page }) => {
  const errors: string[] = []
  runtimeErrors.set(page, errors)
  page.on('pageerror', (error) => errors.push(error.message))
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布', exact: true })).toBeVisible()
})
test.afterEach(async ({ page }) => { expect(runtimeErrors.get(page) ?? []).toEqual([]) })

test('slicer panel creates a bound filter, distinguishes no matches from clearing and restores its snapshot', async ({ page }) => {
  await seedPivot(page, 'region')
  let panel = await openSlicerPanel(page, '切片器')
  await selectOption(page, panel.getByRole('combobox', { name: /切片器源字段/ }), '地区')
  await panel.getByRole('textbox', { name: '切片器标题', exact: true }).fill('地区选择')
  await panel.getByRole('button', { name: '插入切片器', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.length).toBe(1)
  await panel.getByRole('checkbox', { name: '华南', exact: true }).uncheck()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toEqual(['华东', '华北'])
  expect((await state(page)).pivotData.some((row) => row.includes('华南'))).toBe(false)
  await panel.getByRole('button', { name: '全不选', exact: true }).click()
  await panel.getByRole('button', { name: '应用筛选', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toEqual([])
  await panel.getByRole('button', { name: '清除筛选', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toBeUndefined()
  expect((await state(page)).pivotData.some((row) => row.includes('华南'))).toBe(true)
  await closePanel(panel)
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toEqual([])
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toEqual(['华东', '华北'])
  const selected = (await state(page)).workbook.slicers![0]!
  await reopenSnapshot(page, 'workbook')
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]).toEqual(selected)
  panel = await openSlicerPanel(page, '切片器')
  await panel.getByRole('button', { name: /地区选择.*切片器/ }).click()
  await expect(panel.getByRole('checkbox', { name: '华南', exact: true })).not.toBeChecked()
  await expect(panel.getByRole('checkbox', { name: '华东', exact: true })).toBeChecked()
})

test('timeline panel applies day ranges in the 1900 date system including Excel serial sixty', async ({ page }) => {
  await seedPivot(page, 'date')
  const panel = await openSlicerPanel(page, '时间线')
  await selectOption(page, panel.getByRole('combobox', { name: /切片器源字段/ }), '日期')
  await panel.getByRole('button', { name: '插入时间线', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.length).toBe(1)
  await selectOption(page, panel.getByRole('combobox', { name: /时间线粒度/ }), '日')
  await panel.getByRole('textbox', { name: '时间线开始日期', exact: true }).fill('1900-01-01')
  await panel.getByRole('textbox', { name: '时间线结束日期', exact: true }).fill('1900-01-01')
  await panel.getByRole('button', { name: '应用日期范围', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toEqual(['1'])
  await panel.getByRole('textbox', { name: '时间线开始日期', exact: true }).fill('1900-02-29')
  await panel.getByRole('textbox', { name: '时间线结束日期', exact: true }).fill('1900-02-29')
  await panel.getByRole('button', { name: '应用日期范围', exact: true }).click()
  await expect.poll(async () => (await state(page)).workbook.slicers?.[0]?.selectedItems).toEqual(['60'])
  const beforeInvalid = (await state(page)).workbook
  await panel.getByRole('textbox', { name: '时间线开始日期', exact: true }).fill('1900-02-30')
  await panel.getByRole('button', { name: '应用日期范围', exact: true }).click()
  await expect(panel.getByRole('alert')).toContainText('有效日期')
  expect((await state(page)).workbook).toEqual(beforeInvalid)
})

test('content control panel edits values, enforces content locks and preserves undo and internal persistence', async ({ page }) => {
  const panel = await addTextControl(page, '原客户')
  await panel.getByRole('textbox', { name: '控件文本', exact: true }).fill('新客户')
  await panel.getByRole('button', { name: '应用文本', exact: true }).click()
  await expect.poll(async () => (await state(page)).documentText).toBe('新客户')
  await selectOption(page, panel.getByRole('combobox', { name: /属性锁定状态/ }), '锁定内容')
  await panel.getByRole('button', { name: '保存属性', exact: true }).click()
  await expect(panel.getByRole('textbox', { name: '控件文本', exact: true })).toBeDisabled()
  const locked = (await state(page)).document
  await panel.getByRole('button', { name: '应用文本', exact: true }).click()
  await expect(page.locator('.document-notice')).toContainText('locked')
  expect((await state(page)).document).toEqual(locked)
  await closePanel(panel)
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).document.contentControls?.[0]?.lock).toBe('unlocked')
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).documentText).toBe('原客户')
  await history(page, '重做')
  await expect.poll(async () => (await state(page)).documentText).toBe('新客户')
  const savedControl = (await state(page)).document.contentControls![0]!
  await reopenSnapshot(page, 'document')
  await expect.poll(async () => (await state(page)).document.contentControls?.[0]).toEqual(savedControl)
  const reopened = await openControlPanel(page)
  await expect(reopened.getByRole('textbox', { name: '控件文本', exact: true })).toHaveValue('新客户')
})

test('Canvas paragraph splitting relocates a content control and undo restores its original anchor', async ({ page }) => {
  const panel = await addTextControl(page, '客户正文')
  const original = (await state(page)).document.contentControls![0]!
  await closePanel(panel)
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.press('Control+Home')
  await bridge.press('Enter')
  await expect.poll(async () => (await state(page)).document.blocks.length).toBe(2)
  const split = (await state(page)).document
  const body = split.blocks.find((block) => block.runs.map((run) => run.text).join('') === '客户正文')!
  expect(split.contentControls![0]).toMatchObject({ id: original.id, blockId: body.id, startOffset: 0, endOffset: 4 })
  expect(body.id).not.toBe(original.blockId)
  await history(page, '撤销')
  await expect.poll(async () => (await state(page)).document.contentControls?.[0]).toEqual(original)
  await expect.poll(async () => (await state(page)).document.blocks.length).toBe(1)
})
