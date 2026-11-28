import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'
import { COMMIT_SHEET_CHECKBOX_INPUT_COMMAND } from '../src/features/sheet-validation/validation-checkbox-input.ts'
import { SHEET_VALIDATION_DROPDOWN_OPERATIONS } from '../src/features/sheet-validation/validation-dropdown.ts'

// Implementation-phase browser drafts. No browser, Vite server or test runner has been started.
// Canvas2D is forced here; WebGPU, real Windows IME and Office file roundtrips need separate acceptance.
const route = '/__sheet-checkbox-feedback__', fixture = readFileSync(new URL('./fixtures/sheet-validation-cell.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-checkbox-feedback-fixture', configureServer(vite) {
      vite.middlewares.use(route, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(route, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen(); baseURL = `http://127.0.0.1:${server.httpServer.address().port}${route}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1600, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })

const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const checkbox = key => page.getByRole('checkbox', { name: `${key} 复选框`, exact: true })
const cellEditor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const feedback = (mode, key) => page.locator(`[data-sheet-validation-feedback="${mode}:${key}"]`)
const listPopup = key => page.getByRole('dialog', { name: `选择列表项 ${key}`, exact: true })
const listSearch = key => listPopup(key).getByRole('textbox', { name: '搜索列表项', exact: true })
const snapshot = () => page.evaluate(() => window.validationCellFixture.snapshot())
const history = () => page.evaluate(() => window.validationCellFixture.history())
const cell = async (row, column = 2) => (await snapshot()).sheets.find(sheet => sheet.id === 'data').cells[`${row}:${column}`]
const activeDropdown = () => page.evaluate(() => window.validationCellFixture.activeDropdown())
const showDropdown = (row, column) => page.evaluate(({ row, column }) => window.validationCellFixture.showDropdown(row, column), { row, column })
const hideDropdown = () => page.evaluate(() => window.validationCellFixture.hideDropdown())
async function select(key) { await nameBox().fill(key); await nameBox().press('Enter') }
async function appearance(variant, dark = false) { await page.evaluate(({ variant, dark }) => window.validationCellFixture.appearance(dark, variant), { variant, dark }) }
async function hoverCell(row, column) {
  const bounds = await grid().boundingBox()
  await page.mouse.move(bounds.x + 44 + column * 180 + 90, bounds.y + 26 + row * 32 + 16)
}

for (const variant of ['classic', 'grid']) {
  test(variant + ': multi-list edits grow the real row in one history entry and forced heights survive later edits and internal reopen', async () => {
    await appearance(variant); await select('B1')
    const before = await snapshot(), oldHistory = await history(), values = 'Red,"Green, dark",It""s blue'
    await page.evaluate(values => { window.validationCellFixture.setCell('B1', values) }, values)
    assert.equal((await cell(0, 1)).value, values)
    const tall = await page.evaluate(() => window.validationCellFixture.rowHeight(0))
    assert.ok(tall > 32); assert.equal(await page.evaluate(() => window.validationCellFixture.rowAutomatic(0)), true)
    assert.equal((await history()).undo, oldHistory.undo + 1)
    const after = await snapshot()
    await page.evaluate(() => window.validationCellFixture.undo()); assert.deepEqual(await snapshot(), before)
    await page.evaluate(() => window.validationCellFixture.redo()); assert.deepEqual(await snapshot(), after)
    await page.evaluate(() => window.validationCellFixture.reopen()); await select('B1')
    assert.equal(await page.evaluate(() => window.validationCellFixture.rowHeight(0)), tall)
    assert.equal((await cell(0, 1)).value, values)
    assert.equal(await page.evaluate(() => {
      const model = window.validationCellFixture.snapshot()
      return window.validationCellFixture.command('sheet.command.set-row-height', { unitId: model.id, sheetId: 'data',
        ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 29 }], value: 38 })
    }), true)
    await page.evaluate(() => { window.validationCellFixture.setCell('B1', 'Red') })
    assert.equal(await page.evaluate(() => window.validationCellFixture.rowHeight(0)), 38)
    assert.equal(await page.evaluate(() => window.validationCellFixture.rowAutomatic(0)), false)
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await cell(0, 1)).value, values)
    assert.equal(await page.evaluate(() => window.validationCellFixture.rowHeight(0)), 38)
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal(await page.evaluate(() => window.validationCellFixture.rowHeight(0)), tall)
    assert.equal(await page.evaluate(() => window.validationCellFixture.rowAutomatic(0)), true)
  })

  test(variant + ': automatic list column widths preserve CSV selections, repaint the Ant trigger and remain reversible after internal reopen', async () => {
    await appearance(variant, true); await select('B1')
    await page.evaluate(() => { window.validationCellFixture.setCell('B1', 'Red,"Green, dark",It""s blue') })
    assert.equal(await page.evaluate(() => {
      const model = window.validationCellFixture.snapshot()
      return window.validationCellFixture.command('sheet.command.set-worksheet-col-width', { unitId: model.id, sheetId: 'data',
        ranges: [{ startRow: 0, endRow: 159, startCol: 1, endCol: 1 }], value: 80 })
    }), true)
    const before = await snapshot(), oldHistory = await history()
    assert.equal(await page.evaluate(() => window.validationCellFixture.columnWidth(1)), 80)
    assert.equal(await page.evaluate(() => {
      const model = window.validationCellFixture.snapshot()
      return window.validationCellFixture.command('sheet.command.set-col-auto-width', { unitId: model.id, sheetId: 'data',
        ranges: [{ startRow: 0, endRow: 159, startCol: 1, endCol: 1 }] })
    }), true)
    const width = await page.evaluate(() => window.validationCellFixture.columnWidth(1))
    assert.ok(width > 80 && width < 500); assert.equal((await history()).undo, oldHistory.undo + 1)
    assert.deepEqual((await cell(0, 1)), before.sheets[0].cells['0:1'])
    await page.evaluate(() => window.validationCellFixture.undo()); assert.deepEqual(await snapshot(), before)
    await page.evaluate(() => window.validationCellFixture.redo()); assert.equal(await page.evaluate(() => window.validationCellFixture.columnWidth(1)), width)
    await page.evaluate(() => window.validationCellFixture.reopen()); await select('B1')
    assert.equal(await page.evaluate(() => window.validationCellFixture.columnWidth(1)), width)
    await page.getByRole('button', { name: '选择列表项 B1', exact: true }).click()
    for (const label of ['Red', 'Green, dark', 'It"s blue']) await expect(listPopup('B1').getByRole('checkbox', { name: label, exact: true })).toBeChecked()
    await listPopup('B1').getByRole('button', { name: '关闭', exact: true }).click()
    await expect(grid()).toBeFocused()
  })

  test.each(['list', 'listMultiple'])(`${variant}: formatted INDEX %s choices use Ant colors and regional text through selection history and internal reopen`, async type => {
    await appearance(variant, true); assert.equal(await page.evaluate(type => window.validationCellFixture.formatListSource(type), type), true)
    await select('B4'); const before = await snapshot(), oldHistory = await history()
    await page.getByRole('button', { name: '选择列表项 B4', exact: true }).click()
    await expect(listPopup('B4').getByRole('checkbox', { name: '25%', exact: true })).toBeChecked()
    await expect(listPopup('B4').locator('.ant-tag').filter({ hasText: '1234,50' })).toHaveCount(1)
    await listPopup('B4').getByRole('checkbox', { name: '1234,50', exact: true }).check()
    const changed = await cell(3, 1), expected = type === 'listMultiple' ? '25%,"1234,50"' : '1234,50'
    assert.equal(changed.value, expected); assert.equal(changed.dataType, 'text'); assert.equal((await history()).undo, oldHistory.undo + 1)
    assert.deepEqual(await cell(0, 3), before.sheets[0].cells['0:3']); assert.deepEqual(await cell(1, 3), before.sheets[0].cells['1:3'])
    if (type === 'listMultiple') await listPopup('B4').getByRole('button', { name: '关闭', exact: true }).click()
    await expect(listPopup('B4')).toBeHidden(); const after = await snapshot()
    await page.evaluate(() => window.validationCellFixture.undo()); assert.deepEqual(await snapshot(), before)
    await page.evaluate(() => window.validationCellFixture.redo()); assert.deepEqual(await snapshot(), after)
    await page.evaluate(() => window.validationCellFixture.reopen()); await select('B4')
    await page.getByRole('button', { name: '选择列表项 B4', exact: true }).click()
    await expect(listPopup('B4').getByRole('checkbox', { name: '1234,50', exact: true })).toBeChecked()
    assert.equal((await cell(1, 3)).value, 1234.5); assert.equal((await cell(3, 1)).value, expected)
  })

  test(`${variant}: real Worker and WASM initialize a bounded rectangle in one rule and scalar history with actual progress`, async () => {
    await appearance(variant); const before = await snapshot()
    assert.equal(await page.evaluate(() => {
      const model = window.validationCellFixture.snapshot()
      return window.validationCellFixture.commandAsync('sheet.command.set-data-validation-async', { unitId: model.id, sheetId: 'data',
        rule: { id: 'worker-checkbox', type: 'checkbox', ranges: [{ startRow: 0, endRow: 159, startCol: 0, endCol: 29 }] } })
    }), true)
    assert.equal((await cell(159,29)).value,0); assert.equal((await history()).undo,1)
    const progress = await page.evaluate(() => window.validationCellFixture.bulkProgress())
    assert.ok(progress.some(value => value.phase === 'computing' && value.compute.unit === 'cells'))
    assert.ok(progress.some(value => value.phase === 'materializing' && value.processed < value.total))
    const after = await snapshot(); await page.evaluate(() => window.validationCellFixture.undo()); assert.deepEqual(await snapshot(),before)
    await page.evaluate(() => window.validationCellFixture.redo()); assert.deepEqual(await snapshot(),after)
    await page.evaluate(() => window.validationCellFixture.reopen()); assert.equal((await cell(159,29)).value,0)
  }, { timeout:60_000 })

  test.each(['computing','materializing'])(`${variant}: cancellation during real bulk %s preserves every rule scalar and history entry`, async phase => {
    await appearance(variant); const before=await snapshot(), oldHistory=await history()
    const result=await page.evaluate(async phase => {
      const model=window.validationCellFixture.snapshot()
      try {
        await window.validationCellFixture.commandAsync('sheet.command.set-data-validation-async',{unitId:model.id,sheetId:'data',
          rule:{id:'cancelled-worker-checkbox',type:'checkbox',ranges:[{startRow:0,endRow:159,startCol:0,endCol:29}]}},phase)
        return 'unexpected success'
      } catch(cause) {return cause.code}
    },phase)
    assert.equal(result,'COMPUTE_CANCELLED'); assert.deepEqual(await snapshot(),before); assert.deepEqual(await history(),oldHistory)
  }, { timeout:60_000 })

  test(`${variant}: adding a checkbox rule initializes its real blank cells and undo restores both the rule and missing cells`, async () => {
    await appearance(variant); await select('G5:G6'); const before = await snapshot()
    assert.equal(await page.evaluate(() => {
      const model = window.validationCellFixture.snapshot()
      return window.validationCellFixture.command('sheet.command.set-data-validation', { unitId: model.id, sheetId: 'data',
        rule: { id: 'new-checkbox', type: 'checkbox', ranges: [{ startRow: 4, endRow: 5, startCol: 6, endCol: 6 }] } })
    }), true)
    await expect(checkbox('G5')).not.toBeChecked(); await expect(checkbox('G6')).not.toBeChecked()
    assert.equal((await cell(4, 6)).value, 0); assert.equal((await cell(5, 6)).value, 0); assert.equal((await history()).undo, 1)
    await page.evaluate(() => window.validationCellFixture.undo()); await expect(checkbox('G5')).toHaveCount(0); await expect(checkbox('G6')).toHaveCount(0)
    assert.deepEqual(await snapshot(), before)
    await page.evaluate(() => window.validationCellFixture.redo()); await expect(checkbox('G5')).not.toBeChecked(); assert.equal((await cell(4, 6)).value, 0)
  })

  test(`${variant}: changing custom operands migrates the scalar in the rule transaction and a later gesture has its own reversible history`, async () => {
    await appearance(variant); await select('C1'); const before = await snapshot()
    assert.equal(await page.evaluate(() => {
      const model = window.validationCellFixture.snapshot(), rule = model.sheets.find(sheet => sheet.id === 'data').dataValidations.find(rule => rule.id === 'checkbox-default')
      return window.validationCellFixture.command('sheet.command.replace-data-validation', { unitId: model.id, sheetId: 'data', rule: { ...rule, formula1: 'Approved', formula2: 'Pending' } })
    }), true)
    assert.equal((await cell(0)).value, 'Pending'); assert.equal((await history()).undo, 1); await expect(checkbox('C1')).not.toBeChecked()
    await checkbox('C1').check(); assert.equal((await cell(0)).value, 'Approved'); assert.equal((await history()).undo, 2)
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await cell(0)).value, 'Pending')
    await page.evaluate(() => window.validationCellFixture.undo()); assert.deepEqual(await snapshot(), before); await expect(checkbox('C1')).not.toBeChecked()
    await page.evaluate(() => { window.validationCellFixture.redo(); window.validationCellFixture.redo(); window.validationCellFixture.reopen() })
    await expect(grid()).toBeVisible(); assert.equal((await cell(0)).value, 'Approved'); await select('C1'); await expect(checkbox('C1')).toBeChecked()
  })

  test(`${variant}: the default Ant Checkbox edits one scalar, preserves other cells and styles, and survives undo/redo and internal reopen`, async () => {
    await appearance(variant); await select('C1'); await expect(checkbox('C1')).not.toBeChecked()
    const before = await snapshot(); await checkbox('C1').check(); await expect(grid()).toBeFocused()
    const after = await snapshot(), changed = await cell(0)
    assert.equal(changed.value, 1); assert.deepEqual(changed.style, before.sheets[0].cells['0:2'].style); assert.equal((await history()).undo, 1)
    const restored = JSON.parse(JSON.stringify(after)); restored.sheets[0].cells['0:2'] = before.sheets[0].cells['0:2']
    assert.deepEqual(restored, before)
    await page.evaluate(() => window.validationCellFixture.undo()); await expect(checkbox('C1')).not.toBeChecked(); assert.deepEqual(await snapshot(), before)
    await page.evaluate(() => window.validationCellFixture.redo()); await expect(checkbox('C1')).toBeChecked(); assert.deepEqual(await snapshot(), after)
    await page.evaluate(() => window.validationCellFixture.reopen()); await expect(grid()).toBeVisible(); assert.deepEqual(await cell(0), changed)
    await select('C1'); await expect(checkbox('C1')).toBeChecked()
  })

  test(`${variant}: custom text, boolean aliases and evaluated formulas retain their actual operand types when toggled`, async () => {
    await appearance(variant)
    await select('C2'); await checkbox('C2').check(); assert.equal((await cell(1)).value, 'Yes')
    await checkbox('C2').uncheck(); assert.equal((await cell(1)).value, 'No')
    await select('C5'); await expect(checkbox('C5')).not.toBeChecked(); await checkbox('C5').check(); assert.equal((await cell(4)).value, 'TRUE')
    await checkbox('C5').uncheck(); assert.equal((await cell(4)).value, 'FALSE')
    await select('C6'); await expect(checkbox('C6')).not.toBeChecked(); assert.equal((await cell(5)).formula, '=1-1')
    await checkbox('C6').check(); assert.equal((await cell(5)).value, 1); assert.equal((await cell(5)).formula, undefined)
    assert.equal((await history()).undo, 5)
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal((await cell(5)).formula, '=1-1'); await expect(checkbox('C6')).not.toBeChecked()
  })

  test(`${variant}: invalid existing content stays visible, while a required blank has an unchecked control and becomes valid only after editing`, async () => {
    await appearance(variant); const before = await snapshot()
    await select('C3'); await expect(checkbox('C3')).toHaveCount(0)
    await expect(grid().getByRole('gridcell')).toHaveText('C3 invalid checkbox'); await expect(grid().getByRole('gridcell')).toHaveAttribute('aria-invalid', 'true')
    assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
    await select('C4'); await expect(checkbox('C4')).not.toBeChecked(); await expect(grid().getByRole('gridcell')).toHaveAttribute('aria-invalid', 'true')
    await checkbox('C4').check(); assert.equal((await cell(3)).value, 1); await expect(grid().getByRole('gridcell')).not.toHaveAttribute('aria-invalid', 'true')
    await page.evaluate(() => window.validationCellFixture.undo()); assert.equal(await cell(3), undefined)
    await expect(grid().getByRole('gridcell')).toHaveAttribute('aria-invalid', 'true'); assert.equal((await history()).undo, 0)
  })

  test(`${variant}: Edit denial disables the Ant control and View denial hides the checkbox, scalar and feedback`, async () => {
    await appearance(variant); await select('C1'); await page.evaluate(() => window.validationCellFixture.protectCheckbox(false))
    await expect(checkbox('C1')).toBeDisabled(); const before = await snapshot(), oldHistory = await history()
    await grid().focus(); await grid().press('Space'); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await page.evaluate(() => window.validationCellFixture.protectCheckbox(true)); await expect(checkbox('C1')).toHaveCount(0)
    await expect(grid().getByRole('gridcell')).toHaveText('C1 '); await expect(feedback('prompt', 'C1')).toHaveCount(0)
    await hoverCell(0, 2); await expect(feedback('invalid', 'C1')).toHaveCount(0)
  })

  test(`${variant}: command cancellation keeps the checkbox and history, and Canvas Space resumes the same command after cancellation is removed`, async () => {
    await appearance(variant); await select('C1'); const before = await snapshot(), oldHistory = await history()
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, true), COMMIT_SHEET_CHECKBOX_INPUT_COMMAND)
    await checkbox('C1').check(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await expect(page.locator('.sheet-validation-checkbox-error')).toContainText('复选框操作已取消')
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, false), COMMIT_SHEET_CHECKBOX_INPUT_COMMAND)
    await grid().focus(); await grid().press('Space'); await expect(checkbox('C1')).toBeChecked(); assert.equal((await cell(0)).value, 1)
    assert.equal((await history()).undo, 1); await expect(page.locator('.sheet-validation-checkbox-error')).toHaveCount(0)
  })

  test(`${variant}: held pointer and Space callbacks cannot overwrite a newer cell value`, async () => {
    await appearance(variant); await select('C1')
    const bounds = await checkbox('C1').boundingBox(); await page.mouse.move(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2)
    await page.mouse.down(); await page.evaluate(() => window.validationCellFixture.setCell('C1', 1))
    const beforePointerRelease = await snapshot(), oldHistory = await history(); await page.mouse.up()
    assert.deepEqual(await snapshot(), beforePointerRelease); assert.deepEqual(await history(), oldHistory)
    await checkbox('C1').focus(); await page.keyboard.down('Space'); await page.evaluate(() => window.validationCellFixture.setCell('C1', 0))
    const beforeKeyRelease = await snapshot(), keyHistory = await history(); await page.keyboard.up('Space')
    assert.deepEqual(await snapshot(), beforeKeyRelease); assert.deepEqual(await history(), keyHistory)
  })

  test(`${variant}: an active Ant cell input and synthetic composition disable the checkbox without committing the input draft`, async () => {
    await appearance(variant); await select('C1'); await grid().focus(); await grid().press('F2')
    await cellEditor().fill('uncommitted'); await expect(checkbox('C1')).toBeDisabled()
    await cellEditor().dispatchEvent('compositionstart', { data: '未' }); await expect(checkbox('C1')).toBeDisabled()
    await cellEditor().press('Space'); assert.equal((await cell(0)).value, 0); assert.equal((await history()).undo, 0)
    await cellEditor().dispatchEvent('compositionend', { data: '未' }); await cellEditor().press('Escape')
    await expect(cellEditor()).toHaveCount(0); await expect(checkbox('C1')).toBeEnabled(); await expect(checkbox('C1')).not.toBeChecked()
    assert.equal((await history()).undo, 0)
  })

  test(`${variant}: a merged checkbox crossing both frozen boundaries remains one complete Ant control through zoom and scrolling`, async () => {
    await appearance(variant); await page.evaluate(() => window.validationCellFixture.freezeCheckbox()); await select('D2')
    await expect(checkbox('C1')).toHaveCount(1); await expect(page.locator('[data-sheet-checkbox-cell="C1"]')).toHaveCount(1)
    for (let index = 0; index < 4; index++) await page.getByRole('button', { name: '放大工作表', exact: true }).click()
    const coverage = await page.locator('[data-sheet-checkbox-cell="C1"] .ant-checkbox-inner').evaluate(element => {
      const bounds = element.getBoundingClientRect(), points = [[0.2, 0.2], [0.8, 0.2], [0.2, 0.8], [0.8, 0.8]]
      return { width: bounds.width, height: bounds.height, hits: points.map(([x, y]) =>
        document.elementsFromPoint(bounds.left + bounds.width * x, bounds.top + bounds.height * y)
          .some(target => target.closest('[data-sheet-checkbox-cell="C1"]'))) }
    })
    assert.ok(coverage.width > 16); assert.equal(coverage.width, coverage.height); assert.deepEqual(coverage.hits, [true, true, true, true])
    await checkbox('C1').check(); assert.equal((await cell(0)).value, 1)
    await page.locator('.canvas-scroll').evaluate(element => { element.scrollLeft = 80; element.scrollTop = 16; element.dispatchEvent(new Event('scroll')) })
    await expect(checkbox('C1')).toHaveCount(1)
    const host = await grid().boundingBox(), clip = await page.locator('[data-sheet-checkbox-cell="C1"]').boundingBox()
    assert.ok(clip.x >= host.x && clip.y >= host.y && clip.x + clip.width <= host.x + host.width && clip.y + clip.height <= host.y + host.height)
  })

  test(`${variant}: valid cells show their input prompt, dismissal lasts for the selected owner, and changing cells restores it`, async () => {
    await appearance(variant); await select('E2'); await expect(feedback('prompt', 'E2')).toContainText('Quantity')
    await expect(feedback('prompt', 'E2')).toContainText('Enter an integer from 1 through 5')
    await expect(grid().getByRole('gridcell')).not.toHaveAttribute('aria-invalid', 'true')
    await feedback('prompt', 'E2').getByRole('button', { name: '关闭输入提示', exact: true }).click(); await expect(feedback('prompt', 'E2')).toHaveCount(0)
    await page.evaluate(() => window.validationCellFixture.setCell('G10', 'Other')); await expect(feedback('prompt', 'E2')).toHaveCount(0)
    await select('G1'); await select('E2'); await expect(feedback('prompt', 'E2')).toBeVisible(); assert.equal((await cell(1, 4)).value, 2)
    assert.equal((await history()).undo, 1)
  })

  test(`${variant}: invalid hover waits 100ms, shows the rule error without replacing content, and clears on scroll`, async () => {
    await appearance(variant); await select('E1'); await expect(grid().getByRole('gridcell')).toHaveAttribute('aria-invalid', 'true')
    await expect(grid().getByRole('gridcell')).toHaveText('E1 1.5'); const before = await snapshot(), oldHistory = await history()
    await page.clock.install(); await page.clock.pauseAt(Date.now()); await hoverCell(0, 4)
    await page.clock.runFor(99); assert.equal(await feedback('invalid', 'E1').count(), 0)
    await page.clock.runFor(1); await expect(feedback('invalid', 'E1')).toContainText('Invalid quantity')
    await expect(feedback('invalid', 'E1')).toContainText('Use an integer from 1 through 5'); await expect(feedback('prompt', 'E1')).toHaveCount(0)
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await page.locator('.canvas-scroll').evaluate(element => { element.scrollTop = 64; element.dispatchEvent(new Event('scroll')) })
    await expect(feedback('invalid', 'E1')).toHaveCount(0)
  })

  test(`${variant}: synthetic composition in the owned cell input blocks prompt dismissal until composition ends`, async () => {
    await appearance(variant); await select('E2'); await grid().focus(); await grid().press('F2')
    await cellEditor().dispatchEvent('compositionstart', { data: '二' })
    await expect(feedback('prompt', 'E2').getByRole('button', { name: '关闭输入提示', exact: true })).toBeDisabled()
    await feedback('prompt', 'E2').dispatchEvent('keydown', { key: 'Escape', isComposing: true }); await expect(feedback('prompt', 'E2')).toBeVisible()
    await cellEditor().dispatchEvent('compositionend', { data: '二' }); await cellEditor().press('Escape')
    await feedback('prompt', 'E2').getByRole('button', { name: '关闭输入提示', exact: true }).click(); await expect(feedback('prompt', 'E2')).toHaveCount(0)
    assert.equal((await history()).undo, 0); assert.equal((await cell(1, 4)).value, 2)
  })

  test(`${variant}: Ant themes update checkbox and feedback colors while their command and prompt behavior remains usable`, async () => {
    await appearance(variant); await select('C1'); await checkbox('C1').check()
    const lightCheckbox = await page.locator('[data-sheet-checkbox-cell="C1"] .ant-checkbox-inner').evaluate(element => getComputedStyle(element).backgroundColor)
    await select('E2'); const lightPrompt = await feedback('prompt', 'E2').locator('.ant-alert').evaluate(element => getComputedStyle(element).backgroundColor)
    await appearance(variant, true)
    await expect(page.locator('[data-sheet-checkbox-cell="C1"] .ant-checkbox-inner')).not.toHaveCSS('background-color', lightCheckbox)
    await expect(feedback('prompt', 'E2').locator('.ant-alert')).not.toHaveCSS('background-color', lightPrompt)
    await feedback('prompt', 'E2').getByRole('button', { name: '关闭输入提示', exact: true }).click(); await select('C1'); await checkbox('C1').uncheck()
    assert.equal((await cell(0)).value, 0); assert.equal((await history()).undo, 2)
  })

  test(`${variant}: public show opens a hidden-arrow dynamic list without moving selection and idempotent show preserves the search draft`, async () => {
    await appearance(variant); await select('A1'); const before = await snapshot(), oldHistory = await history()
    assert.equal(await showDropdown(0, 5), true); await expect(listPopup('F1')).toBeVisible()
    await expect(page.getByRole('button', { name: '选择列表项 F1', exact: true })).toHaveCount(0); await expect(nameBox()).toHaveValue('A1')
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory)
    await listSearch('F1').fill('Green'); const owner = await listPopup('F1').getAttribute('data-sheet-validation-popup-owner')
    assert.equal(await showDropdown(0, 5), true); await expect(listSearch('F1')).toHaveValue('Green')
    assert.equal(await listPopup('F1').getAttribute('data-sheet-validation-popup-owner'), owner)
    await page.evaluate(() => window.validationCellFixture.setCell('F2', false)); await expect(listPopup('F1').getByRole('checkbox')).toHaveCount(0)
    await listSearch('F1').fill('Blue'); await listPopup('F1').getByRole('checkbox', { name: 'Blue', exact: true }).check()
    assert.equal((await cell(0, 5)).value, 'Blue'); assert.equal((await history()).undo, 2); await expect(nameBox()).toHaveValue('A1')
    assert.equal(await hideDropdown(), true); await expect(listPopup('F1')).toBeHidden(); assert.equal(await activeDropdown(), null)
    assert.equal(await hideDropdown(), true); assert.equal((await history()).undo, 2)
  })

  test(`${variant}: public show/hide respects popup composition and command cancellation, then invalidates old popup callbacks`, async () => {
    await appearance(variant); await select('A1'); assert.equal(await showDropdown(0, 5), true)
    await listSearch('F1').fill('Green'); await listSearch('F1').dispatchEvent('compositionstart', { data: '绿' })
    assert.equal(await showDropdown(0, 1), false); assert.equal(await hideDropdown(), false); await expect(listPopup('F1')).toBeVisible()
    await expect(listSearch('F1')).toHaveValue('Green'); await listSearch('F1').dispatchEvent('compositionend', { data: '绿' })
    const oldCheckbox = await listPopup('F1').getByRole('checkbox', { name: 'Green, dark', exact: true }).elementHandle()
    const oldHistory = await history(), before = await snapshot()
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, true), SHEET_VALIDATION_DROPDOWN_OPERATIONS.show)
    const canceled = await page.evaluate(() => { try { return window.validationCellFixture.showDropdown(0, 1) } catch (error) { return error.message } })
    assert.match(canceled, /^Command was canceled:/); await expect(listSearch('F1')).toHaveValue('Green')
    await page.evaluate(id => window.validationCellFixture.cancelCommand(id, false), SHEET_VALIDATION_DROPDOWN_OPERATIONS.show)
    assert.equal(await showDropdown(0, 1), true); await expect(listPopup('B1')).toBeVisible(); await expect(listPopup('F1')).toBeHidden()
    await oldCheckbox.evaluate(element => { element.checked = true; element.dispatchEvent(new Event('change', { bubbles: true })) })
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), oldHistory); assert.equal((await activeDropdown()).column, 1)
    await page.evaluate(() => window.validationCellFixture.replaceSameId()); await expect(listPopup('B1')).toBeHidden(); assert.equal(await activeDropdown(), null)
  })
}
