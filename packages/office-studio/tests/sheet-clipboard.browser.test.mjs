import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

// Draft only. Clipboard I/O is injected; native event/DOM paths and library controls are exercised when this is later run.
const fixturePath = '/__sheet-clipboard-test__'
const fixture = readFileSync(new URL('./fixtures/sheet-clipboard.html', import.meta.url), 'utf8')
let server, browser, page, baseURL, errors
beforeAll(async () => {
  server = await createServer({ configFile: false, root: process.cwd(), plugins: [vue(), {
    name: 'sheet-clipboard-fixture', configureServer(vite) {
      vite.middlewares.use(fixturePath, async (_req, res, next) => {
        try { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(await vite.transformIndexHtml(fixturePath, fixture)) }
        catch (error) { next(error) }
      })
    },
  }], server: { host: '127.0.0.1', port: 0 }, logLevel: 'error' })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = 'http://127.0.0.1:' + address.port + fixturePath
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, { timeout: 60_000 })
afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []; page = await browser.newPage({ viewport: { width: 1440, height: 900 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.addInitScript(() => Object.defineProperty(navigator, 'gpu', { configurable: true, value: undefined }))
  await page.goto(baseURL); await expect(grid()).toBeVisible()
}, { timeout: 60_000 })
afterEach(async () => { try { assert.deepEqual(errors, []) } finally { await page?.close() } })
const grid = () => page.getByRole('grid', { name: '工作表画布', exact: true })
const editor = () => page.locator('.canvas-cell-input input, .canvas-cell-input textarea')
const nameBox = () => page.getByRole('textbox', { name: '名称框', exact: true })
const pasteButton = () => page.locator('[data-ribbon-item="sheet.command.paste"]').getByRole('button', { name: '粘贴', exact: true }).first()
const optionsButton = () => page.locator('.sheet-paste-options')
const snapshot = () => page.evaluate(() => window.sheetClipboardFixture.snapshot())
const history = () => page.evaluate(() => window.sheetClipboardFixture.history())
const cell = async (row, col, sheet = 0) => (await snapshot()).sheets[sheet].cells[row + ':' + col]
async function select(address) { await nameBox().fill(address); await nameBox().press('Enter'); await expect(grid()).toBeFocused() }
async function dispatch(kind, content = {}) {
  return grid().evaluate((element, input) => {
    const data = new DataTransfer()
    if (input.content.text !== undefined) data.setData('text/plain', input.content.text)
    if (input.content.html !== undefined) data.setData('text/html', input.content.html)
    const event = new ClipboardEvent(input.kind, { bubbles: true, cancelable: true, clipboardData: data })
    element.dispatchEvent(event)
    return { prevented: event.defaultPrevented, text: data.getData('text/plain'), html: data.getData('text/html') }
  }, { kind, content })
}
async function copySelection(address = 'A1:B1', cut = false) {
  await select(address); return dispatch(cut ? 'cut' : 'copy')
}
async function optional(mode) {
  await optionsButton().click(); await page.locator('.sheet-paste-options-popup').getByRole('menuitem', { name: mode, exact: true }).click()
}

test('Canvas copy/cut events populate both flavours and a cut moves only when the paste succeeds', async () => {
  const original = await snapshot(), copied = await copySelection('A1:B1', true)
  assert.equal(copied.prevented, true); assert.equal(copied.text, 'rich\t5'); assert.match(copied.html, /data-office-clipboard-id/)
  assert.deepEqual((await snapshot()).sheets[0].cells, original.sheets[0].cells); assert.equal((await history()).undo, 0)
  await select('E5'); await dispatch('paste', copied)
  assert.equal(await cell(0, 0), undefined); assert.deepEqual((await cell(4, 4)).richText, original.sheets[0].cells['0:0'].richText)
  assert.equal((await history()).undo, 1); await grid().press('Control+z')
  assert.deepEqual((await cell(0, 0)).richText, original.sheets[0].cells['0:0'].richText); assert.equal(await cell(4, 4), undefined)
})

test('browser C and X shortcuts reach native events without losing the canvas selection', async () => {
  await select('A1:B1'); await grid().press('Control+c'); await grid().press('Control+x')
  const events = await page.evaluate(() => window.sheetClipboardFixture.nativeEvents())
  assert.ok(events.some(event => event.kind === 'copy' && event.prevented && event.text === 'rich\t5'))
  assert.ok(events.some(event => event.kind === 'cut' && event.prevented)); assert.equal((await history()).undo, 0)
})

test('the active non-last selection supplies clipboard data instead of the bounds of disjoint selections', async () => {
  await page.evaluate(() => window.sheetClipboardFixture.activeRangeList()); await grid().focus()
  const content = await dispatch('copy'); assert.equal(content.text, 'rich')
  const commands = await page.evaluate(() => window.sheetClipboardFixture.commands())
  const copied = commands.findLast(command => command.id === 'sheet.command.copy')
  assert.deepEqual(copied.params.range, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 })
})

test('recent paste uses the library menu to replace the original history redo across all six modes', async () => {
  const content = await copySelection(); await select('E5'); await dispatch('paste', content)
  await expect(optionsButton()).toBeVisible()
  for (const label of ['仅粘贴值', '仅粘贴公式', '仅粘贴格式', '仅粘贴列宽', '保留目标边框粘贴', '全部粘贴']) {
    await optional(label); assert.equal((await history()).undo, 1)
  }
  assert.equal((await cell(4, 5)).formula, '=E6+$A$2'); await grid().press('Control+z'); assert.equal(await cell(4, 4), undefined)
  await grid().press('Control+y'); assert.equal((await cell(4, 5)).formula, '=E6+$A$2')
})

test('optional values paste is disabled when restoring an old width would touch a protected non-target cell in that column', async () => {
  await page.evaluate(() => window.sheetClipboardFixture.columnWidthProtection())
  const copied = await copySelection(); await select('E5'); await dispatch('paste', copied)
  await optional('仅粘贴列宽')
  assert.equal(await page.evaluate(() => window.sheetClipboardFixture.columnWidth(4)), 123)
  assert.equal(await page.evaluate(() => window.sheetClipboardFixture.columnWidth(5)), 156)
  assert.equal(await cell(4, 4), undefined); assert.equal((await history()).undo, 1)
  await page.evaluate(() => window.sheetClipboardFixture.revokeColumnWidthPermission())
  const before = await snapshot(), depth = await history(), notices = await page.evaluate(() => window.sheetClipboardFixture.notifications().length)
  const commands = await page.evaluate(() => window.sheetClipboardFixture.commands().length)
  await optionsButton().click()
  const popup = page.locator('.sheet-paste-options-popup'), values = popup.getByRole('menuitem', { name: '仅粘贴值', exact: true })
  await expect(values).toHaveAttribute('aria-disabled', 'true')
  await values.getByText('仅粘贴值', { exact: true }).hover()
  await expect(page.getByRole('tooltip').filter({ hasText: /受保护/ })).toBeVisible()
  // Deliberately bypass pointer actionability: the actual library/menu dispatch guard must still reject it.
  await values.evaluate(element => element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true })))
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
  assert.equal(await page.evaluate(() => window.sheetClipboardFixture.notifications().length), notices)
  assert.equal(await page.evaluate(() => window.sheetClipboardFixture.commands().length), commands)
  await expect(optionsButton()).toBeVisible()
  await expect(popup.getByRole('menuitem', { name: '仅粘贴列宽', exact: true })).toHaveClass(/ant-menu-item-selected/)
  await page.keyboard.press('Escape'); await expect(grid()).toBeFocused()
})

test('a cross-sheet table cut moves the chart source while preserving its drawing host and optional values restores those resources', async () => {
  await page.evaluate(() => window.sheetClipboardFixture.tableAndChartSource())
  const original = await snapshot(), copied = await copySelection('A1:B2', true)
  const report = page.getByRole('tab', { name: 'Report', exact: true })
  await report.click(); await expect(report).toHaveAttribute('aria-selected', 'true')
  await select('E5'); await dispatch('paste', copied)
  const moved = await snapshot(), targetRange = { startRow: 4, endRow: 5, startCol: 4, endCol: 5 }
  assert.equal(moved.tables[0].sheetId, moved.sheets[1].id); assert.deepEqual(moved.tables[0].range, targetRange)
  assert.equal(moved.charts[0].source.sheetId, moved.sheets[1].id); assert.deepEqual(moved.charts[0].source.range, targetRange)
  assert.equal(moved.charts[0].sheetId, original.sheets[0].id); assert.deepEqual(moved.charts[0].position, original.charts[0].position)
  assert.deepEqual(moved.charts[0].size, original.charts[0].size); assert.equal(await cell(0, 0), undefined)
  assert.equal((await cell(4, 4, 1)).value, 'rich'); assert.equal((await history()).undo, 1)
  await optional('仅粘贴值')
  const restored = await snapshot()
  assert.deepEqual(restored.tables, original.tables); assert.deepEqual(restored.charts, original.charts)
  assert.equal((await cell(0, 0)).value, 'rich'); assert.equal((await cell(4, 4, 1)).value, 'rich')
  assert.equal((await cell(4, 5, 1)).formula, null); assert.equal((await history()).undo, 1)
  await grid().press('Control+z'); assert.deepEqual((await snapshot()).tables, original.tables); assert.deepEqual((await snapshot()).charts, original.charts)
  assert.equal(await cell(4, 4, 1), undefined); assert.equal((await cell(0, 0)).value, 'rich')
  await grid().press('Control+y'); assert.deepEqual((await snapshot()).charts, original.charts); assert.equal((await cell(4, 5, 1)).formula, null)
})

test('external HTML table styles and spans use the shared paste command while text paste keeps destination formatting', async () => {
  await select('E5'); await dispatch('paste', { text: 'Rich\t', html: '<table><tr><td colspan="2" style="font-weight:bold;background-color:#ffeeaa"><span style="font-style:italic">Rich</span></td></tr></table>' })
  assert.equal((await cell(4, 4)).value, 'Rich'); assert.equal((await cell(4, 4)).style.font.bold, true)
  assert.ok((await snapshot()).sheets[0].mergedRanges.some(range => range.startRow === 4 && range.startCol === 4 && range.endCol === 5))
  await select('E5'); await dispatch('paste', { text: 'plain' })
  assert.equal((await cell(4, 4)).value, 'plain'); assert.equal((await cell(4, 4)).style.font.bold, true)
})

test('cell and formula inputs keep their native copy/cut/paste scope and rich input can still insert multiline text', async () => {
  await select('A1'); await grid().press('F2'); await expect(editor()).toBeFocused()
  await editor().evaluate(element => {
    element.setSelectionRange(0, 2); const data = new DataTransfer(); data.setData('text/plain', 'x\ny')
    element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }))
  })
  await expect(editor()).toHaveValue('x\nych'); await editor().press('Enter')
  assert.equal((await cell(0, 0)).value, 'x\nych'); assert.equal((await history()).undo, 1)
  const before = (await page.evaluate(() => window.sheetClipboardFixture.commands())).filter(event => /sheet.command.(copy|cut|paste)/.test(event.id)).length
  await select('B1'); const formula = page.getByRole('textbox', { name: '公式编辑', exact: true })
  await formula.focus(); await expect(formula).toHaveValue('=A2+$A$2')
  const prevented = await formula.evaluate(element => {
    const flags = []
    element.setSelectionRange(1, 3)
    for (const kind of ['copy', 'cut', 'paste']) {
      const data = new DataTransfer(); data.setData('text/plain', 'A3')
      const event = new ClipboardEvent(kind, { bubbles: true, cancelable: true, clipboardData: data })
      element.dispatchEvent(event); flags.push(event.defaultPrevented)
    }
    // A synthetic paste has no browser default insertion; reproduce its native beforeinput/input pair.
    element.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertFromPaste', data: 'A3' }))
    element.setRangeText('A3', 1, 3, 'end')
    element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertFromPaste', data: 'A3' }))
    return flags
  })
  assert.deepEqual(prevented, [false, false, false]); await expect(formula).toHaveValue('=A3+$A$2'); await formula.press('Enter')
  assert.equal((await cell(0, 1)).formula, '=A3+$A$2'); assert.equal((await history()).undo, 2)
  await page.getByRole('textbox', { name: '外部文本', exact: true }).focus(); await page.keyboard.press('Control+c')
  const after = (await page.evaluate(() => window.sheetClipboardFixture.commands())).filter(event => /sheet.command.(copy|cut|paste)/.test(event.id)).length
  assert.equal(after, before)
})

test('async clipboard reads retain the initial target without taking back later selection or focus', async () => {
  await page.evaluate(() => { window.sheetClipboardFixture.setBoard({ text: 'fixed' }); window.sheetClipboardFixture.delay(true) })
  await select('E5'); await pasteButton().click(); await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.reads())).toBe(1)
  await select('H9'); const selected = await page.evaluate(() => window.sheetClipboardFixture.selection())
  await page.evaluate(() => window.sheetClipboardFixture.resolveRead(0, { text: 'fixed' }))
  await expect.poll(async () => (await cell(4, 4))?.value).toBe('fixed')
  assert.deepEqual(await page.evaluate(() => window.sheetClipboardFixture.selection()), selected); await expect(grid()).toBeFocused()
  assert.equal(await cell(8, 7), undefined); assert.equal((await history()).undo, 1)
})

test('an async paste keeps a new external Input focus with an unchanged selection in production and standalone hosts', async () => {
  for (const standalone of [false, true]) {
    if (standalone) { await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible() }
    await page.evaluate(() => window.sheetClipboardFixture.delay(true)); await select('E5'); await pasteButton().click()
    await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.reads())).toBe(1)
    const outside = page.getByRole('textbox', { name: '外部文本', exact: true })
    await outside.focus(); await page.evaluate(() => window.sheetClipboardFixture.resolveRead(0, { text: 'fixed target' }))
    await expect.poll(async () => (await cell(4, 4))?.value).toBe('fixed target')
    await expect(outside).toBeFocused()
    const actual = await history(); assert.equal(actual.hasStudio, !standalone)
    assert.equal(standalone ? actual.updates : actual.undo, 1)
  }
})

test('late reads after sheet, unit, same-ID reopen, studio replacement and new edits never publish the old candidate', async () => {
  for (const action of ['sheet', 'unit', 'same-id', 'studio', 'edit']) {
    await page.goto(baseURL); await expect(grid()).toBeVisible(); await select('E5')
    await page.evaluate(() => window.sheetClipboardFixture.delay(true)); await pasteButton().click()
    await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.reads())).toBe(1)
    await page.evaluate(value => {
      const fixture = window.sheetClipboardFixture
      if (value === 'sheet') fixture.switchSheet(1)
      else if (value === 'unit') fixture.switchUnit()
      else if (value === 'same-id') fixture.reopenSameId()
      else if (value === 'studio') return fixture.replaceHost()
      else fixture.edit()
    }, action)
    const before = await snapshot(), depth = await history()
    await page.evaluate(() => window.sheetClipboardFixture.resolveRead(0, { text: 'late' }))
    await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 0)))
    assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, depth.undo); assert.equal((await history()).updates, depth.updates)
  }
})

test('out-of-order readers and Escape cancellation respect the latest user request', async () => {
  await page.evaluate(() => window.sheetClipboardFixture.delay(true)); await select('E5'); await pasteButton().click()
  await select('H9'); await pasteButton().click(); await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.reads())).toBe(2)
  await page.evaluate(() => window.sheetClipboardFixture.resolveRead(1, { text: 'new request' }))
  await expect.poll(async () => (await cell(8, 7))?.value).toBe('new request')
  await page.evaluate(() => window.sheetClipboardFixture.resolveRead(0, { text: 'old request' })); assert.equal(await cell(4, 4), undefined)
  await select('E5'); await pasteButton().click(); await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.reads())).toBe(3)
  await grid().focus(); await grid().press('Escape'); await page.evaluate(() => window.sheetClipboardFixture.resolveRead(2, { text: 'canceled' }))
  assert.equal(await cell(4, 4), undefined); assert.equal((await history()).undo, 1)
})

test('empty and media-only clipboard reads retain copied content without pasting stale data', async () => {
  await copySelection(); await select('E5')
  for (const content of [{ text: '', html: '' }, { media: true }]) {
    const notices = await page.evaluate(() => window.sheetClipboardFixture.notifications().length)
    await page.evaluate(value => window.sheetClipboardFixture.setBoard(value), content); await pasteButton().click()
    await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.notifications().length)).toBeGreaterThan(notices)
    assert.equal(await cell(4, 4), undefined); assert.equal((await history()).undo, 0)
  }
})

test('command cancellation, stop validation and protection preserve cut source and content history', async () => {
  for (const kind of ['cancel', 'validation', 'protection']) {
    await page.goto(baseURL); await expect(grid()).toBeVisible()
    if (kind !== 'cancel') await page.evaluate(value => window.sheetClipboardFixture[value](), kind)
    const copied = await copySelection('A1:B1', true); await select('E5')
    if (kind === 'cancel') await page.evaluate(() => window.sheetClipboardFixture.cancel(true))
    const before = await snapshot(); await dispatch('paste', copied)
    assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0); assert.equal(await cell(4, 4), undefined)
    assert.ok((await page.evaluate(() => window.sheetClipboardFixture.notifications())).length)
  }
})

test('the paste overlay retains its theme and follows frozen and zoomed geometry with keyboard focus', async () => {
  const copied = await copySelection(); await select('E5'); await dispatch('paste', copied)
  await page.evaluate(() => window.sheetClipboardFixture.appearance(true, 'grid')); await expect(optionsButton()).toBeVisible()
  const darkBackground = await optionsButton().evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => window.sheetClipboardFixture.appearance(false, 'classic')); await expect(optionsButton()).toBeVisible()
  assert.notEqual(await optionsButton().evaluate(element => getComputedStyle(element).backgroundColor), darkBackground)
  await page.evaluate(() => window.sheetClipboardFixture.freeze()); const one = await copySelection('A2'); await select('A1'); await dispatch('paste', one)
  assert.equal((await cell(0, 0)).value, 5); await expect(optionsButton()).toBeVisible()
  const position = () => optionsButton().evaluate(element => ({ left: parseFloat(element.style.left), top: parseFloat(element.style.top) }))
  assert.deepEqual(await position(), { left: 171, top: 58 })
  for (let index = 0; index < 5; index++) await page.getByRole('button', { name: '放大工作表', exact: true }).click()
  await expect.poll(position).toEqual({ left: 232.5, top: 72 })
  await page.locator('.canvas-scroll').evaluate(element => { element.scrollLeft = 600; element.scrollTop = 500; element.dispatchEvent(new Event('scroll')) })
  assert.deepEqual(await position(), { left: 232.5, top: 72 })
  await optionsButton().focus(); await page.keyboard.press('ArrowDown'); await expect(page.locator('.sheet-paste-options-popup')).toBeVisible()
  await expect(page.locator('.sheet-paste-options-popup [role="menu"]')).toBeFocused()
  await page.keyboard.press('Escape'); await expect(page.locator('.sheet-paste-options-popup')).toBeHidden(); await expect(grid()).toBeFocused()
})

test('reopening the paste menu before its close callback runs keeps focus in the new menu', async () => {
  const copied = await copySelection(); await select('E5'); await dispatch('paste', copied)
  await optionsButton().focus(); await page.keyboard.press('ArrowDown')
  const popup = page.locator('.sheet-paste-options-popup'), menu = popup.locator('[role="menu"]')
  await expect(menu).toBeFocused()
  await menu.evaluate(element => {
    element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
    const trigger = document.querySelector('.sheet-paste-options')
    trigger.focus()
    trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true }))
  })
  await expect(popup).toBeVisible(); await expect(menu).toBeFocused()
  assert.equal((await cell(4, 4)).value, 'rich'); assert.equal((await history()).undo, 1)
})

test('moving focus to a separate Input closes the popup and prevents delayed menu focus from taking it back', async () => {
  const copied = await copySelection(); await select('E5'); await dispatch('paste', copied); await optionsButton().click()
  const outside = page.getByRole('textbox', { name: '外部文本', exact: true })
  await outside.focus(); await expect(page.locator('.sheet-paste-options-popup')).toBeHidden()
  await page.keyboard.press('Escape'); await expect(outside).toBeFocused()
  assert.equal((await cell(4, 4)).value, 'rich'); assert.equal((await history()).undo, 1)
})

test('standalone mode keeps its external content source and replaces optional paste without fabricating an OfficeStudio history', async () => {
  await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible()
  await page.evaluate(() => window.sheetClipboardFixture.standaloneSame())
  const copied = await copySelection('A1'); await select('E5'); await dispatch('paste', copied)
  assert.equal((await cell(4, 4)).style.font.bold, false); await optional('仅粘贴值')
  assert.equal((await cell(4, 4)).style.font.bold, true); assert.equal((await history()).hasStudio, false)
  assert.equal((await history()).updates, 2)
})

test('hidePasteOptions toggles the overlay while preserving recent content, selection, history and an external Input focus', async () => {
  for (const standalone of [false, true]) {
    if (standalone) { await page.goto(baseURL + '?standalone'); await expect(grid()).toBeVisible() }
    await page.evaluate(() => window.sheetClipboardFixture.hidePasteOptions(true))
    const copied = await copySelection(); await select('E5'); await dispatch('paste', copied)
    await expect.poll(async () => (await cell(4, 4))?.value).toBe('rich'); await expect(optionsButton()).toBeHidden()
    const before = await snapshot(), depth = await history(), selected = await page.evaluate(() => window.sheetClipboardFixture.selection())
    await page.evaluate(() => window.sheetClipboardFixture.hidePasteOptions(false)); await expect(optionsButton()).toBeVisible()
    assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    assert.deepEqual(await page.evaluate(() => window.sheetClipboardFixture.selection()), selected)
    await optionsButton().click(); await expect(page.locator('.sheet-paste-options-popup')).toBeVisible()
    const outside = page.getByRole('textbox', { name: '外部文本', exact: true }); await outside.focus()
    await page.evaluate(() => window.sheetClipboardFixture.hidePasteOptions(true)); await expect(optionsButton()).toBeHidden()
    await expect(outside).toBeFocused(); await expect(page.locator('.sheet-paste-options-popup')).toBeHidden()
    await page.evaluate(() => window.sheetClipboardFixture.hidePasteOptions(false)); await expect(optionsButton()).toBeVisible()
    await expect(outside).toBeFocused(); assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
    await optional('仅粘贴值'); assert.equal((await cell(4, 5)).formula, null)
    assert.equal((await history()).undo, standalone ? 0 : 1)
  }
})

test('an unavailable reader and independent value permissions update both Ribbon and context modes without blocking permitted formats', async () => {
  await page.goto(baseURL + '?no-reader'); await expect(grid()).toBeVisible(); await expect(pasteButton()).toBeDisabled()
  await copySelection(); await select('E5'); await expect(pasteButton()).toBeEnabled()
  await page.evaluate(() => window.sheetClipboardFixture.permission('setCellValue', false))
  await expect(pasteButton()).toHaveAttribute('aria-haspopup', 'menu'); await expect(pasteButton()).toBeEnabled()
  await pasteButton().click()
  const ribbonMenu = page.locator('.ribbon-command-menu:visible')
  await expect(ribbonMenu.getByRole('menuitem', { name: '仅粘贴值', exact: true })).toHaveAttribute('aria-disabled', 'true')
  const formatItem = ribbonMenu.getByRole('menuitem', { name: '仅粘贴格式', exact: true })
  await expect(formatItem).not.toHaveAttribute('aria-disabled', 'true'); await formatItem.click()
  await expect.poll(async () => (await cell(4, 4))?.style.fill).toBe('#ffeeaa')
  assert.equal((await cell(4, 4)).value, null); assert.equal((await history()).undo, 1)
  assert.equal(await page.evaluate(() => window.sheetClipboardFixture.reads()), 0)
  await grid().click({ button: 'right', position: { x: 560, y: 150 } })
  const context = page.locator('.sheet-context-menu')
  await expect(context.getByRole('button', { name: '粘贴', exact: true })).toBeDisabled()
  await expect(context.getByRole('button', { name: '复制', exact: true })).toBeEnabled()
  await context.getByRole('menuitem', { name: '选择性粘贴', exact: true }).hover()
  const submenu = page.locator('.sheet-context-menu__submenu:visible')
  await expect(submenu.getByRole('menuitem', { name: '仅粘贴值', exact: true })).toHaveAttribute('aria-disabled', 'true')
  await expect(submenu.getByRole('menuitem', { name: '仅粘贴格式', exact: true })).not.toHaveAttribute('aria-disabled', 'true')
})

test('copy and cut permissions appear separately in context quick actions and native events cannot bypass them', async () => {
  await select('A1:B1'); await page.evaluate(() => window.sheetClipboardFixture.permission('copy', false))
  await grid().click({ button: 'right', position: { x: 90, y: 45 } })
  const context = page.locator('.sheet-context-menu')
  await expect(context.getByRole('button', { name: '复制', exact: true })).toBeDisabled()
  await expect(context.getByRole('button', { name: '剪切', exact: true })).toBeDisabled()
  await context.getByRole('button', { name: '复制', exact: true }).locator('..').hover()
  await expect(page.getByRole('tooltip').filter({ hasText: /复制此区域/ })).toBeVisible()
  await page.keyboard.press('Escape'); const before = await snapshot(), notices = await page.evaluate(() => window.sheetClipboardFixture.notifications().length)
  const rejected = await dispatch('copy'); assert.equal(rejected.text, '')
  await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.notifications().length)).toBeGreaterThan(notices)
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
  await page.evaluate(() => { window.sheetClipboardFixture.permission('copy', true); window.sheetClipboardFixture.permission('edit', false) })
  await grid().click({ button: 'right', position: { x: 90, y: 45 } })
  await expect(context.getByRole('button', { name: '复制', exact: true })).toBeEnabled()
  await expect(context.getByRole('button', { name: '剪切', exact: true })).toBeDisabled()
  await page.keyboard.press('Escape'); const copied = await dispatch('copy'); assert.equal(copied.text, 'rich\t5')
  const cut = await dispatch('cut'); assert.equal(cut.text, '')
  assert.deepEqual(await snapshot(), before); assert.equal((await history()).undo, 0)
})

test('explicit exclusive and shape text contexts change clipboard entrances while disjoint selections preserve the special menus', async () => {
  await copySelection('A1'); await page.evaluate(() => window.sheetClipboardFixture.activeRangeList())
  await expect(pasteButton()).toBeVisible()
  await grid().click({ button: 'right', position: { x: 90, y: 45 } })
  const context = page.locator('.sheet-context-menu')
  await expect(context.getByRole('menuitem', { name: '选择性粘贴', exact: true })).toBeVisible()
  await expect(context.getByRole('menuitem', { name: '选择性复制', exact: true })).toBeVisible()
  await page.keyboard.press('Escape'); const before = await snapshot(), depth = await history()
  await page.evaluate(() => window.sheetClipboardFixture.clipboardContext({ exclusiveRange: true }))
  await expect(pasteButton()).toHaveCount(0)
  await grid().click({ button: 'right', position: { x: 90, y: 45 } })
  await expect(context.getByRole('menuitem', { name: '选择性粘贴', exact: true })).toHaveCount(0)
  await expect(context.getByRole('menuitem', { name: '选择性复制', exact: true })).toHaveCount(0)
  await expect(context.getByRole('button', { name: '复制', exact: true })).toBeVisible()
  await expect(context.getByRole('button', { name: '粘贴', exact: true })).toBeVisible()
  await page.keyboard.press('Escape')
  await page.evaluate(() => window.sheetClipboardFixture.clipboardContext({ shapeTextEditing: true }))
  await expect(pasteButton()).toHaveCount(0); await dispatch('copy')
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
})

test('unknown system content can paste plain values after style permission changes but actual HTML still checks style permission', async () => {
  await copySelection(); await select('E5')
  await page.evaluate(() => { window.sheetClipboardFixture.permission('setCellStyle', false); window.sheetClipboardFixture.setBoard({ text: 'external plain' }) })
  await expect(pasteButton()).toBeEnabled(); await pasteButton().click()
  await expect.poll(async () => (await cell(4, 4))?.value).toBe('external plain'); assert.equal((await history()).undo, 1)
  const before = await snapshot(), depth = await history(), notices = await page.evaluate(() => window.sheetClipboardFixture.notifications().length)
  await page.evaluate(() => window.sheetClipboardFixture.setBoard({ text: 'styled', html: '<table><tr><td style="font-weight:bold">styled</td></tr></table>' }))
  await expect(pasteButton()).toBeEnabled(); await pasteButton().click()
  await expect.poll(() => page.evaluate(() => window.sheetClipboardFixture.notifications().length)).toBeGreaterThan(notices)
  assert.deepEqual(await snapshot(), before); assert.deepEqual(await history(), depth)
})
