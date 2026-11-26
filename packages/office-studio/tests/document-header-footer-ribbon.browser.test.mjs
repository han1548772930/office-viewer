import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__header_footer_ribbon__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{width:100%;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { createApp, computed, h, ref } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import WorkbenchRibbon from '/packages/office-studio/src/platform/ribbon/WorkbenchRibbon.vue'
import { createDocumentRibbon } from '/packages/office-studio/src/components/document/document-ribbon'
const variant = ref('grid'), dark = ref(false), locked = ref(false), longTitle = ref(false)
const requestScope = ref(7)
const area = ref('header')
const margins = ref({ marginHeader: 24, marginFooter: 24 })
const tabs = computed(() => {
  const tabs = createDocumentRibbon({
  state: { marks: {}, mixedMarks: [], paragraph: {}, hasSelection: false },
  canUndo: false, canRedo: false, outline: false, focus: false, fullscreen: false,
  paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true, mode: 'traditional',
  headerFooter: {
    visible: true, area: area.value, requestScope: requestScope.value, disabled: locked.value,
    sectionDisabled: locked.value, contentDisabled: locked.value,
    canLinkToPrevious: true, linkedToPrevious: false, differentFirstPage: false, differentOddEven: false,
    ...margins.value,
  },
  }).filter(tab => tab.id === 'header-footer')
  if (longTitle.value) {
    const item = tabs[0].groups.find(group => group.id === 'doc-header-footer-position').items[0]
    item.popupProps = { ...item.popupProps, title: '页眉与页脚距页面边缘的位置 Header and footer position' }
    item.gridLayout = { ...item.gridLayout, width: 68, minWidth: 60 }
  }
  return tabs
})
window.headerFooterFixture = { variant, dark, locked, longTitle, requestScope }
window.headerFooterInvocations = []
createApp({ render: () => h(ConfigProvider, {
  theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#315bc2' } },
}, { default: () => h(WorkbenchRibbon, {
  tabs: tabs.value, variant: variant.value, modelValue: 'header-footer',
  onExecute: (id, value) => {
    const action = typeof value === 'string' ? JSON.parse(value) : value
    window.headerFooterInvocations.push([id, action])
    if (action?.action === 'navigate') area.value = action.area
    if (action?.action === 'set-options') margins.value = { ...margins.value, ...action.patch }
  },
}) }) }).mount('#app')
</script></body></html>`

let server
let browser
let page
let baseURL
let errors

beforeAll(async () => {
  server = await createServer({
    configFile: false, root: process.cwd(), appType: 'custom', logLevel: 'error',
    plugins: [vue(), {
      name: 'header-footer-ribbon-fixture',
      configureServer(vite) {
        vite.middlewares.use(fixturePath, async (_request, response, next) => {
          try {
            response.setHeader('Content-Type', 'text/html; charset=utf-8')
            response.end(await vite.transformIndexHtml(fixturePath, fixture))
          } catch (error) { next(error) }
        })
      },
    }],
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)

afterAll(async () => { await browser?.close(); await server?.close() })
beforeEach(async () => {
  errors = []
  page = await browser.newPage({ viewport: { width: 1440, height: 700 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await page.getByRole('tab', { name: '页眉和页脚', exact: true }).waitFor()
}, 60_000)
afterEach(async () => { await page?.close() })

async function geometry(control) {
  return control.evaluate(element => {
    const host = element.closest('.workbench-ribbon-item').getBoundingClientRect()
    const bounds = element.getBoundingClientRect()
    return {
      host: { left: host.left, right: host.right, top: host.top, bottom: host.bottom },
      height: bounds.height,
      children: Array.from(element.querySelectorAll('button, .ant-checkbox-wrapper, .ant-form-item-row, .ant-input-number')).map(child => {
        const box = child.getBoundingClientRect()
        return { label: child.textContent, left: box.left, right: box.right, top: box.top, bottom: box.bottom, height: box.height }
      }),
    }
  })
}

for (const variant of ['classic', 'grid']) test(`${variant} header/footer controls fit their real Ribbon hosts and preserve navigation focus and read-only commands`, async () => {
  await page.evaluate(value => { window.headerFooterFixture.variant.value = value }, variant)
  const controls = page.locator('.ribbon-visible-groups .header-footer-ribbon-control, .ribbon-grid-group:not(.is-hidden) .header-footer-ribbon-control')
  await expect(controls).toHaveCount(4)
  for (const control of await controls.all()) {
    const layout = await geometry(control)
    assert.ok(layout.height <= (variant === 'grid' ? 52 : 28))
    for (const child of layout.children) {
      assert.ok(child.top >= layout.host.top - 1 && child.bottom <= layout.host.bottom + 1, child.label)
      assert.ok(child.left >= layout.host.left - 1 && child.right <= layout.host.right + 1, child.label)
      assert.ok(child.height <= 24.5, child.label)
    }
  }
  const position = controls.filter({ has: page.locator('.header-footer-ribbon-control__form') })
  const rows = await position.locator('.ant-form-item-row').evaluateAll(elements => elements.map(element => {
    const row = element.getBoundingClientRect()
    const label = element.querySelector('.ant-form-item-label').getBoundingClientRect()
    const input = element.querySelector('.ant-input-number').getBoundingClientRect()
    return { height: row.height, top: row.top, labelTop: label.top, labelBottom: label.bottom, inputTop: input.top, inputBottom: input.bottom }
  }))
  assert.deepEqual(rows.map(row => row.height), [24, 24])
  for (const row of rows) assert.ok(Math.abs((row.labelTop + row.labelBottom) / 2 - (row.inputTop + row.inputBottom) / 2) <= 1)
  if (variant === 'grid') assert.equal(rows[1].top - rows[0].top, 28)
  for (const button of await controls.locator('button').all()) {
    await expect(button).toHaveCSS('border-width', '0px')
    await button.hover()
    await expect(button).toHaveCSS('border-width', '0px')
    await page.mouse.down()
    if (await button.isEnabled()) assert.equal(await button.evaluate(element => element.matches(':active')), true)
    await expect(button).toHaveCSS('border-width', '0px')
    await page.mouse.move(1400, 650)
    await page.mouse.up()
  }
  const input = position.getByRole('spinbutton').first()
  await input.focus()
  const footer = page.getByRole('button', { name: '转到页脚', exact: true })
  await footer.hover()
  await page.mouse.down()
  await expect(input).toBeFocused()
  assert.deepEqual(await page.evaluate(() => window.headerFooterInvocations), [])
  await page.mouse.up()
  const header = page.getByRole('button', { name: '转到页眉', exact: true })
  await expect(header).toBeEnabled()
  await page.evaluate(() => { window.headerFooterFixture.locked.value = true })
  await expect(input).toBeDisabled()
  await expect(header).toBeEnabled()
  for (const button of await controls.locator('button').all()) await expect(button).toHaveCSS('border-width', '0px')
  const tab = page.getByRole('tab', { name: '页眉和页脚', exact: true })
  await tab.focus()
  await header.hover()
  await page.mouse.down()
  await expect(tab).toBeFocused()
  await page.mouse.up()
  assert.deepEqual(await page.evaluate(() => window.headerFooterInvocations), [
    ['doc.operation.header-footer-control', { action: 'navigate', area: 'footer', requestScope: 7 }],
    ['doc.operation.header-footer-control', { action: 'navigate', area: 'header', requestScope: 7 }],
  ])
  assert.deepEqual(errors, [])
})

test('the compact title and chevron share a centered caption and hover stays inside the trigger', async () => {
  await page.setViewportSize({ width: 800, height: 700 })
  const button = page.locator('.ribbon-grid-group:not(.is-hidden) [data-ribbon-item="doc-header-footer-position"]').getByRole('button', { name: '位置', exact: true })
  await expect(button).toBeVisible()
  const metrics = await button.evaluate(element => {
    const trigger = element.getBoundingClientRect()
    const caption = element.querySelector('.header-footer-ribbon-group__caption').getBoundingClientRect()
    const label = element.querySelector('.header-footer-ribbon-group__caption > span').getBoundingClientRect()
    const arrow = element.querySelector('svg').getBoundingClientRect()
    return { height: trigger.height, triggerCenterX: trigger.x + trigger.width / 2, triggerCenterY: trigger.y + trigger.height / 2,
      captionCenterX: caption.x + caption.width / 2, captionCenterY: caption.y + caption.height / 2,
      labelCenterY: label.y + label.height / 2, arrowCenterY: arrow.y + arrow.height / 2, aboveY: trigger.top - 4 }
  })
  assert.equal(metrics.height, 24)
  assert.ok(Math.abs(metrics.triggerCenterX - metrics.captionCenterX) <= 1)
  assert.ok(Math.abs(metrics.triggerCenterY - metrics.captionCenterY) <= 1)
  assert.ok(Math.abs(metrics.labelCenterY - metrics.arrowCenterY) <= 1)
  await button.hover()
  assert.equal(await button.evaluate(element => element.matches(':hover')), true)
  await page.mouse.move(metrics.triggerCenterX, metrics.aboveY)
  assert.equal(await button.evaluate(element => element.matches(':hover')), false)
  assert.equal(await button.locator('..').evaluate(element => getComputedStyle(element).backgroundColor), 'rgba(0, 0, 0, 0)')
  assert.deepEqual(await page.evaluate(() => window.headerFooterInvocations), [])
  assert.deepEqual(errors, [])
})

test('the compact position popup keeps both label/input rows together and forwards one scoped step change', async () => {
  await page.setViewportSize({ width: 800, height: 700 })
  await page.locator('.ribbon-grid-group:not(.is-hidden) [data-ribbon-item="doc-header-footer-position"]').getByRole('button', { name: '位置', exact: true }).click()
  const control = page.locator('.ant-popover-inner .header-footer-ribbon-control[data-control-kind="position"]')
  await expect(control).toBeVisible()
  const rows = control.locator('.ant-form-item-row')
  await expect(rows).toHaveCount(2)
  assert.deepEqual(await rows.evaluateAll(elements => elements.map(element => element.getBoundingClientRect().height)), [24, 24])
  await control.getByRole('spinbutton').first().press('ArrowUp')
  await expect.poll(() => page.evaluate(() => window.headerFooterInvocations)).toEqual([
    ['doc.operation.header-footer-control', { action: 'set-options', patch: { marginHeader: 25 }, requestScope: 7 }],
  ])
  await expect(control).toBeHidden()
  assert.deepEqual(errors, [])
})

test('a constrained compact caption truncates a long title inside the real hover target', async () => {
  await page.setViewportSize({ width: 800, height: 700 })
  await page.evaluate(() => { window.headerFooterFixture.longTitle.value = true })
  const button = page.locator('.ribbon-grid-group:not(.is-hidden) [data-ribbon-item="doc-header-footer-position"] .header-footer-ribbon-group__trigger')
  await expect(button).toBeVisible()
  const metrics = await button.evaluate(element => {
    const trigger = element.getBoundingClientRect(), caption = element.querySelector('.header-footer-ribbon-group__caption')
    const label = caption.querySelector('span'), labelBox = label.getBoundingClientRect(), arrow = caption.querySelector('svg').getBoundingClientRect()
    return { trigger: { left: trigger.left, right: trigger.right, top: trigger.top, bottom: trigger.bottom },
      label: { left: labelBox.left, right: labelBox.right, clientWidth: label.clientWidth, scrollWidth: label.scrollWidth },
      arrow: { left: arrow.left, right: arrow.right, top: arrow.top, bottom: arrow.bottom }, overflow: getComputedStyle(label).textOverflow }
  })
  assert.equal(metrics.overflow, 'ellipsis'); assert.ok(metrics.label.scrollWidth > metrics.label.clientWidth)
  assert.ok(metrics.label.left >= metrics.trigger.left && metrics.label.right <= metrics.trigger.right)
  assert.ok(metrics.arrow.left >= metrics.trigger.left && metrics.arrow.right <= metrics.trigger.right)
  assert.ok(Math.abs((metrics.arrow.top + metrics.arrow.bottom) / 2 - (metrics.trigger.top + metrics.trigger.bottom) / 2) <= 1)
  await button.hover()
  assert.equal(await button.evaluate(element => element.matches(':hover')), true)
  await page.mouse.move(metrics.trigger.right + 3, (metrics.trigger.top + metrics.trigger.bottom) / 2)
  assert.equal(await button.evaluate(element => element.matches(':hover')), false)
  assert.deepEqual(await page.evaluate(() => window.headerFooterInvocations), [])
  assert.deepEqual(errors, [])
})

for (const appearance of [{ variant: 'classic', dark: false }, { variant: 'grid', dark: true }]) test(`${appearance.variant} compact position popups close on permission, request scope and layout changes`, async () => {
  await page.evaluate(value => {
    window.headerFooterFixture.variant.value = value.variant
    window.headerFooterFixture.dark.value = value.dark
  }, appearance)
  await page.setViewportSize({ width: 800, height: 700 })
  const button = page.locator('.ribbon-visible-groups [data-ribbon-item="doc-header-footer-position"], .ribbon-grid-group:not(.is-hidden) [data-ribbon-item="doc-header-footer-position"]')
    .getByRole('button', { name: '位置', exact: true })
  const control = page.locator('.ant-popover-inner .header-footer-ribbon-control[data-control-kind="position"]')
  await button.click()
  await expect(control).toBeVisible()
  await page.evaluate(() => { window.headerFooterFixture.locked.value = true })
  await expect(button).toBeDisabled()
  await expect(button).toHaveAttribute('aria-expanded', 'false')
  await expect(control).toBeHidden()
  await page.evaluate(() => { window.headerFooterFixture.locked.value = false })
  await expect(button).toBeEnabled()
  await expect(control).toBeHidden()
  await button.click()
  await expect(control).toBeVisible()
  await page.evaluate(() => { window.headerFooterFixture.requestScope.value = 8 })
  await expect(button).toHaveAttribute('aria-expanded', 'false')
  await expect(control).toBeHidden()
  await button.click()
  await expect(control).toBeVisible()
  await page.setViewportSize({ width: 1440, height: 700 })
  await expect(button).toBeHidden()
  await expect(control).toBeHidden()
  await page.setViewportSize({ width: 800, height: 700 })
  await expect(button).toBeVisible()
  await expect(button).toHaveAttribute('aria-expanded', 'false')
  await expect(control).toBeHidden()
  assert.deepEqual(await page.evaluate(() => window.headerFooterInvocations), [])
  await button.click()
  await control.getByRole('spinbutton').first().press('ArrowUp')
  await expect.poll(() => page.evaluate(() => window.headerFooterInvocations)).toEqual([
    ['doc.operation.header-footer-control', { action: 'set-options', patch: { marginHeader: 25 }, requestScope: 8 }],
  ])
  await expect(control).toBeHidden()
  assert.deepEqual(errors, [])
})

test('overflow presents the complete position form without reusing the compact trigger height', async () => {
  await page.setViewportSize({ width: 420, height: 700 })
  await page.getByRole('button', { name: '更多功能', exact: true }).click()
  const panel = page.getByRole('dialog', { name: '更多功能', exact: true })
  const control = panel.locator('[data-ribbon-item="doc-header-footer-position"] .header-footer-ribbon-control')
  await expect(control).toBeVisible()
  await expect(control.getByRole('spinbutton')).toHaveCount(2)
  const layout = await geometry(control)
  assert.equal(layout.height, 52)
  for (const child of layout.children) {
    assert.ok(child.top >= layout.host.top - 1 && child.bottom <= layout.host.bottom + 1, child.label)
    assert.ok(child.left >= layout.host.left - 1 && child.right <= layout.host.right + 1, child.label)
  }
  assert.equal(await control.locator('..').evaluate(element => getComputedStyle(element).height), '52px')
  assert.deepEqual(errors, [])
})

test('light/dark hover colors follow the shared tokens and disabled compact controls stay transparent', async () => {
  await page.setViewportSize({ width: 800, height: 700 })
  const button = page.locator('.ribbon-grid-group:not(.is-hidden) [data-ribbon-item="doc-header-footer-position"]').getByRole('button', { name: '位置', exact: true })
  await expect(button).toBeVisible()
  await expect(button).toHaveCSS('border-width', '0px')
  await button.hover()
  await expect(button).toHaveCSS('border-width', '0px')
  await page.mouse.down()
  assert.equal(await button.evaluate(element => element.matches(':active')), true)
  await expect(button).toHaveCSS('border-width', '0px')
  await page.mouse.move(760, 650)
  await page.mouse.up()
  await button.hover()
  const light = await button.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-hover-background').trim())
  await expect.poll(() => button.evaluate(element => getComputedStyle(element).backgroundColor)).toBe(light)
  await page.evaluate(() => { window.headerFooterFixture.dark.value = true })
  await expect.poll(() => button.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-hover-background').trim())).not.toBe(light)
  const dark = await button.evaluate(element => getComputedStyle(element).getPropertyValue('--ribbon-hover-background').trim())
  await expect.poll(() => button.evaluate(element => getComputedStyle(element).backgroundColor)).toBe(dark)
  await expect(button).toHaveCSS('border-width', '0px')
  await page.evaluate(() => { window.headerFooterFixture.locked.value = true })
  await expect(button).toBeDisabled()
  await expect(button).toHaveCSS('border-width', '0px')
  await expect.poll(() => button.evaluate(element => getComputedStyle(element).backgroundColor)).toBe('rgba(0, 0, 0, 0)')
  assert.deepEqual(await page.evaluate(() => window.headerFooterInvocations), [])
  assert.deepEqual(errors, [])
})
