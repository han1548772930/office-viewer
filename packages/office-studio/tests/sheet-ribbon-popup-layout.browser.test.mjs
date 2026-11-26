import assert from 'node:assert/strict'
import { afterAll, afterEach, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__sheet_ribbon_popup_layout__'
const fixture = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{width:100%;min-width:0}</style></head><body><div id="app"></div><script type="module">
import { computed, createApp, h, nextTick, ref, shallowRef } from 'vue'
import { ConfigProvider, theme } from 'antdv-next'
import WorkbenchRibbon from '/packages/office-studio/src/platform/ribbon/WorkbenchRibbon.vue'
import QuickDropdownPicker from '/packages/office-studio/src/components/spreadsheet/QuickDropdownPicker.vue'
import { createSheetRibbon } from '/packages/office-studio/src/components/spreadsheet/sheet-ribbon'
const variant = ref('classic'), width = ref('100%'), active = ref('home'), dark = ref(false)
const smallControlHeight = ref(24), labelOverrides = ref({})
const quickVisible = ref(false), quickDisabled = ref(false)
const slicerControl = shallowRef({ id: 'layout-slicer', type: 'slicer', supportsConnections: true, canConfigure: true, revision: 7 })
const tabs = computed(() => createSheetRibbon({
  canUndo: false, canRedo: false, hasSheet: true, canMerge: true, canUnmerge: false,
  canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true, canSort: true,
  fontFamily: 'Arial', fontSize: 11, numberFormat: 'general', gridlines: true, frozen: false, zoom: 100, fullscreen: false,
  borderOptions: { style: 'thin', color: '#1f2937' }, borderPreset: 'all',
  slicerControl: slicerControl.value,
}, (key, fallback) => labelOverrides.value[key] ?? fallback))
window.sheetPopupLayoutFixture = {
  configure: (layout, overflow, tab) => {
    variant.value = layout; width.value = overflow ? '80px' : '100%'; active.value = tab; dark.value = false
    return nextTick()
  },
  appearance: value => { dark.value = value; return nextTick() },
  smallControls: value => { smallControlHeight.value = value; return nextTick() },
  labels: value => { labelOverrides.value = value; return nextTick() },
  quickControls: value => { quickVisible.value = true; quickDisabled.value = value; return nextTick() },
  slicerControls: value => { slicerControl.value = { ...slicerControl.value, ...value }; return nextTick() },
  commands: [],
}
createApp({ render: () => h(ConfigProvider, {
  theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { controlHeightSM: smallControlHeight.value } },
}, { default: () => h('div', { style: { width: width.value, minWidth: '0' } }, [h(WorkbenchRibbon, {
  tabs: tabs.value, variant: variant.value, modelValue: active.value,
  'onUpdate:modelValue': value => { active.value = value },
  onExecute: (id, params) => window.sheetPopupLayoutFixture.commands.push([id, params]),
}), quickVisible.value ? h(QuickDropdownPicker, {
  disabled: quickDisabled.value, title: '快捷下拉禁用状态', editDropdown: '编辑下拉列表', clearDropdown: '清除下拉列表',
  presets: [
    { id: 'dropdown-yes-no', values: ['是', '否'] },
    { id: 'dropdown-status', values: ['未开始', '进行中', '完成'] },
    { id: 'dropdown-abc', values: ['A', 'B', 'C'] },
  ],
  onInsert: value => window.sheetPopupLayoutFixture.commands.push(['sheet.operation.quick-data-validation', value]),
}) : null]) }) }).mount('#app')
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
      name: 'sheet-ribbon-popup-layout-fixture',
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
  page = await browser.newPage({ viewport: { width: 1440, height: 800 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('tab', { name: '开始', exact: true })).toBeVisible()
}, 60_000)
afterEach(async () => { await page?.close() })

async function openPopup(variant, overflow, tab, id, selector) {
  await page.evaluate(({ variant, overflow, tab }) => window.sheetPopupLayoutFixture.configure(variant, overflow, tab), { variant, overflow, tab })
  if (overflow) {
    const more = page.getByRole('button', { name: '更多功能', exact: true })
    await expect(more).toBeVisible()
    await more.click()
    await expect(page.locator('.ribbon-overflow-panel:visible [data-ribbon-item="' + id + '"]')).toBeVisible()
  }
  const item = page.locator('[data-ribbon-item="' + id + '"]:visible')
  await expect(item).toHaveCount(1)
  if (!overflow) await expect(item).not.toHaveClass(/is-overflow/)
  await item.locator('button[aria-haspopup]').click()
  const panel = page.locator(selector + ':visible')
  await expect(panel).toBeVisible()
  return panel
}

test('quick dropdown disables every preset and management button without publishing an insert', async () => {
  await page.evaluate(() => window.sheetPopupLayoutFixture.quickControls(true))
  const picker = page.getByRole('group', { name: '快捷下拉禁用状态', exact: true })
  const buttons = picker.getByRole('button')
  await expect(buttons).toHaveCount(5)
  for (const dark of [false, true]) {
    await page.evaluate(value => window.sheetPopupLayoutFixture.appearance(value), dark)
    assert.equal(await buttons.evaluateAll(controls => controls.every(button => button.disabled)), true)
    await buttons.evaluateAll(controls => { for (const button of controls) button.click() })
    assert.deepEqual(await page.evaluate(() => window.sheetPopupLayoutFixture.commands), [])
  }
  await page.evaluate(() => window.sheetPopupLayoutFixture.quickControls(false))
  assert.equal(await buttons.evaluateAll(controls => controls.every(button => !button.disabled)), true)
  await picker.getByRole('button', { name: '是 / 否', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.sheetPopupLayoutFixture.commands), [['sheet.operation.quick-data-validation', 'dropdown-yes-no']])
  assert.deepEqual(errors, [])
})

async function previewGeometry(preview) {
  return preview.evaluate(element => {
    const button = element.closest('button')
    const container = element.classList.contains('slicer-quick-sample') ? button : element.parentElement
    const bounds = element.getBoundingClientRect()
    const host = container.getBoundingClientRect()
    const style = getComputedStyle(container)
    const buttonBounds = button.getBoundingClientRect()
    return {
      bounds: { left: bounds.left, right: bounds.right, top: bounds.top, bottom: bounds.bottom, width: bounds.width, height: bounds.height },
      content: {
        left: host.left + parseFloat(style.borderLeftWidth) + parseFloat(style.paddingLeft),
        right: host.right - parseFloat(style.borderRightWidth) - parseFloat(style.paddingRight),
        top: host.top + parseFloat(style.borderTopWidth) + parseFloat(style.paddingTop),
        bottom: host.bottom - parseFloat(style.borderBottomWidth) - parseFloat(style.paddingBottom),
      },
      button: { width: buttonBounds.width, height: buttonBounds.height },
    }
  })
}

function assertPreviewFits(geometry) {
  assert.ok(geometry.bounds.left >= geometry.content.left - 0.5)
  assert.ok(geometry.bounds.right <= geometry.content.right + 0.5)
  assert.ok(geometry.bounds.top >= geometry.content.top - 0.5)
  assert.ok(geometry.bounds.bottom <= geometry.content.bottom + 0.5)
}

for (const variant of ['classic', 'grid']) {
  for (const overflow of [false, true]) test(`${variant} ${overflow ? 'More' : 'toolbar'} popup controls and previews fit their real content before and after hover`, async () => {
    if (overflow) await page.setViewportSize({ width: 360, height: 800 })
    const border = await openPopup(variant, overflow, 'home', 'sheet.menu.borders', '.border-selector')
    const diagonals = border.locator('.border-preview svg')
    await expect(diagonals).toHaveCount(5)
    const lightInk = await border.evaluate(element => getComputedStyle(element).getPropertyValue('--edge-ink'))
    for (const dark of [false, true]) {
      await page.evaluate(value => window.sheetPopupLayoutFixture.appearance(value), dark)
      if (dark) await expect.poll(() => border.evaluate(element => getComputedStyle(element).getPropertyValue('--edge-ink'))).not.toBe(lightInk)
      for (const height of [24, 30]) {
        await page.evaluate(value => window.sheetPopupLayoutFixture.smallControls(value), height)
        await expect(border.locator('.border-color-field .sheet-color-control')).toHaveCSS('height', height + 'px')
        const fields = await border.evaluate(element => ({
          colorHeight: element.querySelector('.border-color-field .sheet-color-control').getBoundingClientRect().height,
          styleHeight: element.querySelector('.border-style-field .ant-select').getBoundingClientRect().height,
        }))
        assert.equal(fields.colorHeight, height)
        assert.equal(fields.styleHeight, height)
      }
      await page.evaluate(() => window.sheetPopupLayoutFixture.smallControls(24))
      for (const diagonal of await diagonals.all()) {
        const before = await previewGeometry(diagonal)
        assertPreviewFits(before)
        assert.equal(before.bounds.width, 16)
        assert.equal(before.bounds.height, 16)
        assert.deepEqual(before.button, { width: 24, height: 24 })
        const button = diagonal.locator('xpath=ancestor::button')
        const name = await button.getAttribute('aria-label')
        const tooltip = page.getByRole('tooltip').filter({ hasText: name })
        await button.hover()
        await expect(tooltip).toHaveText(name)
        const hovered = await previewGeometry(diagonal)
        assertPreviewFits(hovered)
        assert.deepEqual(hovered.button, before.button)
        await page.mouse.move(0, 0)
        await expect(tooltip).toBeHidden()
        await button.focus()
        await expect(tooltip).toHaveText(name)
        const focused = await previewGeometry(diagonal)
        assertPreviewFits(focused)
        assert.deepEqual(focused.button, before.button)
        await button.evaluate(element => element.blur())
        await expect(tooltip).toBeHidden()
      }
    }
    await diagonals.last().locator('xpath=ancestor::button').press('Escape')
    await expect(border).toBeHidden()

    const color = page.locator('[data-ribbon-item="sheet.menu.fill"]:visible')
    for (const dark of [false, true]) {
      await page.evaluate(value => window.sheetPopupLayoutFixture.appearance(value), dark)
      for (const selector of ['.sheet-color-main', '.sheet-color-arrow button']) {
        const button = color.locator(selector)
        await page.mouse.move(0, 0)
        const resting = await button.boundingBox()
        for (const state of ['hover', 'active']) {
          await button.hover()
          if (state === 'active') await page.mouse.down()
          assert.equal(await button.evaluate((element, state) => element.matches(':' + state), state), true)
          for (const edge of ['top', 'right', 'bottom', 'left']) await expect(button).toHaveCSS('border-' + edge + '-width', '0px')
          assert.deepEqual(await button.boundingBox(), resting)
          if (state === 'active') {
            await page.mouse.move(0, 0)
            await page.mouse.up()
          }
        }
      }
    }

    const slicer = await openPopup(variant, overflow, 'slicer-format', 'sheet.menu.slicer-quick-style', '.slicer-quick-styles')
    const sample = slicer.locator('.slicer-quick-sample').first()
    await expect(sample).toBeVisible()
    const lightSurface = await slicer.evaluate(element => getComputedStyle(element).getPropertyValue('--style-surface'))
    for (const dark of [false, true]) {
      await page.evaluate(value => window.sheetPopupLayoutFixture.appearance(value), dark)
      if (dark) await expect.poll(() => slicer.evaluate(element => getComputedStyle(element).getPropertyValue('--style-surface'))).not.toBe(lightSurface)
      const panel = await slicer.evaluate(element => ({ width: element.getBoundingClientRect().width, limit: Math.min(320, innerWidth - 48) }))
      assert.ok(panel.width <= panel.limit + 0.5)
      const before = await previewGeometry(sample)
      assertPreviewFits(before)
      assert.equal(before.button.height, 72)
      const button = sample.locator('xpath=ancestor::button')
      const name = await button.getAttribute('aria-label')
      const tooltip = page.getByRole('tooltip').filter({ hasText: name })
      await button.hover()
      await expect(tooltip).toHaveText(name)
      const hovered = await previewGeometry(sample)
      assertPreviewFits(hovered)
      assert.deepEqual(hovered.button, before.button)
      await page.mouse.move(0, 0)
      await expect(tooltip).toBeHidden()
      await button.focus()
      await expect(tooltip).toHaveText(name)
      const focused = await previewGeometry(sample)
      assertPreviewFits(focused)
      assert.deepEqual(focused.button, before.button)
      await button.evaluate(element => element.blur())
      await expect(tooltip).toBeHidden()
    }
    await sample.locator('xpath=ancestor::button').press('Escape')
    await expect(slicer).toBeHidden()

    const editLabel = '编辑下拉选项并保留完整可访问名称'.repeat(8)
    const clearLabel = '清除下拉选项并保留完整可访问名称'.repeat(8)
    const formatLabel = '长数字格式名称用于检查区域边界'.repeat(8)
    await page.evaluate(labels => window.sheetPopupLayoutFixture.labels(labels), {
      'sheet.validation.quick.edit-dropdown': editLabel,
      'sheet.validation.quick.clear-dropdown': clearLabel,
      'sheet.number-format.currency': formatLabel,
    })
    const dropdown = await openPopup(variant, overflow, 'insert', 'sheet.operation.quick-data-validation.dropdown', '.quick-dropdown')
    const presets = dropdown.locator('.quick-preset')
    await expect(presets).toHaveCount(3)
    for (const preset of await presets.all()) {
      for (const edge of ['top', 'right', 'bottom', 'left']) {
        await expect(preset).toHaveCSS('padding-' + edge, '12px')
        await expect(preset).toHaveCSS('border-' + edge + '-width', '0px')
      }
      await expect(preset).toHaveCSS('font-size', '14px')
      await expect(preset).toHaveCSS('line-height', '20px')
      const values = preset.locator('.preset-values')
      await expect(values).toHaveCSS('row-gap', '6px')
      const before = await previewGeometry(values)
      assertPreviewFits(before)
      assert.ok(before.button.height >= 80)
      await preset.hover()
      assert.deepEqual(await previewGeometry(values), before)
      for (const edge of ['top', 'right', 'bottom', 'left']) await expect(preset).toHaveCSS('border-' + edge + '-width', '0px')
    }
    for (const name of [editLabel, clearLabel]) {
      const action = dropdown.getByRole('button', { name, exact: true })
      const label = action.locator('.quick-action-label')
      await expect(label).toHaveText(name)
      await expect(label).toHaveCSS('text-overflow', 'ellipsis')
      assert.ok(await label.evaluate(element => element.scrollWidth > element.clientWidth))
      const icon = await action.locator('svg').evaluate(element => ({ width: element.getBoundingClientRect().width, height: element.getBoundingClientRect().height }))
      assert.deepEqual(icon, { width: 16, height: 16 })
      const before = await previewGeometry(label)
      assertPreviewFits(before)
      await action.hover()
      const hovered = await previewGeometry(label)
      assertPreviewFits(hovered)
      assert.deepEqual(hovered.button, before.button)
    }
    await dropdown.getByRole('button', { name: clearLabel, exact: true }).press('Escape')
    await expect(dropdown).toBeHidden()

    const formats = await openPopup(variant, overflow, 'home', 'sheet.command.set-number-format', '.sheet-number-format-menu')
    const formatRow = formats.locator('.sheet-number-format-row').filter({ hasText: formatLabel })
    await expect(formatRow.locator('span').first()).toHaveText(formatLabel)
    await expect(formatRow.locator('span').first()).toHaveCSS('text-overflow', 'ellipsis')
    const formatted = await formatRow.evaluate(element => {
      const label = element.firstElementChild, preview = element.lastElementChild
      const labelRect = label.getBoundingClientRect(), previewRect = preview.getBoundingClientRect()
      return { clipped: label.scrollWidth > label.clientWidth,
        previewClipped: preview.scrollWidth > preview.clientWidth,
        labelRight: labelRect.right, previewLeft: previewRect.left, previewRight: previewRect.right,
        rowRight: element.getBoundingClientRect().right, gap: parseFloat(getComputedStyle(element).columnGap) }
    })
    assert.equal(formatted.clipped, true)
    assert.equal(formatted.previewClipped, false)
    assert.ok(formatted.labelRight + formatted.gap <= formatted.previewLeft + .5)
    assert.ok(formatted.previewRight <= formatted.rowRight + .5)
    const formatTrigger = page.locator('.sheet-number-format-trigger:visible')
    await expect(formatTrigger).toHaveAttribute('aria-expanded', 'true')
    await expect(formatTrigger).toHaveCSS('height', variant === 'grid' && !overflow ? '24px' : '28px')
    for (const dark of [false, true]) {
      await page.evaluate(value => window.sheetPopupLayoutFixture.appearance(value), dark)
      const activeBackground = await formatTrigger.evaluate(element => {
        const probe = document.createElement('span')
        probe.style.color = getComputedStyle(element).getPropertyValue('--ribbon-active-background')
        document.body.append(probe)
        const color = getComputedStyle(probe).color
        probe.remove()
        return color
      })
      await expect(formatTrigger).toHaveCSS('background-color', activeBackground)
      const before = await previewGeometry(formatTrigger.locator('svg'))
      assertPreviewFits(before)
      await formatTrigger.hover()
      await expect(formatTrigger).toHaveCSS('background-color', activeBackground)
      for (const edge of ['top', 'right', 'bottom', 'left']) await expect(formatTrigger).toHaveCSS('border-' + edge + '-width', '0px')
      const hovered = await previewGeometry(formatTrigger.locator('svg'))
      assertPreviewFits(hovered)
      assert.deepEqual(hovered.button, before.button)
      const centerOffset = await formatTrigger.evaluate(element => {
        const button = element.getBoundingClientRect(), arrow = element.querySelector('svg').getBoundingClientRect()
        return Math.abs(arrow.top + arrow.height / 2 - button.top - button.height / 2)
      })
      assert.ok(centerOffset <= 0.05)
    }
    assert.deepEqual(await page.evaluate(() => window.sheetPopupLayoutFixture.commands), [])
    assert.deepEqual(errors, [])
  })
}

for (const kind of ['slicer', 'timeline']) test(`${kind} quick styles show authored regions and the effective document default`, async () => {
  const name = 'Authored ribbon style'
  const formatting = background => ({ bg: { rgb: background }, cl: { rgb: '#fedcba' } })
  const style = { name, wholeStyle: formatting('#102030'), headerRowStyle: formatting('#314253'),
    [kind === 'slicer' ? 'slicerStyleElements' : 'timelineStyleElements']: kind === 'slicer' ? [
      { type: 'selectedItemWithData', formatting: formatting('#d82757') },
      { type: 'unselectedItemWithData', formatting: formatting('#eadedf') },
    ] : [
      { type: 'selectionLabel', formatting: formatting('#112233') },
      { type: 'periodLabel1', formatting: formatting('#aabbcc') },
      { type: 'periodLabel2', formatting: formatting('#ddee11') },
      { type: 'selectedTimeBlock', formatting: formatting('#d82757') },
      { type: 'unselectedTimeBlock', formatting: formatting('#eadedf') },
    ],
  }
  await page.evaluate(({ kind, name, style }) => window.sheetPopupLayoutFixture.slicerControls({ type: kind,
    styleOwner: { [kind === 'slicer' ? 'slicerStyles' : 'timelineStyles']: {
      [kind === 'slicer' ? 'defaultSlicerStyle' : 'defaultTimelineStyle']: name, styles: [style],
    } },
  }), { kind, name, style })
  const gallery = await openPopup('classic', false, 'slicer-format', 'sheet.menu.slicer-quick-style', '.slicer-quick-styles')
  await expect(gallery).toHaveAttribute('aria-label', kind === 'slicer' ? '切片器快速样式' : '时间线快速样式')
  await expect(gallery.getByRole('button')).toHaveCount(kind === 'slicer' ? 16 : 14)
  await expect(gallery.locator('[aria-pressed="true"]')).toHaveCount(1)
  const chosen = gallery.getByRole('button', { name, exact: true })
  await expect(chosen).toHaveAttribute('aria-pressed', 'true')
  await expect(gallery.getByRole('button', { name: '文档默认值', exact: true })).toHaveAttribute('aria-pressed', 'false')
  const sample = chosen.locator('.slicer-quick-sample')
  for (const dark of [false, true]) {
    await page.evaluate(value => window.sheetPopupLayoutFixture.appearance(value), dark)
    await expect(sample).toHaveCSS('background-color', 'rgb(16, 32, 48)')
    await expect(sample.locator('.sample-header')).toHaveCSS('background-color', 'rgb(49, 66, 83)')
    await expect(sample.locator('.sample-header i')).toHaveCSS('background-color', 'rgb(254, 220, 186)')
    assertPreviewFits(await previewGeometry(sample))
    if (kind === 'slicer') {
      const rows = sample.locator('.sample-row')
      await expect(rows).toHaveCount(4)
      await expect(rows.nth(1)).toHaveCSS('background-color', 'rgb(216, 39, 87)')
      for (const index of [2, 3]) await expect(rows.nth(index)).toHaveCSS('background-color', 'rgb(234, 222, 223)')
      for (const row of await rows.all()) assertPreviewFits(await previewGeometry(row))
      continue
    }
    await expect(sample.locator('.sample-label')).toHaveCSS('background-color', 'rgb(17, 34, 51)')
    const periods = sample.locator('.sample-periods > span'), blocks = sample.locator('.sample-blocks > span')
    await expect(periods).toHaveCount(5)
    await expect(blocks).toHaveCount(5)
    for (let index = 0; index < 5; index++) {
      await expect(periods.nth(index)).toHaveCSS('background-color', index % 2 === 0 ? 'rgb(170, 187, 204)' : 'rgb(221, 238, 17)')
      await expect(blocks.nth(index)).toHaveCSS('background-color', index >= 1 && index <= 3 ? 'rgb(216, 39, 87)' : 'rgb(234, 222, 223)')
      assertPreviewFits(await previewGeometry(periods.nth(index)))
      assertPreviewFits(await previewGeometry(blocks.nth(index)))
    }
  }
  await chosen.click()
  await expect(gallery).toBeHidden()
  const reopened = await openPopup('classic', false, 'slicer-format', 'sheet.menu.slicer-quick-style', '.slicer-quick-styles')
  await reopened.getByRole('button', { name: '文档默认值', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.sheetPopupLayoutFixture.commands), [
    ['sheet.command.update-slicer-presentation', { slicerId: 'layout-slicer', expectedRevision: 7, patch: { style: name } }],
    ['sheet.command.update-slicer-presentation', { slicerId: 'layout-slicer', expectedRevision: 7, patch: { style: undefined } }],
  ])
  assert.deepEqual(errors, [])
})
