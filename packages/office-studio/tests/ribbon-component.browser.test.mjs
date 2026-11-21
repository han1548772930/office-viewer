import assert from 'node:assert/strict'
import { beforeAll, afterAll, beforeEach, test } from 'vitest'
import { createServer } from 'vite'
import vue from '@vitejs/plugin-vue'
import { chromium, expect } from '@playwright/test'

let server
let browser
let page
let baseURL
const errors = []

const fixture = `
import { createApp, computed, h, ref } from 'vue'
import { Button, ConfigProvider, theme } from 'antdv-next'
import { Bold, Undo2 } from 'lucide-vue-next'
import WorkbenchRibbon from '/packages/office-studio/src/platform/ribbon/WorkbenchRibbon.vue'
const active = ref('home')
const variant = ref('classic')
const locked = ref(false)
const showContext = ref(true)
const menusVisible = ref(true)
const direction = ref('ltr')
const family = ref('Arial')
const bold = ref(false)
const longTabTitles = ref(false)
const popupTools = ref(false)
const dark = ref(false)
const customSlot = ref(false)
const PopupPicker = {
  emits: ['insert'],
  setup(_props, { emit }) {
    return () => h(Button, { onClick: () => emit('insert', 'chosen-symbol') }, () => '插入测试符号')
  },
}
window.ribbonInvocations = []
window.ribbonFixture = { active, variant, locked, showContext, menusVisible, direction, family, bold, longTabTitles, popupTools, dark, customSlot }
const tabs = computed(() => menusVisible.value ? [
  { id: 'home', title: longTabTitles.value ? '开始与格式编辑' : '开始', order: 0, groups: [
    { id: 'history', title: '历史', items: [
      { id: 'undo', type: 'button', title: '撤销', icon: Undo2, commandId: 'history.command.undo', disabled: locked.value, disabledReason: '没有历史记录' },
    ] },
    { id: 'font', title: '字体', items: [
      { id: 'family', type: 'select', title: '字体', commandId: 'doc.command.font', value: family.value, width: 112, options: [{ label: 'Arial', value: 'Arial' }, { label: '宋体', value: '宋体' }, { label: '不可用字体', value: 'locked', disabled: true }] },
      { id: 'bold', type: 'button', title: '加粗', icon: Bold, activated: bold.value, commandId: 'doc.command.bold' },
      { id: 'colors', type: 'menu', title: '字体颜色', showLabel: true, disabled: locked.value, children: [
        { id: 'red', type: 'button', title: '红色', commandId: 'doc.command.color', params: '#ef4444' },
        { id: 'blue', type: 'button', title: '蓝色', commandId: 'doc.command.color', params: '#3b82f6' },
      ] },
    ] },
    { id: 'paragraph', title: '段落', items: [
      ...Array.from({ length: 5 }, (_, n) => ({ id: 'tool-' + n, type: 'button', title: '段落工具' + n, commandId: 'doc.command.tool', params: n, showLabel: true })),
      { id: 'validation-rule', type: 'select', title: '验证规则', commandId: 'sheet.command.set-validation', value: 'whole', options: [
        { label: '整数', value: 'whole', commandId: 'sheet.command.set-validation', params: { rule: { type: 'whole', operator: 'between', formula1: 1, formula2: 9 } } },
        { label: '列表', value: 'list', commandId: 'sheet.command.set-validation-list', params: { values: ['A', 'B'] } },
      ] },
      { id: 'list-toggle', type: 'split', title: '列表', commandId: 'doc.command.list-toggle', params: { type: 'bullet' }, children: [
        { id: 'list-numbered', type: 'button', title: '有序列表', commandId: 'doc.command.set-block-type', params: { type: 'numbered' } },
      ] },
    ] },
  ] },
  { id: 'hidden', title: '隐藏', hidden: true, groups: [{ id: 'hidden-group', items: [{ id: 'hidden-tool', type: 'button', title: '隐藏工具', commandId: 'hidden' }] }] },
  { id: 'empty', title: '空分类', groups: [] },
  { id: 'insert', title: longTabTitles.value ? '插入与对象管理' : '插入', order: 1, groups: [{ id: 'insert', items: [{ id: 'insert-table', type: 'button', title: '插入表格', commandId: 'doc.command.insert-table', showLabel: true }] }] },
  { id: 'context', title: longTabTitles.value ? '表格上下文工具' : '表格', order: 2, contextual: true, hidden: !showContext.value, groups: [{ id: 'table', items: [{ id: 'delete-table', type: 'button', title: '删除表格', commandId: 'doc.command.delete-table', showLabel: true }] }] },
  { id: 'popup-tools', title: '弹层工具', order: 3, hidden: !popupTools.value, groups: [
    { id: 'popup-first', items: [
      { id: 'popup-only', type: 'menu', title: '符号选择', icon: Bold, popupComponent: PopupPicker, popupCommandId: 'doc.command.insert-symbol', disabled: locked.value },
      { id: 'popup-other', type: 'menu', title: '其他选项', icon: Undo2, children: [{ id: 'popup-other-action', type: 'button', title: '执行其他选项', commandId: 'doc.command.other' }] },
      { id: 'popup-split', type: 'split', title: '分离菜单', icon: Bold, commandId: 'doc.command.split-main', children: [
        { id: 'popup-split-first', type: 'button', title: '第一选项', commandId: 'doc.command.split-first' },
        { id: 'popup-split-last', type: 'button', title: '最后选项', commandId: 'doc.command.split-last' },
      ] },
      { id: 'popup-select', type: 'select', title: '选择测试', value: 'one', options: [{ label: '一', value: 'one' }], commandId: 'doc.command.select' },
    ] },
    { id: 'popup-custom', items: [{ id: 'custom-tool', type: 'button', title: '自定义工具', hidden: !customSlot.value, disabled: locked.value }] },
  ] },
] : [])
const gridLayoutById = {
  undo: { row: 1, column: 1, rowSpan: 2, showLabel: true },
  family: { row: 1, column: 1, rowSpan: 2, width: 112, minWidth: 84, showLabel: true },
  bold: { row: 1, column: 2, width: 32 },
  colors: { row: 2, column: 2, width: 32, showLabel: false },
  ...Object.fromEntries(Array.from({ length: 5 }, (_, index) => ['tool-' + index, { row: index % 2 + 1, column: Math.floor(index / 2) + 1, width: 90, minWidth: 64, showLabel: true }])),
  'validation-rule': { row: 1, column: 4, rowSpan: 2, width: 116, minWidth: 92, showLabel: true },
  'list-toggle': { row: 1, column: 5, rowSpan: 2, width: 84, showLabel: true },
  'insert-table': { row: 1, column: 1, rowSpan: 2, width: 72, showLabel: true },
  'delete-table': { row: 1, column: 1, rowSpan: 2, width: 72, showLabel: true },
  'popup-only': { row: 1, column: 1 },
  'popup-other': { row: 2, column: 1 },
  'popup-split': { row: 1, column: 2, rowSpan: 2 },
  'popup-select': { row: 1, column: 3, width: 96 },
  'custom-tool': { row: 1, column: 1, rowSpan: 2, width: 120 },
}
const gridTabs = computed(() => tabs.value.map((tab) => ({ ...tab, groups: tab.groups.map((group) => ({ ...group, items: group.items.map((item) => ({ ...item, gridLayout: gridLayoutById[item.id] })) })) })))
createApp({ render() { return h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#315bc2' } }, direction: direction.value }, { default: () => h(WorkbenchRibbon, { tabs: gridTabs.value, variant: variant.value, modelValue: active.value, direction: direction.value, 'onUpdate:modelValue': (value) => active.value = value, onExecute: (id, params) => { window.ribbonInvocations.push([id, params]); if (id === 'doc.command.font') family.value = params; if (id === 'doc.command.bold') bold.value = !bold.value } }, { 'tab-extra': () => null, 'item-custom-tool': ({ item, execute, overflow }) => h(Button, { disabled: item.disabled, onClick: () => execute('doc.command.custom', { overflow: Boolean(overflow) }) }, () => '自定义动作') }) }) } }).mount('#app')
`

beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    cacheDir: 'node_modules/.vite-ribbon-contract',
    optimizeDeps: { noDiscovery: true, include: ['vue', 'antdv-next', 'lucide-vue-next'] },
    plugins: [vue(), {
      name: 'ribbon-browser-fixture',
      resolveId(id) { if (id === '/__ribbon_fixture.ts') return '\0ribbon-fixture' },
      load(id) { if (id === '\0ribbon-fixture') return fixture },
      configureServer(vite) {
        vite.middlewares.use('/__ribbon_contract/', (_req, res) => {
          res.setHeader('Content-Type', 'text/html; charset=utf-8')
          res.end('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><style>body{margin:0}#app{width:100%;min-width:0}</style></head><body><div id="app"></div><script type="module" src="/__ribbon_fixture.ts"></script></body></html>')
        })
      },
    }],
    appType: 'custom',
    logLevel: 'error',
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)

afterAll(async () => {
  await page?.close()
  await browser?.close()
  await server?.close()
})

beforeEach(async () => {
  await page?.close()
  page = await browser.newPage({ viewport: { width: 1200, height: 700 } })
  errors.length = 0
  page.on('pageerror', (error) => errors.push(error.message))
  await page.goto(`${baseURL}/__ribbon_contract/`)
  await page.getByRole('tab', { name: '开始', exact: true }).waitFor()
}, 60_000)

test('classic geometry, hidden categories and keyboard navigation use a single roving tab stop', async () => {
  assert.equal(await page.getByRole('tab').count(), 3)
  assert.equal(await page.locator('.workbench-ribbon-tabs').evaluate((element) => element.getBoundingClientRect().height), 36)
  assert.equal(await page.locator('.workbench-ribbon-toolbar').evaluate((element) => element.getBoundingClientRect().height), 40)
  assert.equal(await page.locator('[role="tab"][tabindex="0"]').count(), 1)
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await page.keyboard.press('End')
  assert.equal(await page.getByRole('tab', { name: '表格', exact: true }).getAttribute('aria-selected'), 'true')
  await page.keyboard.press('ArrowRight')
  assert.equal(await page.getByRole('tab', { name: '开始', exact: true }).getAttribute('aria-selected'), 'true')
  await page.evaluate(() => { window.ribbonFixture.direction.value = 'rtl' })
  await page.keyboard.press('ArrowRight')
  assert.equal(await page.getByRole('tab', { name: '表格', exact: true }).getAttribute('aria-selected'), 'true')
  await page.keyboard.press('Home')
  assert.equal(await page.getByRole('tab', { name: '开始', exact: true }).getAttribute('aria-selected'), 'true')
  assert.deepEqual(errors, [])
})

test('classic rows center their contents and match reference button, selector, group and icon geometry', async () => {
  await expect(page.getByRole('button', { name: '更多功能', exact: true })).toHaveCount(0)
  const geometry = await page.evaluate(() => {
    const tabs = [...document.querySelectorAll('.workbench-ribbon-tabs [role="tab"]')].map(element => element.getBoundingClientRect())
    const tabRow = document.querySelector('.workbench-ribbon-tabs').getBoundingClientRect()
    const toolbar = document.querySelector('.workbench-ribbon-toolbar').getBoundingClientRect()
    const controls = document.querySelector('.ribbon-visible-groups').getBoundingClientRect()
    const groups = [...document.querySelectorAll('.workbench-ribbon-group')]
    const undo = document.querySelector('[role="toolbar"] [aria-label="撤销"]').getBoundingClientRect()
    const icon = document.querySelector('[role="toolbar"] [aria-label="撤销"] svg').getBoundingClientRect()
    const selector = document.querySelector('[role="toolbar"] .ribbon-select').getBoundingClientRect()
    return {
      tabsHeight: tabRow.height, tabHeights: tabs.map(rect => rect.height), toolbarHeight: toolbar.height,
      tabCenterDelta: (tabs[0].left + tabs.at(-1).right) / 2 - (tabRow.left + tabRow.right) / 2,
      toolbarCenterDelta: (controls.left + controls.right - toolbar.left - toolbar.right) / 2,
      undo: { width: undo.width, height: undo.height }, icon: { width: icon.width, height: icon.height }, selectorHeight: selector.height,
      groupStyles: groups.map(element => { const style = getComputedStyle(element); return { gap: style.columnGap, left: style.paddingInlineStart, right: style.paddingInlineEnd, divider: style.borderInlineStartWidth } }),
      itemGap: groups[1].children[1].getBoundingClientRect().left - groups[1].children[0].getBoundingClientRect().right,
      emptyExtraHidden: getComputedStyle(document.querySelector('.ribbon-tab-extra')).display === 'none',
      selectedTabBackground: getComputedStyle(document.querySelector('[role="tab"][aria-selected="true"]')).backgroundColor,
      unselectedTabBackground: getComputedStyle(document.querySelector('[role="tab"][aria-selected="false"]')).backgroundColor,
    }
  })
  assert.equal(geometry.tabsHeight, 36)
  assert.deepEqual(geometry.tabHeights, [28, 28, 28])
  assert.equal(geometry.toolbarHeight, 40)
  assert.ok(Math.abs(geometry.tabCenterDelta) <= 1, JSON.stringify(geometry))
  assert.ok(Math.abs(geometry.toolbarCenterDelta) <= 1, JSON.stringify(geometry))
  assert.deepEqual(geometry.undo, { width: 24, height: 24 })
  assert.deepEqual(geometry.icon, { width: 16, height: 16 })
  assert.equal(geometry.selectorHeight, 28)
  assert.deepEqual(geometry.groupStyles, [
    { gap: '8px', left: '8px', right: '8px', divider: '0px' },
    { gap: '8px', left: '8px', right: '8px', divider: '1px' },
    { gap: '8px', left: '8px', right: '8px', divider: '1px' },
  ])
  assert.equal(geometry.itemGap, 8)
  assert.equal(geometry.emptyExtraHidden, true)
  assert.notEqual(geometry.selectedTabBackground, geometry.unselectedTabBackground)
  assert.deepEqual(errors, [])
})

test('grid variant is opt-in and uses two-row placements with large controls spanning both rows', async () => {
  await page.evaluate(() => { window.ribbonFixture.variant.value = 'grid' })
  const toolbar = page.getByTestId('ribbon-grid-toolbar')
  await toolbar.waitFor()
  const geometry = await page.evaluate(() => {
    const tabs = document.querySelector('.workbench-ribbon-tabs').getBoundingClientRect()
    const toolbar = document.querySelector('.workbench-ribbon-toolbar').getBoundingClientRect()
    const menu = document.querySelector('[role="tab"][aria-selected="true"]')
    const largeCell = document.querySelector('[data-ribbon-item="undo"]').closest('.ribbon-grid-cell')
    const cells = document.querySelector('.ribbon-grid-cells')
    return {
      tabsHeight: tabs.height,
      toolbarHeight: toolbar.height,
      toolbarInlinePadding: getComputedStyle(toolbar).paddingInlineStart,
      rowTracks: getComputedStyle(cells).gridTemplateRows,
      rowGap: getComputedStyle(cells).rowGap,
      verticalPadding: getComputedStyle(cells).paddingBlock,
      columnGap: getComputedStyle(cells).columnGap,
      tabBorder: getComputedStyle(menu).borderBottomWidth,
      tabIndicatorHeight: getComputedStyle(menu, '::after').height,
      tabIndicatorColor: getComputedStyle(menu, '::after').backgroundColor,
      tabClass: menu.className,
      tabStyle: menu.getAttribute('style'),
      groupJustify: getComputedStyle(document.querySelector('.workbench-ribbon-tabs')).justifyContent,
      largeRow: getComputedStyle(largeCell).gridRow,
      largeHeight: document.querySelector('[data-ribbon-item="undo"] .ribbon-button').getBoundingClientRect().height,
      largeWidth: largeCell.getBoundingClientRect().width,
      largeRowHeight: largeCell.getBoundingClientRect().height,
      regularButtonHeight: document.querySelector('[data-ribbon-item="bold"] .ribbon-button').getBoundingClientRect().height,
      moreWidth: document.querySelector('.ribbon-grid-more-anchor').getBoundingClientRect().width,
      familyLabelVisible: getComputedStyle(document.querySelector('[data-ribbon-item="family"] label')).clipPath !== 'inset(50%)',
    }
  })
  assert.equal(geometry.tabsHeight, 32)
  assert.equal(geometry.toolbarHeight, 69)
  assert.equal(geometry.toolbarInlinePadding, '0px')
  assert.deepEqual(geometry.rowTracks.split(' ').slice(0, 2), ['24px', '24px'])
  assert.equal(geometry.rowGap, '4px')
  assert.equal(geometry.verticalPadding, '8px')
  assert.equal(geometry.columnGap, '4px')
  assert.equal(geometry.tabBorder, '0px', JSON.stringify(geometry))
  assert.equal(geometry.tabIndicatorHeight, '3px', JSON.stringify(geometry))
  assert.equal(geometry.tabIndicatorColor, 'rgb(49, 91, 194)', JSON.stringify(geometry))
  assert.ok(geometry.groupJustify.endsWith('start'), JSON.stringify(geometry))
  assert.match(geometry.largeRow, /span 2/)
  assert.equal(geometry.largeHeight, geometry.largeRowHeight)
  assert.equal(geometry.largeRowHeight, 52)
  assert.equal(geometry.largeWidth, 56)
  assert.equal(geometry.regularButtonHeight, 24)
  assert.equal(geometry.moreWidth, 72)
  assert.equal(geometry.familyLabelVisible, true)
  assert.equal(await page.locator('.ribbon-visible-groups').count(), 0)
  assert.deepEqual(errors, [])
})

test('grid density preserves whole groups and exposes every overflow command exactly once', async () => {
  await page.evaluate(() => { window.ribbonFixture.variant.value = 'grid' })
  await page.setViewportSize({ width: 210, height: 700 })
  const toolbar = page.getByTestId('ribbon-grid-toolbar')
  await toolbar.waitFor()
  await expect.poll(() => toolbar.evaluate(element => element.style.getPropertyValue('--ribbon-grid-density'))).toBe('1')
  const visibleGroupIds = await toolbar.locator('.ribbon-grid-group:not(.is-hidden)').evaluateAll(elements => elements.map(element => element.dataset.groupId))
  assert.ok(visibleGroupIds.length > 0 && visibleGroupIds.length < 3, JSON.stringify(visibleGroupIds))
  const visibleIds = await toolbar.locator('.ribbon-grid-group:not(.is-hidden) [data-ribbon-item]').evaluateAll(elements => elements.map(element => element.dataset.ribbonItem))
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  const triggerGeometry = await page.evaluate(() => {
    const toolbarBounds = document.querySelector('[data-testid="ribbon-grid-toolbar"]').getBoundingClientRect()
    const anchorBounds = document.querySelector('.ribbon-grid-more-anchor').getBoundingClientRect()
    const triggerBounds = document.querySelector('.ribbon-grid-more-button').getBoundingClientRect()
    const visibleGroups = [...document.querySelectorAll('.ribbon-grid-group:not(.is-hidden)')].map(element => element.getBoundingClientRect())
    const visibleRight = Math.max(toolbarBounds.left, ...visibleGroups.map(bounds => bounds.right))
    return { toolbarLeft: toolbarBounds.left, toolbarRight: toolbarBounds.right, anchorLeft: anchorBounds.left, anchorRight: anchorBounds.right, triggerWidth: triggerBounds.width, visibleRight }
  })
  assert.equal(triggerGeometry.anchorRight - triggerGeometry.anchorLeft, 72)
  assert.ok(triggerGeometry.anchorLeft >= triggerGeometry.visibleRight - 1, JSON.stringify(triggerGeometry))
  assert.ok(triggerGeometry.anchorRight <= triggerGeometry.toolbarRight, JSON.stringify(triggerGeometry))
  assert.ok(triggerGeometry.triggerWidth <= 56, JSON.stringify(triggerGeometry))
  await more.focus()
  await page.keyboard.press('ArrowDown')
  const popup = page.getByRole('menu', { name: '更多功能', exact: true })
  await popup.waitFor()
  const popupBounds = await popup.boundingBox()
  assert.ok(popupBounds.x >= 12 && popupBounds.x + popupBounds.width <= 198, JSON.stringify(popupBounds))
  const overflowIds = await popup.locator('[data-ribbon-item]').evaluateAll(elements => elements.map(element => element.dataset.ribbonItem))
  const allIds = [...visibleIds, ...overflowIds]
  const expectedIds = ['undo', 'family', 'bold', 'colors', ...Array.from({ length: 5 }, (_, index) => `tool-${index}`), 'validation-rule', 'list-toggle']
  assert.deepEqual([...allIds].sort(), [...expectedIds].sort())
  assert.equal(new Set(allIds).size, expectedIds.length)
  await popup.getByRole('button', { name: '段落工具4', exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.tool' && value === 4))
  const documentWidth = await page.evaluate(() => ({ scrollWidth: document.documentElement.scrollWidth, viewportWidth: innerWidth, bodyWidth: document.body.scrollWidth, oversized: [...document.querySelectorAll('body *')].map(element => ({ tag: element.tagName, className: String(element.className), left: element.getBoundingClientRect().left, right: element.getBoundingClientRect().right })).filter(element => element.right > innerWidth + 0.5).slice(-8) }))
  assert.ok(documentWidth.scrollWidth <= documentWidth.viewportWidth, JSON.stringify(documentWidth))
  assert.deepEqual(errors, [])
})

test('grid overflow preserves selector params and split main/dropdown commands', async () => {
  await page.evaluate(() => { window.ribbonFixture.variant.value = 'grid' })
  await page.setViewportSize({ width: 210, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.focus()
  await page.keyboard.press('ArrowDown')
  const popup = page.getByRole('menu', { name: '更多功能', exact: true })
  await popup.waitFor()

  const rule = popup.getByRole('combobox', { name: '验证规则', exact: true })
  assert.equal(await rule.count(), 1)
  await rule.click()
  await page.getByTitle('列表', { exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id]) => id === 'sheet.command.set-validation-list'))
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['sheet.command.set-validation-list', { values: ['A', 'B'] }])

  const list = popup.getByRole('button', { name: '列表', exact: true })
  await list.click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id]) => id === 'doc.command.list-toggle'))
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['doc.command.list-toggle', { type: 'bullet' }])

  await more.focus()
  await page.keyboard.press('ArrowDown')
  await popup.waitFor()
  const dropdown = popup.getByRole('button', { name: '列表选项', exact: true })
  await dropdown.click()
  const numbered = page.getByRole('menu', { name: '列表', exact: true }).getByRole('button', { name: '有序列表', exact: true })
  await numbered.click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id]) => id === 'doc.command.set-block-type'))
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['doc.command.set-block-type', { type: 'numbered' }])

  const splitBounds = await popup.locator('.ribbon-split').evaluate(element => {
    const parent = element.getBoundingClientRect()
    const buttons = [...element.querySelectorAll(':scope > button, :scope > span > button')].map(button => button.getBoundingClientRect())
    return { width: parent.width, buttonWidths: buttons.map(button => button.width), rightmost: Math.max(...buttons.map(button => button.right)), parentRight: parent.right }
  })
  assert.ok(splitBounds.rightmost <= splitBounds.parentRight + 1, JSON.stringify(splitBounds))
  assert.deepEqual(errors, [])
})

test('selectors and activated/disabled buttons reflect editor state and dispatch real command parameters', async () => {
  const font = page.getByRole('combobox', { name: '字体', exact: true })
  assert.equal(await font.count(), 1, await page.getByRole('combobox').evaluateAll((elements) => elements.map((element) => element.outerHTML).join('\n')))
  await font.click()
  await page.getByTitle('宋体', { exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.font' && value === '宋体'))
  await page.getByRole('button', { name: '加粗', exact: true }).click()
  assert.equal(await page.getByRole('button', { name: '加粗', exact: true }).getAttribute('aria-pressed'), 'true')
  await page.evaluate(() => { window.ribbonFixture.locked.value = true })
  assert.equal(await page.getByRole('button', { name: '撤销', exact: true }).isDisabled(), true)
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [['doc.command.font', '宋体'], ['doc.command.bold', undefined]])
  assert.deepEqual(errors, [])
})

test('narrow toolbar exposes overflow controls, nested menus and selectors without document overflow', async () => {
  await page.setViewportSize({ width: 210, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.waitFor()
  await more.focus()
  await page.keyboard.press('ArrowDown')
  const popup = page.getByRole('menu', { name: '更多功能', exact: true })
  await popup.waitFor()
  await expect.poll(async () => {
    const bounds = await popup.boundingBox()
    return bounds.x >= 12 && bounds.x + bounds.width <= 198
  }).toBe(true)
  const moreSize = await more.boundingBox()
  assert.equal(moreSize.width, 24)
  assert.equal(moreSize.height, 24)
  await popup.getByRole('button', { name: '段落工具4', exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.tool' && value === 4))
  await expect(popup).toBeVisible()
  await popup.getByRole('button', { name: '字体颜色', exact: true }).click()
  const colors = page.getByRole('menu', { name: '字体颜色', exact: true })
  await colors.getByRole('button', { name: '红色', exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.color' && value === '#ef4444'))
  await expect(popup).toBeVisible()
  await popup.getByRole('combobox', { name: '字体', exact: true }).click()
  await page.getByTitle('宋体', { exact: true }).click()
  await page.waitForFunction(() => window.ribbonInvocations.some(([id, value]) => id === 'doc.command.font' && value === '宋体'))
  await more.focus()
  await page.keyboard.press('ArrowDown')
  await page.keyboard.press('Escape')
  assert.equal(await more.evaluate((element) => element === document.activeElement), true)
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true)
  assert.deepEqual(errors, [])
})

test('classic fitting preserves every control once and centers the visible suffix trigger in both directions', async () => {
  const expected = ['undo', 'family', 'bold', 'colors', ...Array.from({ length: 5 }, (_, index) => `tool-${index}`), 'validation-rule', 'list-toggle']
  for (const direction of ['ltr', 'rtl']) {
    await page.evaluate(value => { window.ribbonFixture.direction.value = value }, direction)
    for (const width of [210, 360, 620, 1200]) {
      await page.setViewportSize({ width, height: 700 })
      await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))))
      const toolbar = page.getByRole('toolbar', { name: '开始', exact: true })
      const visibleIds = await toolbar.locator('[data-ribbon-item]').evaluateAll(elements => elements.map(element => element.dataset.ribbonItem))
      const more = page.getByRole('button', { name: '更多功能', exact: true })
      const overflowIds = []
      if (await more.count()) {
        await more.focus()
        await more.press('ArrowDown')
        const popup = page.getByRole('menu', { name: '更多功能', exact: true })
        await expect(popup).toBeVisible()
        overflowIds.push(...await popup.locator('.ribbon-overflow-group > [data-ribbon-item]').evaluateAll(elements => elements.map(element => element.dataset.ribbonItem)))
        await page.keyboard.press('Escape')
        await expect(more).toBeFocused()
      }
      assert.deepEqual([...visibleIds, ...overflowIds], expected, `${direction}, ${width}px`)
      const bounds = await page.locator('.workbench-ribbon-toolbar').evaluate(element => {
        const rect = element.getBoundingClientRect()
        const visible = document.querySelector('.ribbon-visible-groups').getBoundingClientRect()
        const more = document.querySelector('.ribbon-more-anchor')?.getBoundingClientRect()
        const left = Math.min(visible.left, more?.left ?? visible.left)
        const right = Math.max(visible.right, more?.right ?? visible.right)
        return { width: innerWidth, scrollWidth: document.documentElement.scrollWidth, left, right, centerDelta: (left + right - rect.left - rect.right) / 2 }
      })
      assert.ok(bounds.left >= 12 && bounds.right <= width - 12, JSON.stringify(bounds))
      assert.ok(Math.abs(bounds.centerDelta) <= 1, JSON.stringify(bounds))
      assert.ok(bounds.scrollWidth <= width, JSON.stringify(bounds))
    }
  }
  assert.deepEqual(errors, [])
})

test('overflowing classic tabs remain reachable from both ends using roving keyboard focus', async () => {
  await page.setViewportSize({ width: 210, height: 700 })
  await page.evaluate(() => { window.ribbonFixture.longTabTitles.value = true })
  const tabs = page.getByRole('tab')
  await tabs.first().focus()
  await expect(tabs.first()).toBeInViewport()
  await tabs.first().press('End')
  await expect(tabs.last()).toBeFocused()
  await expect(tabs.last()).toBeInViewport()
  await tabs.last().press('Home')
  await expect(tabs.first()).toBeFocused()
  await expect(tabs.first()).toBeInViewport()
  assert.equal(await page.locator('[role="tab"][tabindex="0"]').count(), 1)
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true)
  assert.deepEqual(errors, [])
})

test('context category removal returns to the previous regular tab without dispatching content edits', async () => {
  await page.getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('tab', { name: '表格', exact: true }).click()
  await page.evaluate(() => { window.ribbonFixture.showContext.value = false })
  await page.waitForFunction(() => window.ribbonFixture.active.value === 'insert')
  assert.equal(await page.getByRole('tab', { name: '插入', exact: true }).getAttribute('aria-selected'), 'true')
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  assert.deepEqual(errors, [])
})

test('a menu closes when the editor context becomes disabled and cannot keep active stale choices', async () => {
  const trigger = page.getByRole('button', { name: '字体颜色', exact: true })
  await trigger.click()
  const popup = page.getByRole('menu', { name: '字体颜色', exact: true })
  await popup.waitFor()
  await page.evaluate(() => { window.ribbonFixture.locked.value = true })
  await popup.waitFor({ state: 'hidden' })
  assert.equal(await trigger.isDisabled(), true)
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  assert.deepEqual(errors, [])
})

test('resize tracking survives temporarily empty menus and does not reopen stale overflow popups', async () => {
  await page.evaluate(() => { window.ribbonFixture.menusVisible.value = false })
  await page.getByTestId('workbench-ribbon').waitFor({ state: 'detached' })
  await page.evaluate(() => { window.ribbonFixture.menusVisible.value = true })
  await page.getByRole('tab', { name: '开始', exact: true }).waitFor()
  await page.setViewportSize({ width: 210, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.click()
  await page.getByRole('menu', { name: '更多功能', exact: true }).waitFor()
  await page.setViewportSize({ width: 1200, height: 700 })
  await more.waitFor({ state: 'detached' })
  await page.setViewportSize({ width: 210, height: 700 })
  await more.waitFor()
  assert.equal(await more.getAttribute('aria-expanded'), 'false')
  assert.equal(await page.getByRole('menu', { name: '更多功能', exact: true }).count(), 0)
  assert.deepEqual(errors, [])
})

test('popup-component menus remain available and dispatch their payload in both Ribbon layouts', async () => {
  for (const variant of ['classic', 'grid']) {
    await page.evaluate(value => { window.ribbonFixture.variant.value = value; window.ribbonFixture.popupTools.value = true }, variant)
    await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
    const trigger = page.getByRole('button', { name: '符号选择', exact: true })
    await trigger.click()
    const menu = page.getByRole('menu', { name: '符号选择', exact: true })
    await expect(menu).toBeVisible()
    await menu.getByRole('button', { name: '插入测试符号', exact: true }).click()
    await expect(menu).toBeHidden()
    assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['doc.command.insert-symbol', 'chosen-symbol'])
    await page.getByRole('tab', { name: '开始', exact: true }).click()
  }
  assert.deepEqual(errors, [])
})

test('sibling menus and selects share one open popup without closing the active sibling on state updates', async () => {
  await page.evaluate(() => { window.ribbonFixture.popupTools.value = true })
  await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
  const first = page.getByRole('button', { name: '符号选择', exact: true })
  const second = page.getByRole('button', { name: '其他选项', exact: true })
  await first.click()
  await expect(page.getByRole('menu', { name: '符号选择', exact: true })).toBeVisible()
  await second.click()
  await expect(page.getByRole('menu', { name: '符号选择', exact: true })).toBeHidden()
  await expect(page.getByRole('menu', { name: '其他选项', exact: true })).toBeVisible()
  await page.evaluate(() => { window.ribbonFixture.locked.value = true })
  await expect(page.getByRole('menu', { name: '其他选项', exact: true })).toBeVisible()
  await page.getByRole('combobox', { name: '选择测试', exact: true }).click()
  await expect(page.getByRole('menu', { name: '其他选项', exact: true })).toBeHidden()
  await expect(page.getByRole('listbox')).toBeVisible()
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await expect(page.getByRole('listbox')).toBeHidden()
  assert.deepEqual(errors, [])
})

test('split dropdown keyboard navigation returns Escape focus to the arrow trigger', async () => {
  await page.evaluate(() => { window.ribbonFixture.popupTools.value = true })
  await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
  const arrow = page.getByRole('button', { name: '分离菜单选项', exact: true })
  await arrow.focus()
  await arrow.press('ArrowUp')
  await expect(page.getByRole('button', { name: '最后选项', exact: true })).toBeFocused()
  await page.keyboard.press('Home')
  await expect(page.getByRole('button', { name: '第一选项', exact: true })).toBeFocused()
  await page.keyboard.press('Escape')
  await expect(arrow).toBeFocused()
  await expect(page.getByRole('menu', { name: '分离菜单', exact: true })).toBeHidden()
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  assert.deepEqual(errors, [])
})

test('grid item slots retain dispatch guards and overflow placement', async () => {
  await page.evaluate(() => { window.ribbonFixture.popupTools.value = true; window.ribbonFixture.customSlot.value = true; window.ribbonFixture.variant.value = 'grid' })
  await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
  await page.getByRole('button', { name: '自定义动作', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['doc.command.custom', { overflow: false }])
  await page.setViewportSize({ width: 210, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.focus()
  await more.press('ArrowDown')
  const popup = page.getByRole('menu', { name: '更多功能', exact: true })
  await popup.getByRole('button', { name: '自定义动作', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['doc.command.custom', { overflow: true }])
  await page.evaluate(() => { window.ribbonFixture.locked.value = true })
  await expect(popup.getByRole('button', { name: '自定义动作', exact: true })).toBeDisabled()
  assert.deepEqual(errors, [])
})

test('ConfigProvider theme and direction reach Ribbon controls and portaled command menus', async () => {
  const root = page.getByTestId('workbench-ribbon')
  const light = await root.evaluate(element => ({ background: getComputedStyle(element).backgroundColor, text: getComputedStyle(element).color }))
  await page.evaluate(() => { window.ribbonFixture.dark.value = true; window.ribbonFixture.direction.value = 'rtl' })
  await expect.poll(() => root.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light.background)
  const darkText = await root.evaluate(element => getComputedStyle(element).color)
  assert.notEqual(darkText, light.text)
  await page.getByRole('button', { name: '字体颜色', exact: true }).click()
  const menu = page.getByRole('menu', { name: '字体颜色', exact: true })
  assert.equal(await menu.getAttribute('dir'), 'rtl')
  const red = menu.getByRole('button', { name: '红色', exact: true })
  assert.equal(await red.evaluate(element => getComputedStyle(element).color), darkText)
  await page.evaluate(() => { window.ribbonFixture.dark.value = false })
  await expect.poll(() => red.evaluate(element => getComputedStyle(element).color)).toBe(light.text)
  assert.deepEqual(errors, [])
})
