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
import DocumentColorControl from '/packages/office-studio/src/components/document/DocumentColorControl.vue'
import SheetColorControl from '/packages/office-studio/src/components/spreadsheet/SheetColorControl.vue'
import SheetNumberFormatControl from '/packages/office-studio/src/components/spreadsheet/SheetNumberFormatControl.vue'
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
const geometryTools = ref(false)
const longCaptions = ref(false)
const PopupPicker = {
  emits: ['insert'],
  setup(_props, { emit }) {
    return () => h(Button, { onClick: () => emit('insert', 'chosen-symbol') }, () => '插入测试符号')
  },
}
window.ribbonInvocations = []
window.ribbonFixture = { active, variant, locked, showContext, menusVisible, direction, family, bold, longTabTitles, popupTools, dark, customSlot, geometryTools, longCaptions }
window.ribbonContextInvocations = []
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
      { id: 'popup-large-menu', type: 'menu', title: '大号菜单', icon: Bold, children: [{ id: 'popup-large-action', type: 'button', title: '大号菜单动作', commandId: 'doc.command.large-action' }] },
    ] },
    { id: 'popup-custom', items: [{ id: 'custom-tool', type: 'button', title: '自定义工具', hidden: !customSlot.value, disabled: locked.value }] },
  ] },
  { id: 'data', title: '数据', order: 4, hidden: !geometryTools.value, groups: [{ id: 'data-alignment', items: [
    { id: 'data-split', type: 'split', title: '数据拆分', icon: Bold, commandId: 'sheet.command.data-main', children: [
      { id: 'data-split-option', type: 'button', title: '数据选项', commandId: 'sheet.command.data-option' },
    ] },
  ] }] },
  { id: 'geometry-tools', title: '控件尺寸', order: 5, hidden: !geometryTools.value, groups: [
    { id: 'geometry-font', items: [
      { id: 'editable-font', type: 'select', edit: { kind: 'font' }, title: '编辑字体', commandId: 'doc.command.font', value: 'Arial', disabled: locked.value, width: 180, options: [{ label: 'Arial', value: 'Arial' }] },
      { id: 'row-zoom', type: 'select', title: '缩放', commandId: 'sheet.operation.zoom', value: 100, disabled: locked.value, width: 88, options: [{ label: '100%', value: 100 }] },
    ] },
    { id: 'geometry-color', items: [
      { id: 'document-color', type: 'split', title: longCaptions.value ? '文档字体颜色与当前选区颜色的长名称' : '文档颜色', icon: Bold, inlineComponent: DocumentColorControl,
        popupCommandId: 'doc.command.format-text', popupProps: { title: longCaptions.value ? '文档字体颜色与当前选区颜色的长名称' : '文档颜色', field: 'color', icon: Bold, defaultColor: '#000000' } },
      { id: 'sheet-color', type: 'split', title: longCaptions.value ? '表格填充颜色与当前选区颜色的长名称' : '表格颜色', icon: Bold, inlineComponent: SheetColorControl,
        popupProps: { title: longCaptions.value ? '表格填充颜色与当前选区颜色的长名称' : '表格颜色', field: 'background', commandId: 'sheet.command.set-background', icon: Bold, defaultColor: '#315bc2' } },
    ] },
    { id: 'geometry-format', items: [
      { id: 'number-format', type: 'select', title: '数字格式', inlineComponent: SheetNumberFormatControl, popupProps: { pattern: '0.00', title: '数字格式' }, popupCommandId: 'sheet.command.set-number-format' },
    ] },
    { id: 'geometry-dropdown', items: [
      { id: 'row-dropdown', type: 'menu', title: '下拉菜单', icon: Bold, children: [{ id: 'row-action', type: 'button', title: '下拉菜单动作', commandId: 'doc.command.row-action' }] },
      { id: 'caption-dropdown', type: 'menu', title: longCaptions.value ? '带有很长标题文字的下拉菜单' : '标题菜单', icon: Bold, children: [{ id: 'caption-action', type: 'button', title: '标题菜单动作', commandId: 'doc.command.caption-action' }] },
    ] },
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
  'popup-large-menu': { row: 1, column: 4, rowSpan: 2, width: 84, showLabel: true },
  'custom-tool': { row: 1, column: 1, rowSpan: 2, width: 120 },
  'data-split': { row: 1, column: 1, width: 56, showLabel: false },
  'editable-font': { row: 1, column: 1, width: 180, minWidth: 120 },
  'row-zoom': { row: 2, column: 1, width: 88, showLabel: true },
  'document-color': { row: 1, column: 1, width: 40, minWidth: 38 },
  'sheet-color': { row: 2, column: 1, width: 40, minWidth: 38 },
  'number-format': { row: 1, column: 1, width: 160 },
  'row-dropdown': { row: 1, column: 1, width: 40 },
  'caption-dropdown': { row: 2, column: 1, width: 112, minWidth: 88, showLabel: true },
}
const gridTabs = computed(() => tabs.value.map((tab) => ({ ...tab, groups: tab.groups.map((group) => ({ ...group, items: group.items.map((item) => ({ ...item, gridLayout: gridLayoutById[item.id] })) })) })))
createApp({ render() { return h(ConfigProvider, { theme: { algorithm: dark.value ? theme.darkAlgorithm : theme.defaultAlgorithm, token: { colorPrimary: '#315bc2' } }, direction: direction.value }, { default: () => h(WorkbenchRibbon, {
  tabs: gridTabs.value, variant: variant.value, modelValue: active.value, direction: direction.value,
  featureSearchScope: document.getElementById('app'),
  shortcuts: [{ title: '重做', shortcut: 'Ctrl+Shift+Z' }, { title: '切换工作表', shortcut: 'Control+PageDown' }],
  featureSearchContext: () => [{ id: 'context', items: [{ id: 'context-style', type: 'select', title: '上下文段落样式', options: [
    { value: 'heading', label: '搜索测试标题', commandId: 'doc.command.set-paragraph-named-style', params: { namedStyle: 'heading1' } },
  ] }] }],
  onContextCommand: (key, commandId, params) => window.ribbonContextInvocations.push([key, commandId, params]),
  'onUpdate:modelValue': (value) => active.value = value,
  onExecute: (id, params) => { window.ribbonInvocations.push([id, params]); if (id === 'doc.command.font') family.value = params; if (id === 'doc.command.bold') bold.value = !bold.value },
}, { 'tab-extra': () => null, 'item-custom-tool': ({ item, execute, overflow }) => h(Button, { disabled: item.disabled, onClick: () => execute('doc.command.custom', { overflow: Boolean(overflow) }) }, () => '自定义动作') }) }) } }).mount('#app')
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
  const selectedTab = page.getByRole('tab', { name: '开始', exact: true })
  const tabBounds = await selectedTab.boundingBox()
  await selectedTab.hover()
  await expect.poll(() => selectedTab.evaluate(element => getComputedStyle(element).backgroundColor)).toBe(geometry.selectedTabBackground)
  assert.deepEqual(await selectedTab.evaluate(element => {
    const style = getComputedStyle(element)
    return [style.borderTopWidth, style.borderRightWidth, style.borderBottomWidth, style.borderLeftWidth]
  }), ['0px', '0px', '0px', '0px'])
  assert.deepEqual(await selectedTab.boundingBox(), tabBounds)
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
  const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
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
  const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
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
  const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
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
        const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
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
  await page.getByRole('dialog', { name: '更多功能', exact: true }).waitFor()
  await page.setViewportSize({ width: 1200, height: 700 })
  await more.waitFor({ state: 'detached' })
  await page.setViewportSize({ width: 210, height: 700 })
  await more.waitFor()
  assert.equal(await more.getAttribute('aria-expanded'), 'false')
  assert.equal(await page.getByRole('dialog', { name: '更多功能', exact: true }).count(), 0)
  assert.deepEqual(errors, [])
})

test('popup-component menus remain available and dispatch their payload in both Ribbon layouts', async () => {
  for (const variant of ['classic', 'grid']) {
    await page.evaluate(value => { window.ribbonFixture.variant.value = value; window.ribbonFixture.popupTools.value = true }, variant)
    await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
    const trigger = page.getByRole('button', { name: '符号选择', exact: true })
    await trigger.click()
    const menu = page.getByRole('dialog', { name: '符号选择', exact: true })
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
  await expect(page.getByRole('dialog', { name: '符号选择', exact: true })).toBeVisible()
  await second.click()
  await expect(page.getByRole('dialog', { name: '符号选择', exact: true })).toBeHidden()
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
  const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
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

function measureRibbonButtons(buttons) {
  return buttons.map(button => {
    const bounds = button.getBoundingClientRect()
    const style = getComputedStyle(button)
    const arrow = button.querySelector('.ribbon-dropdown-chevron, .ribbon-grid-more-label svg')
      ?? (button.closest('.document-color-arrow, .sheet-color-arrow') || button.classList.contains('sheet-number-format-trigger') ? button.querySelector('svg') : null)
    const arrowBounds = arrow?.getBoundingClientRect()
    const captionBounds = (arrow?.closest('.ribbon-dropdown-caption, .ribbon-grid-more-label') ?? button).getBoundingClientRect()
    return {
      title: button.getAttribute('aria-label'), width: bounds.width, height: bounds.height,
      borders: [style.borderTopWidth, style.borderRightWidth, style.borderBottomWidth, style.borderLeftWidth],
      hit: document.elementFromPoint(bounds.left + bounds.width / 2, bounds.top + bounds.height / 2)?.closest('button') === button,
      contents: [...button.querySelectorAll('svg, .ribbon-button-label, .ribbon-dropdown-caption, .ribbon-grid-more-label, .document-color-bar, .sheet-color-bar, .document-color-main-label, .sheet-color-main-label, .sheet-number-format-trigger > span')].map(element => {
        const content = element.getBoundingClientRect()
        return { className: element.getAttribute('class'), within: content.left >= bounds.left - 0.5 && content.right <= bounds.right + 0.5 && content.top >= bounds.top - 0.5 && content.bottom <= bounds.bottom + 0.5 }
      }),
      arrow: arrowBounds ? { width: arrowBounds.width, height: arrowBounds.height, centerOffset: (arrowBounds.top + arrowBounds.bottom - captionBounds.top - captionBounds.bottom) / 2,
        horizontalOffset: button.classList.contains('ribbon-split-arrow') || button.closest('.document-color-arrow, .sheet-color-arrow') ? (arrowBounds.left + arrowBounds.right - bounds.left - bounds.right) / 2 : 0 } : null,
    }
  })
}

function assertRibbonButtonGeometry(buttons) {
  for (const button of buttons) {
    assert.equal(button.hit, true, JSON.stringify(button))
    assert.ok(button.contents.every(content => content.within), JSON.stringify(button))
    if (!button.arrow) continue
    assert.equal(button.arrow.width, 12, JSON.stringify(button))
    assert.equal(button.arrow.height, 12, JSON.stringify(button))
    assert.ok(Math.abs(button.arrow.centerOffset) <= 0.5, JSON.stringify(button))
    assert.ok(Math.abs(button.arrow.horizontalOffset) <= 0.05, JSON.stringify(button))
  }
}

test.each([['classic', 'ltr', false], ['classic', 'rtl', true], ['grid', 'ltr', false], ['grid', 'rtl', true]])('%s %s dark %s dropdown captions and split arrows fit their toolbar and overflow hover rectangles', async (variant, direction, dark) => {
  await page.evaluate(({ variant, direction, dark }) => {
    window.ribbonFixture.variant.value = variant
    window.ribbonFixture.direction.value = direction
    window.ribbonFixture.dark.value = dark
    window.ribbonFixture.popupTools.value = true
  }, { variant, direction, dark })
  await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
  const toolbar = page.getByRole('toolbar', { name: '弹层工具', exact: true })
  const buttons = await toolbar.locator('button:visible').evaluateAll(measureRibbonButtons)
  assert.equal(buttons.length, 5)
  assertRibbonButtonGeometry(buttons)
  for (const button of buttons) assert.deepEqual(button.borders, ['0px', '0px', '0px', '0px'])
  const arrow = toolbar.getByRole('button', { name: '分离菜单选项', exact: true })
  const main = toolbar.getByRole('button', { name: '分离菜单', exact: true })
  const mainBounds = await main.boundingBox()
  const arrowBounds = await arrow.boundingBox()
  assert.equal(arrowBounds.height, mainBounds.height)
  assert.equal(arrowBounds.height, variant === 'grid' ? 52 : 28)
  assert.equal(arrowBounds.width, variant === 'grid' ? 16 : 24)
  const initial = await arrow.evaluate(element => getComputedStyle(element).backgroundColor)
  await arrow.hover()
  await expect.poll(() => arrow.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(initial)
  await expect.poll(() => arrow.evaluate(element => element.getAnimations().length)).toBe(0)
  assert.equal(await main.evaluate(element => element.matches(':hover')), false)
  const hoverBackground = await arrow.evaluate(element => getComputedStyle(element).backgroundColor)
  assert.deepEqual(await arrow.boundingBox(), arrowBounds)
  await page.mouse.down()
  assert.equal(await arrow.evaluate(element => element.matches(':active')), true)
  const pressed = await arrow.evaluateAll(measureRibbonButtons)
  assertRibbonButtonGeometry(pressed)
  assert.deepEqual(pressed[0].borders, ['0px', '0px', '0px', '0px'])
  assert.deepEqual(await arrow.boundingBox(), arrowBounds)
  await page.mouse.move(1, 1)
  await page.mouse.up()
  await arrow.press('ArrowDown')
  await expect(arrow).toHaveAttribute('aria-expanded', 'true')
  await expect.poll(() => arrow.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(hoverBackground)
  await expect.poll(() => arrow.evaluate(element => element.getAnimations().length)).toBe(0)
  const openBackground = await arrow.evaluate(element => getComputedStyle(element).backgroundColor)
  await arrow.hover()
  assert.equal(await arrow.evaluate(element => getComputedStyle(element).backgroundColor), openBackground)
  assert.deepEqual(await arrow.boundingBox(), arrowBounds)
  await arrow.press('Escape')
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  const textSplit = page.getByRole('toolbar', { name: '开始', exact: true }).locator('[data-ribbon-item="list-toggle"]')
  await expect(textSplit).toBeVisible()
  assertRibbonButtonGeometry(await textSplit.locator('button').evaluateAll(measureRibbonButtons))
  const textArrow = textSplit.getByRole('button', { name: '列表选项', exact: true })
  assert.deepEqual(await textArrow.evaluate(element => {
    const style = getComputedStyle(element)
    return [style.paddingInlineStart, style.paddingInlineEnd]
  }), ['0px', '0px'])
  assert.equal((await textArrow.boundingBox()).width, variant === 'grid' ? 16 : 24)
  await textArrow.hover()
  assert.equal(await textSplit.getByRole('button', { name: '列表', exact: true }).evaluate(element => element.matches(':hover')), false)
  await page.getByRole('tab', { name: '弹层工具', exact: true }).click()
  await page.setViewportSize({ width: 200, height: 700 })
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  await more.focus()
  await more.press('ArrowDown')
  const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
  await expect(popup).toBeVisible()
  assertRibbonButtonGeometry(await popup.locator('button:visible').evaluateAll(measureRibbonButtons))
  const overflowArrow = popup.getByRole('button', { name: '分离菜单选项', exact: true })
  const overflowMain = popup.getByRole('button', { name: '分离菜单', exact: true })
  assert.equal((await overflowArrow.boundingBox()).width, 24)
  assert.equal((await overflowArrow.boundingBox()).height, 28)
  const overflowInitial = await overflowArrow.evaluate(element => getComputedStyle(element).backgroundColor)
  await overflowArrow.hover()
  await expect.poll(() => overflowArrow.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(overflowInitial)
  assert.equal(await overflowMain.evaluate(element => element.matches(':hover')), false)
  await page.keyboard.press('Escape')
  await page.getByRole('tab', { name: '开始', exact: true }).click()
  await more.focus()
  await more.press('ArrowDown')
  const overflowTextSplit = popup.locator('[data-ribbon-item="list-toggle"]')
  await expect(overflowTextSplit).toBeVisible()
  assertRibbonButtonGeometry(await overflowTextSplit.locator('button').evaluateAll(measureRibbonButtons))
  assert.deepEqual(await overflowTextSplit.getByRole('button', { name: '列表选项', exact: true }).evaluate(element => {
    const style = getComputedStyle(element)
    return [element.getBoundingClientRect().width, style.paddingInlineStart, style.paddingInlineEnd]
  }), [24, '0px', '0px'])
  assert.deepEqual(errors, [])
})

test.each(['ltr', 'rtl'])('data grid %s aligns the main action to the start and centers its independent split arrow', async direction => {
  await page.evaluate(value => {
    window.ribbonFixture.variant.value = 'grid'
    window.ribbonFixture.direction.value = value
    window.ribbonFixture.geometryTools.value = true
  }, direction)
  await page.getByRole('tab', { name: '数据', exact: true }).click()
  const item = page.locator('.ribbon-grid-toolbar [data-ribbon-item="data-split"]')
  await expect(item).toHaveClass(/is-grid-align-start/)
  assertRibbonButtonGeometry(await item.locator('button').evaluateAll(measureRibbonButtons))
  const main = item.getByRole('button', { name: '数据拆分', exact: true })
  const arrow = item.getByRole('button', { name: '数据拆分选项', exact: true })
  assert.equal(await main.evaluate(element => getComputedStyle(element).justifyContent), 'flex-start')
  assert.equal(await arrow.evaluate(element => getComputedStyle(element).justifyContent), 'center')
  assert.equal((await arrow.boundingBox()).width, 16)
  assert.equal((await arrow.boundingBox()).height, 24)
  await arrow.hover()
  assert.equal(await main.evaluate(element => element.matches(':hover')), false)
  await arrow.click()
  await page.getByRole('menuitem', { name: '数据选项', exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [['sheet.command.data-option', undefined]])
  assert.deepEqual(errors, [])
})

test.each([['classic', 'ltr', 1200], ['classic', 'rtl', 200], ['grid', 'rtl', 1200], ['grid', 'ltr', 200]])('%s %s color presets keep square swatches within their hover padding at %spx', async (variant, direction, width) => {
  await page.setViewportSize({ width, height: 700 })
  await page.evaluate(({ variant, direction }) => {
    window.ribbonFixture.variant.value = variant
    window.ribbonFixture.direction.value = direction
    window.ribbonFixture.geometryTools.value = true
  }, { variant, direction })
  await page.getByRole('tab', { name: '控件尺寸', exact: true }).click()
  if (width === 200) {
    const more = page.getByRole('button', { name: '更多功能', exact: true })
    await more.focus()
    await more.press('ArrowDown')
  }
  const scope = width === 200 ? page.getByRole('dialog', { name: '更多功能', exact: true }) : page.getByRole('toolbar', { name: '控件尺寸', exact: true })
  for (const product of ['document', 'sheet']) {
    await scope.locator('[data-ribbon-item="' + product + '-color"] .' + product + '-color-arrow button').click()
    const panel = page.locator('.' + product + '-color-panel:visible')
    await expect(panel).toBeVisible()
    const presets = panel.locator('.' + product + '-color-preset')
    await expect(presets).toHaveCount(45)
    const firstColor = await presets.first().getAttribute('aria-label')
    assert.equal(await presets.first().getAttribute('title'), null)
    for (const dark of [false, true]) {
      await page.evaluate(value => { window.ribbonFixture.dark.value = value }, dark)
      await page.mouse.move(0, 650)
      const bounds = await presets.first().boundingBox()
      const background = await presets.first().evaluate(element => getComputedStyle(element).backgroundColor)
      await presets.first().hover()
      await expect.poll(() => presets.first().evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(background)
      await expect(page.locator('.ant-tooltip-inner:visible')).toHaveText(firstColor)
      assert.deepEqual(await presets.first().boundingBox(), bounds)
      const geometry = await presets.evaluateAll(buttons => buttons.map(button => {
        const bounds = button.getBoundingClientRect()
        const style = getComputedStyle(button)
        const swatch = button.firstElementChild.getBoundingClientRect()
        return {
          width: bounds.width, height: bounds.height, swatchWidth: swatch.width, swatchHeight: swatch.height,
          withinPadding: swatch.left >= bounds.left + Number.parseFloat(style.paddingLeft) + Number.parseFloat(style.borderLeftWidth) - 0.5
            && swatch.right <= bounds.right - Number.parseFloat(style.paddingRight) - Number.parseFloat(style.borderRightWidth) + 0.5
            && swatch.top >= bounds.top + Number.parseFloat(style.paddingTop) + Number.parseFloat(style.borderTopWidth) - 0.5
            && swatch.bottom <= bounds.bottom - Number.parseFloat(style.paddingBottom) - Number.parseFloat(style.borderBottomWidth) + 0.5,
          centerX: (swatch.left + swatch.right - bounds.left - bounds.right) / 2,
          centerY: (swatch.top + swatch.bottom - bounds.top - bounds.bottom) / 2,
        }
      }))
      for (const preset of geometry) {
        assert.equal(preset.height, 32, JSON.stringify(preset))
        assert.ok(preset.width <= 32.5, JSON.stringify(preset))
        assert.ok(Math.abs(preset.swatchWidth - preset.swatchHeight) <= 0.5, JSON.stringify(preset))
        assert.ok(preset.swatchWidth > 0 && preset.swatchWidth <= 24.5, JSON.stringify(preset))
        assert.equal(preset.withinPadding, true, JSON.stringify(preset))
        assert.ok(Math.abs(preset.centerX) <= 0.5 && Math.abs(preset.centerY) <= 0.5, JSON.stringify(preset))
        if (width === 1200) assert.equal(preset.swatchWidth, 24, JSON.stringify(preset))
      }
      await page.mouse.move(0, 650)
      await presets.first().focus()
      await expect(page.locator('.ant-tooltip-inner:visible')).toHaveText(firstColor)
      await presets.first().evaluate(button => button.blur())
    }
    await page.keyboard.press('Escape')
    await expect(panel).toHaveCount(0)
  }
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  assert.deepEqual(errors, [])
})

test.each(['ltr', 'rtl'])('compressed grid %s keeps editable inputs, row labels and both color controls inside their cells', async direction => {
  await page.setViewportSize({ width: 520, height: 700 })
  await page.evaluate(value => {
    window.ribbonFixture.direction.value = value
    window.ribbonFixture.variant.value = 'grid'
    window.ribbonFixture.geometryTools.value = true
  }, direction)
  await page.getByRole('tab', { name: '控件尺寸', exact: true }).click()
  const toolbar = page.getByTestId('ribbon-grid-toolbar')
  await expect.poll(() => toolbar.evaluate(element => element.style.getPropertyValue('--ribbon-grid-density'))).toBe('1')
  assert.equal(await toolbar.locator('.ribbon-grid-group:not(.is-hidden)').count(), 4)
  const geometry = await toolbar.evaluate(element => {
    const fontElement = element.querySelector('[data-ribbon-item="editable-font"] .ribbon-edit')
    const font = fontElement.getBoundingClientRect()
    const fontStyle = getComputedStyle(fontElement)
    const input = element.querySelector('[data-ribbon-item="editable-font"] .ant-select').getBoundingClientRect()
    const zoom = element.querySelector('[data-ribbon-item="row-zoom"]').getBoundingClientRect()
    const title = element.querySelector('[data-ribbon-item="row-zoom"] label').getBoundingClientRect()
    const select = element.querySelector('[data-ribbon-item="row-zoom"] .ribbon-select').getBoundingClientRect()
    return { fontWidth: font.width, inputWidth: input.width, inputHeight: input.height, fontHeight: font.height,
      fontContentWidth: font.width - Number.parseFloat(fontStyle.borderInlineStartWidth) - Number.parseFloat(fontStyle.borderInlineEndWidth),
      fontContentHeight: font.height - Number.parseFloat(fontStyle.borderTopWidth) - Number.parseFloat(fontStyle.borderBottomWidth),
      labelWithinRow: title.top >= zoom.top && title.bottom <= zoom.bottom,
      selectWithinRow: select.top >= zoom.top && select.bottom <= zoom.bottom,
      rowHeight: zoom.height,
      colors: [...element.querySelectorAll('.document-color-control, .sheet-color-control')].map(control => ({ width: control.getBoundingClientRect().width, height: control.getBoundingClientRect().height })),
    }
  })
  assert.equal(geometry.fontWidth, 120)
  assert.equal(geometry.inputWidth, geometry.fontContentWidth)
  assert.equal(geometry.inputHeight, geometry.fontContentHeight)
  assert.equal(geometry.fontHeight, 24)
  assert.equal(geometry.rowHeight, 24)
  assert.equal(geometry.labelWithinRow, true)
  assert.equal(geometry.selectWithinRow, true)
  assert.deepEqual(geometry.colors, [{ width: 38, height: 24 }, { width: 38, height: 24 }])
  assertRibbonButtonGeometry(await toolbar.locator('button:visible').evaluateAll(measureRibbonButtons))
  const arrow = toolbar.locator('.document-color-arrow button')
  const light = await arrow.evaluate(element => getComputedStyle(element).backgroundColor)
  await arrow.hover()
  await expect.poll(() => arrow.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(light)
  const lightHover = await arrow.evaluate(element => getComputedStyle(element).backgroundColor)
  await page.evaluate(() => { window.ribbonFixture.dark.value = true })
  await expect.poll(() => arrow.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(lightHover)
  assertRibbonButtonGeometry(await toolbar.locator('button:visible').evaluateAll(measureRibbonButtons))
  assert.deepEqual(errors, [])
})

test.each(['classic', 'grid'])('%s disabled editable and standard selectors keep their background while hovered', async variant => {
  await page.evaluate(value => {
    window.ribbonFixture.variant.value = value
    window.ribbonFixture.geometryTools.value = true
    window.ribbonFixture.locked.value = true
  }, variant)
  await page.getByRole('tab', { name: '控件尺寸', exact: true }).click()
  const toolbar = page.getByRole('toolbar', { name: '控件尺寸', exact: true })
  for (const dark of [false, true]) {
    await page.evaluate(value => { window.ribbonFixture.dark.value = value }, dark)
    for (const id of ['editable-font', 'row-zoom']) {
      const selector = toolbar.locator('[data-ribbon-item="' + id + '"] .ribbon-select')
      await page.mouse.move(0, 650)
      const initial = await selector.evaluate(element => getComputedStyle(element).backgroundColor)
      await selector.hover()
      assert.equal(await selector.evaluate(element => getComputedStyle(element).backgroundColor), initial)
      assert.equal(await selector.evaluate(element => element.closest('[data-ribbon-item]').classList.contains('is-disabled')), true)
    }
  }
  await expect(toolbar.getByRole('combobox', { name: '编辑字体', exact: true })).toBeDisabled()
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  assert.deepEqual(errors, [])
})

test.each([['classic', 'ltr', false], ['classic', 'rtl', true], ['grid', 'ltr', false], ['grid', 'rtl', true]])('%s %s overflow contains long captions, selectors, colors and number-format hover areas', async (variant, direction, dark) => {
  await page.setViewportSize({ width: 200, height: 700 })
  await page.evaluate(({ variant, direction, dark }) => {
    window.ribbonFixture.variant.value = variant
    window.ribbonFixture.direction.value = direction
    window.ribbonFixture.dark.value = dark
    window.ribbonFixture.geometryTools.value = true
    window.ribbonFixture.longCaptions.value = true
  }, { variant, direction, dark })
  await page.getByRole('tab', { name: '控件尺寸', exact: true }).click()
  const more = page.getByRole('button', { name: '更多功能', exact: true })
  const initial = await more.evaluate(element => getComputedStyle(element).backgroundColor)
  const normalBounds = await more.boundingBox()
  await more.hover()
  await expect.poll(() => more.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(initial)
  await expect.poll(() => more.evaluate(element => element.getAnimations().length)).toBe(0)
  const hoverBackground = await more.evaluate(element => getComputedStyle(element).backgroundColor)
  const moreBounds = await more.boundingBox()
  assert.deepEqual(moreBounds, normalBounds)
  const tooltip = page.getByRole('tooltip', { name: '更多功能', exact: true })
  await expect(tooltip).toBeVisible()
  assertRibbonButtonGeometry(await more.evaluateAll(measureRibbonButtons))
  assert.deepEqual((await more.evaluateAll(measureRibbonButtons))[0].borders, ['0px', '0px', '0px', '0px'])
  await more.focus()
  await more.press('ArrowDown')
  const popup = page.getByRole('dialog', { name: '更多功能', exact: true })
  await expect(popup).toBeVisible()
  await expect(tooltip).toBeHidden()
  assert.equal(await more.getAttribute('aria-expanded'), 'true')
  await expect.poll(() => more.evaluate(element => getComputedStyle(element).backgroundColor)).not.toBe(hoverBackground)
  await expect.poll(() => more.evaluate(element => element.getAnimations().length)).toBe(0)
  const openBackground = await more.evaluate(element => getComputedStyle(element).backgroundColor)
  await more.hover()
  assert.equal(await more.evaluate(element => getComputedStyle(element).backgroundColor), openBackground)
  assert.deepEqual(await more.boundingBox(), moreBounds)
  await expect(tooltip).toBeHidden()
  assert.equal(await popup.locator('[data-ribbon-item]').count(), 7)
  assertRibbonButtonGeometry(await popup.locator('button').evaluateAll(measureRibbonButtons))
  const colors = await popup.locator('.document-color-control, .sheet-color-control').evaluateAll(elements => elements.map(element => {
    const label = element.querySelector('.document-color-main-label, .sheet-color-main-label')
    return { height: element.getBoundingClientRect().height, ellipsis: getComputedStyle(label).textOverflow, clipped: label.scrollWidth > label.clientWidth }
  }))
  assert.deepEqual(colors, [{ height: 28, ellipsis: 'ellipsis', clipped: true }, { height: 28, ellipsis: 'ellipsis', clipped: true }])
  const selector = await popup.locator('[data-ribbon-item="editable-font"] .ant-select').boundingBox()
  assert.equal(selector.height, 28)
  const numberFormat = popup.getByRole('button', { name: '数字格式', exact: true })
  assert.equal((await numberFormat.boundingBox()).height, 28)
  await numberFormat.click()
  await page.getByRole('menu', { name: '数字格式', exact: true }).getByText('百分比', { exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations.at(-1)), ['sheet.command.set-number-format', { pattern: '0.00%' }])
  await expect(popup).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(more).toBeFocused()
  assert.equal(await more.getAttribute('aria-expanded'), 'false')
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true)
  assert.deepEqual(errors, [])
})

test('feature search forwards the context selector command and complete parameters, and the shortcut panel lists editor shortcuts', async () => {
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await page.keyboard.press('Control+Shift+P')
  const search = page.getByRole('dialog', { name: '搜索功能', exact: true })
  await expect(search).toBeVisible()
  await search.getByRole('combobox', { name: '搜索功能', exact: true }).fill('搜索测试标题')
  await search.getByText('搜索测试标题', { exact: true }).click()
  await expect(search).toBeHidden()
  assert.deepEqual(await page.evaluate(() => window.ribbonContextInvocations), [['context-style', 'doc.command.set-paragraph-named-style', { namedStyle: 'heading1' }]])
  assert.deepEqual(await page.evaluate(() => window.ribbonInvocations), [])
  await page.getByRole('tab', { name: '开始', exact: true }).focus()
  await page.keyboard.press('Control+Backslash')
  const panel = page.getByRole('dialog', { name: '快捷键面板', exact: true })
  await expect(panel).toBeVisible()
  await expect(panel.getByText('重做', { exact: true })).toBeVisible()
  await expect(panel.getByText('Ctrl+Shift+Z', { exact: true })).toBeVisible()
  await expect(panel.getByText('切换工作表', { exact: true })).toBeVisible()
  await expect(panel.getByText('Control+PageDown', { exact: true })).toBeVisible()
  assert.deepEqual(errors, [])
})
