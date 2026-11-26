import assert from 'node:assert/strict'
import { test } from 'vitest'
import { CLASSIC_RIBBON, GRID_RIBBON, fitRibbonGridGroups, fitRibbonItems, groupRibbonGridOverflowItems, nextRibbonTab, placeRibbonGridItems, resolveRibbonTab, ribbonInvocation, ribbonPlacementKey, simpleRibbonTab, visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'
import { collectFeatureSearchResults, featureSearchRibbonItem } from '../src/platform/ribbon/feature-search.ts'

const button = (id, extra = {}) => ({ id, title: id, type: 'button', commandId: `doc.command.${id}`, ...extra })
const group = (id, items, extra = {}) => ({ id, items, ...extra })
const tab = (id, groups, extra = {}) => ({ id, title: id, groups, ...extra })

test('feature search indexes available command variants once and retains live Ribbon parameters', () => {
  const format = { id: 'font', title: 'Font', type: 'select', commandId: 'doc.command.set-font',
    valueParams: value => ({ value }), options: [{ label: 'Arial', value: 'Arial' }, { label: 'Missing', value: 'Missing', disabled: true }] }
  const input = [
    tab('start', [group('format', [format, button('alternate-font', { commandId: format.commandId, params: { value: 'Arial' } }),
      button('only-second-font', { commandId: format.commandId, params: { value: 'Inter' } }),
      { ...button('disabled-menu', { disabled: true }), type: 'menu', children: [button('disabled-child')] },
      featureSearchRibbonItem('Search features')], { title: 'Text' })]),
    tab('view', [group('hidden-group', [button('hidden-command')], { hidden: true })]),
  ]
  const matches = collectFeatureSearchResults(input, 'Ribbon')
  assert.deepEqual(matches.map(({ title, params }) => [title, params]), [['Arial', { value: 'Arial' }], ['only-second-font', { value: 'Inter' }]])
  assert.equal(matches[0].path, 'Ribbon / start / Text / Font')
  assert.match(matches[0].keywords, /doc\.command\.set-font/)
  format.disabled = true
  assert.equal(collectFeatureSearchResults(input, 'Ribbon')[0].title, 'alternate-font')
})

test('simple measures each placement independently while retaining command and slot IDs', () => {
  const tabs = visibleRibbonTabs(['home', 'insert'].map(id => tab(id, [group('tools', [button('shared', { params: id })])])))
  const groups = simpleRibbonTab(tabs).groups
  const measurements = groups.map(g => ({ id: ribbonPlacementKey(g.id, g.items[0].id), groupId: g.id, width: 40 }))
  const fit = fitRibbonItems(measurements, 120)
  assert.deepEqual(fit.visibleIds, [measurements[0].id])
  assert.deepEqual(fit.overflowIds, [measurements[1].id])
  assert.notEqual(measurements[0].id, measurements[1].id)
  assert.deepEqual(groups.map(g => [g.items[0].id, ribbonInvocation(g.items[0]).params]), [['shared', 'home'], ['shared', 'insert']])
})

test('orders tabs/groups/items by explicit order and preserves equal-order registration order without mutation', () => {
  const input = [tab('insert', [group('second', [button('c')], { order: 2 }), group('first', [button('b', { order: 2 }), button('a', { order: 1 }), button('d', { order: 2 })], { order: 1 })], { order: 2 }), tab('start', [group('x', [button('x')])], { order: 1 })]
  const before = JSON.stringify(input)
  const actual = visibleRibbonTabs(input)
  assert.deepEqual(actual.map(({ id }) => id), ['start', 'insert'])
  assert.deepEqual(actual[1].groups.map(({ id }) => id), ['first', 'second'])
  assert.deepEqual(actual[1].groups[0].items.map(({ id }) => id), ['a', 'b', 'd'])
  assert.equal(JSON.stringify(input), before)
})

test('hidden items and empty menus/groups/tabs are removed while disabled items remain visible', () => {
  const input = [tab('hidden', [group('x', [button('x')])], { hidden: true }), tab('empty', [group('x', [button('x', { hidden: true })])]), tab('start', [group('hidden', [button('x')], { hidden: true }), group('format', [button('bold', { disabled: true, activated: true }), { id: 'empty-menu', type: 'menu', title: 'empty', children: [button('hidden', { hidden: true })] }])])]
  const actual = visibleRibbonTabs(input)
  assert.deepEqual(actual.map(({ id }) => id), ['start'])
  assert.deepEqual(actual[0].groups[0].items.map(({ id }) => id), ['bold'])
  assert.equal(actual[0].groups[0].items[0].disabled, true)
  assert.equal(actual[0].groups[0].items[0].activated, true)
})

test('keyboard navigation wraps, honors RTL and excludes hidden/empty categories', () => {
  const tabs = ['start', 'insert', 'view'].map((id) => tab(id, [group(id, [button(id)])]))
  tabs.splice(1, 0, tab('hidden', [], { hidden: true }))
  assert.equal(nextRibbonTab(tabs, 'start', 'ArrowLeft'), 'view')
  assert.equal(nextRibbonTab(tabs, 'start', 'ArrowRight'), 'insert')
  assert.equal(nextRibbonTab(tabs, 'start', 'ArrowRight', 'rtl'), 'view')
  assert.equal(nextRibbonTab(tabs, 'insert', 'Home'), 'start')
  assert.equal(nextRibbonTab(tabs, 'insert', 'End'), 'view')
  assert.equal(nextRibbonTab(tabs, 'start', 'Enter'), undefined)
  assert.equal(nextRibbonTab([], 'start', 'Home'), undefined)
})

test('ending a contextual category restores the previous valid regular category', () => {
  const start = tab('start', [group('a', [button('a')])])
  const view = tab('view', [group('b', [button('b')])])
  const context = tab('table', [group('c', [button('c')])], { contextual: true })
  assert.equal(resolveRibbonTab([start, view, context], 'table', 'view'), 'table')
  assert.equal(resolveRibbonTab([start, view], 'table', 'view'), 'view')
  assert.equal(resolveRibbonTab([start], 'table', 'view'), 'start')
  assert.equal(resolveRibbonTab([], 'table', 'view'), undefined)
})

test('classic toolbar fits measured items, reserves a real more trigger only when needed and preserves order', () => {
  const items = [{ id: 'undo', groupId: 'history', width: 28 }, { id: 'redo', groupId: 'history', width: 28 }, { id: 'font', groupId: 'font', width: 100 }, { id: 'size', groupId: 'font', width: 60 }]
  assert.equal(CLASSIC_RIBBON.tabHeight, 36)
  assert.equal(CLASSIC_RIBBON.tabButtonHeight, 28)
  assert.equal(CLASSIC_RIBBON.toolbarHeight, 40)
  assert.equal(CLASSIC_RIBBON.buttonHeight, 24)
  assert.equal(CLASSIC_RIBBON.selectorHeight, 28)
  assert.equal(CLASSIC_RIBBON.iconSize, 16)
  assert.equal(CLASSIC_RIBBON.itemGap, 8)
  assert.equal(CLASSIC_RIBBON.groupPadding, 8)
  assert.deepEqual(fitRibbonItems(items, 289), { visibleIds: ['undo', 'redo', 'font', 'size'], overflowIds: [] })
  assert.deepEqual(fitRibbonItems(items, 288), { visibleIds: ['undo', 'redo', 'font'], overflowIds: ['size'] })
  assert.deepEqual(fitRibbonItems(items, 254), { visibleIds: ['undo', 'redo', 'font'], overflowIds: ['size'] })
  assert.deepEqual(fitRibbonItems(items, 253), { visibleIds: ['undo', 'redo'], overflowIds: ['font', 'size'] })
  assert.deepEqual(fitRibbonItems(items, 200), { visibleIds: ['undo', 'redo'], overflowIds: ['font', 'size'] })
  assert.deepEqual(fitRibbonItems(items, 20), { visibleIds: [], overflowIds: ['undo', 'redo', 'font', 'size'] })
  assert.deepEqual(fitRibbonItems([], 0), { visibleIds: [], overflowIds: [] })
})

test('overflow dispatch shares the exact command and params; disabled/hidden/stale options cannot execute', () => {
  const bold = button('bold', { params: { bold: true } })
  assert.deepEqual(ribbonInvocation(bold), { commandId: 'doc.command.bold', params: { bold: true } })
  assert.equal(ribbonInvocation({ ...bold, disabled: true }), undefined)
  assert.equal(ribbonInvocation({ ...bold, hidden: true }), undefined)
  const font = { id: 'font', type: 'select', title: '字体', commandId: 'doc.command.font', value: null, options: [{ label: 'Arial', value: 'Arial' }, { label: '禁止', value: 'bad', disabled: true }, { label: '默认', value: 'default', commandId: 'doc.command.clear-font', params: { clear: true } }], valueParams: (value) => ({ family: value }) }
  assert.deepEqual(ribbonInvocation(font, 'Arial'), { commandId: 'doc.command.font', params: { family: 'Arial' } })
  assert.deepEqual(ribbonInvocation(font, 'default'), { commandId: 'doc.command.clear-font', params: { clear: true } })
  assert.equal(ribbonInvocation(font, 'bad'), undefined)
  assert.equal(ribbonInvocation(font, 'stale'), undefined)
  assert.equal(ribbonInvocation(font), undefined)
  assert.equal(ribbonInvocation({ id: 'menu', type: 'menu', title: 'menu', commandId: 'ignored' }), undefined)
})

test('resizing keeps every command accessible exactly once and never overfills the measured toolbar', () => {
  const items = Array.from({ length: 18 }, (_, index) => ({ id: `tool-${index}`, groupId: `group-${Math.floor(index / 3)}`, width: index % 3 ? 28 : 96 }))
  for (const available of [0, 60, 90, 140, 320, 640, 960, 2000]) {
    const fit = fitRibbonItems(items, available)
    assert.deepEqual([...fit.visibleIds, ...fit.overflowIds], items.map(({ id }) => id))
    assert.equal(new Set([...fit.visibleIds, ...fit.overflowIds]).size, items.length)
    const visible = items.slice(0, fit.visibleIds.length)
    const occupied = visible.reduce((sum, item, index) => sum + item.width + (
      index > 0 && item.groupId === visible[index - 1].groupId
        ? CLASSIC_RIBBON.itemGap
        : 2 * CLASSIC_RIBBON.groupPadding + (index > 0 ? CLASSIC_RIBBON.separatorWidth : 0)
    ), 0)
    if (visible.length) assert.ok(occupied + 2 * CLASSIC_RIBBON.horizontalPadding + (fit.overflowIds.length ? CLASSIC_RIBBON.moreWidth + CLASSIC_RIBBON.separatorWidth : 0) <= available)
  }
})

test('group edges and dividers are counted even for a single control per group', () => {
  const items = [{ id: 'undo', groupId: 'history', width: 24 }, { id: 'bold', groupId: 'font', width: 24 }]
  // 24 outer padding + (8 + 24 + 8) + divider + (8 + 24 + 8).
  assert.deepEqual(fitRibbonItems(items, 105), { visibleIds: ['undo', 'bold'], overflowIds: [] })
  assert.deepEqual(fitRibbonItems(items, 104), { visibleIds: ['undo'], overflowIds: ['bold'] })
  // The final visible group keeps its trailing padding before More's divider.
  assert.deepEqual(fitRibbonItems(items, 96), { visibleIds: [], overflowIds: ['undo', 'bold'] })
  assert.deepEqual(fitRibbonItems(items, 97), { visibleIds: ['undo'], overflowIds: ['bold'] })
})

test('split buttons preserve their main action while a hidden-only submenu is removed', () => {
  const split = { id: 'paste', type: 'split', title: '粘贴', commandId: 'doc.command.paste', params: { mode: 'source' }, children: [button('paste-text', { hidden: true })] }
  const visible = visibleRibbonTabs([tab('start', [group('clipboard', [split])])])[0].groups[0].items[0]
  assert.deepEqual(visible.children, [])
  assert.deepEqual(ribbonInvocation(visible), { commandId: 'doc.command.paste', params: { mode: 'source' } })
  const size = { id: 'size', title: '字号', type: 'select', commandId: 'doc.command.size', options: [{ label: '0', value: 0 }, { label: '自动', value: '' }] }
  assert.deepEqual(ribbonInvocation(size, 0), { commandId: 'doc.command.size', params: 0 })
  assert.deepEqual(ribbonInvocation(size, ''), { commandId: 'doc.command.size', params: '' })
})

test('grid placement honors two-row spans, columns and deterministic fallback for conflicts', () => {
  const items = [
    button('large', { gridLayout: { row: 1, column: 1, rowSpan: 2, width: 84, minWidth: 56, showLabel: true } }),
    button('top', { gridLayout: { row: 1, column: 2 } }),
    button('bottom', { gridLayout: { row: 2, column: 2 } }),
    button('fallback'),
    button('collision', { gridLayout: { row: 1, column: 1, width: 48 } }),
    button('outside', { gridLayout: { row: 3, column: 1, width: 48 } }),
  ]
  const placements = placeRibbonGridItems(items)
  assert.deepEqual(placements.map(({ item, row, column, rowSpan }) => [item.id, row, column, rowSpan]), [
    ['large', 1, 1, 2], ['top', 1, 2, 1], ['fallback', 1, 3, 1], ['outside', 1, 4, 1], ['bottom', 2, 2, 1], ['collision', 2, 3, 1],
  ])
  assert.equal(placements[0].width, 84)
  assert.equal(placements[0].minWidth, 56)
  assert.equal(placements[0].showLabel, true)
})

test('grid overflow groups retain row and icon distinctions', () => {
  const items = [
    button('icon-one', { icon: 'icon', gridLayout: { row: 1, column: 1 } }),
    button('icon-two', { icon: 'icon', gridLayout: { row: 1, column: 2 } }),
    button('text', { gridLayout: { row: 1, column: 3 } }),
    button('next-row', { icon: 'icon', gridLayout: { row: 2, column: 1 } }),
  ]
  assert.deepEqual(groupRibbonGridOverflowItems(items).map((group) => group.map(({ id }) => id)), [
    ['icon-one', 'icon-two'], ['text'], ['next-row'],
  ])
})

test('grid density compresses before complete trailing groups move to overflow', () => {
  assert.equal(GRID_RIBBON.tabHeight, 32)
  assert.equal(GRID_RIBBON.toolbarHeight, 69)
  assert.equal(GRID_RIBBON.rowHeight, 24)
  assert.equal(GRID_RIBBON.rowGap, 4)
  assert.equal(GRID_RIBBON.verticalPadding, 8)
  assert.equal(GRID_RIBBON.itemGap, 4)
  assert.equal(GRID_RIBBON.buttonSize, 28)
  assert.equal(GRID_RIBBON.moreWidth, 72)
  const groups = [{ regular: 90, compact: 75 }, { regular: 85, compact: 70 }, { regular: 40, compact: 32 }]
  assert.deepEqual(fitRibbonGridGroups(groups, 180), { visibleCount: 3, compact: true })
  assert.deepEqual(fitRibbonGridGroups(groups, 149), { visibleCount: 1, compact: true })
  assert.deepEqual(fitRibbonGridGroups(groups, 138), { visibleCount: 0, compact: true })
  assert.deepEqual(fitRibbonGridGroups(groups, 260), { visibleCount: 3, compact: false })
  assert.deepEqual(fitRibbonGridGroups([], 0), { visibleCount: 0, compact: false })
})
