import assert from 'node:assert/strict'
import { test } from 'node:test'
import { CLASSIC_RIBBON, fitRibbonItems, nextRibbonTab, resolveRibbonTab, ribbonInvocation, visibleRibbonTabs } from '../src/platform/ribbon/ribbon-schema.ts'

const button = (id, extra = {}) => ({ id, title: id, type: 'button', commandId: `doc.command.${id}`, ...extra })
const group = (id, items, extra = {}) => ({ id, items, ...extra })
const tab = (id, groups, extra = {}) => ({ id, title: id, groups, ...extra })

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
  assert.equal(CLASSIC_RIBBON.toolbarHeight, 40)
  assert.deepEqual(fitRibbonItems(items, 260), { visibleIds: ['undo', 'redo', 'font', 'size'], overflowIds: [] })
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
    const occupied = visible.reduce((sum, item, index) => sum + item.width + (index === 0 ? 0 : item.groupId === visible[index - 1].groupId ? CLASSIC_RIBBON.itemGap : CLASSIC_RIBBON.groupGap), 0)
    if (visible.length) assert.ok(occupied + 2 * CLASSIC_RIBBON.horizontalPadding + (fit.overflowIds.length ? CLASSIC_RIBBON.moreWidth + CLASSIC_RIBBON.itemGap : 0) <= available)
  }
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
