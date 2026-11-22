import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createRibbonMenuModel, ribbonMenuInvocation } from '../src/platform/ribbon/ribbon-command-menu.ts'

const action = (id, extra = {}) => ({ id, type: 'button', title: id, commandId: `doc.command.${id}`, ...extra })

test('a disabled submenu blocks its descendants and split main action at dispatch time', () => {
  const source = [{ id: 'nested', type: 'menu', title: 'nested', disabled: true, children: [
    action('child'),
    { ...action('split'), type: 'split', children: [action('split-child')] },
  ] }]
  const blocked = createRibbonMenuModel(source)
  assert.equal(ribbonMenuInvocation(blocked, 'child'), undefined)
  assert.equal(ribbonMenuInvocation(blocked, 'split::main'), undefined)
  assert.equal(ribbonMenuInvocation(blocked, 'split-child'), undefined)
  assert.equal(blocked.items[0].children[0].disabled, true)
  source[0].disabled = false
  const enabled = createRibbonMenuModel(source)
  assert.deepEqual(ribbonMenuInvocation(enabled, 'child'), { commandId: 'doc.command.child', params: undefined })
  assert.deepEqual(ribbonMenuInvocation(enabled, 'split::main'), { commandId: 'doc.command.split', params: undefined })
})

test('hidden menu paths and stale options cannot dispatch while command-specific option parameters are retained', () => {
  const selector = { id: 'mode', type: 'select', title: 'mode', commandId: 'doc.command.mode', options: [
    { label: 'blocked', value: 'blocked', disabled: true },
    { label: 'text', value: 'text', commandId: 'doc.command.paste-special', params: { value: 'text' } },
  ] }
  const model = createRibbonMenuModel([
    { id: 'hidden-parent', type: 'menu', title: 'hidden', hidden: true, children: [action('hidden-child')] },
    { id: 'locked-parent', type: 'menu', title: 'locked', disabled: true, children: [{ ...selector, id: 'locked-mode' }] },
    selector,
  ])
  assert.equal(ribbonMenuInvocation(model, 'hidden-child'), undefined)
  assert.equal(ribbonMenuInvocation(model, 'locked-mode::option:1'), undefined)
  assert.equal(ribbonMenuInvocation(model, 'mode::option:0'), undefined)
  assert.equal(ribbonMenuInvocation(model, 'mode::option:99'), undefined)
  assert.deepEqual(ribbonMenuInvocation(model, 'mode::option:1'), { commandId: 'doc.command.paste-special', params: { value: 'text' } })
  selector.options.pop()
  assert.equal(ribbonMenuInvocation(createRibbonMenuModel([selector]), 'mode::option:1'), undefined)
})

test('a previously rendered action rechecks the live ancestor path, selector options and tree membership', () => {
  const option = { label: 'one', value: 'one' }
  const selector = { ...action('selector'), type: 'select', options: [option] }
  const child = action('leaf')
  const parent = { id: 'parent', type: 'menu', title: 'parent', children: [selector, child] }
  const source = [parent]
  const menu = createRibbonMenuModel(source)
  assert.equal(ribbonMenuInvocation(menu, 'leaf').commandId, 'doc.command.leaf')
  parent.disabled = true
  assert.equal(ribbonMenuInvocation(menu, 'leaf'), undefined)
  assert.equal(ribbonMenuInvocation(menu, 'selector::option:0'), undefined)
  parent.disabled = false
  parent.hidden = true
  assert.equal(ribbonMenuInvocation(menu, 'leaf'), undefined)
  parent.hidden = false
  option.disabled = true
  assert.equal(ribbonMenuInvocation(menu, 'selector::option:0'), undefined)
  option.disabled = false
  selector.options = []
  assert.equal(ribbonMenuInvocation(menu, 'selector::option:0'), undefined)
  parent.children = [selector]
  assert.equal(ribbonMenuInvocation(menu, 'leaf'), undefined)
  source.splice(0)
  assert.equal(ribbonMenuInvocation(menu, 'selector::option:0'), undefined)
})
