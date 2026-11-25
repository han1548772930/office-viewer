import assert from 'node:assert/strict'
import { test } from 'vitest'
import { chartContentTooltipText, chartSunburstContentTooltip, chartTreemapContentTooltip } from '../src/features/charts/chart-content-tooltip.ts'
import { ChartNativeHierarchyTooltipBinding } from '../src/features/charts/echarts-hierarchy-tooltip.ts'
import { createChartTreemapPointerRuntime, chartTreemapEChartsData, chartTreemapTooltip } from '../src/features/charts/echarts-treemap-option.ts'
import { createChartSunburstPointerRuntime, chartSunburstEChartsData, chartSunburstTooltip } from '../src/features/charts/echarts-sunburst-option.ts'
import { treemapFixture } from './fixtures/chart-treemap-fixtures.mjs'
import { sunburstFixture } from './fixtures/chart-sunburst-fixtures.mjs'

// Written only. These semantic/lifecycle contracts have not run.
class MouseEvents {
  listeners = new Map()
  removed = []
  on(name, handler) { const group = this.listeners.get(name) ?? new Set(); group.add(handler); this.listeners.set(name, group) }
  off(name, handler) { this.removed.push([name, handler]); this.listeners.get(name)?.delete(handler) }
  emit(name, params) { for (const handler of this.listeners.get(name) ?? []) handler(params) }
  count() { return [...this.listeners.values()].reduce((sum, group) => sum + group.size, 0) }
}
const context = { fieldName: 'Weight', total: 22 }
const semanticNode = { name: 'South', path: ['South'], value: 9, selfValue: 3, formattedSelfValue: '3 个', children: [{ name: 'Leaf', value: 6 }] }
const treemapItem = { id: '__univer_treemap_residual__:South', value: 3, semanticNode }
const sunburstItem = { value: 3, semanticPath: ['South'], semanticSelfValue: 3, semanticChildValue: 6 }
const rawEvent = (mode, data, index = 4) => ({ componentType: 'series', seriesType: mode, seriesIndex: 0, dataIndex: index, data })

test('Treemap residuals preserve owning paths, signed semantic totals and source formatting without rewriting facts', () => {
  const before = structuredClone(treemapItem)
  assert.deepEqual(chartTreemapContentTooltip(treemapItem, context), { title: 'South', lines: ['Weight: 9', '自身数值: 3 个', '子节点合计: 6'] })
  const signed = { semanticNode: { name: '<North>', path: ['<North>', 'Leaf\nA'], value: 13, selfValue: -5, children: [{ value: 18 }] } }
  assert.equal(chartContentTooltipText(chartTreemapContentTooltip(signed, context)), '<North> / Leaf\nA\nWeight: 13\n自身数值: -5\n子节点合计: 18')
  assert.equal(chartTreemapContentTooltip({ semanticNode: { name: 'Leaf', value: 0 } }, context).lines.length, 1)
  assert.deepEqual(treemapItem, before)
})

test('Sunburst residuals use their positive own value and whole-chart weight for percentages, keeping signed parent facts', () => {
  assert.deepEqual(chartSunburstContentTooltip(sunburstItem, context), { title: 'South', lines: ['Weight: 3 (13.6%)', '自身数值: 3', '子节点合计: 6'] })
  const branch = { ...sunburstItem, value: 13, semanticPath: ['North'], semanticSelfValue: -5, semanticChildValue: 18, formattedValue: '13 个' }
  assert.equal(chartSunburstContentTooltip(branch, context).lines[0], 'Weight: 13 个 (59.1%)')
  assert.equal(chartSunburstContentTooltip(branch, { ...context, format: '0.00' }).lines[0], 'Weight: 13.00 (59.1%)')
  assert.equal(chartSunburstContentTooltip({ ...sunburstItem, value: 0, semanticChildValue: 0 }, { ...context, total: 0 }).lines[0], 'Weight: 0 (0.0%)')
})

test('malformed or nonfinite raw native facts produce no content, retaining plain-text labels for valid data', () => {
  for (const bad of [null, [], {}, { semanticNode: { name: 'Bad', value: Infinity } }, { semanticNode: { name: 'Bad', value: 1, path: [1] } },
    { semanticNode: { name: 'Bad', value: 1, selfValue: NaN } }, { semanticNode: { name: 'Bad', value: 1, formattedValue: {} } },
    { semanticNode: { name: 'Bad', value: 1e308, selfValue: -1e308, children: [{}] } }]) {
    assert.equal(chartTreemapContentTooltip(bad, context), undefined)
  }
  for (const bad of [null, [], {}, { ...sunburstItem, value: NaN }, { ...sunburstItem, semanticPath: [null] }, { ...sunburstItem, semanticChildValue: Infinity }]) {
    assert.equal(chartSunburstContentTooltip(bad, context), undefined)
  }
  assert.equal(chartContentTooltipText(), '')
})

test('native item events supply exact indices and mode-isolated facts; blank, foreign and leave events cannot reuse a node', () => {
  const events = new MouseEvents(), binding = new ChartNativeHierarchyTooltipBinding(events)
  const treemap = { mode: 'treemap', contentTooltip: item => chartTreemapContentTooltip(item, context) }
  const sunburst = { mode: 'sunburst', contentTooltip: item => chartSunburstContentTooltip(item, context) }
  events.emit('mouseover', rawEvent('treemap', treemapItem))
  assert.deepEqual(binding.resolve(treemap), { title: 'South', lines: ['Weight: 9', '自身数值: 3 个', '子节点合计: 6'], id: 'treemap:0:4', kind: 'hierarchy-node', index: 4 })
  assert.equal(binding.resolve(sunburst), undefined)
  binding.clear(); assert.equal(binding.resolve(treemap), undefined)
  events.emit('mousemove', rawEvent('sunburst', sunburstItem, 7)); assert.equal(binding.resolve(sunburst).index, 7)
  events.emit('mousemove', { ...rawEvent('sunburst', sunburstItem), componentType: 'graphic' }); assert.equal(binding.resolve(sunburst), undefined)
  events.emit('mousemove', { ...rawEvent('sunburst', sunburstItem), dataIndex: -1 }); assert.equal(binding.resolve(sunburst), undefined)
  for (const name of ['mouseout', 'globalout']) {
    events.emit('mousemove', rawEvent('treemap', treemapItem)); events.emit(name, {}); assert.equal(binding.resolve(treemap), undefined)
  }
  binding.dispose()
})

test('bindings are instance-owned, detach only their handlers and do not resurrect captured facts after disposal', () => {
  const first = new MouseEvents(), second = new MouseEvents(), other = () => {}
  first.on('mousemove', other)
  const a = new ChartNativeHierarchyTooltipBinding(first), b = new ChartNativeHierarchyTooltipBinding(second)
  const runtime = { mode: 'treemap', contentTooltip: item => chartTreemapContentTooltip(item, context) }
  first.emit('mousemove', rawEvent('treemap', treemapItem)); assert.equal(a.resolve(runtime).title, 'South'); assert.equal(b.resolve(runtime), undefined)
  a.dispose(); a.dispose(); assert.equal(first.count(), 1); assert.equal(first.listeners.get('mousemove').has(other), true)
  assert.equal(first.removed.length, 4); assert.equal(a.resolve(runtime), undefined)
  second.emit('mousemove', rawEvent('treemap', treemapItem)); assert.equal(b.resolve(runtime).title, 'South')
  b.dispose(); assert.equal(second.count(), 0)
})

test('native adapters hide their tooltip widget and share semantic presentation with the Ant content runtime', () => {
  const theme = { palette: ['#112233', '#ddeeff'], color: '#334155', axisColor: '#64748b', grid: { left: 12, top: 12, right: 12, bottom: 12 } }
  for (const mode of ['treemap', 'sunburst']) {
    const sample = mode === 'treemap' ? treemapFixture() : sunburstFixture(), { data, chart } = sample, before = JSON.stringify(chart)
    const roots = mode === 'treemap' ? chartTreemapEChartsData(data, chart, theme) : chartSunburstEChartsData(data, chart, theme)
    const pointer = mode === 'treemap' ? createChartTreemapPointerRuntime(data) : createChartSunburstPointerRuntime(data, chart)
    const tooltip = mode === 'treemap' ? chartTreemapTooltip(data, theme) : chartSunburstTooltip(data, chart, theme)
    assert.equal(tooltip.showContent, false); assert.equal(tooltip.triggerOn, 'none')
    const pending = [...roots]
    while (pending.length) {
      const item = pending.pop(); assert.equal(tooltip.formatter({ data: item }), chartContentTooltipText(pointer.contentTooltip(item)))
      pending.push(...(item.children ?? []))
    }
    assert.equal(JSON.stringify(chart), before)
  }
})

test('subscription rollback and disposal attempt every owned off even after failure and report the original or first error', () => {
  const subscribeError = new Error('subscribe failed'), firstOffError = new Error('first off failed'), secondOffError = new Error('second off failed')
  class FailingEvents extends MouseEvents {
    constructor(failSubscribe = false) { super(); this.failSubscribe = failSubscribe; this.attempts = [] }
    on(name, handler) { super.on(name, handler); if (this.failSubscribe && name === 'mousemove') throw subscribeError }
    off(name, handler) { this.attempts.push(name); super.off(name, handler); if (name === 'mouseover') throw firstOffError; if (name === 'mouseout') throw secondOffError }
  }
  const rollback = new FailingEvents(true)
  assert.throws(() => new ChartNativeHierarchyTooltipBinding(rollback), error => error === subscribeError)
  assert.deepEqual(rollback.attempts, ['mouseover', 'mousemove']); assert.equal(rollback.count(), 0)
  const events = new FailingEvents(), binding = new ChartNativeHierarchyTooltipBinding(events)
  const runtime = { mode: 'treemap', contentTooltip: item => chartTreemapContentTooltip(item, context) }
  events.emit('mousemove', rawEvent('treemap', treemapItem)); assert.equal(binding.resolve(runtime).title, 'South')
  assert.throws(() => binding.dispose(), error => error === firstOffError)
  assert.deepEqual(events.attempts, ['mouseover', 'mousemove', 'mouseout', 'globalout']); assert.equal(events.count(), 0)
  assert.equal(binding.resolve(runtime), undefined); binding.dispose(); assert.equal(events.attempts.length, 4)
})
