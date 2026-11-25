import assert from 'node:assert/strict'
import { test } from 'vitest'
import { treemapFixture } from './fixtures/chart-treemap-fixtures.mjs'
import { chartTreemapBackgroundPaint, chartTreemapHeaderHeight, chartTreemapLabelBackground, chartTreemapLabelFits, chartTreemapLabelText,
  chartTreemapRenderNodes, chartTreemapShade, chartTreemapTextBlock, chartTreemapTextColor, resolveChartTreemapLabel } from '../src/features/charts/chart-treemap.ts'
import { chartTreemapLayout } from '../src/features/charts/chart-treemap-layout.ts'
import { chartTreemapBackgroundMode, chartTreemapLeafControls, chartTreemapParentControls, clearChartTreemapSettings,
  setChartTreemapBackground, setChartTreemapBackgroundMode } from '../src/features/charts/chart-treemap-settings-controls.ts'
import { chartTreemapHierarchyFields, chartTreemapLabelValues } from '../src/features/charts/chart-treemap-api.ts'
import { chartTreemapEChartsData, createChartTreemapEChartsSeries } from '../src/features/charts/echarts-treemap-option.ts'
import { CHART_TREEMAP_BACKGROUND_KEY, CHART_TREEMAP_FIT_KEY, CHART_TREEMAP_OVERLAP_KEY } from '../src/features/charts/echarts-treemap-layout.ts'
import { chartCommonControls, chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { mergeChartOptions, validateChartOptions } from '../src/features/charts/chart-options.ts'
import { createChartSettingsDraft, chartSettingsPatch } from '../src/features/charts/chart-settings-draft.ts'
import { updateChart } from '../src/features/charts/chart-engine.ts'
import { estimateChartFinalLayoutMemory } from '../src/features/charts/chart-compute-memory.ts'

// P26 implementation drafts. No test, type check, Canvas or benchmark ran.
const json = value => JSON.parse(JSON.stringify(value))
const plot = { x: 0, y: 0, width: 600, height: 400 }
const context = { palette: ['#112233', '#ddeeff'], color: '#64748b', axisColor: '#64748b', splitLineColor: '#eeeeee',
  backgroundColor: '#ffffff', grid: { left: 12, right: 12, top: 12, bottom: 12, containLabel: false } }
const control = (controls, key) => { const value = controls.find(item => item.key === key); assert.ok(value, key); return value }
const gradient = () => ({ type: 'linear-gradient', start: { x: 0, y: 0 }, end: { x: 1, y: 1 }, stops: [{ offset: 0, color: 'red', opacity: 0.4 }, { offset: 1, color: 'blue' }] })

test('positive parent-owned facts become residuals, negative parent facts stay semantic and actual child sums drive areas', () => {
  const { data } = treemapFixture(), roots = chartTreemapRenderNodes(data.hierarchy), before = json(data.hierarchy)
  assert.deepEqual(roots.map(node => [node.name, node.value, node.node.value, node.node.selfValue]), [['North', 18, 13, -5], ['South', 9, 9, 3]])
  assert.equal(roots[0].children.some(node => node.isTreemapResidual), false)
  const residual = roots[1].children.find(node => node.isTreemapResidual)
  assert.ok(residual); assert.equal(residual.value, 3); assert.equal(residual.name, ''); assert.equal(residual.semanticParentId, roots[1].id)
  assert.equal(residual.node, roots[1].node)
  const rectangles = chartTreemapLayout(data.hierarchy, plot, { parentLabelLayout: 'none' }), parents = rectangles.filter(rect => rect.depth === 0)
  assert.ok(Math.abs(parents[0].width * parents[0].height / (plot.width * plot.height) - 2 / 3) < 1e-10)
  assert.ok(Math.abs(parents[1].width * parents[1].height / (plot.width * plot.height) - 1 / 3) < 1e-10)
  assert.deepEqual(json(data.hierarchy), before)
  assert.equal(rectangles.filter(rect => rect.renderNode.isTreemapResidual).length, 1)
})

test('banner reserves a font-based strip, overlap fills it and labels keep source semantic values', () => {
  const { data, chart } = treemapFixture({ treemap: { label: { visible: false, contentType: 10 } } })
  const labels = resolveChartTreemapLabel(chart), roots = chartTreemapRenderNodes(data.hierarchy)
  assert.equal(labels.visible, false); assert.equal(chartTreemapHeaderHeight('banner', 12), 20); assert.equal(chartTreemapHeaderHeight('overlapping', 12), 12)
  assert.equal(chartTreemapLabelText(roots[0], labels), 'North\n13')
  const residual = roots[1].children.find(node => node.isTreemapResidual)
  assert.equal(chartTreemapLabelText(residual, labels), '3')
  const banner = chartTreemapLayout(data.hierarchy, plot, { parentLabelLayout: 'banner', fontSize: 12 })
  const overlap = chartTreemapLayout(data.hierarchy, plot, { parentLabelLayout: 'overlapping', fontSize: 12 })
  const parent = banner.find(rect => rect.node.name === 'North' && rect.depth === 0)
  assert.ok(banner.filter(rect => rect.depth === 1 && rect.rootIndex === parent.rootIndex).every(rect => rect.y >= parent.y + 20))
  const overlappingParent = overlap.find(rect => rect.node.name === 'North' && rect.depth === 0)
  assert.ok(Math.min(...overlap.filter(rect => rect.depth === 1 && rect.rootIndex === overlappingParent.rootIndex).map(rect => rect.y)) < overlappingParent.y + 12)
  const native = createChartTreemapEChartsSeries(data, chart, context)
  assert.equal(native.label.show, false); assert.equal(native.upperLabel.show, true)
  assert.equal(native.data[0].label.show, false); assert.equal(native.data[0].upperLabel.show, true)
  assert.equal(native.animation, false); assert.equal(native.sort, false); assert.equal(native.visibleMin, 0)
  assert.equal(native.nodeClick, false); assert.equal(native.roam, false); assert.equal(native.breadcrumb.show, false); assert.equal(native[CHART_TREEMAP_FIT_KEY], true)
})

test('complete text wraps only when every grapheme and line fits and never truncates or compresses', () => {
  const measure = text => Array.from(text).length * 10
  assert.deepEqual(chartTreemapTextBlock('ABCD\nEF', { width: 20, height: 30 }, '12px sans-serif', 10, measure),
    { lines: ['AB', 'CD', 'EF'], width: 20, height: 30, lineHeight: 10 })
  assert.equal(chartTreemapTextBlock('ABCD\nEF', { width: 20, height: 29 }, '12px sans-serif', 10, measure), undefined)
  assert.equal(chartTreemapTextBlock('A', { width: 9, height: 10 }, '12px sans-serif', 10, measure), undefined)
  assert.equal(chartTreemapLabelFits({ width: 20, height: 10 }, { width: 21, height: 10 }), false)
  assert.equal(chartTreemapLabelFits({ width: 20, height: 10 }, { width: 20, height: 10 }), true)
})

test('fixed depth tint, contrast and pattern mean retain original authoring and native gradients use relative text bounds', () => {
  assert.equal(chartTreemapShade('#000000', 1), '#1f1f1f'); assert.equal(chartTreemapShade('#000000', 20), '#7a7a7a')
  assert.equal(chartTreemapShade('navy', 3), 'navy'); assert.equal(chartTreemapTextColor('#ffffff'), '#000000'); assert.equal(chartTreemapTextColor('#000000'), '#ffffff')
  const pattern = { type: 'pattern', preset: 'pct25', foregroundColor: '#000000', backgroundColor: '#ffffff' }, before = json(pattern)
  assert.equal(chartTreemapLabelBackground(pattern), 'rgba(191, 191, 191, 1)'); assert.deepEqual(pattern, before)
  assert.equal(chartTreemapLabelBackground({ ...pattern, preset: 'cross' }), 'rgba(128, 128, 128, 1)')
  const color = gradient(), { data, chart } = treemapFixture({ treemap: { parentLabelLayout: 'overlapping', label: { backgroundColor: color } } })
  const native = createChartTreemapEChartsSeries(data, chart, context), paint = native.data[0][CHART_TREEMAP_BACKGROUND_KEY]
  assert.equal(native[CHART_TREEMAP_OVERLAP_KEY], true); assert.equal(paint.global, false); assert.deepEqual([paint.x, paint.y, paint.x2, paint.y2], [0, 0, 1, 1])
  assert.deepEqual(chartTreemapBackgroundPaint(color, plot).bounds, plot)
  assert.equal(chartTreemapEChartsData(data, chart, context)[1].children.find(node => node.isTreemapResidual).semanticNode.selfValue, 3)
})

test('native fill opacity is encoded in color while the separate outline keeps its own alpha', () => {
  const { data, chart } = treemapFixture({ allSeriesStyle: { fillOpacity: 0.4, border: { visible: true, color: 'red', opacity: 0.75, width: 3 } } })
  const roots = chartTreemapEChartsData(data, chart, context), north = roots[0], leaf = roots[1].children.find(node => node.name === 'Leaf A')
  assert.equal(north.itemStyle.color, 'rgba(221, 238, 255, 0.4)'); assert.equal(leaf.itemStyle.color, 'rgba(46, 61, 75, 0.4)')
  for (const node of [north, leaf]) {
    assert.equal(node.itemStyle.opacity, 1); assert.equal(node.itemStyle.borderColor, 'rgba(255, 0, 0, 0.75)'); assert.equal(node.itemStyle.borderWidth, 3)
  }
})

test('Treemap exposes only active category/value label choices, conditional number formats and independent default resets', () => {
  const { workbook, chart } = treemapFixture({ treemap: { parentLabelLayout: 'overlapping', label: { visible: true, contentType: 27, format: '0.00', fontSize: 18 } } })
  const draft = createChartSettingsDraft(chart), leaf = chartTreemapLeafControls()
  assert.deepEqual(control(leaf, 'treemap.label.contentType').choices.map(choice => choice.value), ['category', 'value'])
  assert.equal(leaf.some(item => item.key.endsWith('.position')), false)
  assert.deepEqual(control(leaf, 'treemap.label.contentType').read(draft.config), ['category', 'value'])
  control(leaf, 'treemap.label.contentType').write(draft.config, ['value'])
  assert.equal(draft.config.options.treemap.label.contentType, 25)
  assert.equal(control(leaf, 'treemap.label.format').visible(draft.config), true)
  control(leaf, 'treemap.label.contentType').write(draft.config, ['category'])
  assert.equal(control(leaf, 'treemap.label.format').visible(draft.config), false)
  control(leaf, 'treemap.label.visible').write(draft.config, false)
  assert.equal(control(leaf, 'treemap.label.contentType').visible(draft.config), false)
  control(chartTreemapParentControls(), 'treemap.parentLabelLayout').reset(draft.config)
  assert.equal(draft.config.options.treemap.parentLabelLayout, undefined); assert.equal(draft.config.options.treemap.label.visible, false)
  clearChartTreemapSettings(draft.config, 'label'); updateChart(workbook, chart, chartSettingsPatch(draft))
  assert.deepEqual(json(chart.options.treemap), {}); assert.equal(resolveChartTreemapLabel(chart).visible, true)
  assert.equal(resolveChartTreemapLabel(chart).fontSize, 12)
  assert.equal(chartCommonControls(chart).some(item => item.key.startsWith('legend.')), false)
  for (const item of chartSeriesControls().filter(item => item.key.includes('.label.') || item.key.includes('.point.'))) assert.equal(item.visible(chart), false)
})

test('background replacements are atomic, patterns detach from callers and malformed background/API inputs cannot commit', () => {
  const bare = { type: 'treemap' }, emptyBefore = json(bare)
  assert.throws(() => setChartTreemapBackground(bare, { type: 'unknown' })); assert.deepEqual(bare, emptyBefore)
  const { chart } = treemapFixture(), color = gradient(), pattern = { type: 'pattern', preset: 'pct50', foregroundColor: 'black', backgroundColor: 'white' }
  setChartTreemapBackground(chart, color); assert.equal(chartTreemapBackgroundMode(chart), 'color')
  setChartTreemapBackgroundMode(chart, 'pattern'); assert.equal(chartTreemapBackgroundMode(chart), 'pattern')
  const merged = mergeChartOptions({ treemap: { label: { backgroundColor: color, fontSize: 20 } } }, { treemap: { label: { backgroundColor: pattern } } })
  assert.deepEqual(merged.treemap.label.backgroundColor, pattern); assert.equal(merged.treemap.label.fontSize, 20)
  assert.equal(Object.hasOwn(merged.treemap.label.backgroundColor, 'stops'), false)
  const detached = chartTreemapLabelValues({ backgroundColor: pattern, color: '#000000' }); pattern.foregroundColor = 'red'
  assert.equal(detached.backgroundColor.foregroundColor, 'black')
  assert.doesNotThrow(() => validateChartOptions({ treemap: { label: { backgroundColor: { ...pattern, preset: 'importedPattern' } } } }))
  for (const value of [{ type: 'pattern', preset: '', foregroundColor: 'red', backgroundColor: 'blue' },
    { type: 'pattern', preset: 'cross', foregroundColor: 'invalid', backgroundColor: 'blue' }, { type: 'unknown' }]) assert.throws(() => chartTreemapLabelValues({ backgroundColor: value }))
  assert.throws(() => chartTreemapLabelValues({ position: 'left' })); assert.throws(() => chartTreemapHierarchyFields([])); assert.throws(() => chartTreemapHierarchyFields([0, 0])); assert.throws(() => chartTreemapHierarchyFields(Array(2)))
})

test('final-scene admission includes descendant rectangles and parent residuals even with one category at the root', () => {
  const { data, chart } = treemapFixture(), shallow = { ...data, hierarchy: undefined }
  const extra = estimateChartFinalLayoutMemory(data, chart, {}) - estimateChartFinalLayoutMemory(shallow, chart, {})
  assert.ok(extra >= 5 * 2048 + 1024)
})
