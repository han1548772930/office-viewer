import assert from 'node:assert/strict'
import { test } from 'vitest'
import * as echarts from 'echarts'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart } from '../src/features/charts/chart-engine.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { ChartNativeElementIndex } from '../src/features/charts/echarts-element-interaction.ts'
import { chartCanvasElementKey, chartCanvasSeriesBindings } from '../src/features/charts/chart-element-interaction.ts'

// P19 deferred geometry drafts. SVG SSR supplies the real ECharts model/view
// and ZRender containment, transforms and clip paths. It does not accept the
// Canvas renderer, DOM input, fonts or browser gestures; those have separate drafts.
const size = { width: 600, height: 420 }
function authored(type = 'line', rows = [['Category', 'Original'], ['A', 3], ['B', 8], ['C', 5]], config = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  sheet.cells = Object.fromEntries(rows.flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, { type, position: { x: 0, y: 0 }, size,
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: rows.length - 1, endCol: rows[0].length - 1 } },
    categoryField: 0, valueFields: [1], legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear', title: 'Categories', line: { visible: true }, tick: { visible: true } },
    yAxis: { scaleType: 'linear', title: 'Values', line: { visible: true }, tick: { visible: true }, gridLine: { visible: true } },
    ...config }, 'native-elements')
  return chartEChartsOption(workbook, chart)
}
function render(input) {
  const { officeChartElements: runtime = { series: [], axes: [] }, officeChartKey, officeChartPointer, officeChartLegend, officeChartPlotArea, ...option } = input
  const instance = echarts.init(null, null, { renderer: 'svg', ssr: true, ...size })
  instance.setOption(option, { notMerge: true, lazyUpdate: false })
  const index = () => { instance.renderToSVGString(); return new ChartNativeElementIndex(instance, option, runtime, size.width, size.height) }
  return { instance, option, runtime, index, dispose: () => instance.dispose() }
}
function measuredHit(index, element, seriesOnly = false) {
  const target = index.target(element); assert.ok(target, `Missing painted ${chartCanvasElementKey(element)}`)
  const { bounds } = target
  const matches = point => { const hit = index.hit(point, seriesOnly); return hit?.key === target.key ? { hit, point } : undefined }
  const center = matches({ x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 }); if (center) return center
  for (let y = bounds.y; y <= bounds.y + bounds.height; y += 0.5) for (let x = bounds.x; x <= bounds.x + bounds.width; x += 0.5) {
    const found = matches({ x, y }); if (found) return found
  }
  assert.fail(`The measured target is shielded or fully clipped: ${target.key}`)
}
const binding = () => chartCanvasSeriesBindings({ series: [{ field: 1, name: 'Renamed display', sourceName: 'Original', values: [5, null, 8] }] })[0]

test('actual Cartesian axis displayables and authored axis titles supply measured native hits, including secondary axes', () => {
  const value = render(authored('combination', [['Category', 'Original', 'Comparison'], ['A', 3, 7], ['B', 8, 2], ['C', 5, 9]], {
    valueFields: [1, 2], options: { series: { 2: { type: 'line', rightAxis: true } }, rightYAxis: { scaleType: 'linear', title: 'Secondary', line: { visible: true } } },
  }))
  try {
    const index = value.index()
    for (const axisId of ['xAxis', 'yAxis', 'rightYAxis']) {
      measuredHit(index, { kind: 'axisTitle', axisId }); measuredHit(index, { kind: 'axis', axisId, part: 'label' })
    }
    measuredHit(index, { kind: 'gridLine', axisId: 'yAxis' }); measuredHit(index, { kind: 'tick', axisId: 'xAxis' })
    for (const { bounds } of index.list()) {
      assert.ok(Object.values(bounds).every(Number.isFinite)); assert.ok(bounds.x >= 0 && bounds.y >= 0)
      assert.ok(bounds.x + bounds.width <= size.width && bounds.y + bounds.height <= size.height)
    }
    assert.equal(index.hit({ x: -1, y: -1 }), undefined)
  } finally { value.dispose() }
})

test('native filtered points and attached labels map raw indices to semantic points and source selectors after a display rename', () => {
  const value = render(authored('scatter', [['X', 'Original'], [1, 5], [2, null], [4, 8]], {
    options: { series: { 1: { name: 'Renamed display', point: { shape: 'circle', size: 16 }, label: { visible: true, position: 'top' } } } },
  }))
  try {
    assert.deepEqual(value.runtime.series[0].dataIndices, [0, 2]); assert.deepEqual(value.runtime.series[0].binding, binding())
    const index = value.index(), point = { kind: 'dataPoint', binding: binding(), dataIndex: 2 }
    const measured = measuredHit(index, point), promoted = index.hit(measured.point, true)
    assert.equal(measured.hit.element.dataIndex, 2); assert.equal(promoted.element.kind, 'series')
    assert.equal(promoted.element.binding.selector, 1); assert.equal(promoted.element.binding.sourceName, 'Original')
    measuredHit(index, { ...point, kind: 'dataLabel' })
    assert.equal(index.target({ ...point, dataIndex: 1 }), undefined)
  } finally { value.dispose() }
})

test('real graphic group clipping excludes invisible targets and clips transformed selection bounds without widening native hits', () => {
  const value = render({ animation: false, graphic: { elements: [
    { type: 'group', id: 'clipped', x: 10, y: 15, clipPath: { type: 'rect', shape: { x: 30, y: 20, width: 20, height: 20 } }, children: [
      { type: 'rect', id: 'visible', shape: { x: 0, y: 0, width: 100, height: 100 }, style: { fill: '#123456' },
        info: { officeChartElement: { kind: 'gridLine', axisId: 'xAxis' } } },
      { type: 'rect', id: 'outside', shape: { x: 80, y: 80, width: 10, height: 10 }, style: { fill: '#abcdef' },
        info: { officeChartElement: { kind: 'tick', axisId: 'xAxis' } } },
    ] },
    { type: 'rect', id: 'hidden', invisible: true, shape: { x: 150, y: 20, width: 20, height: 20 }, style: { fill: '#000000' },
      info: { officeChartElement: { kind: 'axisTitle', axisId: 'xAxis' } } },
  ] } })
  try {
    const index = value.index(), target = index.target({ kind: 'gridLine', axisId: 'xAxis' })
    assert.deepEqual(target.bounds, { x: 40, y: 35, width: 20, height: 20 })
    assert.equal(index.hit({ x: 50, y: 45 }).key, target.key); assert.equal(index.hit({ x: 20, y: 25 }), undefined)
    assert.equal(index.target({ kind: 'tick', axisId: 'xAxis' }), undefined); assert.equal(index.target({ kind: 'axisTitle', axisId: 'xAxis' }), undefined)
  } finally { value.dispose() }
})

test('native scrolling legend pages rebuild their measured selection target and keep page controls outside author hits', () => {
  const names = Array.from({ length: 18 }, (_value, index) => `Long series ${index + 1}`)
  const value = render({ animation: false, legend: { type: 'scroll', data: names, left: 80, top: 10, width: 200, selectedMode: true },
    grid: { left: 60, right: 40, top: 80, bottom: 40 }, xAxis: { type: 'category', data: ['A', 'B'] }, yAxis: { type: 'value' },
    series: names.map((name, index) => ({ type: 'line', name, data: [index + 1, index + 2] })),
  })
  try {
    const before = value.index(), first = measuredHit(before, { kind: 'legend' })
    const initial = first.hit.element.name
    value.instance.dispatchAction({ type: 'legendScroll', scrollDataIndex: 14 })
    const after = value.index(), last = measuredHit(after, { kind: 'legend' })
    assert.notEqual(last.hit.element.name, initial)
    assert.ok(last.hit.bounds.x >= 80 && last.hit.bounds.x + last.hit.bounds.width <= 280)
    for (const element of value.instance.getZr().storage.getDisplayList(true)) if (element.name === 'pagePrev' || element.name === 'pageNext') {
      const rect = element.getBoundingRect(), transform = element.getComputedTransform(), x = rect.x + rect.width / 2, y = rect.y + rect.height / 2
      const point = transform ? { x: transform[0] * x + transform[2] * y + transform[4], y: transform[1] * x + transform[3] * y + transform[5] } : { x, y }
      assert.equal(after.hit(point), undefined)
    }
  } finally { value.dispose() }
})
