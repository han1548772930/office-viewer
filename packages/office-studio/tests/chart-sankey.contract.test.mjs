import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { chartGraphData, readSpecialChartData } from '../src/features/charts/chart-data.ts'
import { createChart, readChartData, updateChart } from '../src/features/charts/chart-engine.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { chartColorChannels } from '../src/features/charts/chart-color.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { chartSankeyLayout, moveChartSankeyNode } from '../src/features/charts/chart-sankey-layout.ts'
import { chartSankeyLabelAnchor, chartSankeyNodePercentages, createChartSankeyScene, moveChartSankeyScene } from '../src/features/charts/chart-sankey-draw.ts'
import { chartSankeyFocusedCommands, chartSankeyTrajectory, createChartSankeyHitIndex, createChartSankeyTopology, hitChartSankey } from '../src/features/charts/chart-sankey-interaction.ts'
import { chartSeriesControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'

// P20 drafts only. No runner, browser, type check, build or benchmark has run.
const palette = ['#112233', '#223344', '#334455', '#445566', '#556677', '#667788']
const plot = { x: 10, y: 20, width: 420, height: 260 }
const rows = () => [['R', 'A', 10], ['A', 'B', 4], ['A', 'C', 6], ['B', 'D', 4], ['C', 'E', 6]]
const records = entries => ({ fields: ['Source', 'Target', 'Weight'], records: entries })
function model(options = {}) {
  return { id: 'sankey-draft', sheetId: 'sheet', revision: 0, type: 'sankey', categoryField: 0, valueFields: [1, 2],
    source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 5, endCol: 2 } },
    position: { x: 0, y: 0 }, size: { width: 480, height: 340 }, palette, legend: { visible: false, position: 'bottom' },
    xAxis: { scaleType: 'linear' }, yAxis: { scaleType: 'linear' }, options: { mapping: { edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } }, ...options } }
}
const data = (entries = rows(), chart = model()) => readSpecialChartData(records(entries), chart)
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} ≈ ${expected}`)
const center = node => ({ x: node.x + node.width / 2, y: node.y + node.height / 2 })
function geometry(entries = rows(), chart = model()) {
  const source = data(entries, chart), derived = deriveChartData(source, chart), layout = layoutDerivedChartData(derived, chart)
  return { source, derived, layout, scene: createChartSankeyScene(source, chart, layout.sankey, '#abcdef', palette, layout.frame) }
}
function polygonArea(points) {
  let area = 0
  for (let index = 0; index < points.length; index++) {
    const current = points[index], next = points[(index + 1) % points.length]
    area += current.x * next.y - next.x * current.y
  }
  return Math.abs(area / 2)
}

test('duplicate directed edges aggregate while transit nodes carry max(input, output) and source identity stays stable', () => {
  const source = records([[' A ', 'B', '2'], ['A', 'B', 3], ['B', 'C', 5], ['', 'D', 9], ['missing', 'D', null], ['zero', 'D', 0]])
  const before = structuredClone(source), graph = chartGraphData(source, model())
  assert.deepEqual(graph.nodes, [{ name: 'A', value: 5 }, { name: 'B', value: 5 }, { name: 'C', value: 5 }])
  assert.deepEqual(graph.links, [{ source: 'A', target: 'B', value: 5 }, { source: 'B', target: 'C', value: 5 }])
  assert.equal(graph.cyclic, false)
  assert.deepEqual(source, before)
  const relation = chartGraphData(source, { ...model(), type: 'relation' })
  assert.equal(relation.nodes.find(node => node.name === 'B').value, 10)
})

test('self loops count one carried flow and complete cycle detection includes cycles longer than a reciprocal pair', () => {
  const self = chartGraphData(records([['A', 'A', 5]]), model())
  assert.deepEqual(self.nodes, [{ name: 'A', value: 5 }]); assert.equal(self.cyclic, true)
  const graph = chartGraphData(records([['A', 'B', 3], ['B', 'C', 3], ['C', 'A', 3]]), model())
  assert.equal(graph.cyclic, true); assert.deepEqual(graph.nodes.map(node => node.value), [3, 3, 3])
  const chord = chartGraphData(records([['A', 'B', 2], ['B', 'A', 3]]), { ...model(), type: 'chord' })
  assert.deepEqual(chord.links, [{ source: 'A', target: 'B', value: 5 }])
})

test('missing weight counts occurrences, numeric endpoints are names, and negative or overflowing node flows fail explicitly', () => {
  const chart = model({ mapping: { edge: { sourceIndex: 0, targetIndex: 1 } } })
  assert.deepEqual(chartGraphData(records([[0, 'B'], [0, 'B']]), chart).links, [{ source: '0', target: 'B', value: 2 }])
  assert.throws(() => chartGraphData(records([['A', 'B', -1]]), model()), RangeError)
  assert.throws(() => chartGraphData(records([['A', 'B', Number.MAX_VALUE], ['A', 'C', Number.MAX_VALUE]]), model()), RangeError)
  assert.throws(() => chartGraphData(records(rows()), model({ mapping: { edge: { sourceIndex: 0, targetIndex: 0, valueIndex: 2 } } })), RangeError)
})

test('balanced maximum finite flows retain real values and finite geometry without input-plus-output overflow', () => {
  const graph = chartGraphData(records([['A', 'B', Number.MAX_VALUE], ['B', 'C', Number.MAX_VALUE]]), model())
  const layout = chartSankeyLayout(graph, plot)
  assert.deepEqual(layout.nodes.map(node => node.value), [Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE])
  assert.equal(layout.nodes[1].inValue, Number.MAX_VALUE); assert.equal(layout.nodes[1].outValue, Number.MAX_VALUE)
  for (const link of layout.links) {
    assert.ok(link.width > 0); assert.ok(polygonArea(link.band) > 0)
    assert.equal([...link.points, ...link.band].every(point => Number.isFinite(point.x) && Number.isFinite(point.y)), true)
  }
})

test('justify puts disconnected short-branch terminals in the final column without changing source indices', () => {
  const graph = chartGraphData(records([['A', 'B', 6], ['B', 'C', 6], ['X', 'Y', 4]]), model()), before = structuredClone(graph)
  const layout = chartSankeyLayout(graph, plot), byName = new Map(layout.nodes.map(node => [node.name, node]))
  assert.equal(byName.get('B').layer, 1)
  assert.equal(byName.get('C').layer, 2); assert.equal(byName.get('Y').layer, 2)
  near(byName.get('Y').x, byName.get('C').x)
  assert.deepEqual(layout.nodes.map(node => [node.name, node.index]), graph.nodes.map((node, index) => [node.name, index]))
  for (const layer of new Set(layout.nodes.map(node => node.layer))) {
    const column = layout.nodes.filter(node => node.layer === layer).sort((a, b) => a.y - b.y)
    for (let index = 1; index < column.length; index++) assert.ok(column[index].y >= column[index - 1].y + column[index - 1].height + layout.gap - 1e-7)
  }
  assert.deepEqual(graph, before)
})

test('forward bands meet exact vertical ports and their widths share one node-flow scale', () => {
  const graph = chartGraphData(records(rows()), model()), layout = chartSankeyLayout(graph, plot)
  for (const link of layout.links) {
    const source = layout.nodes[link.sourceIndex], target = layout.nodes[link.targetIndex]
    near(link.band[0].x, source.x + source.width); near(link.band[16].x, target.x)
    near(link.band.at(-1).y - link.band[0].y, link.width)
    near(link.band[17].y - link.band[16].y, link.width)
    assert.equal(link.feedback, false); assert.ok(polygonArea(link.band) > 0)
  }
  const a = layout.nodes.find(node => node.name === 'A')
  near(layout.links.filter(link => link.sourceIndex === a.index).reduce((sum, link) => sum + link.width, 0), a.height)
})

test('self and three-node cycles produce bounded real feedback bands instead of an empty state', () => {
  for (const entries of [[['A', 'A', 5]], [['A', 'B', 5], ['B', 'C', 5], ['C', 'A', 5]]]) {
    const graph = chartGraphData(records(entries), model()), layout = chartSankeyLayout(graph, plot)
    assert.equal(layout.cyclic, true); assert.equal(layout.links.length, entries.length)
    for (const link of layout.links) {
      assert.equal(link.feedback, true); assert.ok(polygonArea(link.band) > 0)
      for (const point of link.band) {
        assert.ok(point.x >= plot.x - 1e-7 && point.x <= plot.x + plot.width + 1e-7)
        assert.ok(point.y >= plot.y - 1e-7 && point.y <= plot.y + plot.height + 1e-7)
      }
      near(link.band[0].x, layout.nodes[link.sourceIndex].x + layout.nodes[link.sourceIndex].width)
      near(link.band[0].y - link.band.at(-1).y, link.width)
    }
  }
})

test('layout validates endpoint identity, resource budgets and finite nonnegative plot rectangles', () => {
  const graph = chartGraphData(records(rows()), model())
  assert.throws(() => chartSankeyLayout(graph, { ...plot, width: -1 }), RangeError)
  assert.throws(() => chartSankeyLayout(graph, { ...plot, x: Infinity }), RangeError)
  assert.throws(() => chartSankeyLayout(graph, plot, { maxNodes: 2, maxLinks: 10, maxPoints: 1000 }), RangeError)
  assert.throws(() => chartSankeyLayout(graph, plot, { maxNodes: 10, maxLinks: 10, maxPoints: 50 }), RangeError)
  assert.throws(() => chartSankeyLayout({ ...graph, nodes: [...graph.nodes, graph.nodes[0]] }, plot), TypeError)
  assert.throws(() => chartSankeyLayout({ ...graph, links: [{ source: 'unknown', target: 'A', value: 1 }] }, plot), TypeError)
  assert.deepEqual(chartSankeyLayout(graph, { ...plot, width: 0 }).links, [])
})

test('iterative SCC discovery supports a long source chain and deterministic identities', () => {
  const entries = Array.from({ length: 1200 }, (_item, index) => [`N${index}`, `N${index + 1}`, 1])
  const layout = chartSankeyLayout(chartGraphData(records(entries), model()), plot)
  assert.equal(layout.nodes.length, 1201); assert.equal(layout.links.length, 1200)
  assert.equal(layout.nodes.at(-1).layer, 1200)
  assert.deepEqual(layout.nodes.map(node => node.index), Array.from({ length: 1201 }, (_item, index) => index))
})

test('prepared Canvas layout is reused for drawing and source point identities survive vertical relaxation', () => {
  const { source, layout, scene } = geometry()
  assert.deepEqual(layout.series[0].points.map(point => point.categoryIndex), layout.sankey.nodes.map(node => node.index))
  const commands = drawChartData(source, model()), bands = commands.filter(command => command.kind === 'polygon')
  assert.deepEqual(bands.slice(0, scene.layout.links.length).map(command => command.points), scene.layout.links.map(link => link.band))
  const option = chartEChartsOption(createBlankWorkbook(), model(), {}, { preparedData: source })
  assert.equal(option.officeChartPointer.mode, 'sankey'); assert.equal(option.series, undefined)
  const group = option.graphic.elements.find(element => element.id === 'sankey-draft:sankey-content')
  assert.equal(group.type, 'group'); assert.ok(group.children.length > 0)
  assert.deepEqual(option.officeChartPointer.scene.layout, layout.sankey)
})

test('node paints inherit shared, field and named styles while bands use source colors and independent fixed alpha', () => {
  const chart = model({ appearance: { font: { family: 'Georgia', italic: true } },
    allSeriesStyle: { color: 'red', fillOpacity: 0.5, label: { visible: true, contentType: 14, format: '0.00' } },
    series: { 2: { color: 'blue', border: { color: 'black', width: 2, opacity: 0.8 } } },
    namedSeries: [{ name: 'Weight', style: { dataPoints: { 2: { color: '#00ff00' } } } }] })
  const { scene } = geometry(rows(), chart), b = scene.layout.nodes.find(node => node.name === 'B'), mark = scene.marks.find(mark => mark.kind === 'node' && mark.index === b.index)
  assert.equal(mark.commands[0].fill, 'rgba(0, 255, 0, 0.5)')
  const band = scene.marks.find(mark => mark.kind === 'link' && scene.layout.links.find(link => link.index === mark.index).sourceIndex === b.index)
  near(chartColorChannels(band.commands[0].fill)[3], 0.1)
  const label = scene.commands.find(command => command.kind === 'text' && command.text === 'Weight,B,4.00')
  assert.equal(label.font, 'italic 10px Georgia')
  assert.ok(scene.commands.some(command => command.kind === 'line' && Math.abs(chartColorChannels(command.color)[3] - 0.8) < 1e-7))
})

test('all authored label positions have real anchors and default terminal labels face inward', () => {
  const { layout } = geometry(), node = layout.sankey.nodes.find(node => node.name === 'D'), frame = layout.frame
  const anchor = position => chartSankeyLabelAnchor(position, node, layout.sankey, 10, frame)
  assert.equal(anchor('auto').align, 'right'); assert.equal(anchor('outside').align, 'right')
  assert.ok(anchor('left').x < node.x); assert.ok(anchor('right').x > node.x + node.width)
  assert.ok(anchor('top').y < node.y); assert.ok(anchor('bottom').y > node.y + node.height)
  near(anchor('inside').x, node.x + node.width / 2); near(anchor('center').y, node.y + node.height / 2)
  assert.equal(anchor('insideBase').align, 'left'); assert.equal(anchor('insideEnd').align, 'right')
  assert.ok(anchor('insideBase').x < anchor('insideEnd').x)
  assert.deepEqual(anchor('inside').clip, { x: node.x, y: node.y, width: node.width, height: node.height })
})

test('Sankey label defaults match controls, combined content and percentage are real, and inactive markers stay authored', () => {
  const chart = model({ allSeriesStyle: { point: { shape: 'diamond', size: 40, color: 'purple' } } }), controls = chartSeriesControls()
  const visible = controls.find(control => control.key === 'allSeriesStyle.label.visible'), content = controls.find(control => control.key === 'allSeriesStyle.label.contentType')
  assert.equal(visible.read(chart), true); assert.deepEqual(content.read(chart), ['category'])
  for (const control of controls.filter(control => control.key.startsWith('allSeriesStyle.point.'))) assert.equal(control.visible(chart), false)
  const saved = structuredClone(chart.options.allSeriesStyle.point)
  visible.write(chart, false); assert.equal(geometry(rows(), chart).scene.commands.some(command => command.kind === 'text'), false)
  assert.deepEqual(chart.options.allSeriesStyle.point, saved)
  const combined = model({ allSeriesStyle: { label: { visible: true, contentType: 30, format: '0.00%' } } }), result = geometry(rows(), combined)
  assert.ok(result.scene.commands.some(command => command.kind === 'text' && command.text === 'Weight,B,400.00%,40.00%'))
  near(chartSankeyNodePercentages(result.layout.sankey)[result.layout.sankey.nodes.find(node => node.name === 'B').index], 40)
  const empty = geometry(rows(), model({ allSeriesStyle: { label: { visible: true, contentType: 0 } } }))
  assert.equal(empty.scene.commands.some(command => command.kind === 'text'), false)
})

test('directed trajectories exclude sibling forks and terminate on cycles', () => {
  const { layout } = geometry(), topology = createChartSankeyTopology(layout.sankey)
  const index = name => layout.sankey.nodes.find(node => node.name === name).index
  const link = layout.sankey.links.find(link => link.source === 'A' && link.target === 'B')
  const selected = chartSankeyTrajectory(topology, { kind: 'link', index: link.index })
  assert.deepEqual([...selected.nodes].sort((a, b) => a - b), ['R', 'A', 'B', 'D'].map(index).sort((a, b) => a - b))
  assert.equal(selected.nodes.has(index('C')), false); assert.equal(selected.nodes.has(index('E')), false)
  const cycle = chartSankeyLayout(chartGraphData(records([['A', 'B', 1], ['B', 'C', 1], ['C', 'A', 1]]), model()), plot)
  const cyclic = chartSankeyTrajectory(createChartSankeyTopology(cycle), { kind: 'node', index: 0 })
  assert.equal(cyclic.nodes.size, 3); assert.equal(cyclic.links.size, 3)
})

test('node priority and filled-band hit tests reject bounding-box false positives and nonfinite points', () => {
  const { layout } = geometry(), hitIndex = createChartSankeyHitIndex(layout.sankey), node = layout.sankey.nodes[0]
  assert.deepEqual(hitChartSankey(hitIndex, center(node)), { kind: 'node', index: node.index })
  const triangle = { plot: { x: 0, y: 0, width: 100, height: 100 }, nodes: [], links: [{ index: 7,
    bounds: { x: 0, y: 0, width: 100, height: 100 }, band: [{ x: 0, y: 0 }, { x: 100, y: 100 }, { x: 0, y: 100 }] }] }
  assert.equal(hitChartSankey(triangle, { x: 90, y: 10 }), undefined)
  assert.deepEqual(hitChartSankey(triangle, { x: 10, y: 50 }), { kind: 'link', index: 7 })
  assert.deepEqual(hitChartSankey(triangle, { x: 0, y: 50 }), { kind: 'link', index: 7 })
  assert.equal(hitChartSankey(triangle, { x: NaN, y: 50 }), undefined)
})

test('hover multiplies existing alpha and emphasizes the selected edge without mutating the base scene', () => {
  const { scene } = geometry(), before = structuredClone(scene.commands), topology = createChartSankeyTopology(scene.layout)
  const link = scene.layout.links.find(link => link.source === 'A' && link.target === 'B'), hit = { kind: 'link', index: link.index }
  const trajectory = chartSankeyTrajectory(topology, hit), focused = chartSankeyFocusedCommands(scene, trajectory, hit)
  const selected = focused.find(command => command.kind === 'polygon' && command.points === link.band)
  near(chartColorChannels(selected.fill)[3], 0.5)
  const unrelated = scene.layout.links.find(link => link.source === 'A' && link.target === 'C')
  const faded = focused.find(command => command.kind === 'polygon' && command.points === unrelated.band)
  near(chartColorChannels(faded.fill)[3], 0.03)
  assert.deepEqual(scene.commands, before); assert.equal(chartSankeyFocusedCommands(scene), scene.commands)
})

test('view drag clamps positions, rebuilds exact ports and preserves author paint, flow identity and input geometry', () => {
  const { scene } = geometry(), before = structuredClone(scene.layout), node = scene.layout.nodes.find(node => node.name === 'D')
  const moved = moveChartSankeyScene(scene, node.index, { x: node.x - 40, y: node.y + 20 })
  assert.notEqual(moved.layout.nodes[node.index].x, node.x); assert.equal(moved.paint, scene.paint)
  assert.deepEqual(scene.layout, before)
  for (const link of moved.layout.links) {
    near(link.points[0].x, moved.layout.nodes[link.sourceIndex].x + moved.layout.nodes[link.sourceIndex].width)
    near(link.points.at(-1).x, moved.layout.nodes[link.targetIndex].x)
  }
  const clamped = moveChartSankeyNode(moved.layout, node.index, { x: -1e6, y: 1e6 }).nodes[node.index]
  near(clamped.x, scene.layout.nodeArea.x); near(clamped.y + clamped.height, scene.layout.nodeArea.y + scene.layout.nodeArea.height)
  assert.throws(() => moveChartSankeyNode(scene.layout, -1, { x: 0, y: 0 }), RangeError)
})

test('source edits reproject flows and malformed mapping updates leave the chart model unchanged', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0], table = [['Source', 'Target', 'Weight'], ...rows()]
  sheet.cells = Object.fromEntries(table.flatMap((row, r) => row.map((value, c) => [`${r}:${c}`, createCell(r, c, value)])))
  const chart = createChart(workbook, sheet.id, model()), before = structuredClone(chart)
  assert.throws(() => updateChart(workbook, chart, { options: { mapping: { edge: { sourceIndex: 0, targetIndex: 0, valueIndex: 2 } } } }), RangeError)
  assert.deepEqual(chart, before)
  sheet.cells['2:2'] = createCell(2, 2, 8)
  assert.equal(readChartData(workbook, chart).graph.nodes.find(node => node.name === 'A').value, 14)
  assert.deepEqual(chart, before)
})
