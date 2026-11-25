import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChart } from '../src/features/charts/chart-engine.ts'
import { readSpecialChartData } from '../src/features/charts/chart-data.ts'
import { deriveChartData } from '../src/features/charts/chart-derived.ts'
import { layoutDerivedChartData } from '../src/features/charts/chart-layout.ts'
import { drawChartData } from '../src/features/charts/chart-draw.ts'
import { chartEChartsOption } from '../src/features/charts/echarts-option.ts'
import { chartChordLayout } from '../src/features/charts/chart-chord-layout.ts'
import { createChartChordScene, chartChordRibbonGradient } from '../src/features/charts/chart-chord-draw.ts'
import { chartChordContentTooltip, chartChordFocusedCommands, chartChordFocusSelection, createChartChordHitIndex, createChartChordTopology, hitChartChord } from '../src/features/charts/chart-chord-interaction.ts'
import { createChartChordPlanSteps, chartChordResultAsync } from '../src/features/charts/chart-chord-data.ts'
import { chartChordAngleParameters, decodeChartChordSteps, executeChartChordTaskSteps } from '../src/features/charts/chart-chord-compute.ts'
import { consumeChartPreparationSync } from '../src/features/charts/chart-compute-lifecycle.ts'
import { chartChordLayoutValues, chartChordLayoutPatch, chartChordMappingValues, chartChordNodeValues, chartChordLabelValues, chartChordRibbonValues, chartChordRibbonPatch, chartChordEmphasisValues } from '../src/features/charts/chart-chord-api.ts'
import { resolveChartChordLayout, resolveChartChordOptions } from '../src/features/charts/chart-chord.ts'
import { CHART_CHORD_COLOR_TOKENS, resolveChartChordColor, validateChartChordColor } from '../src/features/charts/chart-chord-color.ts'
import { chartColorChannels } from '../src/features/charts/chart-color.ts'
import { cloneChartOptions, mergeChartOptions } from '../src/features/charts/chart-options.ts'
import { chartCommonControls } from '../src/features/charts/chart-settings-controls.ts'
import { chartChordLayoutControls, chartChordNodeControls, chartChordLabelControls, chartChordRibbonControls, chartChordEmphasisControls,
  chartChordMappingControls, chartChordSettingsRadiiError, chartChordSettingsMappingError, clearChartChordSettings, setChartChordSettingsField } from '../src/features/charts/chart-chord-settings-controls.ts'
import { cloneChartDraftConfig } from '../src/features/charts/chart-settings-draft.ts'
import { estimateChartDataMemory, estimateChartDecodedMemory, estimateChartFinalLayoutMemory } from '../src/features/charts/chart-compute-memory.ts'
import { estimateChartTaskMemory, validateChartTask } from '../src/compute/chart-tasks.ts'
import { CHORD_CASES, CHORD_PALETTE, chordGoldenOutput, chordGoldenPlan, chordGoldenTask, chordModel, chordRecords, goldenValue } from './fixtures/chart-chord-fixtures.mjs'

// P28 drafts only. No runner, compiler, server, browser, WASM or benchmark ran.
const consume = consumeChartPreparationSync
const near = (actual, expected, tolerance = 1e-9) => assert.ok(Math.abs(actual - expected) <= Math.max(1, Math.abs(actual), Math.abs(expected)) * tolerance, `${actual} ≈ ${expected}`)
const alpha = command => chartColorChannels(command.kind === 'line' || command.kind === 'text' ? command.color : command.fill)[3]
const plot = { x: 10, y: 20, width: 200, height: 200 }
function prepared(records = chordRecords(), chart = chordModel()) {
  const data = readSpecialChartData(records, chart), derived = deriveChartData(data, chart), layout = layoutDerivedChartData(derived, chart, CHORD_PALETTE)
  return { data, derived, layout, scene: createChartChordScene(data, chart, layout.chord, {}, CHORD_PALETTE, layout.frame) }
}
const layoutControl = key => chartChordLayoutControls().find(control => control.key === `chord.layout.${key}`)

for (const fixture of CHORD_CASES) test(`Chord portable independent golden: ${fixture.name}`, () => {
  const task = chordGoldenTask(fixture), original = structuredClone(task), output = consume(executeChartChordTaskSteps(task))
  assert.deepEqual(Array.from(output.indices), fixture.indices)
  assert.equal(output.values.length, fixture.values.length)
  fixture.values.forEach((value, index) => Array.isArray(value) ? near(output.values[index], goldenValue(value)) : assert.equal(output.values[index], value))
  assert.equal(output.segments[0].validSampleCount, fixture.validSampleCount); assert.equal(output.segments[0].available, fixture.available)
  const decoded = consume(decodeChartChordSteps(chordGoldenPlan(fixture), output))
  assert.equal(decoded.nodes.length, fixture.indices[0]); assert.equal(decoded.links.length, fixture.indices[1])
  assert.deepEqual(task, original)
})

test('display text owns identity without trimming; only positive finite numeric weights create nodes', () => {
  const records = chordRecords([[' A ', 'B', 2], ['B', ' A ', 3], ['A', 'B', 4], [' ', 'B', 1], ['', 'B', 9],
    ['B', 'B', 5], ['String', 'B', '2'], ['Zero', 'B', 0], ['Negative', 'B', -2], ['Invalid', 'B', Infinity], ['Missing', 'B', null]])
  const before = structuredClone(records), data = readSpecialChartData(records, chordModel())
  assert.deepEqual(data.chord.nodes.map(node => [node.name, node.value]), [[' A ',5],['B',10],['A',4],[' ',1]])
  assert.deepEqual(data.chord.links.map(link => [link.source, link.target, link.value]), [[' A ','B',5],['A','B',4],[' ','B',1]])
  assert.equal(data.chord.validSampleCount, 4); assert.deepEqual(records, before)
  assert.equal(data.graph.nodes, data.chord.nodes); assert.equal(data.graph.links, data.chord.links)
})

test('formatted endpoint text merges raw IDs and count mode preserves whitespace and numeric labels', () => {
  const records = chordRecords([[1,2,7],[3,4,8],[0,' ',null],['',' ',4],[0,0,9]],
    [['One','Two','7'],['One','Two','8'],['0',' ',null],['',' ','4'],['0','0','9']])
  const weighted = readSpecialChartData(records, chordModel())
  assert.deepEqual(weighted.chord.nodes.map(node => [node.name,node.value]), [['One',15],['Two',15]])
  const count = readSpecialChartData(records, chordModel({ mapping: { chord: { sourceIndex: 0, targetIndex: 1 } } }))
  assert.equal(count.chord.weightMode, 'count')
  assert.deepEqual(count.chord.links.map(link => [link.source,link.target,link.value]), [['One','Two',2],['0',' ',1]])
  assert.equal(count.chord.validSampleCount, 3)
})

test('weights and incident totals format once after undirected aggregation using the first positive display format', () => {
  const records = chordRecords([['A','B',1.25],['B','A',2.5],['A','C',0.75]], [['A','B','$1.25'],['B','A','$2.50'],['A','C','$0.75']])
  const { scene, data } = prepared(records)
  assert.deepEqual(data.chord.links.map(link => link.formattedValue), ['$3.75','$0.75'])
  assert.deepEqual(data.chord.nodes.map(node => node.formattedValue), ['$4.50','$3.75','$0.75'])
  assert.deepEqual(chartChordContentTooltip(scene, { kind: 'link', index: 0 }), { id: 'chord-link:0', kind: 'link', index: 0, title: 'A — B', lines: ['Weight：$3.75'] })
  assert.equal(chartChordContentTooltip(scene, { kind: 'node', index: 0 }).lines[0], 'Weight：$4.50')
})

test('all skipped rows remain a valid empty chart, while link and node aggregate overflow reject explicitly', () => {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  const table = [['Source','Target','Weight'],['A','A',4],['A','B',-1],['A','B','2']]
  sheet.cells = Object.fromEntries(table.flatMap((row,r) => row.map((value,c) => [`${r}:${c}`,createCell(r,c,value)])))
  const chart = createChart(workbook, sheet.id, chordModel())
  assert.equal(readSpecialChartData(chordRecords(table.slice(1)), chart).chord.nodes.length, 0)
  const empty = chartEChartsOption(workbook, chart)
  assert.equal(empty.officeChartPointer, undefined); assert.equal(empty.series, undefined)
  assert.ok(empty.graphic.elements.length > 0)
  for (const rows of [[['A','B',Number.MAX_VALUE],['B','A',Number.MAX_VALUE]], [['A','B',Number.MAX_VALUE],['A','C',Number.MAX_VALUE]]]) {
    assert.throws(() => readSpecialChartData(chordRecords(rows), chordModel()), RangeError)
  }
})

test('fixed defaults, legacy angle translation and nested author precedence resolve per leaf', () => {
  assert.deepEqual(resolveChartChordLayout(), { startAngle:90,clockwise:true,padAngle:2,minAngle:0,innerRadiusRatio:0.58,outerRadiusRatio:0.78 })
  const options = resolveChartChordOptions()
  assert.equal(options.node.opacity,1); assert.equal(options.node.borderColor,'#FFFFFF'); assert.equal(options.node.borderWidth,1)
  assert.deepEqual(options.ribbon,{color:'gradient',opacity:0.65,width:0,type:'solid'}); assert.equal(options.emphasis.focus,'adjacency')
  assert.equal(options.label.visible,true); assert.equal(options.label.position,'outside'); assert.equal(options.label.fontSize,12)
  assert.equal(options.label.rotation,0); assert.equal(options.label.distance,8); assert.equal(options.label.family,'Arial')
  const legacy = { startAngle:-35, clockwise:false,padAngle:7,minAngle:2,innerRadiusRatio:0.5,outerRadiusRatio:0.9,ribbonOpacity:0.27,layout:{padAngle:5},ribbon:{color:'target'} }
  assert.deepEqual(resolveChartChordLayout(legacy), {startAngle:35,clockwise:false,padAngle:5,minAngle:2,innerRadiusRatio:0.5,outerRadiusRatio:0.9})
  assert.equal(resolveChartChordOptions(legacy).ribbon.opacity,0.27)
})

test('nullable API leaves reject unknown shapes and values, accept font zero and defer paired radius validation to final merge', () => {
  assert.deepEqual(chartChordLayoutValues({innerRadiusRatio:0.9,startAngle:null}),{innerRadiusRatio:0.9,startAngle:undefined})
  assert.equal(chartChordLabelValues({fontSize:0}).fontSize,0); assert.equal(chartChordLayoutValues({padAngle:720,minAngle:500}).padAngle,720)
  for (const [validator,value] of [[chartChordLayoutValues,{padAngle:-1}], [chartChordNodeValues,{opacity:1.1}], [chartChordLabelValues,{position:'left'}],
    [chartChordRibbonValues,{color:'red'}], [chartChordEmphasisValues,{focus:'trajectory'}], [chartChordLabelValues,{family:' '}], [chartChordLayoutValues,{startAngle:Infinity}],
    [chartChordLabelValues,{font:{fontSize:12}}], [chartChordNodeValues,{color:'gray.999'}], [chartChordNodeValues,Object.assign({},{[Symbol('unknown')]:1})]]) assert.throws(() => validator(value))
  for (const value of [{sourceIndex:0,targetIndex:0},{sourceIndex:0,targetIndex:1,valueIndex:1},{sourceIndex:0.5,targetIndex:1},{sourceIndex:0,targetIndex:1,valueIndex:null}]) assert.throws(() => chartChordMappingValues(value))
  assert.throws(() => cloneChartOptions({chord:{layout:{innerRadiusRatio:0.9}}}),RangeError)
  const current = {chord:{layout:{innerRadiusRatio:0.9,outerRadiusRatio:0.95}}}, before = structuredClone(current)
  assert.equal(mergeChartOptions(current,{chord:{layout:chartChordLayoutValues({innerRadiusRatio:0.92})}}).chord.layout.outerRadiusRatio,0.95)
  assert.throws(() => mergeChartOptions(current,{chord:{layout:{outerRadiusRatio:0.8}}}),RangeError); assert.deepEqual(current,before)
})

test('known theme tokens resolve by current theme while original authors and independent CSS alpha are retained', () => {
  assert.equal(CHART_CHORD_COLOR_TOKENS.length,140)
  for (const token of CHART_CHORD_COLOR_TOKENS) assert.doesNotThrow(() => validateChartChordColor(token))
  const author = {color:'blue.500',borderColor:'gray.700'}, before = {...author}
  assert.notEqual(resolveChartChordColor(author.color,{backgroundColor:'#ffffff'}),resolveChartChordColor(author.color,{backgroundColor:'#000000'}))
  assert.equal(resolveChartChordColor('gray.700',{textColor:'#abcdef'}),'#abcdef')
  assert.equal(resolveChartChordColor('blue.500',{tokenColors:{'blue.500':'rgba(1,2,3,0.4)'}}),'rgba(1,2,3,0.4)')
  assert.deepEqual(author,before); assert.throws(() => validateChartChordColor('unknown.500'))
})

test('source planning uses three numeric columns, bounded ownership, genuine async task dispatch and cancellation', async () => {
  const chart = chordModel({chord:{layout:{padAngle:0}}}), records = chordRecords(), plan = consume(createChartChordPlanSteps(records,chart))
  assert.deepEqual(plan.labels,['B','A']); assert.deepEqual(Array.from(plan.task.values),[2,3,0,0,1,0,1,0,0])
  assert.deepEqual(Array.from(plan.task.states),[1,1,0,1,1,0,1,1,0])
  assert.throws(() => consume(createChartChordPlanSteps(records,chart,{memoryBudgetBytes:1})),/预算|内存/)
  let calls = 0
  const result = await chartChordResultAsync(records,chart,{runChartTask:async task => {
    calls++; assert.deepEqual(task.params,plan.task.params)
    const output = chordGoldenOutput(CHORD_CASES[0]); output.indices = new Uint32Array([2,1,0,1,0,1])
    return output
  }})
  assert.equal(calls,1); assert.equal(result.links[0].value,5)
  const controller = new AbortController(); controller.abort()
  await assert.rejects(() => chartChordResultAsync(records,chart,{signal:controller.signal,runChartTask:async () => {calls++;throw new Error('unexpected dispatch')}}))
  assert.equal(calls,1)
})

test('decoder rejects forged identity, ownership, counts, ports, arc closure and very small inconsistent totals', () => {
  const fixture = CHORD_CASES[0], plan = chordGoldenPlan(fixture)
  const changes = [output => {output.indices[3]=output.indices[2]}, output => {output.indices[5]=output.indices[4]},
    output => {output.values[2]+=0.1}, output => {output.values[7]+=0.1}, output => {output.values[6]=9},
    output => {output.values[0]=Infinity}, output => {output.segments[0].available=false}, output => {output.segments[0].validSampleCount=99},
    output => {output.values=new Float64Array(output.values.buffer,8)}, output => {output.segments[0].field=999}]
  for (const change of changes) {const output=chordGoldenOutput(fixture);change(output);assert.throws(() => consume(decodeChartChordSteps(plan,output)))}
  const tinyPlan = chordGoldenPlan(fixture), tiny = chordGoldenOutput(fixture)
  for (const at of [0,3,6]) tiny.values[at] = 1e-240
  tinyPlan.task.params.jobs[0].padAngle = 180
  for (const at of [1,2,7,8]) tiny.values[at] = -Math.PI/2
  for (const at of [4,5,9,10]) tiny.values[at] = Math.PI/2
  tiny.values[0] = 2e-240
  assert.throws(() => consume(decodeChartChordSteps(tinyPlan,tiny)),/合计|数值/)
})

test('padding/minimum saturate to one complete circle and both directions remain finite for extreme finite authors', () => {
  for (const clockwise of [true,false]) for (const [padAngle,minAngle] of [[2,10],[720,0],[1e308,1e308],[0,0]]) {
    const fixture = {...CHORD_CASES[5],layout:{startAngle:1e308,clockwise,padAngle,minAngle}}, output = consume(executeChartChordTaskSteps(chordGoldenTask(fixture)))
    const options = chartChordAngleParameters(fixture.layout,4), spans = [0,1,2,3].map(index => (output.values[index*3+2]-output.values[index*3+1])*options.direction)
    for (const span of spans) {assert.ok(Number.isFinite(span));assert.ok(span>=options.minimum-1e-9)}
    near(spans.reduce((sum,span)=>sum+span,0)+options.padding*4,Math.PI*2)
  }
})

test('viewport layout only maps retained angles, rejects stale options and never reparses source rows', () => {
  const chart = chordModel(), data = readSpecialChartData(chordRecords(),chart), before = structuredClone(data.chord)
  const layout = chartChordLayout(data.chord,chart,plot)
  assert.deepEqual(layout.center,{x:110,y:120});near(layout.innerRadius,58);near(layout.outerRadius,78)
  assert.deepEqual(layout.nodes.map(node=>[node.start,node.end]),data.chord.nodes.map(node=>[node.start,node.end]))
  assert.throws(() => chartChordLayout(data.chord,{...chart,options:{...chart.options,chord:{layout:{startAngle:91}}}},plot),/过期/)
  assert.deepEqual(data.chord,before)
})

test('ribbon fill, node fill and outlines have independent alpha, with real dash/point geometry', () => {
  const chart = chordModel({chord:{node:{color:'rgba(255,0,0,0.8)',opacity:0.2,borderColor:'rgba(0,0,0,0.7)',borderWidth:2,borderType:'dashed'},ribbon:{color:'source',opacity:0.5,width:1,type:'dotted'}},
    allSeriesStyle:{fillOpacity:0.3,border:{opacity:0.6}}})
  const {scene}=prepared(chordRecords(),chart), node=scene.marks.find(mark=>mark.kind==='node'&&mark.commands.some(command=>command.kind==='polygon'))
  near(alpha(node.commands.find(command=>command.kind==='polygon')),0.16)
  near(alpha(node.commands.find(command=>command.kind==='line')),0.42)
  const ribbon=scene.marks.find(mark=>mark.kind==='link')
  near(alpha(ribbon.commands.find(command=>command.kind==='polygon')),0.4)
  assert.ok(ribbon.commands.some(command=>command.kind==='line'));assert.ok(node.commands.filter(command=>command.kind==='line').length>4)
})

test('global endpoint-average gradients have analytical quarter-arc coordinates and independent stop alpha', () => {
  const fixture=CHORD_CASES[4], data=consume(decodeChartChordSteps(chordGoldenPlan(fixture),chordGoldenOutput(fixture)))
  const chart=chordModel({chord:{layout:{startAngle:0,padAngle:0}}}), layout=chartChordLayout(data,chart,plot)
  const gradient=chartChordRibbonGradient(layout,layout.links[0],'rgba(255,0,0,0.8)','rgba(0,0,255,0.6)',0.5)
  near(gradient.start.x,0.645);near(gradient.start.y,0.645);near(gradient.end.x,0.355);near(gradient.end.y,0.645)
  assert.deepEqual(gradient.bounds,plot);near(chartColorChannels(gradient.stops[0].color)[3],0.4);near(chartColorChannels(gradient.stops[1].color)[3],0.3)
})

test('labels really rotate, honor inside/outside and font zero, while zero inner radius renders nodes without degenerate bands', () => {
  const chart=chordModel({chord:{layout:{padAngle:0},label:{rotation:25,family:'Georgia',position:'outside',distance:8}}}), {scene}=prepared(chordRecords(),chart)
  const labels=scene.commands.filter(command=>command.kind==='text');assert.equal(labels.length,2)
  assert.equal(labels[0].align,'left');assert.equal(labels[0].baseline,'top');assert.equal(labels[0].rotation,25);assert.equal(labels[0].font,'12px Georgia')
  const inside=prepared(chordRecords(),chordModel({chord:{label:{position:'inside'}}})).scene.commands.filter(command=>command.kind==='text')
  assert.ok(inside.every(label=>label.align==='center'&&label.baseline==='middle'))
  assert.equal(prepared(chordRecords(),chordModel({chord:{label:{fontSize:0}}})).scene.commands.some(command=>command.kind==='text'),false)
  const zero=prepared(chordRecords(),chordModel({chord:{layout:{innerRadiusRatio:0}}})).scene
  assert.equal(zero.marks.filter(mark=>mark.kind==='link').flatMap(mark=>mark.commands).length,0);assert.ok(zero.commands.length>0)
})

test('annular arc hits wrap correctly in either direction, give node priority, and reject the padded ring gaps', () => {
  for (const clockwise of [true,false]) {
    const {scene}=prepared(chordRecords(),chordModel({chord:{layout:{clockwise}}})), index=createChartChordHitIndex(scene.layout), node=scene.layout.nodes[0]
    const angle=(node.start+node.end)/2,radius=(scene.layout.innerRadius+scene.layout.outerRadius)/2
    assert.deepEqual(hitChartChord(index,{x:scene.layout.center.x+Math.cos(angle)*radius,y:scene.layout.center.y+Math.sin(angle)*radius}),{kind:'node',index:0})
    const gap=node.end+(clockwise?1:-1)*Math.PI/180
    assert.equal(hitChartChord(index,{x:scene.layout.center.x+Math.cos(gap)*radius,y:scene.layout.center.y+Math.sin(gap)*radius}),undefined)
  }
})

test('ribbon hit uses its filled polygon and measured rotated labels respect actual clip instead of a guessed text width', () => {
  const index={plot:{x:0,y:0,width:100,height:100},center:{x:50,y:50},innerRadius:20,outerRadius:30,nodes:[],labels:[],
    links:[{index:7,bounds:{x:0,y:0,width:100,height:100},band:[{x:0,y:0},{x:100,y:100},{x:0,y:100}]}]}
  assert.equal(hitChartChord(index,{x:90,y:10}),undefined);assert.deepEqual(hitChartChord(index,{x:10,y:60}),{kind:'link',index:7})
  assert.equal(hitChartChord(index,{x:NaN,y:60}),undefined)
  index.labels=[{index:3,polygon:[{x:105,y:20},{x:115,y:30},{x:105,y:40},{x:95,y:30}],clip:{x:100,y:0,width:20,height:60}}]
  assert.deepEqual(hitChartChord(index,{x:105,y:30}),{kind:'node',index:3});assert.equal(hitChartChord(index,{x:114,y:21}),undefined)
  assert.notDeepEqual(hitChartChord(index,{x:96,y:30}),{kind:'node',index:3})
})

test('adjacency is one undirected hop; self keeps only its mark and none returns the immutable base', () => {
  const {scene}=prepared(chordRecords([['A','B',1],['B','C',1],['C','D',1],['C','E',1]])), topology=createChartChordTopology(scene.layout)
  const node=scene.layout.nodes.find(node=>node.name==='B'), selected=chartChordFocusSelection(topology,'adjacency',{kind:'node',index:node.index})
  assert.deepEqual([...selected.nodes].map(index=>scene.layout.nodes[index].name).sort(),['A','B','C'])
  assert.deepEqual([...selected.links].sort(),[0,1]);assert.equal(topology.incident.length,5)
  const link=chartChordFocusSelection(topology,'adjacency',{kind:'link',index:1});assert.deepEqual([...link.nodes].sort(),[1,2]);assert.deepEqual([...link.links],[1])
  const self=chartChordFocusSelection(topology,'self',{kind:'link',index:1});assert.equal(self.nodes.size,0);assert.deepEqual([...self.links],[1])
  assert.equal(chartChordFocusSelection(topology,'none',{kind:'node',index:1}),undefined)
  assert.equal(chartChordFocusedCommands(scene),scene.commands)
})

test('focus multiplies independent paint alpha, keeps selected base paint unchanged and restores by reference without source edits', () => {
  const {scene}=prepared(chordRecords([['A','B',2],['C','D',3]])), before=structuredClone(scene.commands), hit={kind:'link',index:0}
  const selection=chartChordFocusSelection(createChartChordTopology(scene.layout),'adjacency',hit), focused=chartChordFocusedCommands(scene,selection)
  let at=0
  for (const mark of scene.marks) for (const command of mark.commands) {
    const selected=(mark.kind==='node'?selection.nodes:selection.links).has(mark.index)
    if (selected) assert.equal(focused[at],command)
    else near(alpha(focused[at]),alpha(command)*0.15)
    at++
  }
  assert.deepEqual(scene.commands,before);assert.equal(chartChordFocusedCommands(scene),scene.commands)
})

test('shared Canvas/ECharts adapter owns one Chord scene and suppresses legend even for an authored visible flag', () => {
  const chart=chordModel();chart.legend.visible=true
  const {data}=prepared(chordRecords(),chart), option=chartEChartsOption(createBlankWorkbook(),chart,{}, {preparedData:data})
  assert.equal(option.officeChartPointer.mode,'chord');assert.equal(option.series,undefined)
  assert.equal(option.graphic.elements.find(element=>element.id==='chord-draft:chord-content').type,'group')
  assert.deepEqual(option.officeChartPointer.scene.layout.nodes.map(node=>node.value),[5,5])
  assert.equal(chartCommonControls(chart).some(control=>control.key.startsWith('legend.')),false)
  const commands=drawChartData(data,chart);assert.equal(commands.some(command=>command.kind==='text'&&command.text==='Weight'),false)
})

test('all 28 setting values exist, label visibility hides eight values and temporary radii can be repaired before Apply', () => {
  const chart=chordModel(), controls=[...chartChordMappingControls(3),...chartChordLayoutControls(),...chartChordNodeControls(),...chartChordLabelControls(),...chartChordRibbonControls(),...chartChordEmphasisControls()]
  assert.equal(controls.length,28);assert.equal(new Set(controls.map(control=>control.key)).size,28)
  assert.equal(chartChordLabelControls().filter(control=>control.visible?.(chart)!==false).length,9)
  chart.options.chord={label:{visible:false}};assert.equal(chartChordLabelControls().filter(control=>control.visible?.(chart)!==false).length,1)
  const next=cloneChartDraftConfig(chart);layoutControl('innerRadiusRatio').write(next,0.9)
  assert.ok(chartChordSettingsRadiiError(next));assert.throws(()=>cloneChartOptions(next.options),RangeError)
  const repair=cloneChartDraftConfig(next);layoutControl('outerRadiusRatio').write(repair,0.95)
  assert.equal(chartChordSettingsRadiiError(repair),undefined);assert.doesNotThrow(()=>cloneChartOptions(repair.options))
  assert.equal(chart.options.chord.layout,undefined)
})

test('partial settings migrate old authored layout/ribbon leaves and group resets remove both canonical and flat values', () => {
  const chart=chordModel({chord:{startAngle:-35,clockwise:false,padAngle:7,minAngle:2,innerRadiusRatio:0.5,outerRadiusRatio:0.99,ribbonOpacity:0.27}})
  layoutControl('innerRadiusRatio').write(chart,0.94)
  assert.deepEqual(resolveChartChordLayout(chart.options.chord),{startAngle:35,clockwise:false,padAngle:7,minAngle:2,innerRadiusRatio:0.94,outerRadiusRatio:0.99})
  assert.equal(chart.options.chord.startAngle,undefined);assert.equal(chart.options.chord.outerRadiusRatio,undefined)
  chartChordRibbonControls().find(control=>control.key==='chord.ribbon.color').write(chart,'target')
  assert.equal(chart.options.chord.ribbon.opacity,0.27);assert.equal(chart.options.chord.ribbonOpacity,undefined)
  clearChartChordSettings(chart,'layout');assert.deepEqual(resolveChartChordLayout(chart.options.chord),resolveChartChordLayout())
  assert.equal(chart.options.chord.ribbon.color,'target');clearChartChordSettings(chart,'ribbon');assert.equal(resolveChartChordOptions(chart.options.chord).ribbon.opacity,0.65)
  clearChartChordSettings(chart,'all');assert.equal(chart.options.chord,undefined)
})

test('shared partial patches retain nested precedence, migrate undefined legacy leaves and never author resolved defaults', () => {
  assert.deepEqual(chartChordLayoutPatch(undefined, { padAngle: 4 }).layout, { padAngle: 4 })
  assert.deepEqual(chartChordRibbonPatch(undefined, { color: 'target' }).ribbon, { color: 'target' })
  const legacy = { startAngle: -35, padAngle: 7, outerRadiusRatio: 0.99,
    layout: { startAngle: 80, padAngle: undefined }, ribbonOpacity: 0.27, ribbon: { opacity: undefined, type: 'dotted' } }
  const layout = chartChordLayoutPatch(legacy, { innerRadiusRatio: 0.9 })
  assert.deepEqual(layout.layout, { startAngle: 80, padAngle: 7, outerRadiusRatio: 0.99, innerRadiusRatio: 0.9 })
  assert.equal(layout.startAngle, undefined); assert.equal(layout.padAngle, undefined)
  assert.equal(chartChordLayoutPatch(legacy, { padAngle: null }).layout.padAngle, undefined)
  assert.deepEqual(chartChordRibbonPatch(legacy, { color: 'target' }).ribbon, { opacity: 0.27, type: 'dotted', color: 'target' })
  assert.equal(chartChordRibbonPatch(legacy, { opacity: null }).ribbon.opacity, undefined)
  assert.equal(legacy.layout.padAngle, undefined); assert.equal(legacy.ribbonOpacity, 0.27)
})

test('canonical mapping count clears the old edge weight, diagnoses temporary duplicate roles and allows their repair', () => {
  const chart=chordModel({mapping:{edge:{sourceIndex:0,targetIndex:1,valueIndex:2}}})
  setChartChordSettingsField(chart,'valueIndex',undefined,3)
  assert.deepEqual(chart.options.mapping.chord,{sourceIndex:0,targetIndex:1});assert.equal(chart.options.mapping.edge,undefined);assert.deepEqual(chart.valueFields,[1])
  setChartChordSettingsField(chart,'sourceIndex',1,3);assert.ok(chartChordSettingsMappingError(chart,3))
  setChartChordSettingsField(chart,'targetIndex',0,3);assert.equal(chartChordSettingsMappingError(chart,3),undefined)
  const before=structuredClone(chart);assert.throws(()=>setChartChordSettingsField(chart,'valueIndex',3,3),RangeError);assert.deepEqual(chart,before)
})

test('protocol admission rejects overlapping roles/invalid angles and budgets decoded facts plus retained bands/hover staging', async () => {
  const fixture=CHORD_CASES[1], task=chordGoldenTask(fixture), estimate=estimateChartTaskMemory(task.params)
  assert.ok(estimate.estimatedPeakBytes>task.values.byteLength+task.states.byteLength+task.categoryIds.byteLength)
  for (const patch of [{sourceColumn:0},{targetColumn:1},{labelCount:999},{startAngle:Infinity},{padAngle:-1},{clockwise:1}]) {
    const malformed=structuredClone(task);Object.assign(malformed.params.jobs[0],patch);assert.throws(()=>validateChartTask(malformed))
  }
  const output=chordGoldenOutput(fixture);assert.ok(estimateChartDecodedMemory(output)>=3*512+3*512)
  const {data}=prepared(chordRecords([['A','B',2],['B','C',1],['C','A',4]])), bytes=await estimateChartDataMemory(data)
  assert.ok(bytes>data.chord.nodes.length*256+data.chord.links.length*384)
  const sceneBytes=estimateChartFinalLayoutMemory(data,chordModel(),{preparedData:data})
  assert.ok(sceneBytes>4*1024*1024+data.categories.length*128)
})
