import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { CHORD_TABLE, chartChordGoldenTransport } from './fixtures/chart-chord-fixtures.mjs'

// P28 drafts only. No server, compiler, runner, browser or WASM ran.
let server, api
beforeAll(async () => {
  server=await createServer({configFile:resolve(process.cwd(),'vite.config.ts'),root:process.cwd(),server:{middlewareMode:true,hmr:false},appType:'custom',logLevel:'error'})
  api=await server.ssrLoadModule('/packages/office-studio/src/index.ts')
})
afterAll(async () => {await server?.close()})
const json=value=>JSON.parse(JSON.stringify(value))
const methods=['setChordMapping','setChordLayout','resetChordLayout','setChordNode','resetChordNode','setChordLabel','resetChordLabel','setChordRibbon','resetChordRibbon','setChordEmphasis','resetChordEmphasis']
function base(studio) {
  const workbook=api.FOffice.newAPI(studio).getActiveWorkbook(),sheet=workbook.getActiveSheet()
  sheet.getRange('A1:C4').setValues(CHORD_TABLE);sheet.getRange('C2:C4').setNumberFormat('0.00')
  return {workbook,sheet}
}
function builder(sheet) {return sheet.newChart('chord').setSource('A1:C4').setChordMapping({sourceIndex:0,targetIndex:1,valueIndex:2})
  .setChordLayout({padAngle:0}).setAbsolutePosition(20,30).setSize(560,400)}
async function setup() {
  const compute=chartChordGoldenTransport(),studio=new api.OfficeStudio({computeClient:compute}),{workbook,sheet}=base(studio)
  return {compute,studio,workbook,sheet,chart:await sheet.insertChart(builder(sheet).build())}
}

test('all eleven Builder methods chain, detach inputs and expose three fixed Chord enums', () => {
  const studio=new api.OfficeStudio(),{sheet}=base(studio)
  try {
    const enums=api.FOffice.newAPI(studio).Enum
    assert.deepEqual(enums.ChartChordLabelPosition,{Inside:'inside',Outside:'outside'})
    assert.deepEqual(enums.ChartChordRibbonColorMode,{Source:'source',Target:'target',Gradient:'gradient'})
    assert.deepEqual(enums.ChartChordFocusMode,{Adjacency:'adjacency',Self:'self',None:'none'})
    const value=builder(sheet),mapping={sourceIndex:0,targetIndex:1,valueIndex:2},layout={innerRadiusRatio:0.9,outerRadiusRatio:0.95,padAngle:0}
    const node={color:'blue.500',opacity:0.4,borderColor:'gray.700',borderWidth:2,borderType:'dashed'},label={visible:true,position:'inside',fontSize:0,rotation:22,family:'Georgia'}
    const ribbon={color:'target',opacity:0.3,width:2,type:'dotted'},emphasis={focus:'self'}
    for(const [name,arg] of [['setChordMapping',mapping],['setChordLayout',layout],['setChordNode',node],['setChordLabel',label],['setChordRibbon',ribbon],['setChordEmphasis',emphasis]]) assert.equal(value[name](arg),value)
    const before=value.build();mapping.targetIndex=0;layout.outerRadiusRatio=0.2;node.color='red';label.family='Arial';ribbon.opacity=1;emphasis.focus='none'
    assert.deepEqual(value.build(),before)
    assert.equal(value.setChordLayout({innerRadiusRatio:0.92}),value);assert.equal(value.build().options.chord.layout.outerRadiusRatio,0.95)
    assert.equal(value.setChordLabel({fontSize:null,rotation:null}),value);assert.equal(value.build().options.chord.label.fontSize,undefined)
    for(const name of methods.filter(name=>name.startsWith('reset'))) assert.equal(value[name](),value)
    const after=value.build();for(const key of ['layout','node','label','ribbon','emphasis']) assert.equal(after.options.chord[key],undefined)
    assert.equal(after.options.mapping.chord.valueIndex,2)
    for(const name of methods) {
      const args=name==='setChordMapping'?[{sourceIndex:0,targetIndex:1}]:name.startsWith('set')?[{}]:[]
      assert.throws(()=>sheet.newChart('line')[name](...args),TypeError)
    }
  }finally{studio.dispose()}
})

test('Builder partial legacy edits preserve untouched authored layout/ribbon values and reject invalid final radii atomically', () => {
  const studio=new api.OfficeStudio(),{sheet}=base(studio)
  try {
    const value=builder(sheet).resetChordLayout().setOptions({chord:{startAngle:-35,clockwise:false,padAngle:7,minAngle:2,innerRadiusRatio:0.5,outerRadiusRatio:0.99,ribbonOpacity:0.27}})
    value.setChordLayout({innerRadiusRatio:0.94});value.setChordRibbon({color:'target'})
    const result=value.build().options.chord
    assert.equal(result.layout.startAngle,35);assert.equal(result.layout.clockwise,false);assert.equal(result.layout.padAngle,7);assert.equal(result.layout.minAngle,2)
    assert.equal(result.layout.outerRadiusRatio,0.99);assert.equal(result.layout.innerRadiusRatio,0.94);assert.equal(result.ribbon.opacity,0.27)
    assert.equal(result.startAngle,undefined);assert.equal(result.ribbonOpacity,undefined)
    const before=value.build();assert.throws(()=>value.setChordLayout({outerRadiusRatio:0.9}),RangeError);assert.deepEqual(value.build(),before)
    value.resetChordLayout();value.resetChordRibbon();assert.equal(value.build().options.chord.layout,undefined);assert.equal(value.build().options.chord.ribbon,undefined)
  }finally{studio.dispose()}
})

test('one live Apply shares chart history, then independent resets and internal reopen retain author-only Chord options', async () => {
  const {compute,studio,workbook,chart}=await setup();let reopened
  try {
    const raw=await chart.getComputedData();assert.deepEqual(raw.chord.nodes.map(node=>[node.name,node.value,node.formattedValue]),[['B',5,'5.00'],['A',5,'5.00']])
    assert.deepEqual(raw.chord.links.map(link=>[link.source,link.target,link.value]),[['B','A',5]])
    const before=json(chart.getInfo()),history=studio.editor.history.captureSavepoint(workbook.getId())
    await chart.update({options:{chord:{node:{opacity:0.4,color:'blue.500'},label:{visible:false,rotation:25},ribbon:{color:'target',opacity:0.3},emphasis:{focus:'self'}}}})
    const applied=json(chart.getInfo()),appliedHistory=studio.editor.history.captureSavepoint(workbook.getId())
    assert.equal(applied.revision,before.revision+1);assert.equal((await chart.getComputedRenderCommands()).some(command=>command.kind==='text'),false)
    workbook.undo();assert.deepEqual(json(chart.getInfo()),before);assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()),history)
    workbook.redo();assert.deepEqual(json(chart.getInfo()),applied);assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()),appliedHistory)
    assert.equal(await chart.setChordLayout({innerRadiusRatio:0.9,outerRadiusRatio:0.95}),chart)
    assert.equal(await chart.setChordLayout({innerRadiusRatio:0.92}),chart);assert.equal(chart.getOptions().chord.layout.outerRadiusRatio,0.95)
    assert.equal(await chart.setChordLabel({visible:true,fontSize:0}),chart);assert.equal((await chart.getComputedRenderCommands()).some(command=>command.kind==='text'),false)
    assert.equal(await chart.resetChordLabel(),chart);assert.ok((await chart.getComputedRenderCommands()).some(command=>command.kind==='text'&&command.text==='B'))
    assert.equal(chart.getOptions().chord.node.color,'blue.500')
    for(const [method,key] of [['resetChordNode','node'],['resetChordRibbon','ribbon'],['resetChordEmphasis','emphasis'],['resetChordLayout','layout']]) {
      assert.equal(await chart[method](),chart);assert.equal(chart.getOptions().chord[key],undefined)
    }
    await chart.setChordNode({color:'blue.500',opacity:0.5});await chart.setChordRibbon({opacity:0.4});await chart.setChordEmphasis({focus:'none'})
    const saved=workbook.save(),authors=JSON.stringify(saved.charts)
    for(const field of ['angleOptions','band','hitIndex','topology','sourceVersion','formattedValue']) assert.equal(authors.includes(`"${field}"`),false)
    assert.equal(authors.includes('chord-node:'),false)
    const detached=await chart.getComputedData();detached.chord.nodes[0].value=99;assert.equal((await chart.getComputedData()).chord.nodes[0].value,5)
    reopened=api.FOffice.newAPI(studio).createWorkbook({...saved,id:'chord-reopened-draft'},{makeCurrent:false})
    const restored=reopened.getActiveSheet().getCharts()[0]
    assert.deepEqual(restored.getOptions(),chart.getOptions());assert.deepEqual(json(await restored.getComputedData()),json(await chart.getComputedData()))
    assert.deepEqual(await restored.getComputedRenderCommands(),await chart.getComputedRenderCommands())
    assert.ok(compute.calls.length>0);assert.ok(compute.calls.some(call=>call.options.documentId!==compute.calls[0].options.documentId))
  }finally{reopened?.dispose();studio.dispose()}
})

test('live partial legacy layout/ribbon edits retain untouched author values and fail final radii atomically', async () => {
  const { studio, workbook, chart } = await setup()
  try {
    await chart.resetChordLayout(); await chart.resetChordRibbon()
    await chart.setOptions({ chord: { startAngle: -35, clockwise: false, padAngle: 7, minAngle: 2,
      innerRadiusRatio: 0.5, outerRadiusRatio: 0.99, ribbonOpacity: 0.27 } })
    const legacy = json(chart.getInfo())
    await chart.setChordLayout({ innerRadiusRatio: 0.94 }); await chart.setChordRibbon({ color: 'target' })
    assert.deepEqual(chart.getOptions().chord.layout, { startAngle: 35, clockwise: false, padAngle: 7, minAngle: 2,
      innerRadiusRatio: 0.94, outerRadiusRatio: 0.99 })
    assert.deepEqual(chart.getOptions().chord.ribbon, { opacity: 0.27, color: 'target' })
    const before = json(chart.getInfo()), history = studio.editor.history.captureSavepoint(workbook.getId())
    await assert.rejects(chart.setChordLayout({ outerRadiusRatio: 0.9 }), RangeError)
    assert.deepEqual(json(chart.getInfo()), before); assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()), history)
    workbook.undo(); workbook.undo(); assert.deepEqual(json(chart.getInfo()), legacy)
    workbook.redo(); workbook.redo(); assert.deepEqual(json(chart.getInfo()), before)
    await chart.setChordLayout({ padAngle: null }); await chart.setChordRibbon({ opacity: null })
    assert.equal(chart.getOptions().chord.layout.padAngle, undefined); assert.equal(chart.getOptions().chord.ribbon.opacity, undefined)
    assert.equal(chart.getOptions().chord.layout.outerRadiusRatio, 0.99)
    assert.equal(chart.getOptions().chord.ribbon.color, 'target')
  } finally { studio.dispose() }
})

test('source changes refresh weighted aggregates; count mapping clears the prior weight and survives undo/redo', async () => {
  const {studio,workbook,sheet,chart}=await setup()
  try {
    const before=json(chart.getInfo());sheet.getRange('C2').setValue(8)
    assert.equal((await chart.getComputedData()).chord.links[0].value,11);assert.deepEqual(json(chart.getInfo()),before)
    workbook.undo();assert.equal((await chart.getComputedData()).chord.links[0].value,5);workbook.redo();assert.equal((await chart.getComputedData()).chord.links[0].value,11)
    assert.equal(await chart.setChordMapping({sourceIndex:0,targetIndex:1}),chart)
    assert.equal(chart.getOptions().mapping.chord.valueIndex,undefined);assert.equal(chart.getOptions().mapping.edge,undefined)
    assert.equal((await chart.getComputedData()).chord.weightMode,'count');assert.equal((await chart.getComputedData()).chord.links[0].value,2)
    workbook.undo();assert.equal((await chart.getComputedData()).chord.links[0].value,11);workbook.redo();assert.equal((await chart.getComputedData()).chord.links[0].value,2)
  }finally{studio.dispose()}
})

test('invalid live mapping/style/type inputs preserve model, revision, history and Worker allocation', async () => {
  const {compute,studio,workbook,chart}=await setup()
  try {
    const before=json(chart.getInfo()),history=studio.editor.history.captureSavepoint(workbook.getId()),calls=compute.calls.length
    for(const operation of [()=>chart.setChordMapping({sourceIndex:0,targetIndex:0}),()=>chart.setChordMapping({sourceIndex:0,targetIndex:1,valueIndex:1}),
      ()=>chart.setChordMapping({sourceIndex:0,targetIndex:3,valueIndex:2}),()=>chart.setChordLayout({innerRadiusRatio:0.9}),()=>chart.setChordNode({opacity:2}),
      ()=>chart.setChordLabel({position:'left'}),()=>chart.setChordLabel({fontSize:-1}),()=>chart.setChordRibbon({color:'red'}),()=>chart.setChordEmphasis({focus:'trajectory'})]) {
      await assert.rejects(async()=>operation());assert.deepEqual(json(chart.getInfo()),before);assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()),history)
    }
    assert.equal(compute.calls.length,calls)
    await chart.update({type:'line',valueFields:[2],options:{mapping:undefined}})
    const inactive=json(chart.getInfo()),inactiveHistory=studio.editor.history.captureSavepoint(workbook.getId())
    for(const name of methods) {
      const args=name==='setChordMapping'?[{sourceIndex:0,targetIndex:1}]:name.startsWith('set')?[{}]:[]
      await assert.rejects(async()=>chart[name](...args),TypeError)
    }
    assert.deepEqual(json(chart.getInfo()),inactive);assert.deepEqual(studio.editor.history.captureSavepoint(workbook.getId()),inactiveHistory)
  }finally{studio.dispose()}
})

test('worksheet edit protection blocks all eleven Chord operations and the direct chart command before Worker allocation', async () => {
  const {studio,workbook,sheet,chart}=await setup();let visitor
  try {
    studio.executeCommandSync('sheet.command.add-worksheet-protection',{unitId:workbook.getId(),sheetId:sheet.getSheetId(),rule:{id:'chord-protection',permissionId:'chord-permission',unitType:'worksheet',ownerUserId:'owner',editState:'onlyMe',viewState:'othersCanView',allowedUserIds:[]}})
    const compute=chartChordGoldenTransport();visitor=new api.OfficeStudio({workbook:json(workbook.save()),currentUserId:'visitor',computeClient:compute})
    const targetWorkbook=api.FOffice.newAPI(visitor).getActiveWorkbook(),target=targetWorkbook.getActiveSheet().getChartById(chart.getChartId())
    const before=json(target.getInfo()),history=visitor.editor.history.captureSavepoint(targetWorkbook.getId())
    for(const name of methods) {
      const args=name==='setChordMapping'?[{sourceIndex:0,targetIndex:1}]:name.startsWith('set')?[{}]:[]
      await assert.rejects(async()=>target[name](...args),/编辑权限/)
    }
    await assert.rejects(async()=>visitor.executeCommand('sheet.command.update-chart',{unitId:targetWorkbook.getId(),chartId:chart.getChartId(),patch:{title:'Blocked Chord edit'}}),/编辑权限/)
    assert.deepEqual(json(target.getInfo()),before);assert.deepEqual(visitor.editor.history.captureSavepoint(targetWorkbook.getId()),history);assert.equal(compute.calls.length,0)
  }finally{visitor?.dispose();studio.dispose()}
})
