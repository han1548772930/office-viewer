import assert from 'node:assert/strict'
import { beforeAll, afterAll, test } from 'vitest'
import { createServer } from 'vite'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { cloneModel } from '../src/domain/serialization.ts'
import { ComputeError } from '../src/compute/errors.ts'
import { createSheetDataValidationRule } from '../src/features/sheet-validation/validation-engine.ts'
import { planSheetCheckboxRuleCellsAsync } from '../src/features/sheet-validation/validation-checkbox-rule-cells-async.ts'
import { SHEET_DATA_VALIDATION_ASYNC_COMMANDS } from '../src/features/sheet-validation/validation-rule-command.ts'
// Fixed-result transport fixtures exercise admission, batching, cancellation and actual one-history publication. Rust correctness has independent native/binding drafts.
let vite,OfficeStudio,FOffice
beforeAll(async()=>{vite=await createServer({configFile:'./vite.config.ts',server:{middlewareMode:true,hmr:false},appType:'custom',logLevel:'error'});
  ({OfficeStudio,FOffice}=await vite.ssrLoadModule('/packages/office-studio/src/index.ts'))},{timeout:60_000})
afterAll(async()=>{await vite?.close()})
const area={startRow:0,endRow:79,startCol:0,endCol:79}
function fixture(configure=()=>{}) {
  const model=createBlankWorkbook(),sheet=model.sheets[0]; model.id='bulk-checkbox';model.activeSheetId=sheet.id='data';model.calculationMode='manual'
  sheet.rowCount=80;sheet.columnCount=80;sheet.cells={};sheet.dataValidations=[];configure(model,sheet)
  const next=createSheetDataValidationRule({id:'checkbox',type:'checkbox',formula1:'Yes',formula2:'No'},[area])
  return {model,sheet,next}
}
function clientFor(output={kind:'checkboxCells',runs:new Uint32Array([0,79,0,79,1]),changedCells:6400},beforeResolve=()=>{}) {
  const submitted=[];return {submitted,forgetDocument(){},setDocumentRevision(){},submit(task,options){
    const captured=structuredClone(task,{transfer:[task.ranges.buffer,task.merges.buffer,task.positions.buffer,task.actions.buffer]});submitted.push({task:captured,options})
    const promise=Promise.resolve().then(()=>{beforeResolve(captured,options);return {type:'result',requestId:'fixed-bulk',documentId:options.documentId,modelRevision:options.modelRevision,
      output:structuredClone(output),metrics:{elapsedMs:1,inputBytes:16,estimatedPeakBytes:1,boundary:{},transport:{}}}})
    return {requestId:'fixed-bulk',promise,cancel(){}}
  }}
}
const plan=(f,compute,options={})=>planSheetCheckboxRuleCellsAsync(f.model,f.sheet,f.next,f.sheet.dataValidations[0],'local-user','zh-CN',
  {compute,modelRevision:4,assertCurrent(){},...options})
test('a large blank rectangle is classified once, transferred without scalars and materialized in bounded host batches',async()=>{
  const f=fixture(),compute=clientFor(),before=cloneModel(f.model),progress=[]
  const result=await plan(f,compute,{onProgress:value=>progress.push(value)});assert.equal(result.cells.length,6400);assert.ok(result.cells.every(cell=>cell.value==='No'))
  assert.deepEqual(f.model,before);assert.equal(compute.submitted.length,1);assert.equal(compute.submitted[0].task.params.existingCount,0)
  assert.equal(compute.submitted[0].options.priority,'interactive');assert.ok(progress.some(value=>value.phase==='materializing'&&value.processed<6400))
  assert.equal(result.metrics.preparedBulkBytes,16);assert.ok(result.metrics.materializationEstimatedBytes>0)
})
test('typed sparse actions share the checked migration semantics and preserve formula contents',async()=>{
  const f=fixture((_model,sheet)=>{sheet.dataValidations=[createSheetDataValidationRule({id:'checkbox',type:'checkbox'},[area])];
    sheet.cells['0:0']=createCell(0,0,true);sheet.cells['1:1']=createCell(1,1,0);sheet.cells['1:1'].formula='=1-1'})
  const output={kind:'checkboxCells',runs:new Uint32Array([0,0,0,0,2,0,0,1,79,1,1,1,0,0,1,1,1,2,79,1,2,79,0,79,1]),changedCells:6399}
  const compute=clientFor(output),before=cloneModel(f.model),result=await plan(f,compute)
  assert.deepEqual([...compute.submitted[0].task.actions],[2,0]);assert.equal(result.cells.find(cell=>cell.id==='0:0').value,'Yes')
  assert.ok(!result.cells.some(cell=>cell.id==='1:1'));assert.deepEqual(f.model,before)
})
test.each(['before','classifying','materializing'])('cancellation at %s never modifies the live model',async phase=>{
  const f=fixture(),controller=new AbortController(),compute=clientFor(),before=cloneModel(f.model)
  if(phase==='before')controller.abort()
  await assert.rejects(()=>plan(f,compute,{signal:controller.signal,onProgress:value=>{if(value.phase===phase)controller.abort()}}),error=>error.code==='COMPUTE_CANCELLED')
  assert.deepEqual(f.model,before);if(phase!=='materializing')assert.equal(compute.submitted.length,0)
})
test.each(['cells','large','invalid'])('admission rejects %s memory demand before partial output publication',async kind=>{
  const f=fixture(),compute=clientFor(),before=cloneModel(f.model),options={}
  if(kind==='cells')options.maxMaterializationBytes=1
  if(kind==='invalid')options.maxMaterializationBytes=NaN
  if(kind==='large'){f.sheet.rowCount=1_048_576;f.next.ranges=[{startRow:0,endRow:1_048_575,startCol:0,endCol:79}]}
  const unchanged=cloneModel(f.model)
  await assert.rejects(()=>plan(f,compute,options),error=>error.code==='COMPUTE_MEMORY_LIMIT');assert.deepEqual(f.model,unchanged)
  assert.deepEqual(before.sheets[0].cells,f.sheet.cells);if(kind!=='cells')assert.equal(compute.submitted.length,0)
})
test('complete View is checked before classification reads hidden target getters',async()=>{
  const f=fixture((_model,sheet)=>{const cell=createCell(0,0,null);Object.defineProperty(cell,'value',{get(){throw new Error('hidden scalar read')}});sheet.cells['0:0']=cell;
    sheet.rangeProtectionRules=[{id:'hidden',ranges:[area],ownerUserId:'peer',allowedUserIds:[],viewState:'noOneElseCanView',editState:'onlyMe'}]})
  const compute=clientFor();await assert.rejects(()=>plan(f,compute),/查看/);assert.equal(compute.submitted.length,0)
})
test('a source revision or permission change after a Worker result rejects all detached cells',async()=>{
  const f=fixture(),before=cloneModel(f.model);let revoked=false
  const compute=clientFor(undefined,()=>{revoked=true})
  await assert.rejects(()=>plan(f,compute,{assertCurrent(){if(revoked)throw new ComputeError('COMPUTE_STALE_RESULT','revoked')}}),error=>error.code==='COMPUTE_STALE_RESULT')
  assert.deepEqual(f.model,before)
})
test('the actual async command commits rules and bulk cells once with isolated undo redo and internal reopen',async()=>{
  const f=fixture(),compute=clientFor(),studio=new OfficeStudio({workbook:f.model,currentUserId:'local-user',computeClient:compute}),events=[]
  const book=FOffice.newAPI(studio).getActiveWorkbook(),subscription=studio.events.addEvent('SheetValueChanged',value=>events.push(cloneModel(value)))
  try{
    const before=cloneModel(studio.getWorkbookModel(f.model.id)),generation=studio.editor.history.captureSavepoint(f.model.id).generation
    const changed=await studio.executeCommand(SHEET_DATA_VALIDATION_ASYNC_COMMANDS.set,{unitId:f.model.id,sheetId:'data',rule:f.next,
      expectedRevision:studio.getWorkbookRevision(f.model.id),expectedHistoryGeneration:generation,expectedUserId:'local-user'})
    assert.equal(changed,true);const after=cloneModel(studio.getWorkbookModel(f.model.id));assert.equal(Object.keys(after.sheets[0].cells).length,6400)
    assert.equal(studio.editor.history.units[f.model.id].undo.length,1);assert.equal(events.length,1);book.undo();assert.deepEqual(studio.getWorkbookModel(f.model.id),before)
    book.redo();assert.deepEqual(studio.getWorkbookModel(f.model.id),after)
    const reopened=new OfficeStudio({workbook:book.save()});try{assert.deepEqual(reopened.getWorkbookModel(f.model.id).sheets[0],after.sheets[0])}finally{reopened.dispose()}
  }finally{subscription.dispose();studio.dispose()}
})
test('actual async materialization cancellation keeps rules cells value observers and history unchanged',async()=>{
  const f=fixture(),compute=clientFor(),studio=new OfficeStudio({workbook:f.model,currentUserId:'local-user',computeClient:compute}),controller=new AbortController(),events=[]
  const subscription=studio.events.addEvent('SheetValueChanged',value=>events.push(value)),before=cloneModel(studio.getWorkbookModel(f.model.id))
  try{
    await assert.rejects(()=>studio.executeCommand(SHEET_DATA_VALIDATION_ASYNC_COMMANDS.replace,{unitId:f.model.id,sheetId:'data',rule:f.next,signal:controller.signal,
      onProgress:value=>{if(value.phase==='materializing')controller.abort()}}),error=>error.code==='COMPUTE_CANCELLED')
    assert.deepEqual(studio.getWorkbookModel(f.model.id),before);assert.equal(studio.editor.history.units[f.model.id]?.undo.length??0,0);assert.deepEqual(events,[])
  }finally{subscription.dispose();studio.dispose()}
})
