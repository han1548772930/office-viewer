import assert from 'node:assert/strict'
import { test } from 'vitest'
import { OfficeComputeClient } from '../src/compute/client.ts'
import { cloneTask, estimateTaskMemory, taskViews, transferTaskBuffers, validateTask } from '../src/compute/tasks.ts'
import { readResult, progressUnits } from '../src/compute/guards.ts'
import { readCheckboxCellsOutput } from '../src/compute/checkbox-cell-results.ts'
import { createWasmTaskControl } from '../src/platform/wasm/task-protocol.ts'
// Protocol/transport drafts only: the fake Worker returns an independently fixed output, never an alternative checkbox engine.
const task = () => ({ kind: 'checkboxCells', ranges: new Uint32Array([0,1,0,2]), merges: new Uint32Array([]),
  positions: new Uint32Array([0,0,1,2]), actions: new Uint8Array([2,0]),
  params: { rowCount: 2, columnCount: 3, rangeCount: 1, mergeCount: 0, existingCount: 2, coveredCells: 6, maxOutputRuns: 6 } })
const fixedRuns = () => new Uint32Array([0,0,0,0,2,0,0,1,2,1,1,1,0,1,1])
function response(input, control) {
  const runs = fixedRuns(), memory = estimateTaskMemory(input)
  return { type: 'result', context: control, engineVersion: 'checkbox-contract', output: { kind: 'checkboxCells', runs, changedCells: 5 },
    metrics: { elapsedMs: 1, ...memory, boundary: { inputCopies: 3, inputBytes: 34, outputCopies: 1, outputBytes: runs.byteLength,
      controlDtoConversions: 2, jsonConversions: 0, rustBufferClones: 0 }, transport: { inputTransfers: 4, inputBytes: 34,
      outputTransfers: 1, outputBytes: runs.byteLength, structuredResultMessages: 1 } } }
}
class WorkerFixture {
  messages = []; terminated = false; onmessage = null; onerror = null; onmessageerror = null
  postMessage(message, transfer = []) { this.messages.push(structuredClone(message, { transfer })) }
  send(message) { this.onmessage?.({ data: message }) }
  terminate() { this.terminated = true }
  ready() { this.send({ type: 'ready', protocolVersion: 2, engineVersion: 'checkbox-contract' }) }
}
test('bulk task uses four independent transferable typed inputs and copy ownership retains detached source values', () => {
  const input = task(), copy = cloneTask(input); validateTask(input); assert.equal(taskViews(input).length,4)
  assert.deepEqual(progressUnits(input),{ total:6,unit:'cells' }); assert.deepEqual(estimateTaskMemory(input),{ inputBytes:34,estimatedPeakBytes:8*1024*1024+34*2+2*96+6*160+32 })
  const moved = structuredClone(input,{ transfer:transferTaskBuffers(input) }); assert.deepEqual([...moved.actions],[2,0]); assert.deepEqual([...copy.positions],[0,0,1,2])
  assert.throws(()=>validateTask(input),error=>error.code==='COMPUTE_INVALID_INPUT'); assert.notEqual(copy.params,moved.params)
})
test.each(['params','ranges','actions','budget','shared','slice'])('bulk admission rejects malformed %s without transferring caller data', kind => {
  const input = task()
  if(kind==='params') input.params.unrecognized=1
  if(kind==='ranges') input.ranges=new Uint32Array([0,0])
  if(kind==='actions') input.actions=new Uint8Array([1])
  if(kind==='shared') input.actions=new Uint8Array(new SharedArrayBuffer(2))
  if(kind==='slice') input.positions=new Uint32Array(new ArrayBuffer(32),8,4)
  if(kind==='slice') assert.throws(()=>transferTaskBuffers(input),error=>error.code==='COMPUTE_INVALID_INPUT')
  else assert.throws(()=>validateTask(input,kind==='budget'?1:256*1024*1024))
  assert.equal(input.ranges.buffer.byteLength>0,true)
})
test.each(['action','extent','count','length','slice','detached'])('output rejects %s and preserves complete-run ownership', kind => {
  const p=task().params, output={kind:'checkboxCells',runs:fixedRuns(),changedCells:5}; assert.ok(readCheckboxCellsOutput(output,p))
  if(kind==='action')output.runs[4]=0
  if(kind==='extent')output.runs[1]=5
  if(kind==='count')output.changedCells=4
  if(kind==='length')output.runs=new Uint32Array([0])
  if(kind==='slice')output.runs=new Uint32Array(new ArrayBuffer(80),4,15)
  if(kind==='detached')structuredClone(output,{transfer:[output.runs.buffer]})
  assert.equal(readCheckboxCellsOutput(output,p),null)
})
test.each(['inputCopies','inputBytes','outputCopies','outputBytes','controlDtoConversions','jsonConversions','rustBufferClones'])('result guard rejects false boundary metric %s', field => {
  const input=task(), control=createWasmTaskControl('checkboxCells','request','book',4), packet=response(input,control)
  assert.ok(readResult(packet,{requestId:'request',documentId:'book',modelRevision:4},control,input,estimateTaskMemory(input),2))
  packet.metrics.boundary[field]=field==='controlDtoConversions'?0:packet.metrics.boundary[field]+1
  assert.equal(readResult(packet,{requestId:'request',documentId:'book',modelRevision:4},control,input,estimateTaskMemory(input),2),null)
})
test('actual client dispatches the nonnumeric task, detaches all four inputs and accepts a matching one-use result', async () => {
  const worker=new WorkerFixture(), client=new OfficeComputeClient({workerFactory:()=>worker}), input=task()
  try {
    const job=client.submit(input,{documentId:'book',modelRevision:4,ownership:'transfer'}); worker.ready()
    assert.equal(input.ranges.buffer.byteLength,0); assert.equal(input.positions.buffer.byteLength,0)
    const request=worker.messages.find(message=>message.type==='run'); assert.equal(request.context.kind,'checkboxCells')
    worker.send(response(request.task,request.context)); const result=await job.promise
    assert.equal(result.output.changedCells,5); assert.equal(result.metrics.ownershipTransfers,4)
    worker.send(response(request.task,request.context)); assert.equal(result.output.changedCells,5)
  } finally {client.dispose()}
})
test.each(['abort','revision','close'])('actual queued client rejects %s and never accepts a stale bulk result', async reason => {
  const worker=new WorkerFixture(), client=new OfficeComputeClient({workerFactory:()=>worker}), controller=new AbortController()
  try {
    const job=client.submit(task(),{documentId:'book',modelRevision:4,ownership:'transfer',signal:controller.signal})
    if(reason==='abort')controller.abort(); if(reason==='revision')client.setDocumentRevision('book',5); if(reason==='close')client.forgetDocument('book')
    await assert.rejects(job.promise,error=>error.code===(reason==='abort'?'COMPUTE_CANCELLED':'COMPUTE_STALE_RESULT'))
  } finally {client.dispose()}
})
