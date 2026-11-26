import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createCell } from '../src/domain/model.ts'
import { getRecordValueKey } from '../src/domain/record-value.ts'
import { setSlicerRecordSelection } from '../src/features/slicer/slicer-engine.ts'
import { SlicerMemberQueryService, normalizeSlicerMemberWindowRequest } from '../src/features/slicer/slicer-member-query.ts'
import { slicerRecordSelectionWorkbook, slicerTableRecordSelectionWorkbook } from './fixtures/pivot-filter-fixture.mjs'
import { createSlicerMemberTestClient } from './fixtures/slicer-member-compute-fixture.mjs'

// Bounded members, source identity and owner lifecycle drafts. No tests have been executed.
const setup = kind => {
  const workbook = kind === 'table' ? slicerTableRecordSelectionWorkbook() : slicerRecordSelectionWorkbook(), slicer = workbook.slicers[0]
  return { workbook, slicer, context: { unitId: workbook.id, baseRevision: 0, assertCurrent() {} } }
}
async function withService(work, options) {
  const service = new SlicerMemberQueryService({ computeClient: createSlicerMemberTestClient(), ...options })
  try { await work(service) } finally { service.dispose() }
}

test('member windows default to 0/100, cap count, reject malformed requests and preserve caller objects', () => {
  assert.deepEqual(normalizeSlicerMemberWindowRequest(), { start: 0, count: 100 })
  const request = { start: 10, count: 500 }, before = structuredClone(request)
  assert.deepEqual(normalizeSlicerMemberWindowRequest(request), { start: 10, count: 200 }); assert.deepEqual(request, before)
  for (const invalid of [null, [], { start: -1 }, { start: .5 }, { start: Number.MAX_SAFE_INTEGER + 1 },
    { count: 0 }, { count: Infinity }, { count: null }, { signal: {} }, { unknown: true }]) {
    assert.throws(() => normalizeSlicerMemberWindowRequest(invalid), error => error.code === 'SLICER_INVALID_MEMBER_WINDOW')
  }
})

for (const kind of ['pivot', 'table']) test(`${kind} async raw windows retain all typed identities, defensive results and filter state`, async () => {
  const f = setup(kind), source = f.workbook.sheets.find(sheet => sheet.id === 'data')
  const rows = [1, 1, '1', '#N/A', '#N/A', null, '']
  rows.forEach((value, index) => {
    const cell = createCell(index + 3, 6, value)
    if (index === 1) cell.style.numberFormat = 'yyyy-mm-dd'
    if (index === 4) { cell.dataType = 'e'; cell.metadata = { error: '#N/A' } }
    source.cells[`${index + 3}:6`] = cell
  })
  setSlicerRecordSelection(f.workbook, f.slicer, { mode: 'include', values: [{ type: 'date', value: 1 }] })
  const before = structuredClone(f.workbook)
  await withService(async service => {
    const result = await service.getItems(f.workbook, f.slicer, {}, f.context)
    assert.equal(result.total, 7); assert.equal(result.items.length, 7)
    assert.deepEqual(new Set(result.items.map(item => item.key)), new Set([{ type: 'number', value: 1 }, { type: 'date', value: 1 },
      { type: 'string', value: '1' }, { type: 'string', value: '#N/A' }, { type: 'error', value: '#N/A' },
      { type: 'blank' }, { type: 'string', value: '' }].map(getRecordValueKey)))
    assert.equal(result.items.every(item => item.hasData), true)
    assert.deepEqual(f.workbook, before)
    result.items[1].value.value = 'changed'
    assert.deepEqual(await service.getItems(f.workbook, f.slicer, { start: 7, count: 4 }, f.context), { total: 7, items: [] })
    assert.deepEqual(f.workbook, before)
  })
})

test('real distinct totals, maximum windows and tail pages are independent of startItem and selection', async () => {
  const f = setup('table'), table = f.workbook.tables.find(table => table.id === 'typed-table'), source = f.workbook.sheets.find(sheet => sheet.id === 'data')
  table.range.endRow = 302
  for (let index = 0; index < 300; index++) source.cells[`${index + 3}:6`] = createCell(index + 3, 6, index)
  f.slicer.startItem = 30
  await withService(async service => {
    const first = await service.getItems(f.workbook, f.slicer, {}, f.context), capped = await service.getItems(f.workbook, f.slicer, { count: 999 }, f.context)
    assert.equal(first.total, 300); assert.equal(first.items.length, 100); assert.equal(capped.items.length, 200)
    const tail = await service.getItems(f.workbook, f.slicer, { start: 290, count: 20 }, f.context)
    assert.equal(tail.total, 300); assert.deepEqual(tail.items.map(item => item.value.value), Array.from({ length: 10 }, (_, i) => i + 290))
    assert.equal(f.slicer.startItem, 30)
  })
})

for (const action of ['signal', 'unit', 'service', 'source']) test(`${action} invalidation rejects a yielded query and releases its owner`, async () => {
  const f = setup('table'), service = new SlicerMemberQueryService({ computeClient: createSlicerMemberTestClient() }), controller = new AbortController()
  let invalidated = false
  const context = { unitId: f.workbook.id, baseRevision: 0, assertCurrent() {
    if (invalidated) { const error = new Error('source changed'); error.code = 'SLICER_SOURCE_INVALIDATED'; throw error }
  } }
  try {
    const pending = service.getItems(f.workbook, f.slicer, { signal: controller.signal }, context)
    if (action === 'signal') controller.abort()
    if (action === 'unit') service.forgetDocument(f.workbook.id)
    if (action === 'service') service.dispose()
    if (action === 'source') invalidated = true
    const code = { signal: 'SLICER_RPC_ABORTED', unit: 'SLICER_RPC_UNIT_DISPOSED', service: 'SLICER_RPC_DISPOSED', source: 'SLICER_SOURCE_INVALIDATED' }[action]
    await assert.rejects(pending, error => error.code === code)
    if (action !== 'service') { invalidated = false; assert.equal((await service.getItems(f.workbook, f.slicer, {}, f.context)).total, 7) }
  } finally { service.dispose() }
})

test('date metadata exposes source name, grouped unit and serial bounds without changing authored state', async () => {
  const f = setup('pivot'), dateView = { ...f.slicer, id: 'date-query', sourceFieldId: 'date-source', cacheId: undefined, type: 'timeline' }
  f.workbook.slicers.push(dateView)
  for (const report of f.workbook.pivotTables) report.fields.find(field => field.sourceFieldId === 'date-source').dateGroupType = 'month'
  const before = structuredClone(f.workbook)
  await withService(async service => {
    const metadata = await service.getSourceMetadata(f.workbook, dateView, f.context)
    assert.equal(metadata.fieldId, 'date-source'); assert.equal(metadata.name, 'Date'); assert.equal(metadata.ready, true)
    assert.deepEqual(metadata.dateGroup, { type: 3, dateType: 'month' }); assert.equal(metadata.timeline.interactive, true)
    assert.ok(metadata.timeline.minDate < metadata.timeline.maxDate); assert.deepEqual(f.workbook, before)
  })
})

test('metadata bounds include cached typed Date and exclude untyped numbers, ISO text and cached errors', async () => {
  const f = setup('pivot'), source = f.workbook.sheets.find(sheet => sheet.id === 'data')
  const dateView = { ...f.slicer, id: 'date-only-metadata', sourceFieldId: 'date-source', cacheId: undefined, type: 'timeline' }
  f.workbook.slicers.push(dateView)
  const values = [100,1000000,'1900-01-01','-99','#N/A',-2.25,null]
  values.forEach((value, index) => {
    const cell = createCell(index+3, 4, value)
    if (index === 0) { cell.formula = '=DATE(1900,4,9)'; cell.cachedDataType = 'date' }
    if (index === 4) { cell.formula = '=NA()'; cell.cachedDataType = 'error'; cell.metadata = { error: '#N/A' } }
    if (index === 5) cell.dataType = 'date'
    source.cells[`${index+3}:4`] = cell
  })
  await withService(async service => {
    const metadata = await service.getSourceMetadata(f.workbook, dateView, f.context)
    assert.deepEqual(metadata.timeline, { minDate: -2.25, maxDate: 100, interactive: true })
  })
})

test('a source with no typed Date has no guessed bounds and a noninteractive Timeline', async () => {
  const f = setup('pivot'), source = f.workbook.sheets.find(sheet => sheet.id === 'data')
  const dateView = { ...f.slicer, id: 'no-date-metadata', sourceFieldId: 'date-source', cacheId: undefined, type: 'timeline' }
  f.workbook.slicers.push(dateView)
  for (let row = 3; row <= 9; row++) source.cells[`${row}:4`] = createCell(row,4,row%2 ? 45000 : '2026-01-01')
  await withService(async service => {
    assert.deepEqual((await service.getSourceMetadata(f.workbook,dateView,f.context)).timeline, { interactive: false })
  })
})

test('source and permission invalidation while awaiting compute aborts the owner before publishing late rows', async () => {
  const f = setup('table'), client = createSlicerMemberTestClient({ delayMs: 500 }), service = new SlicerMemberQueryService({ computeClient: client })
  let invalidated = false
  const context = { ...f.context, assertCurrent() { if (invalidated) { const error = new Error('permission/source invalidated'); error.code = 'SLICER_SOURCE_INVALIDATED'; throw error } } }
  try {
    const pending = service.getItems(f.workbook,f.slicer,{},context)
    const rejected = assert.rejects(pending, error => error.code === 'SLICER_SOURCE_INVALIDATED')
    while (!client.submitted.length) await new Promise(resolve => setTimeout(resolve,0))
    invalidated = true
    await rejected
    assert.equal(client.submitted[0].options.signal.aborted, true)
    assert.equal(client.isCurrent({ documentId: f.workbook.id, modelRevision: 0 }), true)
  } finally { service.dispose(); client.dispose() }
})

test('query deadlines include source preparation and waiting, and pending counts remain bounded', async () => {
  const f = setup('table'), client = createSlicerMemberTestClient({ delayMs: 1000 })
  const service = new SlicerMemberQueryService({ computeClient: client, timeoutMs: 20, maxPendingQueries: 1 })
  try {
    const pending = service.getItems(f.workbook,f.slicer,{},f.context)
    const timeout = assert.rejects(pending, error => error.code === 'COMPUTE_TIMEOUT')
    await assert.rejects(() => service.getItems(f.workbook,f.slicer,{},f.context), error => error.code === 'COMPUTE_MEMORY_LIMIT')
    await timeout
  } finally { service.dispose(); client.dispose() }
})

test('unsafe or missing source revisions reject before any compute job is submitted', async () => {
  const f = setup('table'), client = createSlicerMemberTestClient(), service = new SlicerMemberQueryService({ computeClient: client })
  try {
    for (const baseRevision of [undefined,-1,0.5,Number.MAX_SAFE_INTEGER+1]) {
      await assert.rejects(() => service.getItems(f.workbook,f.slicer,{}, { ...f.context,baseRevision }), error => error.code === 'COMPUTE_INVALID_INPUT')
    }
    assert.equal(client.submitted.length, 0)
  } finally { service.dispose(); client.dispose() }
})

test('local display windows preserve first typed-Date raw captions with bounded source lookup and leave formal items unchanged', async () => {
  const f = setup('table'), source = f.workbook.sheets.find(sheet => sheet.id === 'data')
  const values = ['1900-04-09',100,100,'100',null,'','#N/A']
  values.forEach((value, index) => {
    const cell = createCell(index + 3, 6, value)
    if (index === 0 || index === 2) cell.dataType = 'date'
    if (index === 6) { cell.dataType = 'e'; cell.metadata = { error: '#N/A' } }
    source.cells[`${index + 3}:6`] = cell
  })
  const reads = [], retainedCells = source.cells
  source.cells = new Proxy(retainedCells, { get(target, key, receiver) {
    if (typeof key === 'string' && /^\d+:6$/.test(key)) reads.push(key)
    return Reflect.get(target, key, receiver)
  } })
  const client = createSlicerMemberTestClient({ onSubmit() { reads.length = 0 } })
  const service = new SlicerMemberQueryService({ computeClient: client })
  try {
    const display = await service.getDisplayMemberWindow(f.workbook, f.slicer, { start: 1, count: 2 }, f.context)
    assert.equal(display.total, 4); assert.deepEqual(display.items.map(item => item.label), ['100','1900-04-09'])
    assert.deepEqual(display.items.map(item => item.value), [{ type: 'number', value: 100 },{ type: 'date', value: 100 }])
    assert.deepEqual(reads, ['4:6','3:6'])
    const raw = await service.getItems(f.workbook, f.slicer, { start: 1, count: 2 }, f.context)
    assert.deepEqual(raw, { total: 6, items: display.items.map(({ key, value, hasData }) => ({ key, value, hasData })) })
    assert.deepEqual(Object.keys(raw.items[0]).sort(), ['hasData','key','value']); assert.deepEqual(reads, ['4:6','3:6'])
  } finally { service.dispose(); client.dispose(); source.cells = retainedCells }
})

test('local Canvas windows fold caption identities globally and return full/partial state without publishing extra formal fields', async () => {
  const f = setup('table'), client = createSlicerMemberTestClient(), service = new SlicerMemberQueryService({ computeClient: client })
  setSlicerRecordSelection(f.workbook,f.slicer,{ mode:'include',values:[{ type:'number',value:1 }] })
  const before = structuredClone(f.workbook)
  try {
    const first = await service.getDisplayMemberWindow(f.workbook,f.slicer,{ start:0,count:2 },f.context)
    const second = await service.getDisplayMemberWindow(f.workbook,f.slicer,{ start:2,count:2 },f.context)
    assert.equal(first.total,4); assert.equal(second.total,4)
    const captions = [...first.items,...second.items]; assert.deepEqual(captions.map(item => item.label),['','1','beta','true'])
    const one = captions.find(item => item.label === '1'); assert.equal(one.selected,false); assert.equal(one.partial,true)
    assert.ok(captions.filter(item => item !== one).every(item => item.selected === false && item.partial === false))
    assert.equal(client.submitted[0].task.params.mode,'canvas'); assert.equal(client.submitted[0].task.textIds.length,14)
    const formal = await service.getItems(f.workbook,f.slicer,{},f.context)
    assert.equal(formal.total,7); assert.ok(formal.items.every(item => !Object.hasOwn(item,'label') && !Object.hasOwn(item,'selected')))
    assert.deepEqual(f.workbook,before)
  } finally { service.dispose();client.dispose() }
})

test('metadata cancellation signal rejects preparation and leaves the injected shared client reusable', async () => {
  const f = setup('table'), client = createSlicerMemberTestClient(), service = new SlicerMemberQueryService({ computeClient:client })
  const controller = new AbortController()
  try {
    const pending = service.getSourceMetadata(f.workbook,f.slicer,f.context,controller.signal); controller.abort()
    await assert.rejects(pending,error => error.code === 'SLICER_RPC_ABORTED')
    assert.equal(client.submitted.length,0)
    assert.equal((await service.getItems(f.workbook,f.slicer,{},f.context)).total,7)
  } finally { service.dispose();client.dispose() }
})
