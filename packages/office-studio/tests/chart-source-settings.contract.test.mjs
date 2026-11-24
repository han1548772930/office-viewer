import assert from 'node:assert/strict'
import { afterEach, test } from 'vitest'
import { createBlankWorkbook, createCell } from '../src/domain/model.ts'
import { createChartDefinition } from '../src/features/charts/chart-engine.ts'
import { readChartRecordsAsync } from '../src/features/charts/chart-source-compute.ts'
import { projectChartRecords, projectChartRecordsAsync } from '../src/features/charts/chart-record-projection.ts'
import { ChartStatisticsService } from '../src/features/charts/services/chart-statistics-service.ts'

// Deferred implementation contracts. No Worker/statistical engine is replaced.
const services = new Set()
afterEach(() => { for (const service of services) service.dispose(); services.clear() })
function fixture(samples = 3, type = 'boxplot', options = {}) {
  const workbook = createBlankWorkbook(), sheet = workbook.sheets[0]
  workbook.id = 'source-settings'; sheet.rowCount = samples + 1; sheet.columnCount = 2
  sheet.cells = { '0:0': createCell(0, 0, 'Label'), '0:1': createCell(0, 1, 'Value') }
  for (let row = 1; row <= samples; row++) {
    sheet.cells[row + ':0'] = createCell(row, 0, 'Item ' + row)
    sheet.cells[row + ':1'] = createCell(row, 1, row)
  }
  const config = { type, source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: samples, endCol: 1 } },
    position: { x: 0, y: 0 }, size: { width: 420, height: 280 }, categoryField: 0, valueFields: [1], options }
  return { workbook, sheet, config, chart: createChartDefinition(workbook, sheet.id, config, 'unpublished-draft') }
}
function serviceFor(workbook, options) {
  const revisions = new Map(), calls = []
  const compute = { calls, setDocumentRevision(id, revision) { revisions.set(id, revision) },
    isCurrent(context) { return revisions.get(context.documentId) === context.modelRevision },
    submit() { calls.push(true); throw new Error('Source reads must not submit statistics') } }
  const service = new ChartStatisticsService(compute, options)
  service.setDocumentRevision(workbook.id, 4); services.add(service)
  return { service, compute, context: { documentId: workbook.id, modelRevision: 4 } }
}

test('unpublished large-source settings return prepared data and numeric identities without history, cache or statistics', async () => {
  const { workbook, config, chart } = fixture(32769), before = structuredClone(config)
  const { service, compute, context } = serviceFor(workbook), notices = []
  service.subscribe(event => notices.push(event))
  const data = await service.requestSourceData(workbook, chart, context).promise
  assert.equal(data.records.length, 32769); assert.equal(data.series[0].values.length, 32769)
  assert.equal(data.series[0].values.at(-1), 32769)
  assert.deepEqual(data.sourceNumericFields, [1])
  assert.deepEqual(data.fields, ['Label', 'Value'])
  assert.deepEqual(notices, []); assert.deepEqual(compute.calls, [])
  assert.equal(service.getCached(chart, context), undefined)
  assert.equal(workbook.charts?.length ?? 0, 0); assert.deepEqual(config, before)
  data.records[0][1] = 999; data.series[0].values[0] = 999
  assert.equal(workbook.sheets[0].cells['1:1'].value, 1)
})

test('source data reads share queue admission, and cancelling a queued draft releases its slot', async () => {
  const { workbook, chart } = fixture(32769), { service, context } = serviceFor(workbook, { maxPendingCharts: 2 })
  const active = service.requestSourceData(workbook, chart, context)
  const queued = service.requestSourceData(workbook, { ...chart, id: 'queued' }, context)
  const rejection = assert.rejects(queued.promise, error => error.code === 'COMPUTE_CANCELLED')
  assert.throws(() => service.requestSourceData(workbook, { ...chart, id: 'full' }, context), error => error.code === 'COMPUTE_MEMORY_LIMIT')
  queued.cancel()
  const replacement = service.requestSourceData(workbook, { ...chart, id: 'replacement' }, context)
  await rejection; await active.promise
  assert.equal((await replacement.promise).series[0].values.at(-1), 32769)
})

test('abort and source revision messages stop large reads before a result reaches settings', async () => {
  for (const kind of ['abort', 'revision', 'close']) {
    const { workbook, chart } = fixture(32769), { service, context } = serviceFor(workbook)
    const controller = new AbortController()
    const request = service.requestSourceData(workbook, chart, { ...context, signal: controller.signal })
    const rejection = assert.rejects(request.promise, error => error.code === (kind === 'abort' ? 'COMPUTE_CANCELLED' : 'COMPUTE_STALE_RESULT'))
    setTimeout(() => {
      if (kind === 'abort') controller.abort()
      else if (kind === 'close') service.forgetDocument(workbook.id)
      else service.setDocumentRevision(workbook.id, 5)
    }, 0)
    await rejection
    assert.equal(service.getCached(chart, { ...context, modelRevision: 5 }), undefined)
  }
})

test('source permission is checked before reading and again after yielding', async () => {
  const { workbook, sheet, chart } = fixture(32769), { service, context } = serviceFor(workbook)
  const rule = { id: 'private', permissionId: 'permission', ranges: [chart.source.range], description: '',
    ownerUserId: 'other-user', allowedUserIds: [], editState: 'onlyMe', viewState: 'noOneElseCanView' }
  sheet.rangeProtectionRules = [rule]
  assert.throws(() => service.requestSourceData(workbook, chart, context), error => error.code === 'COMPUTE_INVALID_INPUT')
  sheet.rangeProtectionRules = []
  const pending = service.requestSourceData(workbook, chart, context)
  const rejection = assert.rejects(pending.promise, error => error.code === 'COMPUTE_INVALID_INPUT')
  setTimeout(() => { sheet.rangeProtectionRules = [rule] }, 0)
  await rejection
})

test('cooperative projection preserves source numeric roles and the small synchronous contract', async () => {
  const { workbook, chart } = fixture(3, 'line')
  const source = await readChartRecordsAsync(workbook, chart)
  assert.deepEqual(await projectChartRecordsAsync(source, chart), projectChartRecords(source, chart))
  const pivot = { fields: ['Numeric ID', 'Value'], records: [[42, 3], [43, 4]],
    sourceNumericFields: [0, 1], sourceValueFields: [1], sourceCategoryFields: [0] }
  const data = await projectChartRecordsAsync(pivot, chart)
  assert.deepEqual(data.sourceNumericFields, [0, 1]); assert.deepEqual(data.sourceValueFields, [1])
  assert.deepEqual(data.series.map(series => series.field), [1])
})

test('masked, graph and hierarchy projections yield to cancellation without returning partial data', async () => {
  for (const type of ['line', 'sankey', 'treemap']) {
    const controller = new AbortController(), records = Array.from({ length: 24000 }, (_unused, index) => ['A', 'B' + index, 1])
    const source = { fields: ['From', 'To', 'Value'], records, sourceNumericFields: [2],
      ...(type === 'line' ? { visibleFields: [0, 2] } : {}) }
    const chart = { type, categoryField: 0, valueFields: [2], options: type === 'sankey'
      ? { mapping: { edge: { sourceIndex: 0, targetIndex: 1, valueIndex: 2 } } }
      : type === 'treemap' ? { mapping: { hierarchyIndexes: [0, 1] } } : {} }
    const pending = projectChartRecordsAsync(source, chart, { signal: controller.signal })
    const rejection = assert.rejects(pending, error => error.code === 'COMPUTE_CANCELLED')
    setTimeout(() => controller.abort(), 0)
    await rejection
  }
})

test('grouped bubble projection checks its real aligned-array peak before creating excessive groups', async () => {
  const source = { fields: ['X', 'Y', 'Size', 'Group'], sourceNumericFields: [0, 1, 2],
    records: Array.from({ length: 1024 }, (_unused, index) => [index, index, 1, 'Group ' + index]) }
  const chart = { type: 'bubble', categoryField: 0, valueFields: [1, 2], options: {
    mapping: { bubble: { xIndex: 0, yIndex: 1, sizeIndex: 2, seriesIndex: 3 } } } }
  await assert.rejects(projectChartRecordsAsync(source, chart, { memoryBudgetBytes: 12 * 1024 * 1024 }), error => error.code === 'COMPUTE_MEMORY_LIMIT')
})
