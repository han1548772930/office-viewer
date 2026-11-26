import { OfficeComputeClient } from '../../src/compute/client.ts'
import { estimateTaskMemory, taskViews } from '../../src/compute/tasks.ts'
import { createCell } from '../../src/domain/model.ts'
import { setAutoFilter, setFilterCriteria } from '../../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { SlicerMemberInputBuilder } from '../../src/features/slicer/slicer-member-input.ts'
import { SlicerMemberQueryError, SlicerMemberQueryService, getSlicerDisplayMembers } from '../../src/features/slicer/slicer-member-query.ts'
import { patchSlicerCacheOptions } from '../../src/features/slicer/slicer-cache-options.ts'
import { setSlicerRecordSelection } from '../../src/features/slicer/slicer-engine.ts'
import { filterReports, slicerRecordSelectionWorkbook, slicerTableRecordSelectionWorkbook } from './pivot-filter-fixture.mjs'
import { referenceSlicerMemberWindow } from './slicer-member-compute-fixture.mjs'
import { fixedSlicerWorkerRows, fixedSlicerWorkerAvailability, generateSlicerWorkerRows,
  slicerWorkerDefaults, slicerWorkerTableWorkbook } from './slicer-member-worker-data.mjs'

const yieldTask = () => new Promise(resolve => setTimeout(resolve, 0))
export function plainSlicerWorkerOutput(output) {
  return { kind: output.kind, indices: Array.from(output.indices), hasData: Array.from(output.hasData), total: output.total,
    dateCount: output.dateCount, minDate: output.minDate, maxDate: output.maxDate }
}
export async function packSlicerWorkerRows(rows, availability = rows.map(() => true), options = {}) {
  const builder = new SlicerMemberInputBuilder(rows.length, { ...slicerWorkerDefaults, ...options })
  for (let index = 0; index < rows.length; index++) {
    builder.append(rows[index], availability[index] === true || availability[index] === 1)
    if ((index + 1) % 2048 === 0) await yieldTask()
  }
  return builder.finish(yieldTask)
}

/** Real Worker observation only: every calculation and result comes from the
 * production Worker and the required local Rust/WASM artifact. No JS double.
 */
export function observedSlicerWorkerClient(options = {}) {
  const workers = [], events = [], listeners = new Set()
  const client = new OfficeComputeClient({ ...options, workerFactory() {
    const worker = new Worker(new URL('../../src/compute/compute.worker.ts', import.meta.url), { type: 'module', name: 'slicer-member-real-worker' })
    workers.push(worker)
    worker.addEventListener('message', event => {
      const observed = { data: event.data, atMs: performance.now() }
      events.push(observed)
      for (const listener of [...listeners]) listener(observed)
    })
    return worker
  } })
  return { client, workers, events, observe(listener) { listeners.add(listener); return () => listeners.delete(listener) } }
}
function detached(view) { try { new Uint8Array(view.buffer, 0, 0); return false } catch { return true } }
const errorInfo = error => ({ code: error?.code ?? error?.name, name: error?.name, message: error?.message })

export function createSlicerMemberWorkerRuntime() {
  const observed = observedSlicerWorkerClient(), { client, workers, events } = observed
  const services = new Set()
  let revision = 0
  async function raw(rows, availability, options = {}, context = {}) {
    const task = await packSlicerWorkerRows(rows, availability, options), views = taskViews(task)
    const job = client.submit(task, { documentId: context.documentId ?? 'slicer-real-binding', modelRevision: context.modelRevision ?? revision,
      ownership: 'transfer', timeoutMs: 30_000, ...context })
    const callerDetached = views.every(detached), result = await job.promise
    return { output: plainSlicerWorkerOutput(result.output), metrics: result.metrics, engineVersion: result.engineVersion,
      context: result.context, callerDetached }
  }
  const warm = () => raw(fixedSlicerWorkerRows(), fixedSlicerWorkerAvailability)
  return {
    async fixed(options = {}) { return raw(fixedSlicerWorkerRows(), fixedSlicerWorkerAvailability, options) },
    async text(locale, customListSort, sortOrder) {
      const long = new Intl.DateTimeFormat(locale, { month: 'long', timeZone: 'UTC' })
      const short = new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' })
      const values = [long.format(new Date(Date.UTC(2020, 1, 1))), short.format(new Date(Date.UTC(2020, 0, 1))),
        long.format(new Date(Date.UTC(2020, 0, 1))), 'item10', 'item2', 'ALPHA', 'alpha', '\ud800', '\ud801', 'a\0b', '😀']
      const rows = [...values.map(value => ({ type: 'string', value })), { type: 'boolean', value: true }, { type: 'boolean', value: false }]
      const availability = rows.map(() => true), options = { locale, customListSort, sortOrder }
      const baseline = referenceSlicerMemberWindow(rows, Uint8Array.from(availability, value => value ? 1 : 0), { ...slicerWorkerDefaults, ...options })
      return { ...await raw(rows, availability, options), baseline: plainSlicerWorkerOutput(baseline), values }
    },
    async sourceWindow(kind, options, request = {}) {
      const workbook = kind === 'table' ? slicerTableRecordSelectionWorkbook() : slicerRecordSelectionWorkbook(), slicer = workbook.slicers[0]
      if (kind === 'table') {
        const table = workbook.tables.find(candidate => candidate.id === 'typed-table'), sheet = workbook.sheets.find(candidate => candidate.id === 'data')
        const filterSheet = { ...sheet, autoFilter: undefined }; setAutoFilter(filterSheet, table.range)
        setFilterCriteria(filterSheet, 3, { kind: 'condition', operator: 'equals', value: 'North' }); table.autoFilter = filterSheet.autoFilter
      } else for (const report of filterReports(workbook)) {
        report.fields.find(field => field.sourceFieldId === 'region-source').filter = { mode: 'include', values: ['North'] }
      }
      setSlicerRecordSelection(workbook, slicer, { mode: 'include', values: [{ type: 'number', value: 1 }] })
      patchSlicerCacheOptions(workbook, slicer, options); slicer.startItem = 3
      const before = JSON.stringify(workbook), metrics = []
      const baselineMembers = getSlicerDisplayMembers(workbook, slicer, 'en-US')
      const start = request.start ?? 0, count = Math.min(request.count ?? 100, 200)
      const baseline = { total: baselineMembers.length, items: baselineMembers.slice(start, start + count).map(({ key, value, hasData }) => ({ key, value, hasData })) }
      const service = new SlicerMemberQueryService({ computeClient: client, onMetrics: value => metrics.push(value) }); services.add(service)
      try {
        const window = await service.getItems(workbook, slicer, request, { unitId: workbook.id, baseRevision: 0, locale: 'en-US', assertCurrent() {} })
        return { window, baseline, unchanged: JSON.stringify(workbook) === before, metrics,
          engineVersion: events.findLast(event => event.data.type === 'result' && event.data.context.unitId === workbook.id)?.data.engineVersion }
      } finally { service.dispose(); services.delete(service) }
    },
    async metadata(noDates = false) {
      const workbook = slicerRecordSelectionWorkbook(), source = workbook.sheets.find(sheet => sheet.id === 'data')
      const slicer = { ...workbook.slicers[0], id: 'slicer-real-timeline', type: 'timeline', sourceFieldId: 'date-source', cacheId: undefined }
      workbook.slicers.push(slicer)
      for (const report of filterReports(workbook)) report.fields.find(field => field.sourceFieldId === 'date-source').dateGroupType = 'month'
      const values = noDates ? [100, 1_000_000, '1900-01-01', '-99', '#N/A', 45000, null] : [100, 1_000_000, '1900-01-01', '-99', '#N/A', -2.25, null]
      values.forEach((value, index) => {
        const cell = createCell(index + 3, 4, value)
        if (index === 0 && !noDates) { cell.formula = '=DATE(1900,4,9)'; cell.cachedDataType = 'date' }
        if (index === 4) { cell.formula = '=NA()'; cell.cachedDataType = 'error'; cell.metadata = { error: '#N/A' } }
        if (index === 5 && !noDates) cell.dataType = 'date'
        source.cells[`${index + 3}:4`] = cell
      })
      const before = JSON.stringify(workbook), metrics = []
      const service = new SlicerMemberQueryService({ computeClient: client, onMetrics: value => metrics.push(value) }); services.add(service)
      try {
        const metadata = await service.getSourceMetadata(workbook, slicer, { unitId: workbook.id, baseRevision: 0, assertCurrent() {} })
        return { metadata, metrics, unchanged: JSON.stringify(workbook) === before }
      } finally { service.dispose(); services.delete(service) }
    },
    async empty(mode) { return raw([], [], { mode }) },
    async canvas(options = {}, request = {}) {
      const { workbook, slicer } = slicerWorkerTableWorkbook(fixedSlicerWorkerRows(), fixedSlicerWorkerAvailability, options)
      const service = new SlicerMemberQueryService({ computeClient: client }); services.add(service)
      try {
        const window = await service.getDisplayMemberWindow(workbook, slicer, request, { unitId: workbook.id, baseRevision: 0, locale: 'en-US', assertCurrent() {} })
        return { window, engineVersion: events.findLast(event => event.data.type === 'result')?.data.engineVersion }
      } finally { service.dispose(); services.delete(service) }
    },
    async canvasDateCaptions(hideItemsWithNoData = false) {
      const rows = [{ type:'date',value:100 },{ type:'date',value:100 },{ type:'number',value:100 },{ type:'string',value:'100' }]
      const { workbook,slicer } = slicerWorkerTableWorkbook(rows,[false,true,false,false],{ hideItemsWithNoData })
      const sheet = workbook.sheets.find(sheet => sheet.id === 'data'); sheet.cells['3:6'].value = '1900-04-09'
      setSlicerRecordSelection(workbook,slicer,{ mode:'include',values:[{ type:'date',value:100 },{ type:'number',value:100 }] })
      const service = new SlicerMemberQueryService({ computeClient: client }); services.add(service)
      try {
        const owner = { unitId:workbook.id,baseRevision:0,locale:'en-US',assertCurrent() {} }
        const display = await service.getDisplayMemberWindow(workbook,slicer,{},owner)
        const formal = await service.getItems(workbook,slicer,{},owner)
        return { display,formal,engineVersion:events.findLast(event => event.data.type === 'result')?.data.engineVersion }
      } finally { service.dispose();services.delete(service) }
    },
    async lifecycle(action) {
      await warm()
      const generated = generateSlicerWorkerRows(65_536), { workbook, slicer } = slicerWorkerTableWorkbook(generated.rows, generated.availability)
      const service = new SlicerMemberQueryService({ computeClient: client }); services.add(service)
      const controller = new AbortController(); let invalidated = false, triggered = false, triggerAtMs
      client.setDocumentRevision(workbook.id, 0)
      const remove = observed.observe(({ data }) => {
        if (triggered || data.type !== 'progress' || data.context.unitId !== workbook.id || data.ratio !== 0) return
        triggered = true; triggerAtMs = performance.now()
        if (action === 'cancel' || action === 'canvas' || action === 'metadata') controller.abort()
        else if (action === 'close') service.forgetDocument(workbook.id)
        else if (action === 'service') service.dispose()
        else if (action === 'revision') client.setDocumentRevision(workbook.id, 1)
        else invalidated = true
      })
      try {
        const owner = { unitId: workbook.id, baseRevision: 0, assertCurrent() {
          if (invalidated || !client.isCurrent({ documentId: workbook.id, modelRevision: 0 })) {
            throw new SlicerMemberQueryError('SLICER_SOURCE_INVALIDATED', action === 'permission' ? 'permission guard invalidated' : 'source guard invalidated')
          }
        } }
        const query = action === 'metadata' ? service.getSourceMetadata(workbook,slicer,owner,controller.signal)
          : action === 'canvas' ? service.getDisplayMemberWindow(workbook,slicer,{ signal:controller.signal },owner)
          : service.getItems(workbook,slicer,{ signal:controller.signal },owner)
        const pending = query.then(() => ({ code:'unexpected-success' }),errorInfo)
        const rejected = await pending, rejectedAtMs = performance.now()
        remove(); service.dispose(); services.delete(service)
        const recovered = await warm()
        return { rejected, triggered, recovery: recovered.output, engineVersion: recovered.engineVersion, workerCount: workers.length,
          rejectionMs: triggered ? rejectedAtMs - triggerAtMs : null, recoveryMs: triggered ? performance.now() - triggerAtMs : null }
      } finally { remove(); service.dispose(); services.delete(service) }
    },
    async directCancellation() {
      await warm()
      const generated = generateSlicerWorkerRows(65_536), preparationStartedAt = performance.now()
      const task = await packSlicerWorkerRows(generated.rows, generated.availability), memory = estimateTaskMemory(task)
      const preparationMs = performance.now() - preparationStartedAt
      let triggered = false, cancelledAtMs, requestId
      const job = client.submit(task, { documentId: 'slicer-direct-cancel', modelRevision: 0, ownership: 'transfer', timeoutMs: 30_000,
        onProgress(progress) { if (!triggered && progress.ratio === 0) { triggered = true; cancelledAtMs = performance.now(); job.cancel() } } })
      requestId = job.requestId
      const rejected = await job.promise.then(() => ({ code: 'unexpected-success' }), errorInfo)
      const rejectedAtMs = performance.now(), recovered = await warm()
      const acknowledgement = events.find(event => event.data.type === 'cancelled' && event.data.context.requestId === requestId)
      return { triggered, rejected, workerCount: workers.length, recovered: recovered.output, engineVersion: recovered.engineVersion,
        rowCount: generated.rows.length, seed: generated.seed, datasetVersion: generated.version, preparationMs,
        inputBytes: memory.inputBytes, estimatedPeakBytes: memory.estimatedPeakBytes, recoveredMetrics: recovered.metrics,
        rejectionMs: triggered ? rejectedAtMs - cancelledAtMs : null, acknowledgementMs: acknowledgement ? acknowledgement.atMs - cancelledAtMs : null,
        recoveryMs: triggered ? performance.now() - cancelledAtMs : null }
    },
    async lateResult() {
      await warm()
      const packet = events.findLast(event => event.data.type === 'result').data
      const generated = generateSlicerWorkerRows(16_384), task = await packSlicerWorkerRows(generated.rows, generated.availability)
      const baseline = referenceSlicerMemberWindow(generated.rows, Uint8Array.from(generated.availability, value => value ? 1 : 0), task.params)
      let replayed = false
      const job = client.submit(task, { documentId: 'slicer-real-late', modelRevision: 0, ownership: 'transfer', timeoutMs: 30_000,
        onProgress(progress) {
          if (!replayed && progress.ratio === 0) {
            // Delivery replay of an actual previous Worker packet. Calculation
            // and current result still come from the real production WASM.
            replayed = true; workers.at(-1).onmessage?.(new MessageEvent('message', { data: packet }))
          }
        } })
      const result = await job.promise
      return { replayed, output: plainSlicerWorkerOutput(result.output), baseline: plainSlicerWorkerOutput(baseline),
        requestId: result.requestId, previousRequestId: packet.context.requestId, engineVersion: result.engineVersion, workerCount: workers.length }
    },
    async release() {
      const task = await packSlicerWorkerRows(fixedSlicerWorkerRows(), fixedSlicerWorkerAvailability)
      const result = await client.submit(task, { documentId: 'slicer-output-owner', modelRevision: 0, ownership: 'transfer' }).promise
      const output = result.output, before = plainSlicerWorkerOutput(output)
      client.dispose()
      const independent = output.indices.buffer !== output.hasData.buffer && output.indices.byteOffset === 0 && output.hasData.byteOffset === 0
      const received = structuredClone(output, { transfer: [output.indices.buffer, output.hasData.buffer] })
      return { before, after: plainSlicerWorkerOutput(received), independent,
        callerDetached: detached(output.indices) && detached(output.hasData), engineVersion: result.engineVersion }
    },
    stats() { return { workerCount: workers.length, readyCount: events.filter(event => event.data.type === 'ready').length } },
    dispose() { for (const service of services) service.dispose(); services.clear(); client.dispose(); revision++ },
  }
}
