import { ComputeError } from '../../src/compute/errors.ts'
import { estimateTaskMemory, taskViews, transferTaskBuffers } from '../../src/compute/tasks.ts'
import { SLICER_NO_TEXT } from '../../src/compute/slicer-member-protocol.ts'
import { createWasmTaskControl } from '../../src/platform/wasm/task-protocol.ts'
import { getRecordValueKey } from '../../src/domain/record-value.ts'
import { applySlicerCacheItemOptions } from '../../src/features/slicer/slicer-cache-options.ts'

// Explicit test-only/reference baseline. This does NOT execute Rust/WASM and
// must never be registered as a production fallback or reported as WASM proof.
function decodeSlicerTextDictionary(task) {
  const texts = []
  for (let id = 0; id < task.params.dictionaryCount; id++) {
    const units = []
    for (let offset = task.textOffsets[id]; offset < task.textOffsets[id + 1]; offset += 2) {
      units.push(task.textBytes[offset] | task.textBytes[offset + 1] << 8)
    }
    let text = ''
    for (let index = 0; index < units.length; index += 4096) text += String.fromCharCode(...units.slice(index, index + 4096))
    texts.push(text)
  }
  return texts
}
export function decodeSlicerMemberCaptions(task) {
  if (task.params.mode !== 'canvas') throw new Error('Captions are only present in the Canvas fixture mode')
  const texts = decodeSlicerTextDictionary(task)
  return Array.from(task.states, (_kind, row) => texts[task.textIds[task.params.rowCount + row]])
}
export function decodeSlicerMemberTask(task) {
  const texts = decodeSlicerTextDictionary(task)
  return Array.from(task.states, (kind, row) => {
    if (kind === 0) return { type: 'blank' }
    if (kind === 1 || kind === 2) return { type: kind === 1 ? 'number' : 'date', value: task.values[row] }
    if (kind === 4) return { type: 'boolean', value: task.values[row] === 1 }
    if ((kind === 3 || kind === 5) && task.textIds[row] !== SLICER_NO_TEXT) {
      return { type: kind === 3 ? 'string' : 'error', value: texts[task.textIds[row]] }
    }
    throw new Error('Unsupported test fixture type')
  })
}

/** Reproducible pre-migration algorithm for comparison, with the same current
 * locale comparator. Independent fixed expectations also exist in Rust tests.
 */
export function referenceSlicerMemberWindow(values, availability, params, captions) {
  let dateCount = 0, minDate = null, maxDate = null
  const members = new Map()
  values.forEach((value, row) => {
    if (value.type === 'date') {
      dateCount++
      minDate = minDate === null ? value.value : Math.min(minDate, value.value)
      maxDate = maxDate === null ? value.value : Math.max(maxDate, value.value)
    }
    if (params.mode === 'metadata') return
    const key = getRecordValueKey(value), previous = members.get(key), hasData = (availability[row] & 1) === 1
    if (previous) previous.hasData ||= hasData
    else members.set(key, { key, value, hasData, row })
  })
  let projected = params.mode === 'metadata' ? [] : applySlicerCacheItemOptions([...members.values()], paramsOptions(params), params.locale)
  if (params.mode === 'canvas') {
    if (!captions || captions.length !== values.length) throw new Error('Canvas reference needs every exact source caption')
    const states = new Map(), folded = new Map()
    captions.forEach((caption, row) => {
      const state = states.get(caption) ?? { selected: false, unselected: false }
      if (availability[row] & 2) state.selected = true; else state.unselected = true
      states.set(caption, state)
    })
    projected.forEach(member => {
      const caption = captions[member.row], previous = folded.get(caption)
      if (previous) previous.hasData ||= member.hasData
      else folded.set(caption, { ...member, caption })
    })
    projected = [...folded.values()].map(member => {
      const state = states.get(member.caption)
      return { ...member, flags: (member.hasData ? 1 : 0) | (state.selected && !state.unselected ? 2 : 0)
        | (state.selected && state.unselected ? 4 : 0) }
    })
  }
  const window = projected.slice(params.start, params.start + params.count)
  return { kind: 'slicerMembers', total: projected.length, indices: Uint32Array.from(window, member => member.row),
    hasData: Uint8Array.from(window, member => params.mode === 'canvas' ? member.flags : member.hasData ? 1 : 0), dateCount, minDate, maxDate }
}
function paramsOptions(params) {
  return { sortOrder: params.sortOrder, customListSort: params.customListSort,
    crossFilter: params.crossFilter, hideItemsWithNoData: params.hideItemsWithNoData }
}

export function slicerComputeFixtureResult(task, options, requestId) {
  const output = referenceSlicerMemberWindow(decodeSlicerMemberTask(task), task.availability, task.params,
    task.params.mode === 'canvas' ? decodeSlicerMemberCaptions(task) : undefined)
  const memory = estimateTaskMemory(task), views = taskViews(task), outputBytes = output.indices.byteLength + output.hasData.byteLength
  return { type: 'result', requestId, documentId: options.documentId, modelRevision: options.modelRevision,
    context: createWasmTaskControl('slicerMembers', requestId, options.documentId, options.modelRevision),
    engineVersion: 'test-fixture-no-wasm', output, metrics: { elapsedMs: 0, ...memory,
      boundary: { inputCopies: views.filter(view => view.byteLength).length, inputBytes: memory.inputBytes,
        outputCopies: outputBytes ? 2 : 0, outputBytes, rustBufferClones: 0, controlDtoConversions: 2, jsonConversions: 0 },
      transport: { inputTransfers: 7, inputBytes: memory.inputBytes, outputTransfers: 2, outputBytes, structuredResultMessages: 1 } } }
}

/** Deterministic client double for source packing, lifecycle and Facade drafts.
 * The real OfficeComputeClient protocol and real WASM are tested separately.
 */
export function createSlicerMemberTestClient({ delayMs = 0, onSubmit } = {}) {
  let sequence = 0, disposed = false
  const revisions = new Map(), jobs = new Set(), submitted = []
  const client = {
    submitted,
    isCurrent(context) { return !disposed && revisions.get(context.documentId) === context.modelRevision },
    setDocumentRevision(unitId, revision) {
      revisions.set(unitId, revision)
      for (const job of jobs) if (job.options.documentId === unitId && job.options.modelRevision !== revision) job.reject(new ComputeError('COMPUTE_STALE_RESULT', 'test revision invalidated'))
    },
    forgetDocument(unitId) {
      for (const job of jobs) if (job.options.documentId === unitId) job.reject(new ComputeError('COMPUTE_STALE_RESULT', 'test unit closed'))
      revisions.delete(unitId)
    },
    submit(task, options) {
      const requestId = `slicer-fixture-${++sequence}`
      const rejected = error => ({ requestId, promise: Promise.reject(error), cancel() {} })
      if (disposed) return rejected(new ComputeError('COMPUTE_DISPOSED', 'test fixture disposed'))
      if (options.signal?.aborted) return rejected(new ComputeError('COMPUTE_CANCELLED', 'test query cancelled'))
      if (task.kind !== 'slicerMembers') return { requestId, promise: Promise.reject(new ComputeError('COMPUTE_UNSUPPORTED_TYPE', 'fixture supports slicerMembers only')), cancel() {} }
      const revision = revisions.get(options.documentId)
      if (revision !== undefined && options.modelRevision < revision) return rejected(new ComputeError('COMPUTE_STALE_RESULT', 'test old revision rejected'))
      const owned = options.ownership === 'transfer' ? structuredClone(task, { transfer: transferTaskBuffers(task) }) : structuredClone(task)
      submitted.push({ task: owned, options: { ...options }, requestId })
      let resolve, reject, timer, settled = false
      const promise = new Promise((yes, no) => { resolve = yes; reject = no })
      const cleanup = () => { clearTimeout(timer); options.signal?.removeEventListener('abort', abort); jobs.delete(job) }
      const job = { options, reject(error) { if (!settled) { settled = true; cleanup(); reject(error) } } }
      const abort = () => job.reject(new ComputeError('COMPUTE_CANCELLED', 'test query cancelled'))
      jobs.add(job)
      options.signal?.addEventListener('abort', abort, { once: true })
      client.setDocumentRevision(options.documentId, options.modelRevision)
      timer = setTimeout(() => {
        if (disposed || options.signal?.aborted) { abort(); return }
        try { const result = slicerComputeFixtureResult(owned, options, requestId); settled = true; cleanup(); resolve(result) }
        catch (error) { job.reject(error) }
      }, delayMs)
      if (disposed || options.signal?.aborted) abort()
      try { onSubmit?.(submitted.at(-1)) } catch (error) { job.reject(error) }
      return { requestId, promise, cancel: abort }
    },
    dispose() {
      disposed = true
      for (const job of [...jobs]) job.reject(new ComputeError('COMPUTE_DISPOSED', 'test fixture disposed'))
      revisions.clear()
    },
  }
  return client
}
