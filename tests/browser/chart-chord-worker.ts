import { OfficeComputeClient, type ComputeWorkerPort } from '../../packages/office-studio/src/compute/client.ts'
import { ComputeError } from '../../packages/office-studio/src/compute/errors.ts'
import { isRecord, type ChartStatisticsTask, type ComputeProgress, type ComputeResult } from '../../packages/office-studio/src/compute/protocol.ts'

// P28 browser drafts. Each port forwards to a native module Worker running the
// production compute.worker.ts and the newly built protocol 2 WASM. Tracing
// records postMessage/terminate only; it supplies no ready/progress/results.
interface WorkerPost {
  type: string; requestId?: string; transferCount: number; delivered: boolean
}
interface WorkerLifecycle { posts: WorkerPost[]; terminated: boolean }
interface WorkerTrace { workers: WorkerLifecycle[]; failedCancelPosts: number }

function trackedWorkers(failFirstCancel = false) {
  const trace: WorkerTrace = { workers: [], failedCancelPosts: 0 }
  let cancelFailurePending = failFirstCancel
  const workerFactory = (): ComputeWorkerPort => {
    const worker = new Worker(new URL('../../packages/office-studio/src/compute/compute.worker.ts', import.meta.url), { type: 'module' })
    const lifecycle: WorkerLifecycle = { posts: [], terminated: false }
    trace.workers.push(lifecycle)
    return {
      get onmessage() { return worker.onmessage },
      set onmessage(listener) { worker.onmessage = listener },
      get onerror() { return worker.onerror },
      set onerror(listener) { worker.onerror = listener },
      get onmessageerror() { return worker.onmessageerror },
      set onmessageerror(listener) { worker.onmessageerror = listener },
      postMessage(message: unknown, transferOrOptions?: Transferable[] | StructuredSerializeOptions) {
        const type = isRecord(message) && typeof message.type === 'string' ? message.type : 'unknown'
        const context = isRecord(message) && isRecord(message.context) ? message.context : undefined
        const transfer = Array.isArray(transferOrOptions) ? transferOrOptions : transferOrOptions?.transfer ?? []
        const post: WorkerPost = { type, transferCount: transfer.length, delivered: false,
          ...(context && typeof context.requestId === 'string' ? { requestId: context.requestId } : {}) }
        lifecycle.posts.push(post)
        // Deliberately fail the cancellation channel once. This verifies the
        // client's immediate kill/recreate branch, not synchronous kernel time.
        if (type === 'cancel' && cancelFailurePending) {
          cancelFailurePending = false; trace.failedCancelPosts++
          throw new DOMException('Injected Chord cancel postMessage failure', 'DataCloneError')
        }
        if (Array.isArray(transferOrOptions)) worker.postMessage(message, transferOrOptions)
        else worker.postMessage(message, transferOrOptions)
        post.delivered = true
      },
      terminate() { lifecycle.terminated = true; worker.terminate() },
    }
  }
  return { trace, workerFactory }
}

function chordTask(errorWeight = false): ChartStatisticsTask {
  const states = new Uint8Array([1, 1, 1, 1, 1, 1])
  if (errorWeight) states[0] = 3
  return {
    kind: 'chartStatistics', values: new Float64Array([2, 3, 0, 1, 1, 0]), states, categoryIds: new Uint32Array([0, 0]),
    params: { rowCount: 2, columnCount: 3, categoryCount: 1,
      series: [{ field: 2, seriesPosition: 0, column: 0 }],
      jobs: [{ kind: 'chord', seriesIndex: 0, slot: 0, sourceColumn: 1, targetColumn: 2, labelCount: 2,
        startAngle: 90, clockwise: true, padAngle: 0, minAngle: 0 }] },
  }
}

function readOutput(result: ComputeResult) {
  if (result.output.kind !== 'chartStatistics') throw new Error(`Unexpected Chord output: ${result.output.kind}`)
  const { values, indices, segments } = result.output
  return { requestId: result.requestId, documentId: result.documentId, modelRevision: result.modelRevision,
    context: result.context, engineVersion: result.engineVersion, metrics: result.metrics,
    values: Array.from(values), indices: Array.from(indices), segments,
    independent: values.buffer !== indices.buffer && values.byteOffset === 0 && indices.byteOffset === 0
      && values.byteLength === values.buffer.byteLength && indices.byteLength === indices.buffer.byteLength }
}

async function requireRejection(promise: Promise<ComputeResult>) {
  return promise.then(() => { throw new Error('Rejected Chord request published a successful result') }, (error: unknown) => {
    if (!(error instanceof ComputeError)) throw error
    return { code: error.code, message: error.message }
  })
}
function checkpoint(progress: ComputeProgress): boolean { return progress.phase === 'computing' && progress.ratio === 0 }
const submitContext = (documentId: string, modelRevision = 0) => ({ documentId, modelRevision, ownership: 'transfer' as const, timeoutMs: 15_000 })

async function runTransfer() {
  const tracked = trackedWorkers(), client = new OfficeComputeClient({ workerFactory: tracked.workerFactory }), task = chordTask()
  const callerBuffers = [task.values.buffer, task.states.buffer, task.categoryIds.buffer]
  const inputBytes = callerBuffers.map(buffer => buffer.byteLength)
  try {
    const job = client.submit(task, submitContext('chord-worker-transfer'))
    const callerBytesAfterSubmit = callerBuffers.map(buffer => buffer.byteLength)
    const first = await job.promise, output = readOutput(first), secondTask = chordTask()
    secondTask.values[0] = 7; secondTask.values[1] = 1
    const second = readOutput(await client.submit(secondTask, submitContext('chord-worker-transfer')).promise)
    if (first.output.kind !== 'chartStatistics') throw new Error('Original Chord result changed kind')
    return { inputBytes, callerBytesAfterSubmit, output, second,
      firstValuesAfterReuse: Array.from(first.output.values), firstIndicesAfterReuse: Array.from(first.output.indices), trace: tracked.trace }
  } finally { client.dispose() }
}

async function runCancellation(failCancelPost: boolean) {
  const tracked = trackedWorkers(failCancelPost), client = new OfficeComputeClient({ workerFactory: tracked.workerFactory, cancellationGraceMs: 500 })
  const controller = new AbortController(), progress: ComputeProgress[] = []
  try {
    const job = client.submit(chordTask(), { ...submitContext('chord-worker-cancel'), signal: controller.signal,
      onProgress(value) { progress.push(value); if (checkpoint(value)) controller.abort() } })
    const rejected = await requireRejection(job.promise)
    const recovered = readOutput(await client.submit(chordTask(), submitContext('chord-worker-cancel')).promise)
    return { rejected, recovered, progress, aborted: controller.signal.aborted, trace: tracked.trace }
  } finally { client.dispose() }
}

async function runStaleRevision() {
  const tracked = trackedWorkers(), client = new OfficeComputeClient({ workerFactory: tracked.workerFactory, cancellationGraceMs: 500 })
  const documentId = 'chord-worker-revision', progress: ComputeProgress[] = []
  try {
    const job = client.submit(chordTask(), { ...submitContext(documentId),
      onProgress(value) { progress.push(value); if (checkpoint(value)) client.setDocumentRevision(documentId, 1) } })
    const rejected = await requireRejection(job.promise)
    const recovered = readOutput(await client.submit(chordTask(), submitContext(documentId, 1)).promise)
    const oldTask = chordTask(), oldBuffers = [oldTask.values.buffer, oldTask.states.buffer, oldTask.categoryIds.buffer]
    const oldAdmission = await requireRejection(client.submit(oldTask, submitContext(documentId)).promise)
    return { rejected, recovered, progress, oldAdmission, oldCallerBytes: oldBuffers.map(buffer => buffer.byteLength),
      oldCurrent: client.isCurrent({ documentId, modelRevision: 0 }), newCurrent: client.isCurrent({ documentId, modelRevision: 1 }), trace: tracked.trace }
  } finally { client.dispose() }
}

async function runInvalidWeight() {
  const tracked = trackedWorkers(), client = new OfficeComputeClient({ workerFactory: tracked.workerFactory })
  try {
    // State 3 is a globally valid spreadsheet error state. Only the Chord
    // Rust job rejects this weight, after receiving the real bulk input.
    const rejected = await requireRejection(client.submit(chordTask(true), submitContext('chord-worker-invalid')).promise)
    const recovered = readOutput(await client.submit(chordTask(), submitContext('chord-worker-invalid')).promise)
    return { rejected, recovered, trace: tracked.trace }
  } finally { client.dispose() }
}

async function runDispose() {
  const tracked = trackedWorkers(), client = new OfficeComputeClient({ workerFactory: tracked.workerFactory }), progress: ComputeProgress[] = []
  try {
    const job = client.submit(chordTask(), { ...submitContext('chord-worker-dispose'),
      onProgress(value) { progress.push(value); if (checkpoint(value)) client.dispose() } })
    const rejected = await requireRejection(job.promise)
    const afterDispose = await requireRejection(client.submit(chordTask(), submitContext('chord-worker-dispose')).promise)
    return { rejected, afterDispose, progress, trace: tracked.trace }
  } finally { client.dispose() }
}

const runtime = {
  runTransfer, runCancellation: () => runCancellation(false), runCancelPostFailure: () => runCancellation(true),
  runStaleRevision, runInvalidWeight, runDispose,
}
declare global { interface Window { chordWorkerRuntime: typeof runtime } }
window.chordWorkerRuntime = runtime
const status = document.getElementById('status')
if (!status) throw new Error('Chord Worker fixture status element is missing')
status.textContent = 'ready'
