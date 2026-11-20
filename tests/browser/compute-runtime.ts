import { OfficeComputeClient } from '../../packages/office-studio/src/compute/client.ts'
import type { ComputeProgress, ComputeTask } from '../../packages/office-studio/src/compute/protocol.ts'

type RuntimeCheck = { indices?: number[]; engineVersion?: string; cancelled?: string; stale?: string; disposed?: string }

declare global {
  interface Window {
    computeRuntime: {
      runSortAndFilter(): Promise<{ sort: RuntimeCheck; filter: RuntimeCheck }>
    runAggregate(): Promise<{ statistics: Record<string, number | boolean | null>; engineVersion: string }>
      runCancellation(): Promise<RuntimeCheck>
      runStaleRevision(): Promise<RuntimeCheck>
      runDispose(): Promise<RuntimeCheck>
    }
  }
}

const status = document.querySelector<HTMLElement>('#status')
if (!status) throw new Error('Compute runtime status element is missing')
status.textContent = 'ready'

window.computeRuntime = {
  async runSortAndFilter() {
    const client = new OfficeComputeClient()
    client.setDocumentRevision('runtime-workbook', 1)
    try {
      const sort = await client.submit({
        kind: 'sort',
        values: Float64Array.from([9, 2, 5, -1]),
        states: new Uint8Array([1, 1, 1, 1]),
        rowCount: 4,
        directions: new Uint8Array([0]),
        nullsLast: true,
      }, { documentId: 'runtime-workbook', modelRevision: 1, ownership: 'transfer', timeoutMs: 10_000 }).promise
      if (sort.output.kind !== 'sort') throw new Error(`Unexpected sort output: ${sort.output.kind}`)

      const filter = await client.submit({
        kind: 'filter',
        values: Float64Array.from([2, 3, 7, 8, 9]),
        states: new Uint8Array([1, 1, 1, 1, 1]),
        rowCount: 5,
        operators: new Uint8Array([6]),
        lower: Float64Array.from([3]),
        upper: Float64Array.from([8]),
        match: 'all',
      }, { documentId: 'runtime-workbook', modelRevision: 1, ownership: 'transfer', timeoutMs: 10_000 }).promise
      if (filter.output.kind !== 'filter') throw new Error(`Unexpected filter output: ${filter.output.kind}`)
      return {
        sort: { indices: Array.from(sort.output.indices), engineVersion: sort.engineVersion },
        filter: { indices: Array.from(filter.output.indices), engineVersion: filter.engineVersion },
      }
    } finally {
      client.dispose()
    }
  },

  async runAggregate() {
    const client = new OfficeComputeClient()
    client.setDocumentRevision('aggregate-workbook', 1)
    try {
      const result = await client.submit({
        kind: 'aggregate',
        values: Float64Array.from([4, 2, 7, Number.POSITIVE_INFINITY, 99]),
        states: new Uint8Array([1, 1, 1, 1, 3]),
      }, { documentId: 'aggregate-workbook', modelRevision: 1, ownership: 'transfer', timeoutMs: 10_000 }).promise
      if (result.output.kind !== 'aggregate') throw new Error(`Unexpected aggregate output: ${result.output.kind}`)
      return { statistics: result.output.statistics, engineVersion: result.engineVersion }
    } finally {
      client.dispose()
    }
  },

  async runCancellation() {
    const client = new OfficeComputeClient()
    client.setDocumentRevision('cancel-workbook', 4)
    let resolveProgress!: (progress: ComputeProgress) => void
    const firstProgress = new Promise<ComputeProgress>((resolve) => { resolveProgress = resolve })
    const job = client.submit(largeSortTask(800_000), {
      documentId: 'cancel-workbook', modelRevision: 4, ownership: 'transfer', timeoutMs: 20_000,
      onProgress: (progress) => resolveProgress(progress),
    })
    try {
      await firstProgress
      job.cancel()
      try {
        await job.promise
        throw new Error('Cancelled computation unexpectedly produced a result')
      } catch (error) {
        const cancelled = errorCode(error)
        const recovered = await client.submit({
          kind: 'sort',
          values: Float64Array.from([4, 1, 3]),
          states: new Uint8Array([1, 1, 1]),
          rowCount: 3,
          directions: new Uint8Array([0]),
          nullsLast: true,
        }, { documentId: 'cancel-workbook', modelRevision: 4, ownership: 'transfer', timeoutMs: 10_000 }).promise
        if (recovered.output.kind !== 'sort') throw new Error(`Unexpected recovery output: ${recovered.output.kind}`)
        return { cancelled, recoveredIndices: Array.from(recovered.output.indices) }
      }
    } finally {
      client.dispose()
    }
  },

  async runStaleRevision() {
    const client = new OfficeComputeClient()
    client.setDocumentRevision('stale-workbook', 10)
    let resolveProgress!: (progress: ComputeProgress) => void
    const firstProgress = new Promise<ComputeProgress>((resolve) => { resolveProgress = resolve })
    const job = client.submit(largeSortTask(800_000), {
      documentId: 'stale-workbook', modelRevision: 10, ownership: 'transfer', timeoutMs: 20_000,
      onProgress: (progress) => resolveProgress(progress),
    })
    try {
      await firstProgress
      client.setDocumentRevision('stale-workbook', 11)
      try {
        await job.promise
        throw new Error('Stale computation unexpectedly produced a result')
      } catch (error) {
        return { stale: errorCode(error) }
      }
    } finally {
      client.dispose()
    }
  },

  async runDispose() {
    const client = new OfficeComputeClient()
    client.setDocumentRevision('dispose-workbook', 2)
    let resolveProgress!: (progress: ComputeProgress) => void
    const firstProgress = new Promise<ComputeProgress>((resolve) => { resolveProgress = resolve })
    const job = client.submit(largeSortTask(800_000), {
      documentId: 'dispose-workbook', modelRevision: 2, ownership: 'transfer', timeoutMs: 20_000,
      onProgress: (progress) => resolveProgress(progress),
    })
    await firstProgress
    client.dispose()
    try {
      await job.promise
      throw new Error('Disposed computation unexpectedly produced a result')
    } catch (error) {
      return { disposed: errorCode(error) }
    }
  },
}

function largeSortTask(rowCount: number): ComputeTask {
  const values = new Float64Array(rowCount)
  for (let row = 0; row < rowCount; row += 1) values[row] = (rowCount - row) % 8191
  return { kind: 'sort', values, states: new Uint8Array(rowCount).fill(1), rowCount, directions: new Uint8Array([0]), nullsLast: true }
}

function errorCode(error: unknown): string {
  return error && typeof error === 'object' && 'code' in error ? String(error.code) : String(error)
}
