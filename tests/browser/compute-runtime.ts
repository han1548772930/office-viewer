import { OfficeComputeClient } from '../../packages/office-studio/src/compute/client.ts'
import type { ComputeMetrics, ComputeProgress, ComputeTask, NumericStatistics } from '../../packages/office-studio/src/compute/protocol.ts'
import type { ChartStatisticsSegment, ChartStatisticsTask } from '../../packages/office-studio/src/compute/chart-protocol.ts'
import { ChartStatisticsService } from '../../packages/office-studio/src/features/charts/services/chart-statistics-service.ts'
import { createChart } from '../../packages/office-studio/src/features/charts/chart-engine.ts'
import { createBlankWorkbook, createCell } from '../../packages/office-studio/src/domain/model.ts'

type RuntimeCheck = { indices?: number[]; recoveredIndices?: number[]; engineVersion?: string; cancelled?: string; stale?: string; disposed?: string; recoveredChart?: number[] }
type ChartRuntimeCheck = { values: number[]; indices: number[]; segments: ChartStatisticsSegment[]; metrics: ComputeMetrics; callerDetached: boolean }
type ChartServiceCheck = { categories: string[]; values: Array<number | null>; counts: number[]; updatedCounts: number[]; cached: boolean; invalidated: boolean; metrics?: ComputeMetrics }

declare global {
  interface Window {
    computeRuntime: {
      runSortAndFilter(): Promise<{ sort: RuntimeCheck; filter: RuntimeCheck }>
      runAggregate(): Promise<{ statistics: NumericStatistics; engineVersion: string }>
      runChartStatistics(): Promise<ChartRuntimeCheck>
      runChartService(): Promise<ChartServiceCheck>
      runChartCancellation(): Promise<RuntimeCheck>
      runChartStaleRevision(): Promise<RuntimeCheck>
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

  async runChartStatistics() {
    const client = new OfficeComputeClient(), task = smallChartTask()
    try {
      const job = client.submit(task, { documentId: 'chart-binding-book', modelRevision: 11, ownership: 'transfer', timeoutMs: 10_000 })
      const callerDetached = [task.values, task.states, task.categoryIds].every(view => view.byteLength === 0)
      const result = await job.promise
      if (result.output.kind !== 'chartStatistics') throw new Error(`Unexpected chart output: ${result.output.kind}`)
      // Verify transferred result data remains readable after the Worker owner
      // and client have been released, as it does in the real chart host.
      client.dispose()
      return { values: Array.from(result.output.values), indices: Array.from(result.output.indices),
        segments: result.output.segments, metrics: result.metrics, callerDetached }
    } finally { client.dispose() }
  },

  async runChartService() {
    const client = new OfficeComputeClient(), service = new ChartStatisticsService(client)
    const workbook = { ...createBlankWorkbook(), id: 'chart-source-book' }, sheet = workbook.sheets[0]!
    sheet.cells = { '0:0': createCell('Label'), '0:1': createCell('Value') }
    for (const [index, value] of [4, 7, 9].entries()) {
      sheet.cells[`${index + 1}:0`] = createCell(['A', 'B', 'C'][index]!)
      sheet.cells[`${index + 1}:1`] = createCell(value)
    }
    const chart = createChart(workbook, sheet.id, { type: 'histogram',
      source: { orientation: 'columns', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 } },
      position: { x: 20, y: 20 }, size: { width: 360, height: 240 }, categoryField: 0, valueFields: [1],
      options: { histogram: { binning: { mode: 'count', count: 2 } } },
    }, 'runtime-histogram')
    workbook.charts = [chart]
    service.setDocumentRevision(workbook.id, 1)
    const context = { documentId: workbook.id, modelRevision: 1, priority: 'interactive' as const, timeoutMs: 10_000 }
    try {
      const first = await service.request(workbook, chart, context).promise
      const cached = service.getCached(chart, context) === first
      sheet.cells['3:1'] = createCell(14)
      service.setDocumentRevision(workbook.id, 2)
      const nextContext = { ...context, modelRevision: 2 }
      const invalidated = service.getCached(chart, nextContext) === undefined
      const updated = await service.request(workbook, chart, nextContext).promise
      return { categories: first.statistics.preparedData?.categories ?? [], values: first.statistics.preparedData?.series[0]?.values ?? [],
        counts: first.statistics.histogram?.bins.map(bin => bin.count) ?? [], updatedCounts: updated.statistics.histogram?.bins.map(bin => bin.count) ?? [],
        cached, invalidated, metrics: first.metrics }
    } finally { service.dispose(); client.dispose() }
  },

  async runChartCancellation() {
    const client = new OfficeComputeClient()
    let resolveProgress!: (progress: ComputeProgress) => void
    const progress = new Promise<ComputeProgress>(resolve => { resolveProgress = resolve })
    const job = client.submit(largeChartTask(800_000), { documentId: 'cancel-chart-book', modelRevision: 3,
      ownership: 'transfer', timeoutMs: 20_000, onProgress: value => resolveProgress(value) })
    try {
      await Promise.race([progress, job.promise.then(() => { throw new Error('Chart completed before cancellation checkpoint') })])
      job.cancel()
      const cancelled = await job.promise.then(() => { throw new Error('Cancelled chart returned a result') }, error => errorCode(error))
      const recovered = await client.submit(smallChartTask(), { documentId: 'cancel-chart-book', modelRevision: 3,
        ownership: 'transfer', timeoutMs: 10_000 }).promise
      if (recovered.output.kind !== 'chartStatistics') throw new Error(`Unexpected chart recovery: ${recovered.output.kind}`)
      return { cancelled, recoveredChart: Array.from(recovered.output.values) }
    } finally { client.dispose() }
  },

  async runChartStaleRevision() {
    const client = new OfficeComputeClient()
    let resolveProgress!: (progress: ComputeProgress) => void
    const progress = new Promise<ComputeProgress>(resolve => { resolveProgress = resolve })
    const job = client.submit(largeChartTask(800_000), { documentId: 'stale-chart-book', modelRevision: 5,
      ownership: 'transfer', timeoutMs: 20_000, onProgress: value => resolveProgress(value) })
    try {
      await Promise.race([progress, job.promise.then(() => { throw new Error('Chart completed before revision checkpoint') })])
      client.setDocumentRevision('stale-chart-book', 6)
      const stale = await job.promise.then(() => { throw new Error('Stale chart returned a result') }, error => errorCode(error))
      const recovered = await client.submit(smallChartTask(), { documentId: 'stale-chart-book', modelRevision: 6,
        ownership: 'transfer', timeoutMs: 10_000 }).promise
      if (recovered.output.kind !== 'chartStatistics') throw new Error(`Unexpected current chart: ${recovered.output.kind}`)
      return { stale, recoveredChart: Array.from(recovered.output.values) }
    } finally { client.dispose() }
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

function smallChartTask(): ChartStatisticsTask {
  return { kind: 'chartStatistics', values: new Float64Array([4, 7, 9]), states: new Uint8Array([1, 1, 1]),
    categoryIds: new Uint32Array([0, 1, 2]), params: { rowCount: 3, columnCount: 1, categoryCount: 3,
      series: [{ field: 7, seriesPosition: 0, column: 0 }], jobs: [
        { kind: 'histogram', seriesIndex: 0, slot: 2, binning: { mode: 'count', count: 2 } },
        { kind: 'boxplot', seriesIndex: 0, slot: 5, quartileMethod: 'inclusive', whiskerMultiplier: 1.5 },
      ] } }
}

function largeChartTask(rowCount: number): ChartStatisticsTask {
  const values = new Float64Array(rowCount)
  for (let row = 0; row < rowCount; row++) values[row] = (rowCount - row) % 8191
  return { kind: 'chartStatistics', values, states: new Uint8Array(rowCount).fill(1), categoryIds: new Uint32Array(rowCount),
    params: { rowCount, columnCount: 1, categoryCount: 1, series: [{ field: 7, seriesPosition: 0, column: 0 }],
      jobs: Array.from({ length: 8 }, (_, slot) => ({ kind: 'histogram' as const, seriesIndex: 0, slot,
        binning: { mode: 'count' as const, count: 64 } })) } }
}

function errorCode(error: unknown): string {
  return error && typeof error === 'object' && 'code' in error ? String(error.code) : String(error)
}
