import { nextTick, onActivated, onDeactivated, onMounted, onUnmounted, shallowRef, watch } from 'vue'
import type { OfficeStudio } from '../../../api/office-studio'
import type { WorkbookModel } from '../../../domain/model'
import { OfficeComputeClient } from '../../../compute/client.ts'
import { ComputeError, normalizeComputeError } from '../../../compute/errors.ts'
import type { ComputeProgress } from '../../../compute/protocol.ts'
import type { ChartComputedStatistics } from '../chart-derived'
import { ChartStatisticsService, chartStatisticsSignature, type ChartStatisticsRequest, type ChartStatisticsSnapshot } from '../services/chart-statistics-service'
import { assertChartSourcePresent, readChartSourceNotFoundDiagnostic, type ChartSourceNotFoundDiagnostic } from '../chart-source-diagnostic'

export interface ChartStatisticsViewHost {
  studio(): OfficeStudio | undefined
  workbook(): WorkbookModel | undefined
  activeSheetId(): string
  /** Optional current-revision results supplied by an owning host. */
  externalStatistics(): Readonly<Record<string, ChartComputedStatistics>> | undefined
}

/** Large immutable statistics stay shallow. UI deactivation cancels this
 * view's consumers; the Studio can retain a valid shared cache for remounts. */
export function useChartStatistics(host: ChartStatisticsViewHost) {
  const statistics = shallowRef<Readonly<Record<string, ChartComputedStatistics>>>({})
  const snapshots = shallowRef<Readonly<Record<string, ChartStatisticsSnapshot>>>({})
  const pending = shallowRef<Readonly<Record<string, ComputeProgress | true>>>({})
  const errors = shallowRef<Readonly<Record<string, string>>>({})
  const diagnostics = shallowRef<Readonly<Record<string, ChartSourceNotFoundDiagnostic>>>({})
  const requests = new Set<ChartStatisticsRequest>()
  const standaloneRevisions = new Map<string, number>()
  let local: { service: ChartStatisticsService; client: OfficeComputeClient } | undefined
  let connectedStudio: OfficeStudio | undefined, subscription: { dispose(): void } | undefined
  let active = false, generation = 0, scheduled = false

  function cancelRequests(): void {
    generation++
    for (const request of requests) request.cancel()
    requests.clear(); pending.value = {}
  }

  function connectService(studio: OfficeStudio | undefined): ChartStatisticsService {
    if (studio) {
      if (connectedStudio !== studio) {
        subscription?.dispose(); connectedStudio = studio
        subscription = studio.chartStatisticsService.subscribe(event => {
          if (event.documentId !== host.workbook()?.id) return
          cancelRequests(); statistics.value = {}; snapshots.value = {}
          scheduleRefresh()
        })
      }
      return studio.chartStatisticsService
    }
    subscription?.dispose(); subscription = undefined; connectedStudio = undefined
    if (!local) {
      const client = new OfficeComputeClient()
      local = { client, service: new ChartStatisticsService(client) }
    }
    return local.service
  }

  function scheduleRefresh(): void {
    if (!active || scheduled) return
    scheduled = true
    void nextTick().then(() => { scheduled = false; if (active) refresh() })
  }

  function refresh(): void {
    if (!active) return
    cancelRequests()
    const viewGeneration = generation, studio = host.studio(), workbook = host.workbook(), sheetId = host.activeSheetId()
    statistics.value = {}; snapshots.value = {}; errors.value = {}; diagnostics.value = {}
    if (!workbook?.id) return
    const documentId = workbook.id, charts = (workbook.charts ?? []).filter(chart => chart.sheetId === sheetId)
    let service: ChartStatisticsService
    try { service = connectService(studio) }
    catch (error) { errors.value = { host: normalizeComputeError(error).message }; return }
    // Empty sheets still need revision notifications when their first chart is inserted.
    if (!charts.length) return
    const modelRevision = studio ? studio.getWorkbookRevision(documentId) : standaloneRevisions.get(documentId) ?? 0
    if (!studio) service.setDocumentRevision(documentId, modelRevision)
    const external = host.externalStatistics()
    for (const chart of charts) {
      try { assertChartSourcePresent(workbook, chart) }
      catch (error) {
        const failure = normalizeComputeError(error), diagnostic = readChartSourceNotFoundDiagnostic(failure)
        errors.value = { ...errors.value, [chart.id]: failure.message }
        if (diagnostic) diagnostics.value = { ...diagnostics.value, [chart.id]: diagnostic }
        continue
      }
      if (external?.[chart.id]?.preparedData) { statistics.value = { ...statistics.value, [chart.id]: external[chart.id]! }; continue }
      const current = () => active && viewGeneration === generation && host.studio() === studio && host.activeSheetId() === sheetId
        && host.workbook()?.id === documentId && service.isCurrent({ documentId, modelRevision })
      const assertCurrent = () => {
        const latest = host.workbook()?.charts?.find(candidate => candidate.id === chart.id)
        if (!current() || !latest || chartStatisticsSignature(latest) !== chartStatisticsSignature(chart)) throw new ComputeError('COMPUTE_STALE_RESULT', '图表视图或来源已更新')
        assertChartSourcePresent(host.workbook()!, latest)
      }
      const context = { documentId, modelRevision, userId: studio?.currentUserId ?? 'local-user', assertCurrent }
      try {
        const cached = service.getCached(chart, context, workbook)
        if (cached) {
          assertCurrent()
          statistics.value = { ...statistics.value, [chart.id]: cached.statistics }
          snapshots.value = { ...snapshots.value, [chart.id]: cached }
          continue
        }
        pending.value = { ...pending.value, [chart.id]: true }
        const request = service.request(workbook, chart, { ...context, onProgress: progress => {
          if (current()) pending.value = { ...pending.value, [chart.id]: progress }
        } })
        requests.add(request)
        request.promise.then(snapshot => {
          if (!current()) return
          assertCurrent()
          statistics.value = { ...statistics.value, [chart.id]: snapshot.statistics }
          snapshots.value = { ...snapshots.value, [chart.id]: snapshot }
        }).catch(error => {
          if (!current()) return
          const failure = normalizeComputeError(error)
          const diagnostic = readChartSourceNotFoundDiagnostic(failure)
          if (diagnostic) {
            const retained = { ...statistics.value }, retainedSnapshots = { ...snapshots.value }
            delete retained[chart.id]; delete retainedSnapshots[chart.id]
            statistics.value = retained; snapshots.value = retainedSnapshots
            diagnostics.value = { ...diagnostics.value, [chart.id]: diagnostic }
          }
          if (failure.code !== 'COMPUTE_CANCELLED' && failure.code !== 'COMPUTE_STALE_RESULT') errors.value = { ...errors.value, [chart.id]: failure.message }
        }).finally(() => {
          requests.delete(request)
          if (viewGeneration !== generation) return
          const remaining = { ...pending.value }; delete remaining[chart.id]; pending.value = remaining
        })
      } catch (error) {
        const failure = normalizeComputeError(error), diagnostic = readChartSourceNotFoundDiagnostic(failure)
        errors.value = { ...errors.value, [chart.id]: failure.message }
        if (diagnostic) diagnostics.value = { ...diagnostics.value, [chart.id]: diagnostic }
        const remaining = { ...pending.value }; delete remaining[chart.id]; pending.value = remaining
      }
    }
  }

  // Studio revisions arrive through its service. Only standalone callers need
  // to observe their external model; the production path never traverses all cells.
  watch(() => host.studio() ? undefined : host.workbook(), workbook => {
    if (workbook?.id) standaloneRevisions.set(workbook.id, (standaloneRevisions.get(workbook.id) ?? 0) + 1)
    scheduleRefresh()
  }, { deep: true, flush: 'sync' })
  watch([host.studio, () => host.workbook()?.id, host.activeSheetId, host.externalStatistics], scheduleRefresh, { flush: 'post' })

  function activate(): void { if (!active) { active = true; scheduleRefresh() } }
  function deactivate(): void {
    active = false; cancelRequests(); statistics.value = {}; snapshots.value = {}; errors.value = {}; diagnostics.value = {}
    subscription?.dispose(); subscription = undefined; connectedStudio = undefined
    local?.service.dispose(); local?.client.dispose(); local = undefined
  }
  onMounted(activate); onActivated(activate); onDeactivated(deactivate); onUnmounted(deactivate)
  return { statistics, snapshots, pending, errors, diagnostics }
}
