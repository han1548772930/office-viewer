import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import type { ChartConfig, ChartData } from '../chart-engine'
import { chartConfigFromDraft, type ChartSettingsDraft, type ChartSettingsSourceSheet } from '../chart-settings-draft'
import { chartSettingsBindingKey, chartSettingsDateAxisError, chartSettingsPreparedData, chartSettingsPreparedStatistics, chartSettingsProvidedValue, projectChartSettings,
  type ChartSettingsProjection, type ChartSettingsProvidedData, type ChartSettingsSourceLoader, type ChartSettingsSourceRevision, type ChartSettingsStatisticsLoader } from '../chart-settings-data'

export interface ChartSettingsSourceInputs {
  active: () => boolean
  draft: () => ChartSettingsDraft | undefined
  sheets?: () => readonly ChartSettingsSourceSheet[] | undefined
  hostSheetId?: () => string | undefined
  provided?: () => ChartData | ChartSettingsProvidedData | undefined
  date1904?: () => boolean | undefined
  sourceRevision?: () => ChartSettingsSourceRevision | undefined
  load?: () => ChartSettingsSourceLoader | undefined
  loadStatistics?: () => ChartSettingsStatisticsLoader | undefined
}

type SettingsReader = { kind: 'source'; load: ChartSettingsSourceLoader } | { kind: 'statistics'; load: ChartSettingsStatisticsLoader }
interface SourceSnapshot {
  key: string
  revision?: ChartSettingsSourceRevision
  reader?: SettingsReader
  provided?: ChartData | ChartSettingsProvidedData
  projection: ChartSettingsProjection
}

/** One temporary source per panel. Model/session changes invalidate it before
 * the next microtask, while paint edits reuse the prepared sample arrays. */
export function useChartSettingsSource(inputs: ChartSettingsSourceInputs) {
  const snapshot = shallowRef<SourceSnapshot>()
  const loading = ref(false), failure = ref('')
  let controller: AbortController | undefined, generation = 0, disposed = false
  const date1904 = (): boolean | undefined => inputs.date1904?.() ?? chartSettingsProvidedValue(inputs.provided?.())?.date1904
  const identity = () => ({ hostSheetId: inputs.hostSheetId?.(), sourceRevision: inputs.sourceRevision?.() })
  const reader = computed<SettingsReader | undefined>(() => {
    if (inputs.draft()?.config.type === 'pareto' || inputs.draft()?.config.type === 'histogram') {
      const load = inputs.loadStatistics?.()
      return load ? { kind: 'statistics', load } : undefined
    }
    const load = inputs.load?.()
    return load ? { kind: 'source', load } : undefined
  })
  const binding = computed<{ key?: string; error?: string }>(() => {
    const draft = inputs.draft()
    if (!draft) return {}
    try { return { key: chartSettingsBindingKey(draft, inputs.hostSheetId?.(), date1904()) } }
    catch (reason) { return { error: errorMessage(reason) } }
  })
  const projection = computed<ChartSettingsProjection>(() => {
    const draft = inputs.draft()
    if (!inputs.active() || !draft) return {}
    if (binding.value.error) return { error: binding.value.error }
    if (!reader.value) {
      const sheet = inputs.sheets?.()?.find(item => item.id === (draft.config.source.sheetId ?? inputs.hostSheetId?.()))
      const projected = projectChartSettings(draft, sheet, inputs.provided?.(), date1904(), identity())
      return draft.config.type === 'pareto' && !projected.error && !projected.statistics?.pareto
        ? { ...projected, needsPreparation: true, error: '尚未取得当前帕累托统计结果，请重新读取数据源。' }
        : draft.config.type === 'histogram' && !projected.error && !projected.statistics?.histogram
          ? { ...projected, needsPreparation: true, error: '尚未取得当前直方图分箱，请重新读取数据源。' } : projected
    }
    if (failure.value) return { error: failure.value, needsPreparation: true }
    const current = snapshot.value
    return current?.key === binding.value.key && current.revision === inputs.sourceRevision?.()
      ? current.projection : { needsPreparation: true }
  })
  const dateAxisError = computed(() => {
    const draft = inputs.draft()
    return draft && inputs.active() ? chartSettingsDateAxisError(draft.config, projection.value.data, date1904(),
      projection.value.prepared ? 'options' : 'source') : undefined
  })

  function abort(): void {
    generation += 1
    controller?.abort()
    controller = undefined
    loading.value = false
  }
  function refresh(force = false): void {
    const draft = inputs.draft(), key = binding.value.key, revision = inputs.sourceRevision?.(), capturedReader = reader.value, provided = inputs.provided?.()
    if (disposed || !inputs.active() || !draft || !key) {
      abort(); snapshot.value = undefined; failure.value = ''
      return
    }
    const current = snapshot.value
    if (!force && current?.key === key && current.revision === revision && sameReader(current.reader, capturedReader) && current.provided === provided
      && !current.projection.error) return
    abort(); failure.value = ''
    const sheet = inputs.sheets?.()?.find(item => item.id === (draft.config.source.sheetId ?? inputs.hostSheetId?.()))
    if (provided && 'kind' in provided && provided.kind === 'prepared-source') {
      const projected = projectChartSettings(draft, sheet, provided, date1904(), identity())
      if (!projected.error && (draft.config.type !== 'pareto' || projected.statistics?.pareto)
        && (draft.config.type !== 'histogram' || projected.statistics?.histogram)) {
        snapshot.value = { key, revision, reader: capturedReader, provided, projection: projected }; return
      }
    }
    if (!capturedReader) {
      snapshot.value = undefined
      return
    }
    let config: ChartConfig
    try { config = chartConfigFromDraft(draft) }
    catch (reason) { snapshot.value = undefined; failure.value = errorMessage(reason); return }
    const preparedDraft = { config, sourceText: draft.sourceText }, capturedIdentity = identity(), capturedDate1904 = date1904()
    const request = ++generation, activeController = new AbortController()
    controller = activeController
    snapshot.value = { key, revision, reader: capturedReader, provided, projection: { needsPreparation: true } }
    loading.value = true
    const currentRequest = (): boolean => !disposed && !activeController.signal.aborted && request === generation && inputs.active()
      && binding.value.key === key && inputs.sourceRevision?.() === revision && sameReader(reader.value, capturedReader) && inputs.provided?.() === provided
    // Synchronous watchers invalidate immediately; a microtask starts only the
    // final source request when one interaction changes several binding fields.
    void Promise.resolve().then(async () => {
      if (!currentRequest()) return
      try {
        let bound: ChartSettingsProvidedData
        if (capturedReader.kind === 'statistics') {
          const statistics = await capturedReader.load(config, { signal: activeController.signal })
          if (!currentRequest()) return
          bound = chartSettingsPreparedStatistics(preparedDraft, statistics, capturedIdentity, capturedDate1904)
        } else {
          const data = await capturedReader.load(config, { signal: activeController.signal })
          if (!currentRequest()) return
          bound = chartSettingsPreparedData(preparedDraft, data, capturedIdentity, capturedDate1904)
        }
        const projected = projectChartSettings(preparedDraft, sheet, bound, capturedDate1904, capturedIdentity)
        if (projected.error) throw new Error(projected.error)
        snapshot.value = { key, revision, reader: capturedReader, provided, projection: projected }
      } catch (reason) { if (currentRequest()) { snapshot.value = undefined; failure.value = errorMessage(reason) } }
      finally { if (currentRequest()) { controller = undefined; loading.value = false } }
    })
  }
  function assertReady(): ChartSettingsProjection {
    if (loading.value) throw new Error('图表数据源仍在读取，请稍候再应用。')
    if (projection.value.error) throw new Error(projection.value.error)
    if (!projection.value.data) throw new Error('尚未取得当前图表数据源，请重新读取数据源。')
    if (dateAxisError.value) throw new RangeError(dateAxisError.value)
    return projection.value
  }

  watch([inputs.active, () => binding.value.key, () => binding.value.error, () => inputs.sourceRevision?.(), () => reader.value?.kind, () => reader.value?.load,
    () => inputs.provided?.(), () => inputs.draft()?.config, () => inputs.sheets?.()], () => refresh(), { immediate: true, flush: 'sync' })
  onScopeDispose(() => { disposed = true; abort(); snapshot.value = undefined; failure.value = '' })
  return { projection, loading, dateAxisError, assertReady, reload: () => refresh(true) }
}

function errorMessage(reason: unknown): string { return reason instanceof Error ? reason.message : String(reason) }
function sameReader(first?: SettingsReader, second?: SettingsReader): boolean { return first?.kind === second?.kind && first?.load === second?.load }
