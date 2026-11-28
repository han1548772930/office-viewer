import { computed, onScopeDispose, shallowRef, watch } from 'vue'
import { OfficeComputeClient } from '../../compute/client.ts'
import { ComputeError, normalizeComputeError } from '../../compute/errors.ts'
import type { ComputeMetrics } from '../../compute/protocol.ts'
import { filterSheetValidationListOptions, prepareSheetValidationListQuery, querySheetValidationList,
  type PreparedSheetValidationListQuery, type SheetValidationListQueryProgress } from '../../features/sheet-validation/validation-list-query.ts'
import type { SheetValidationListOption } from '../../features/sheet-validation/validation-list-input.ts'

interface ListQueryState {
  owner?: object
  options: readonly SheetValidationListOption[]
  query: string
  active: boolean
  unitId: string
  baseRevision: number
  current(): boolean
}
interface ListQueryOptions {
  state(): ListQueryState
  compute?: Pick<OfficeComputeClient, 'submit'>
  minOptions?: number
  minTextUnits?: number
  memoryBudgetBytes?: number
}

/** One current source and query result, one lazy bounded Worker client. A new
 * query/owner/revision cancels its predecessor; closed/removed hosts release
 * packed text and all derived rows. The owning component disposes the client.
 */
export function useSheetValidationListQuery(options: ListQueryOptions) {
  const view = shallowRef<{ matches: SheetValidationListOption[]; busy: boolean; error: string; errorCode?: string;
    progress?: SheetValidationListQueryProgress; metrics?: ComputeMetrics }>({ matches: [], busy: false, error: '' })
  let sequence = 0, stopped = false, pending: AbortController | undefined, owned: OfficeComputeClient | undefined
  let prepared: PreparedSheetValidationListQuery | undefined, cached: { options: readonly SheetValidationListOption[]; query: string;
    matches: SheetValidationListOption[]; metrics?: ComputeMetrics } | undefined
  const minOptions = options.minOptions ?? 1024, minTextUnits = options.minTextUnits ?? 65_536
  if (![minOptions, minTextUnits].every(value => Number.isSafeInteger(value) && value > 0)) throw new RangeError('列表查询路径阈值必须是正安全整数。')
  function needsWorker(state: ListQueryState): boolean {
    if (state.options.length >= minOptions || state.query.length >= minTextUnits) return true
    let units = 0
    for (const option of state.options) { units += option.label.length; if (units >= minTextUnits) return true }
    return false
  }
  function owns(token: number, controller: AbortController): boolean {
    return !stopped && !controller.signal.aborted && token === sequence && pending === controller
  }
  function current(initial: ListQueryState, token: number, controller: AbortController): boolean {
    if (!owns(token, controller)) return false
    const latest = options.state()
    const valid = owns(token, controller) && latest.active && latest.owner === initial.owner && latest.options === initial.options && latest.query === initial.query
      && latest.unitId === initial.unitId && latest.baseRevision === initial.baseRevision && latest.current()
    // Both the state getter and capability callback can synchronously replace
    // the task, including a change away and back to identical field values.
    return valid && owns(token, controller)
  }
  async function refresh(): Promise<void> {
    const token = ++sequence, previous = pending
    pending = undefined; previous?.abort()
    if (stopped || token !== sequence) return
    const controller = new AbortController()
    pending = controller
    const initial = options.state()
    if (!owns(token, controller)) return
    const active = initial.active && !!initial.owner && initial.current()
    if (!owns(token, controller)) return
    if (!active) {
      pending = undefined
      prepared = undefined; cached = undefined
      view.value = { matches: [], busy: false, error: '' }; return
    }
    const assertCurrent = () => { if (!current(initial, token, controller)) throw new ComputeError('COMPUTE_STALE_RESULT', '列表查询的原目标已更改') }
    const publish = (patch: Partial<typeof view.value>): boolean => {
      if (!current(initial, token, controller)) return false
      view.value = { ...view.value, ...patch }
      return current(initial, token, controller)
    }
    try {
      assertCurrent()
      if (cached?.options === initial.options && cached.query === initial.query) {
        publish({ matches: cached.matches, metrics: cached.metrics, busy: false, error: '', errorCode: undefined, progress: undefined }); return
      }
      if (!needsWorker(initial)) {
        const filtered = filterSheetValidationListOptions(initial.options, initial.query)
        assertCurrent(); prepared = undefined; cached = { options: initial.options, query: initial.query, matches: filtered }
        publish({ matches: filtered, metrics: undefined, busy: false, error: '', errorCode: undefined, progress: undefined }); return
      }
      if (!publish({ matches: [], metrics: undefined, busy: true, error: '', errorCode: undefined, progress: undefined })) return
      const guard = { signal: controller.signal, assertCurrent, memoryBudgetBytes: options.memoryBudgetBytes,
        onProgress: (value: SheetValidationListQueryProgress) => { publish({ progress: value }); assertCurrent() } }
      if (prepared?.options !== initial.options) {
        // Old rows must not remain available if the candidate source changes.
        cached = undefined; prepared = undefined
        const candidate = await prepareSheetValidationListQuery(initial.options, guard)
        assertCurrent(); prepared = candidate
      }
      const source = prepared
      if (!source) throw new ComputeError('COMPUTE_STALE_RESULT', '列表查询的候选来源已释放')
      const compute = options.compute ?? (owned ??= new OfficeComputeClient({ memoryBudgetBytes: options.memoryBudgetBytes }))
      const result = await querySheetValidationList(source, initial.query, { ...guard, compute, unitId: initial.unitId, baseRevision: initial.baseRevision })
      assertCurrent()
      cached = { options: initial.options, query: initial.query, matches: result.options, metrics: result.metrics }
      publish({ matches: result.options, metrics: result.metrics })
    } catch (cause) {
      if (current(initial, token, controller)) {
        const failure = normalizeComputeError(cause)
        if (!current(initial, token, controller)) return
        cached = undefined
        // A watcher can replace this task when the failure is published. The
        // obsolete refresh must resolve rather than throw from its catch path.
        publish({ errorCode: failure.code, error: failure.code === 'COMPUTE_MEMORY_LIMIT' ? '列表查询超过当前内存预算。' : '列表查询未能完成，请重试。', matches: [] })
      }
    } finally {
      if (owns(token, controller)) {
        const valid = current(initial, token, controller)
        // Recheck after calling the capability getter before clearing shared
        // source/cache state. It may have started a successor synchronously.
        if (!owns(token, controller)) return
        pending = undefined
        if (valid) view.value = { ...view.value, busy: false, progress: undefined }
        else { prepared = undefined; cached = undefined; view.value = { matches: [], busy: false, error: '' } }
      }
    }
  }
  watch(() => { const state = options.state()
    return [state.owner, state.options, state.query, state.active, state.unitId, state.baseRevision] as const
  }, () => { void refresh() }, { immediate: true, flush: 'sync' })
  onScopeDispose(() => {
    stopped = true; sequence++; pending?.abort(); pending = undefined; prepared = undefined; cached = undefined
    owned?.dispose(); view.value = { matches: [], busy: false, error: '' }
  })
  return { matches: computed(() => view.value.matches), busy: computed(() => view.value.busy), error: computed(() => view.value.error), errorCode: computed(() => view.value.errorCode),
    progress: computed(() => view.value.progress), metrics: computed(() => view.value.metrics), refresh }
}
