import type { SlicerModel, SpreadsheetCell, WorkbookModel } from '../../domain/model.ts'
import { getRecordValueKey, normalizeRecordValue, type RecordValue } from '../../domain/record-value.ts'
import { sheetCellRecordValue } from '../sheet-workbench/sheet-record-value.ts'
import { compileSheetCellFilterCriteria } from '../sheet-sort-filter/sort-filter-engine.ts'
import { getPivotSourceAvailableRows } from '../pivot/services/pivot-engine.ts'
import { compilePivotRecordSelection } from '../pivot/services/pivot-filter-predicate.ts'
import { getSlicerCache, getSlicerConnectionReports, getSlicerSourceCellRange, getSlicerTableNativeFilter } from './slicer-connections.ts'
import { slicerRecordSelectionFromNativeValues } from './slicer-record-selection.ts'
import { applySlicerCacheItemOptions, effectiveSlicerCacheOptions } from './slicer-cache-options.ts'
import { ComputeError } from '../../compute/errors.ts'
import { COMPUTE_LIMITS } from '../../compute/protocol.ts'
import type { ComputeMetrics } from '../../compute/protocol.ts'
import type { ComputeClientOptions } from '../../compute/client.ts'
import { SlicerMemberInputBuilder, type SlicerMemberInputOptions } from './slicer-member-input.ts'
import { SlicerMemberQueryTransport, type SlicerMemberComputeClient } from './slicer-member-transport.ts'

export interface ISheetSlicerItemsRequest { start?: number; count?: number; signal?: AbortSignal }
export interface ISlicerSourceMemberWindow { total: number; items: Array<{ key: string; value: RecordValue; hasData: boolean }> }
export interface ISlicerSourceMetadata {
  fieldId: string
  name: string
  ready: boolean
  dateGroup?: { type: 3; dateType: 'date' | 'year' | 'quarter' | 'month';
    excelLike?: { unit: 'day' | 'year' | 'quarter' | 'month'; autoStart: boolean; autoEnd: boolean } }
  timeline?: { minDate?: number; maxDate?: number; interactive: boolean }
}
export interface SlicerMemberQueryMetrics {
  unitId: string
  baseRevision: number
  mode: 'items' | 'canvas' | 'metadata'
  sourceRows: number
  sourcePackedBytes: number
  sourcePreparationMs: number
  totalMs: number
  compute: ComputeMetrics
}
export interface SlicerMemberQueryOptions {
  maxMemberWindowSize?: number
  computeClient?: SlicerMemberComputeClient
  computeOptions?: ComputeClientOptions
  memoryBudgetBytes?: number
  maxPendingQueries?: number
  timeoutMs?: number
  onMetrics?: (metrics: SlicerMemberQueryMetrics) => void
}
export interface SlicerMemberQueryContext { unitId: string; baseRevision: number; locale?: string; assertCurrent(): void }
export type SlicerMemberQueryErrorCode = 'SLICER_RPC_ABORTED' | 'SLICER_RPC_DISPOSED' | 'SLICER_RPC_UNIT_DISPOSED'
  | 'SLICER_SOURCE_INVALIDATED' | 'SLICER_INVALID_MEMBER_WINDOW'

export class SlicerMemberQueryError extends Error {
  constructor(readonly code: SlicerMemberQueryErrorCode, message: string) { super(message); this.name = code === 'SLICER_RPC_ABORTED' ? 'AbortError' : 'SlicerMemberQueryError' }
}

export function normalizeSlicerMemberWindowRequest(request: ISheetSlicerItemsRequest = {}, maximum = 200): { start: number; count: number; signal?: AbortSignal } {
  if (!Number.isSafeInteger(maximum) || maximum <= 0) throw new RangeError('Slicer member window maximum must be a positive integer')
  if (!request || typeof request !== 'object' || Array.isArray(request) || Object.keys(request).some(key => !['start', 'count', 'signal'].includes(key))) {
    throw new SlicerMemberQueryError('SLICER_INVALID_MEMBER_WINDOW', '切片器成员窗口参数无效。')
  }
  const start = request.start === undefined ? 0 : request.start, count = request.count === undefined ? 100 : request.count
  if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(count) || count <= 0) {
    throw new SlicerMemberQueryError('SLICER_INVALID_MEMBER_WINDOW', '成员起始位置须为非负整数，数量须为正整数。')
  }
  const signal = request.signal
  if (signal !== undefined && (!signal || typeof signal.aborted !== 'boolean'
    || typeof signal.addEventListener !== 'function' || typeof signal.removeEventListener !== 'function')) {
    throw new SlicerMemberQueryError('SLICER_INVALID_MEMBER_WINDOW', '成员查询取消信号无效。')
  }
  return { start, count: Math.min(count, maximum), ...(signal ? { signal } : {}) }
}

/** hasData ignores this field's own filter and reuses every other native/report predicate. */
export function slicerSourceRowHasData(workbook: WorkbookModel, slicer: SlicerModel): (row: number) => boolean {
  const cache = getSlicerCache(workbook, slicer), { sheet, column } = getSlicerSourceCellRange(workbook, slicer), source = cache.source
  if (source.kind === 'table') {
    const table = workbook.tables?.find(candidate => candidate.id === source.tableId)
    if (!table) throw new Error('切片器表格来源不存在。')
    const predicates = Object.entries(table.autoFilter?.columns ?? {}).filter(([key]) => Number(key) !== column)
      .map(([key, criteria]) => ({ column: Number(key), matches: compileSheetCellFilterCriteria(criteria) }))
    return row => predicates.every(({ column: sourceColumn, matches }) => matches(sheet.cells[`${row}:${sourceColumn}`]))
  }
  const reports = getSlicerConnectionReports(workbook, slicer)
  if (!reports.length) return () => true
  const available = new Set(reports.flatMap(report => getPivotSourceAvailableRows(workbook, report.pivot, report.field.id)))
  return row => available.has(row)
}

/** Compile the same native/retained predicate once; selected/partial caption
 * aggregation itself belongs to Rust and sees every source row, even hidden.
 */
export function slicerSourceRowSelected(workbook: WorkbookModel, slicer: SlicerModel): (cell: SpreadsheetCell | undefined, value: RecordValue) => boolean {
  const cache = getSlicerCache(workbook, slicer)
  if (cache.source.kind === 'table') {
    const filter = getSlicerTableNativeFilter(workbook, slicer)
    if (filter && filter.kind !== 'values') {
      const matches = compileSheetCellFilterCriteria(filter)
      return cell => matches(cell)
    }
    const selection = filter ? slicerRecordSelectionFromNativeValues(filter) : undefined
    if (!selection) return () => true
    const matches = compilePivotRecordSelection(selection)
    return (cell, value) => matches(cell?.value ?? null, value)
  }
  if (cache.recordSelection !== undefined) {
    const matches = compilePivotRecordSelection(cache.recordSelection)
    return (cell, value) => matches(cell?.value ?? null, value)
  }
  if (!cache.selectedItems) return () => true
  const selected = new Set(cache.selectedItems)
  return cell => selected.has(String(cell?.value ?? ''))
}

export interface SlicerDisplayMember { key: string; value: RecordValue; label: string; hasData: boolean; selected?: boolean; partial?: boolean }
export interface ISlicerDisplayMemberWindow { total: number; items: SlicerDisplayMember[] }

/** Canvas labels and public windows share the typed identities, availability and cache projection. */
export function getSlicerDisplayMembers(workbook: WorkbookModel, slicer: SlicerModel, locale = 'zh-CN'): SlicerDisplayMember[] {
  const cache = getSlicerCache(workbook, slicer), { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer)
  const hasData = slicer.type === 'slicer' && effectiveSlicerCacheOptions(cache.options).crossFilter === 'none'
    ? () => true : slicerSourceRowHasData(workbook, slicer)
  const members = new Map<string, SlicerDisplayMember>()
  for (let row = startRow; row <= endRow; row++) {
    const cell = sheet.cells[`${row}:${column}`]
    const value = sheetCellRecordValue(cell, workbook.date1904)
    if (slicer.type === 'timeline' && value.type !== 'date') continue
    const key = getRecordValueKey(value), existing = members.get(key)
    if (existing) existing.hasData ||= hasData(row)
    else members.set(key, { key, value, label: String(cell?.value ?? ''), hasData: hasData(row) })
  }
  return applySlicerCacheItemOptions([...members.values()], slicer.type === 'slicer' ? cache.options : {}, locale)
}

interface RunningQuery { unitId: string; controller: AbortController; reservedBytes: number }

interface QueryExecution {
  check(): void
  signal: AbortSignal
  startedAt: number
  deadlineAt: number
  reserve(bytes: number): void
}

/** Bounded query ownership is independent of presentation, filters and content history. */
export class SlicerMemberQueryService {
  readonly maxMemberWindowSize: number
  private readonly running = new Set<RunningQuery>()
  private readonly transport: SlicerMemberQueryTransport
  private readonly memoryBudgetBytes: number
  private readonly maxPendingQueries: number
  private readonly timeoutMs: number
  private readonly onMetrics?: (metrics: SlicerMemberQueryMetrics) => void
  private disposed = false

  constructor(options: SlicerMemberQueryOptions = {}) {
    const maximum = options.maxMemberWindowSize ?? 200
    if (!Number.isSafeInteger(maximum) || maximum <= 0) throw new RangeError('Slicer maxMemberWindowSize must be a positive integer')
    this.maxMemberWindowSize = maximum
    this.memoryBudgetBytes = options.memoryBudgetBytes ?? options.computeOptions?.memoryBudgetBytes ?? COMPUTE_LIMITS.defaultMemoryBudget
    this.maxPendingQueries = options.maxPendingQueries ?? COMPUTE_LIMITS.maxPendingTasks
    this.timeoutMs = options.timeoutMs ?? 30_000
    if (!Number.isSafeInteger(this.memoryBudgetBytes) || this.memoryBudgetBytes <= 0 || this.memoryBudgetBytes > 0xffff_ffff
      || !Number.isSafeInteger(this.maxPendingQueries) || this.maxPendingQueries <= 0 || this.maxPendingQueries > COMPUTE_LIMITS.maxPendingTasks
      || !Number.isFinite(this.timeoutMs) || this.timeoutMs <= 0 || this.timeoutMs > 2_147_483_647) {
      throw new RangeError('Slicer member query budget, queue or deadline is invalid')
    }
    this.transport = new SlicerMemberQueryTransport(options.computeClient, { ...options.computeOptions, memoryBudgetBytes: this.memoryBudgetBytes })
    this.onMetrics = options.onMetrics
  }

  async getItems(workbook: WorkbookModel, slicer: SlicerModel, request: ISheetSlicerItemsRequest = {}, context: SlicerMemberQueryContext): Promise<ISlicerSourceMemberWindow> {
    const window = await this.queryMemberWindow(workbook, slicer, request, context)
    return { total: window.total, items: window.items.map(({ key, value, hasData }) => ({ key, value, hasData })) }
  }

  /** Local Canvas adapter only. Preserve the representative source's raw
   * caption without widening the formal defensive typed getItems contract.
   */
  getDisplayMemberWindow(workbook: WorkbookModel, slicer: SlicerModel, request: ISheetSlicerItemsRequest = {}, context: SlicerMemberQueryContext): Promise<ISlicerDisplayMemberWindow> {
    return this.queryMemberWindow(workbook, slicer, request, context, 'canvas')
  }

  private queryMemberWindow(workbook: WorkbookModel, slicer: SlicerModel, request: ISheetSlicerItemsRequest, context: SlicerMemberQueryContext, mode: 'items' | 'canvas' = 'items'): Promise<ISlicerDisplayMemberWindow> {
    const window = normalizeSlicerMemberWindowRequest(request, this.maxMemberWindowSize)
    return this.run(context, window.signal, async execution => {
      const { check } = execution
      await queryYield(check)
      const cache = getSlicerCache(workbook, slicer), options = slicer.type === 'slicer' ? cache.options : {}
      const { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer)
      const effective = effectiveSlicerCacheOptions(options)
      const hasData = effective.crossFilter === 'none' ? () => true : slicerSourceRowHasData(workbook, slicer)
      const selected = mode === 'canvas' ? slicerSourceRowSelected(workbook, slicer) : undefined
      check()
      const input: SlicerMemberInputOptions = { mode, start: window.start,
        count: Math.min(window.count, COMPUTE_LIMITS.maxRows), locale: context.locale ?? 'zh-CN', ...effective }
      const builder = new SlicerMemberInputBuilder(Math.max(0, endRow - startRow + 1), input, this.memoryBudgetBytes, execution.reserve)
      for (let row = startRow; row <= endRow; row++) {
        const cell = sheet.cells[`${row}:${column}`], value = sheetCellRecordValue(cell, workbook.date1904)
        builder.append(value, hasData(row), selected ? { caption: String(cell?.value ?? ''), selected: selected(cell, value) } : undefined)
        if ((row - startRow + 1) % 2048 === 0) await queryYield(check)
      }
      const task = await builder.finish(() => queryYield(check)), sourcePreparationMs = performance.now() - execution.startedAt
      const sourcePackedBytes = task.values.byteLength + task.states.byteLength + task.textIds.byteLength + task.availability.byteLength
        + task.textOffsets.byteLength + task.textBytes.byteLength + task.textRanks.byteLength
      const result = await this.transport.query(task, { unitId: context.unitId, baseRevision: context.baseRevision,
        signal: execution.signal, timeoutMs: Math.max(1, execution.deadlineAt - performance.now()) })
      check()
      const items = Array.from(result.output.indices, (offset, index) => {
        const cell = sheet.cells[`${startRow + offset}:${column}`]
        const value = normalizeRecordValue(sheetCellRecordValue(cell, workbook.date1904))
        const flags = result.output.hasData[index]
        return { key: getRecordValueKey(value), value, label: String(cell?.value ?? ''), hasData: (flags & 1) === 1,
          ...(mode === 'canvas' ? { selected: (flags & 2) !== 0, partial: (flags & 4) !== 0 } : {}) }
      })
      this.onMetrics?.({ unitId: context.unitId, baseRevision: context.baseRevision, mode, sourceRows: task.params.rowCount,
        sourcePackedBytes, sourcePreparationMs, totalMs: performance.now() - execution.startedAt, compute: result.metrics })
      check()
      return { total: result.output.total, items }
    })
  }

  async getSourceMetadata(workbook: WorkbookModel, slicer: SlicerModel, context: SlicerMemberQueryContext, signal?: AbortSignal): Promise<ISlicerSourceMetadata> {
    return this.run(context, signal, async execution => {
      const { check } = execution
      await queryYield(check)
      const cache = getSlicerCache(workbook, slicer), { sheet, column, startRow, endRow } = getSlicerSourceCellRange(workbook, slicer)
      const source = cache.source, fieldId = source.kind === 'table' ? source.columnId : source.sourceFieldId
      const table = source.kind === 'table' ? workbook.tables?.find(candidate => candidate.id === source.tableId) : undefined
      const name = table?.columns.find(candidate => candidate.id === fieldId)?.name ?? String(sheet.cells[`${startRow - 1}:${column}`]?.value ?? fieldId)
      const groupedReport = source.kind === 'pivot' ? getSlicerConnectionReports(workbook, slicer).find(report => report.field.dateGroupType !== undefined) : undefined
      const dateType = groupedReport?.field.dateGroupType
      const dateGroup: ISlicerSourceMetadata['dateGroup'] = dateType ? { type: 3, dateType: dateType === 'day' ? 'date' : dateType,
        ...(groupedReport?.pivot.dateGroupingMode === 'excelLike' ? { excelLike: { unit: dateType, autoStart: true, autoEnd: true } } : {}) } : undefined
      const input: SlicerMemberInputOptions = { mode: 'metadata', start: 0, count: 1, locale: context.locale ?? 'zh-CN',
        ...effectiveSlicerCacheOptions({}) }
      const builder = new SlicerMemberInputBuilder(Math.max(0, endRow - startRow + 1), input, this.memoryBudgetBytes, execution.reserve)
      for (let row = startRow; row <= endRow; row++) {
        builder.append(sheetCellRecordValue(sheet.cells[`${row}:${column}`], workbook.date1904))
        if ((row - startRow + 1) % 2048 === 0) await queryYield(check)
      }
      const task = await builder.finish(() => queryYield(check)), sourcePreparationMs = performance.now() - execution.startedAt
      const sourcePackedBytes = task.values.byteLength + task.states.byteLength + task.textIds.byteLength + task.availability.byteLength
        + task.textOffsets.byteLength + task.textBytes.byteLength + task.textRanks.byteLength
      const result = await this.transport.query(task, { unitId: context.unitId, baseRevision: context.baseRevision,
        signal: execution.signal, timeoutMs: Math.max(1, execution.deadlineAt - performance.now()) })
      check()
      const { minDate, maxDate } = result.output
      this.onMetrics?.({ unitId: context.unitId, baseRevision: context.baseRevision, mode: 'metadata', sourceRows: task.params.rowCount,
        sourcePackedBytes, sourcePreparationMs, totalMs: performance.now() - execution.startedAt, compute: result.metrics })
      check()
      const timeline: ISlicerSourceMetadata['timeline'] = minDate !== null && maxDate !== null
        ? { minDate, maxDate, interactive: true } : slicer.type === 'timeline' ? { interactive: false } : undefined
      return { fieldId, name, ready: true, ...(dateGroup ? { dateGroup } : {}), ...(timeline ? { timeline } : {}) }
    })
  }

  forgetDocument(unitId: string): void {
    for (const task of this.running) if (task.unitId === unitId) task.controller.abort(new SlicerMemberQueryError('SLICER_RPC_UNIT_DISPOSED', '切片器所属工作簿已关闭。'))
    this.transport.forgetDocument(unitId)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    for (const task of this.running) task.controller.abort(new SlicerMemberQueryError('SLICER_RPC_DISPOSED', '切片器成员查询服务已释放。'))
    this.transport.dispose()
  }

  private async run<T>(context: SlicerMemberQueryContext, signal: AbortSignal | undefined, work: (execution: QueryExecution) => Promise<T>): Promise<T> {
    if (this.disposed) throw new SlicerMemberQueryError('SLICER_RPC_DISPOSED', '切片器成员查询服务已释放。')
    if (!context || typeof context.unitId !== 'string' || !context.unitId.trim() || !Number.isSafeInteger(context.baseRevision)
      || context.baseRevision < 0 || typeof context.assertCurrent !== 'function') {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器查询必须提供 unit 和安全整数 source revision')
    }
    if (signal !== undefined && (!signal || typeof signal.aborted !== 'boolean'
      || typeof signal.addEventListener !== 'function' || typeof signal.removeEventListener !== 'function')) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器查询取消信号无效')
    }
    if (this.running.size >= this.maxPendingQueries) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '切片器查询队列已满')
    const startedAt = performance.now(), deadlineAt = startedAt + this.timeoutMs
    const controller = new AbortController(), task: RunningQuery = { unitId: context.unitId, controller, reservedBytes: 0 }
    const abort = () => controller.abort(new SlicerMemberQueryError('SLICER_RPC_ABORTED', '切片器成员查询已取消。'))
    if (signal?.aborted) abort(); else signal?.addEventListener('abort', abort, { once: true })
    const check = () => {
      if (controller.signal.aborted) throw controller.signal.reason
      if (this.disposed) throw new SlicerMemberQueryError('SLICER_RPC_DISPOSED', '切片器成员查询服务已释放。')
      context.assertCurrent()
    }
    this.running.add(task)
    // Keep permission/source guards live while awaiting a Worker result, not
    // only during packing. Polling is owner-bound and is not a progress timer.
    const invalidation = setInterval(() => {
      if (controller.signal.aborted) return
      try { check() } catch (error) { controller.abort(error) }
    }, 50)
    const deadline = setTimeout(() => controller.abort(new ComputeError('COMPUTE_TIMEOUT', '切片器成员查询超时')), this.timeoutMs)
    const reserve = (bytes: number) => {
      const other = [...this.running].reduce((sum, query) => sum + (query === task ? 0 : query.reservedBytes), 0)
      if (other + bytes > this.memoryBudgetBytes) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '并发切片器来源准备超过服务内存预算')
      task.reservedBytes = bytes
    }
    try {
      check()
      const result = await work({ check, signal: controller.signal, reserve, startedAt, deadlineAt })
      check()
      return result
    } catch (error) {
      check()
      if (error instanceof ComputeError && error.code === 'COMPUTE_STALE_RESULT') {
        throw new SlicerMemberQueryError('SLICER_SOURCE_INVALIDATED', '切片器来源版本已失效。')
      }
      throw error
    } finally {
      clearInterval(invalidation); clearTimeout(deadline)
      signal?.removeEventListener('abort', abort); this.running.delete(task)
    }
  }
}

/** Timer tasks allow cancellation and source invalidation to be observed between source chunks. */
async function queryYield(check: () => void): Promise<void> {
  check()
  await new Promise<void>(resolve => setTimeout(resolve, 0))
  check()
}
