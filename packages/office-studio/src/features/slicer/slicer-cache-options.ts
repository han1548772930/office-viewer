import type { SheetSlicerCacheOptions, SlicerCacheModel, SlicerModel, WorkbookModel } from '../../domain/model.ts'
import { normalizeRecordValue, type RecordValue } from '../../domain/record-value.ts'
import { equivalentModelData } from '../../platform/commands/model-equivalence.ts'
import { ensureSlicerCache, getSlicerCache, SlicerConnectionError } from './slicer-connections.ts'
import { createSlicerMemberValueComparator } from './slicer-member-order.ts'

export const SLICER_CACHE_SORT_ORDERS = ['ascending', 'descending'] as const
export const SLICER_CACHE_CROSS_FILTERS = ['none', 'showItemsWithDataAtTop', 'showItemsWithNoData'] as const
const keys = ['sortOrder', 'customListSort', 'crossFilter', 'hideItemsWithNoData'] as const
const record = (input: unknown): input is Record<string, unknown> => !!input && typeof input === 'object' && !Array.isArray(input)

/** Validation is presence-aware: an own undefined clears, while null and unknown properties refuse. */
export function normalizeSlicerCacheOptionsPatch(input: unknown): Partial<SheetSlicerCacheOptions> {
  if (!record(input) || Object.keys(input).some(key => !(keys as readonly string[]).includes(key))) {
    throw new SlicerConnectionError('invalid-target', '切片器共享缓存选项无效。')
  }
  const result: Partial<SheetSlicerCacheOptions> = {}
  for (const key of keys) if (Object.hasOwn(input, key)) {
    const value = input[key]
    if (value !== undefined && (key === 'sortOrder' ? !(SLICER_CACHE_SORT_ORDERS as readonly unknown[]).includes(value)
      : key === 'crossFilter' ? !(SLICER_CACHE_CROSS_FILTERS as readonly unknown[]).includes(value) : typeof value !== 'boolean')) {
      throw new SlicerConnectionError('invalid-target', `切片器共享缓存选项 ${key} 无效。`)
    }
    Object.assign(result, { [key]: value })
  }
  return result
}

/** Authored settings remain optional; effective query defaults are applied separately. */
export function normalizeSlicerCacheOptions(input: unknown): SheetSlicerCacheOptions {
  const patch = normalizeSlicerCacheOptionsPatch(input === undefined ? {} : input), result: SheetSlicerCacheOptions = {}
  for (const key of keys) if (patch[key] !== undefined) Object.assign(result, { [key]: patch[key] })
  return result
}

export function readSlicerCacheOptions(workbook: WorkbookModel, slicer: SlicerModel): SheetSlicerCacheOptions {
  if (slicer.type !== 'slicer') throw new SlicerConnectionError('invalid-target', '时间线不使用切片器缓存排序选项。')
  return normalizeSlicerCacheOptions(getSlicerCache(workbook, slicer).options)
}

/** Mutates only shared author options and revisions. Selection, source, connections and reports stay authoritative. */
export function patchSlicerCacheOptions(workbook: WorkbookModel, slicer: SlicerModel, input: Partial<SheetSlicerCacheOptions>): boolean {
  const patch = normalizeSlicerCacheOptionsPatch(input), previous = readSlicerCacheOptions(workbook, slicer), next = { ...previous }
  for (const key of keys) if (Object.hasOwn(patch, key)) {
    if (patch[key] === undefined) delete next[key]
    else Object.assign(next, { [key]: patch[key] })
  }
  if (equivalentModelData(previous, next)) return false
  const retained = ensureSlicerCache(workbook, slicer), shared: SlicerModel[] = []
  for (const control of workbook.slicers ?? []) {
    try { if (getSlicerCache(workbook, control).id === retained.id) shared.push(control) }
    catch (error) { if (!(error instanceof SlicerConnectionError)) throw error }
  }
  if (Object.keys(next).length) retained.options = next
  else delete retained.options
  retained.revision += 1
  for (const control of shared) control.revision += 1
  return true
}

export function effectiveSlicerCacheOptions(input: SheetSlicerCacheOptions = {}): Required<SheetSlicerCacheOptions> {
  const options = normalizeSlicerCacheOptions(input)
  return { sortOrder: options.sortOrder ?? 'ascending', customListSort: options.customListSort === true,
    crossFilter: options.crossFilter ?? 'showItemsWithNoData', hideItemsWithNoData: options.hideItemsWithNoData === true }
}

export interface SlicerTypedMember { key: string; value: RecordValue; hasData: boolean }

/** Typed kind ordering matches the fixed raw-member comparator; localized calendar lists require explicit true. */
export function createSlicerMemberComparator(input: SheetSlicerCacheOptions = {}, locale = 'zh-CN'):
  (first: Pick<SlicerTypedMember, 'value'>, second: Pick<SlicerTypedMember, 'value'>) => number {
  const options = effectiveSlicerCacheOptions(input)
  return createSlicerMemberValueComparator(options.sortOrder, options.customListSort, locale)
}

/** Apply shared display behavior before window slicing; hasData comes from the existing source predicates. */
export function applySlicerCacheItemOptions<T extends SlicerTypedMember>(members: readonly T[], input: SheetSlicerCacheOptions = {}, locale = 'zh-CN'): T[] {
  const options = effectiveSlicerCacheOptions(input), ordered = members.map(item => ({ ...item,
    value: normalizeRecordValue(item.value), hasData: options.crossFilter === 'none' || item.hasData }))
    .sort(createSlicerMemberComparator(options, locale))
  if (options.crossFilter === 'none') return ordered
  if (options.hideItemsWithNoData) return ordered.filter(item => item.hasData)
  if (options.crossFilter === 'showItemsWithDataAtTop') return [...ordered.filter(item => item.hasData), ...ordered.filter(item => !item.hasData)]
  return ordered
}

/** A cache-level snapshot projection for query/UI clients that already resolved the cache. */
export function slicerCacheMemberOptions(cache: Pick<SlicerCacheModel, 'options'>): SheetSlicerCacheOptions {
  return normalizeSlicerCacheOptions(cache.options)
}
