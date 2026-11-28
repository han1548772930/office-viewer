import type { OfficeComputeClient } from '../../compute/client.ts'
import { ComputeError } from '../../compute/errors.ts'
import { COMPUTE_LIMITS, type ComputeMetrics, type ComputeProgress } from '../../compute/protocol.ts'
import type { ValidationListQueryTask } from '../../compute/validation-list-query-protocol.ts'
import { estimateValidationListQueryMemory } from '../../compute/validation-list-query-tasks.ts'
import { readValidationListQueryOutput } from '../../compute/validation-list-query-results.ts'
import type { SheetValidationListOption } from './validation-list-input.ts'

export interface SheetValidationListQueryProgress {
  phase: 'preparing' | 'querying'; processed: number; total: number; ratio: number
}
export interface SheetValidationListQueryGuard {
  signal?: AbortSignal
  assertCurrent(): void
  memoryBudgetBytes?: number
  onProgress?(progress: SheetValidationListQueryProgress): void
}
export interface PreparedSheetValidationListQuery {
  options: readonly SheetValidationListOption[]
  offsets: Uint32Array
  textBytes: Uint8Array
  preparation: { elapsedMs: number; encodedBytes: number; normalizedCodeUnits: number; yieldCount: number }
}
export interface SheetValidationListQueryContext extends SheetValidationListQueryGuard {
  compute: Pick<OfficeComputeClient, 'submit'>
  unitId: string
  baseRevision: number
}
export interface SheetValidationListQueryResult {
  // Listy's public prop requires an Array. This is the sole result array,
  // frozen in place before publication; no UI row/callback array is duplicated.
  options: SheetValidationListOption[]
  total: number
  metrics: ComputeMetrics
  preparation: PreparedSheetValidationListQuery['preparation'] & { queryEncodedBytes: number }
}

function assertCurrent(guard: SheetValidationListQueryGuard): void {
  if (guard.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '列表查询已取消')
  guard.assertCurrent()
}
function budget(guard: SheetValidationListQueryGuard): number {
  const value = guard.memoryBudgetBytes ?? COMPUTE_LIMITS.defaultMemoryBudget
  if (!Number.isSafeInteger(value) || value <= 0 || value > 0xffff_ffff) throw new ComputeError('COMPUTE_INVALID_INPUT', '列表查询内存预算无效')
  return value
}
async function yieldAndCheck(guard: SheetValidationListQueryGuard): Promise<void> {
  assertCurrent(guard)
  await new Promise<void>(resolve => setTimeout(resolve, 0))
  assertCurrent(guard)
}
function encodeUnits(value: string, bytes: Uint8Array, offset: number, start: number, end: number): void {
  for (let index = start; index < end; index++) {
    const unit = value.charCodeAt(index), position = offset + index * 2
    bytes[position] = unit & 255; bytes[position + 1] = unit >>> 8
  }
}

/** Small queries preserve the original UI's case-fold/substring/order contract.
 * Large queries use the same folded strings at the explicit locale boundary,
 * while Rust owns matching and windows. The threshold is provisional until the
 * requested same-dataset benchmark; no performance claim follows from it.
 */
export function filterSheetValidationListOptions(options: readonly SheetValidationListOption[], query: string): SheetValidationListOption[] {
  const folded = query.toLocaleLowerCase()
  const matches = options.filter(option => !folded || option.label.toLocaleLowerCase().includes(folded))
  Object.freeze(matches)
  return matches
}

/** One current source retains its packed dictionary across query changes.
 * Strings/row metadata have conservative admission estimates, separate from
 * measured binary encoding and the Worker's JS/WASM copy/transfer counters.
 */
export async function prepareSheetValidationListQuery(options: readonly SheetValidationListOption[],
  guard: SheetValidationListQueryGuard): Promise<PreparedSheetValidationListQuery> {
  assertCurrent(guard)
  const started = performance.now(), maximum = budget(guard), count = options.length
  if (!Number.isSafeInteger(count) || count > COMPUTE_LIMITS.maxRows) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '列表候选数量超过查询预算')
  const fixedBytes = (count + 1) * 4 + count * 96 + 8 * 1024 * 1024
  if (fixedBytes > maximum) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '列表候选索引超过内存预算')
  const folded: string[] = [], offsets = new Uint32Array(count + 1)
  let textBytes = 0, yieldCount = 0
  for (let index = 0; index < count; index++) {
    if (index % 1024 === 0) assertCurrent(guard)
    const option = options[index]!
    if (typeof option.label !== 'string' || typeof option.value !== 'string') throw new ComputeError('COMPUTE_INVALID_INPUT', '列表候选必须是文本')
    const label = option.label.toLocaleLowerCase()
    textBytes += label.length * 2
    if (!Number.isSafeInteger(textBytes) || textBytes > COMPUTE_LIMITS.maxElements * 2 || fixedBytes + textBytes * 3 > maximum) {
      throw new ComputeError('COMPUTE_MEMORY_LIMIT', '列表候选文字超过内存预算')
    }
    folded.push(label); offsets[index + 1] = textBytes
    if ((index + 1) % 1024 === 0) {
      assertCurrent(guard)
      guard.onProgress?.({ phase: 'preparing', processed: index + 1, total: count * 2, ratio: (index + 1) / Math.max(1, count * 2) })
      yieldCount++; await yieldAndCheck(guard)
    }
  }
  const packed = new Uint8Array(textBytes)
  for (let index = 0; index < count; index++) {
    const label = folded[index]!
    for (let start = 0; start < label.length; start += 32_768) {
      encodeUnits(label, packed, offsets[index]!, start, Math.min(label.length, start + 32_768))
      if (label.length > 32_768) { yieldCount++; await yieldAndCheck(guard) }
    }
    if ((index + 1) % 1024 === 0) {
      guard.onProgress?.({ phase: 'preparing', processed: count + index + 1, total: count * 2, ratio: (count + index + 1) / Math.max(1, count * 2) })
      yieldCount++; await yieldAndCheck(guard)
    }
  }
  assertCurrent(guard)
  guard.onProgress?.({ phase: 'preparing', processed: count * 2, total: count * 2, ratio: 1 })
  assertCurrent(guard)
  return { options, offsets, textBytes: packed, preparation: { elapsedMs: performance.now() - started,
    encodedBytes: packed.byteLength + offsets.byteLength, normalizedCodeUnits: textBytes / 2, yieldCount } }
}

export async function querySheetValidationList(source: PreparedSheetValidationListQuery, query: string,
  context: SheetValidationListQueryContext, start = 0, count = Math.max(1, source.options.length)): Promise<SheetValidationListQueryResult> {
  assertCurrent(context)
  const maximum = budget(context), folded = query.toLocaleLowerCase()
  if (!Number.isSafeInteger(start) || start < 0 || start > 0xffff_ffff || !Number.isSafeInteger(count) || count <= 0 || count > COMPUTE_LIMITS.maxRows) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '列表查询窗口无效')
  }
  if (folded.length > COMPUTE_LIMITS.maxElements) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '列表搜索文字超过查询预算')
  const params = { rowCount: source.options.length, textByteLength: source.textBytes.length, queryByteLength: folded.length * 2, start, count }
  if (estimateValidationListQueryMemory(params).estimatedPeakBytes > maximum) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '列表查询超过内存预算')
  const queryBytes = new Uint8Array(folded.length * 2)
  for (let first = 0; first < folded.length; first += 32_768) {
    encodeUnits(folded, queryBytes, 0, first, Math.min(folded.length, first + 32_768))
    if (folded.length > 32_768) await yieldAndCheck(context)
  }
  const task: ValidationListQueryTask = { kind: 'validationListQuery', offsets: source.offsets, textBytes: source.textBytes, queryBytes, params }
  // The packed source remains owned here. The client performs/counts one
  // independent input clone per buffer before transfer; WASM then counts its
  // necessary inbound copies and one independent output-window copy.
  const job = context.compute.submit(task, { documentId: context.unitId, modelRevision: context.baseRevision, ownership: 'copy',
    priority: 'interactive', signal: context.signal, onProgress: (progress: ComputeProgress) => {
      assertCurrent(context)
      context.onProgress?.({ phase: 'querying', processed: progress.processed, total: progress.total, ratio: progress.ratio })
    } })
  const result = await job.promise
  assertCurrent(context)
  if (result.documentId !== context.unitId || result.modelRevision !== context.baseRevision || result.output.kind !== 'validationListQuery') {
    throw new ComputeError('COMPUTE_STALE_RESULT', '列表查询结果与当前目标不一致')
  }
  const output = readValidationListQueryOutput(result.output, params)
  if (!output) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '列表查询没有返回完整、独立的候选窗口')
  const options: SheetValidationListOption[] = [], indices = output.indices
  let previous = -1
  for (let index = 0; index < indices.length; index++) {
    const position = indices[index]!, option = source.options[position]
    if (!option || position <= previous) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '列表查询返回了无效或重复候选索引')
    previous = position; options.push(option)
    if ((index + 1) % 2048 === 0) await yieldAndCheck(context)
  }
  assertCurrent(context)
  Object.freeze(options)
  return { options, total: output.total, metrics: result.metrics,
    preparation: { ...source.preparation, queryEncodedBytes: queryBytes.byteLength } }
}
