import { ComputeError } from './errors.ts'
import { CELL_STATE, COMPUTE_LIMITS } from './protocol.ts'
import type { NumericColumn } from './protocol.ts'

export type NumericInputValue = number | string | boolean | null | undefined | { error: string }
/** Feature adapters pass raw typed/cached values. No text-to-number coercion;
 * range de-duplication, hidden rows and formula authority stay in the feature. */
export async function packNumericColumn(values: Iterable<NumericInputValue>, length: number, signal?: AbortSignal): Promise<NumericColumn> {
  if (!Number.isSafeInteger(length) || length < 0 || length > COMPUTE_LIMITS.maxElements) throw new ComputeError('COMPUTE_MEMORY_LIMIT', '列长度超出计算预算')
  if (signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '数据准备已取消')
  const column: NumericColumn = { values: new Float64Array(length), states: new Uint8Array(length) }
  let index = 0
  for (const value of values) {
    if (signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '数据准备已取消')
    if (index >= length) throw new ComputeError('COMPUTE_INVALID_INPUT', '迭代值数量超过声明长度')
    if (value === null || value === undefined) column.states[index] = CELL_STATE.empty
    else if (typeof value === 'number') { column.values[index] = value; column.states[index] = CELL_STATE.number }
    else if (typeof value === 'string' || typeof value === 'boolean') column.states[index] = CELL_STATE.other
    else if (typeof value === 'object' && typeof value.error === 'string') column.states[index] = CELL_STATE.error
    else throw new ComputeError('COMPUTE_INVALID_INPUT', '列值类型不合法')
    index += 1
    if (index % 8192 === 0) await new Promise<void>((resolve) => setTimeout(resolve, 0))
  }
  if (signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '数据准备已取消')
  if (index !== length) throw new ComputeError('COMPUTE_INVALID_INPUT', '迭代值数量小于声明长度')
  return column
}

