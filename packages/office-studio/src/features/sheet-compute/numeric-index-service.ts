import { OfficeComputeClient } from '../../compute/client.ts'
import { ComputeError } from '../../compute/errors.ts'
import { CELL_STATE, COMPUTE_LIMITS, FILTER_OPERATOR } from '../../compute/protocol.ts'
import type { ComputeMetrics, ComputeProgress, FilterTask, NumericColumn, SortTask } from '../../compute/protocol.ts'
import type { NumericInputValue } from '../../compute/adapters.ts'

export type NumericFilterOperator = keyof typeof FILTER_OPERATOR

export interface NumericIndexContext {
  documentId: string
  modelRevision: number
  signal?: AbortSignal
  onProgress?: (progress: ComputeProgress) => void
  timeoutMs?: number
}

export interface NumericSortInput {
  rowCount: number
  columns: readonly (readonly NumericInputValue[])[]
  directions?: readonly ('ascending' | 'descending')[]
  nullsLast?: boolean
}

export interface NumericFilterPredicate {
  operator: NumericFilterOperator
  lower?: number
  upper?: number
}

export interface NumericFilterInput {
  rowCount: number
  columns: readonly (readonly NumericInputValue[])[]
  predicates: readonly NumericFilterPredicate[]
  match?: 'all' | 'any'
}

export interface NumericIndexResult {
  documentId: string
  modelRevision: number
  indices: Uint32Array
  metrics: ComputeMetrics
}

export interface NumericIndexServiceOptions {
  /** Routing gate only; it is not a performance claim or measured break-even. */
  minRows?: number
}

const VALID_OPERATORS = new Set<string>(Object.keys(FILTER_OPERATOR))

/** Encodes and dispatches numeric-only table index work to the Rust/WASM Worker.
 * Returns null for small inputs or unsupported sort keys so the owning feature
 * can use its authoritative TypeScript comparator without changing semantics.
 */
export class NumericIndexService {
  private readonly client: OfficeComputeClient
  private readonly minRows: number

  constructor(client: OfficeComputeClient, options: NumericIndexServiceOptions = {}) {
    this.client = client
    this.minRows = options.minRows ?? 4096
    if (!Number.isSafeInteger(this.minRows) || this.minRows < 0 || this.minRows > COMPUTE_LIMITS.maxRows) {
      throw new RangeError('minRows must be a non-negative supported row count')
    }
  }

  setDocumentRevision(documentId: string, modelRevision: number): void {
    this.client.setDocumentRevision(documentId, modelRevision)
  }

  isCurrent(context: Pick<NumericIndexContext, 'documentId' | 'modelRevision'>): boolean {
    return this.client.isCurrent(context)
  }

  sortRows(input: NumericSortInput, context: NumericIndexContext): Promise<NumericIndexResult | null> {
    assertContext(context)
    assertShape(input.rowCount, input.columns)
    const directions = input.directions ?? input.columns.map(() => 'ascending' as const)
    if (directions.length !== input.columns.length || directions.some((direction) => direction !== 'ascending' && direction !== 'descending')) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', 'Sort direction count must match the numeric columns')
    }
    if (!shouldDispatch(input.rowCount, input.columns.length, this.minRows)) return Promise.resolve(null)
    if (input.columns.some((column) => column.some((value) => typeof value === 'string' || typeof value === 'boolean'))) return Promise.resolve(null)

    return this.packColumns(input.columns, input.rowCount, context.signal).then((packed) => {
      const task: SortTask = {
        kind: 'sort', ...packed, rowCount: input.rowCount,
        directions: Uint8Array.from(directions, (direction) => direction === 'descending' ? 1 : 0),
        nullsLast: input.nullsLast ?? true,
      }
      return this.submit(task, context)
    })
  }

  filterRows(input: NumericFilterInput, context: NumericIndexContext): Promise<NumericIndexResult | null> {
    assertContext(context)
    assertShape(input.rowCount, input.columns)
    if (input.predicates.length !== input.columns.length || input.predicates.some((predicate) => !predicate || !VALID_OPERATORS.has(predicate.operator))) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', 'Each numeric filter column must have one supported predicate')
    }
    if (input.match !== undefined && input.match !== 'all' && input.match !== 'any') throw new ComputeError('COMPUTE_INVALID_INPUT', 'Filter match must be all or any')
    if (!shouldDispatch(input.rowCount, input.columns.length, this.minRows)) return Promise.resolve(null)
    const operators = input.predicates.map((predicate) => FILTER_OPERATOR[predicate.operator])
    const lower = input.predicates.map((predicate) => predicate.lower ?? 0)
    const upper = input.predicates.map((predicate) => predicate.upper ?? 0)
    for (let index = 0; index < input.predicates.length; index += 1) {
      const operator = operators[index]!
      if (operator < FILTER_OPERATOR.isEmpty && !Number.isFinite(lower[index])) throw new ComputeError('COMPUTE_INVALID_INPUT', 'Numeric filter operands must be finite')
      if ((operator === FILTER_OPERATOR.between || operator === FILTER_OPERATOR.notBetween)
        && (!Number.isFinite(upper[index]) || upper[index]! < lower[index]!)) {
        throw new ComputeError('COMPUTE_INVALID_INPUT', 'Between filter upper bound must be finite and no lower than its lower bound')
      }
    }

    return this.packColumns(input.columns, input.rowCount, context.signal).then((packed) => {
      const task: FilterTask = {
        kind: 'filter', ...packed, rowCount: input.rowCount,
        operators: Uint8Array.from(operators), lower: Float64Array.from(lower), upper: Float64Array.from(upper),
        match: input.match ?? 'all',
      }
      return this.submit(task, context)
    })
  }

  private async packColumns(columns: readonly (readonly NumericInputValue[])[], rowCount: number, signal?: AbortSignal): Promise<NumericColumn> {
    const length = rowCount * columns.length
    const values = new Float64Array(length)
    const states = new Uint8Array(length)
    for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
      const column = columns[columnIndex]!
      const offset = columnIndex * rowCount
      for (let row = 0; row < rowCount; row += 1) {
        if ((row & 8191) === 0) {
          if (signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', 'Numeric computation was cancelled while preparing columns')
          if (row > 0) await yieldToEventLoop()
        }
        const value = column[row]
        const index = offset + row
        if (value === null || value === undefined) states[index] = CELL_STATE.empty
        else if (typeof value === 'number') { values[index] = value; states[index] = CELL_STATE.number }
        else if (typeof value === 'string' || typeof value === 'boolean') states[index] = CELL_STATE.other
        else if (typeof value === 'object' && typeof value.error === 'string') states[index] = CELL_STATE.error
        else throw new ComputeError('COMPUTE_INVALID_INPUT', `Invalid numeric cell at row ${row}, column ${columnIndex}`)
      }
    }
    if (signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', 'Numeric computation was cancelled while preparing columns')
    return { values, states }
  }

  private async submit(task: SortTask | FilterTask, context: NumericIndexContext): Promise<NumericIndexResult> {
    const result = await this.client.submit(task, {
      documentId: context.documentId,
      modelRevision: context.modelRevision,
      ownership: 'transfer',
      priority: 'interactive',
      signal: context.signal,
      timeoutMs: context.timeoutMs,
      onProgress: context.onProgress,
    }).promise
    if (!this.client.isCurrent(context)) throw new ComputeError('COMPUTE_STALE_RESULT', 'Numeric result belongs to an older workbook revision')
    if (result.output.kind === 'aggregate') throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', 'Numeric index task returned aggregate output')
    return { documentId: context.documentId, modelRevision: context.modelRevision, indices: result.output.indices, metrics: result.metrics }
  }
}

function assertContext(context: NumericIndexContext): void {
  if (!context || typeof context.documentId !== 'string' || !context.documentId.trim()
    || !Number.isSafeInteger(context.modelRevision) || context.modelRevision < 0) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', 'A workbook id and non-negative model revision are required')
  }
}

function assertShape(rowCount: number, columns: readonly (readonly NumericInputValue[])[]): void {
  if (!Number.isSafeInteger(rowCount) || rowCount < 0 || rowCount > COMPUTE_LIMITS.maxRows
    || !Array.isArray(columns) || columns.length === 0 || columns.length > COMPUTE_LIMITS.maxKeys
    || rowCount * columns.length > COMPUTE_LIMITS.maxElements || columns.some((column) => !Array.isArray(column) || column.length !== rowCount)) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', 'Numeric table must have supported row/key dimensions and equal-length columns')
  }
}

function shouldDispatch(rowCount: number, keyCount: number, minRows: number): boolean {
  return rowCount >= minRows && rowCount > 0 && rowCount * keyCount >= minRows
}

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0))
}
