import { ComputeError, normalizeComputeError } from './errors.ts'
import { isRecord } from './protocol.ts'
import { assertWasmProtocol, requireWasmMemory } from '../platform/wasm/task-protocol.ts'
import type { WasmTaskControl } from '../platform/wasm/task-protocol.ts'
import type { ChartStatisticsParams } from './chart-protocol.ts'
import type { SlicerMembersParams } from './slicer-member-protocol.ts'
import type { CheckboxCellsParams } from './checkbox-cell-protocol.ts'
import type { ValidationListQueryParams } from './validation-list-query-protocol.ts'

export interface WasmAggregate {
  step(budget: number): boolean
  progress(): number
  result(): unknown
  free(): void
}
export interface WasmIndexResult { copy_indices(): Uint32Array; metadata(): unknown; free(): void }
export interface WasmChartResult { copy_values(): Float64Array; copy_indices(): Uint32Array; metadata(): unknown; free(): void }
export interface WasmSlicerMemberResult { copy_indices(): Uint32Array; copy_has_data(): Uint8Array; metadata(): unknown; free(): void }
export interface WasmCheckboxCellsResult { copy_runs(): Uint32Array; metadata(): unknown; free(): void }
export interface WasmStepper<Result = WasmIndexResult> { step(budget: number): boolean; progress(): number; finish(): Result; free(): void }
export interface ComputeWasmModule {
  default(): Promise<unknown>
  compute_protocol_version(): number
  compute_version(): string
  compute_yield(): Promise<void>
  create_validation_list_query(offsets: Uint32Array, textBytes: Uint8Array, queryBytes: Uint8Array,
    control: { context: WasmTaskControl; params: ValidationListQueryParams }): WasmStepper
  create_checkbox_cells(ranges: Uint32Array, merges: Uint32Array, positions: Uint32Array, actions: Uint8Array,
    control: { context: WasmTaskControl; params: CheckboxCellsParams }): WasmStepper<WasmCheckboxCellsResult>
  create_numeric_aggregate(values: Float64Array, states: Uint8Array, context: WasmTaskControl): WasmAggregate
  create_numeric_sort(values: Float64Array, states: Uint8Array, control: {
    context: WasmTaskControl; rowCount: number; directions: number[]; nullsLast: boolean
  }): WasmStepper
  create_numeric_filter(values: Float64Array, states: Uint8Array, control: {
    context: WasmTaskControl; rowCount: number; operators: number[]; lower: number[]; upper: number[]; matchAll: boolean
  }): WasmStepper
  create_chart_statistics(values: Float64Array, states: Uint8Array, categoryIds: Uint32Array,
    control: { context: WasmTaskControl; params: ChartStatisticsParams }): WasmStepper<WasmChartResult>
  create_slicer_members(values: Float64Array, states: Uint8Array, textIds: Uint32Array, availability: Uint8Array,
    textOffsets: Uint32Array, textBytes: Uint8Array, textRanks: Uint32Array,
    control: { context: WasmTaskControl; params: SlicerMembersParams }): WasmStepper<WasmSlicerMemberResult>
}
export interface LoadedComputeWasm { module: ComputeWasmModule; memory: WebAssembly.Memory }
export function isComputeWasmModule(value: unknown): value is ComputeWasmModule {
  return isRecord(value) && ['default', 'compute_protocol_version', 'compute_version', 'compute_yield',
    'create_numeric_aggregate', 'create_numeric_sort', 'create_numeric_filter', 'create_chart_statistics',
    'create_slicer_members', 'create_checkbox_cells', 'create_validation_list_query'].every((key) => typeof value[key] === 'function')
}
export async function initializeComputeWasm(value: unknown): Promise<LoadedComputeWasm> {
  if (!isComputeWasmModule(value)) throw new ComputeError('COMPUTE_WASM_API_MISSING', 'Rust 计算模块缺少协议 2 接口，请执行 build:compute')
  try {
    const memory = requireWasmMemory(await value.default())
    assertWasmProtocol(value.compute_protocol_version(), '计算模块')
    return { module: value, memory }
  } catch (error) { throw normalizeComputeError(error, 'COMPUTE_WASM_LOAD_FAILED') }
}
export async function loadComputeWasm(moduleUrl: string): Promise<LoadedComputeWasm> {
  let module: unknown
  try { module = await import(/* @vite-ignore */ moduleUrl) }
  catch (error) { throw new ComputeError('COMPUTE_WASM_LOAD_FAILED', 'Rust 计算模块加载失败，请先执行 build:compute', String(error)) }
  return initializeComputeWasm(module)
}
