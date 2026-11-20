import { ComputeError } from './errors.ts'
import { isRecord } from './protocol.ts'

export interface WasmAggregate { push(values: Float64Array, states: Uint8Array): void; result(): Float64Array; free(): void }
export interface WasmStepper { step(budget: number): boolean; progress(): number; finish(): Uint32Array; free(): void }
export interface ComputeWasmModule {
  default(): Promise<unknown>
  compute_version(): string
  NumericAggregator: new () => WasmAggregate
  NumericIndexSort: new (values: Float64Array, states: Uint8Array, rows: number, directions: Uint8Array, nullsLast: boolean) => WasmStepper
  NumericIndexFilter: new (values: Float64Array, states: Uint8Array, rows: number, operators: Uint8Array, lower: Float64Array, upper: Float64Array, all: boolean) => WasmStepper
}

function isComputeModule(value: unknown): value is ComputeWasmModule {
  return isRecord(value) && ['default', 'compute_version', 'NumericAggregator', 'NumericIndexSort', 'NumericIndexFilter'].every((key) => typeof value[key] === 'function')
}

export async function loadComputeWasm(moduleUrl: string): Promise<ComputeWasmModule> {
  let module: unknown
  try { module = await import(/* @vite-ignore */ moduleUrl) }
  catch (error) { throw new ComputeError('COMPUTE_WASM_LOAD_FAILED', 'Rust 计算模块加载失败，请先执行 build:compute', String(error)) }
  if (!isComputeModule(module)) throw new ComputeError('COMPUTE_WASM_API_MISSING', 'Rust 计算模块接口不完整')
  try { await module.default() }
  catch (error) { throw new ComputeError('COMPUTE_WASM_LOAD_FAILED', 'Rust 计算 WASM 初始化失败', String(error)) }
  return module
}

