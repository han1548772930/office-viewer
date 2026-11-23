import { assertWasmProtocol, isWasmRecord, requireWasmMemory, WasmProtocolError } from '../platform/wasm/task-protocol.ts'
import type { WasmTaskControl } from '../platform/wasm/task-protocol.ts'

export interface ParserWasmModule {
  default(): Promise<unknown>
  parser_protocol_version(): number
  parse_office(bytes: Uint8Array, fileName: string, context: WasmTaskControl): Promise<unknown>
}
export interface WriterWasmResult {
  copy_bytes(): Uint8Array
  metadata(): unknown
  free(): void
}
export interface WriterWasmModule {
  default(): Promise<unknown>
  writer_protocol_version(): number
  write_office(kind: string, data: string, context: WasmTaskControl): Promise<WriterWasmResult>
}
export interface LoadedWasmModule<T> { module: T; memory: WebAssembly.Memory }

export function isParserWasmModule(value: unknown): value is ParserWasmModule {
  return isWasmRecord(value) && typeof value.default === 'function'
    && typeof value.parser_protocol_version === 'function' && typeof value.parse_office === 'function'
}
export function isWriterWasmModule(value: unknown): value is WriterWasmModule {
  return isWasmRecord(value) && typeof value.default === 'function'
    && typeof value.writer_protocol_version === 'function' && typeof value.write_office === 'function'
}
export function isWriterWasmResult(value: unknown): value is WriterWasmResult {
  return isWasmRecord(value) && typeof value.copy_bytes === 'function'
    && typeof value.metadata === 'function' && typeof value.free === 'function'
}

async function importWasmModule(moduleUrl: string, domain: string): Promise<unknown> {
  try { return await import(/* @vite-ignore */ moduleUrl) }
  catch (error) { throw new WasmProtocolError('WASM_LOAD_FAILED', domain + ' WASM 加载失败，请先构建匹配的新产物',
    error instanceof Error ? error.message : String(error)) }
}
export async function initializeParserWasm(value: unknown): Promise<LoadedWasmModule<ParserWasmModule>> {
  if (!isParserWasmModule(value)) throw new WasmProtocolError('WASM_API_MISSING', '解析器缺少协议 2 接口，请执行 build:wasm')
  const memory = requireWasmMemory(await value.default())
  assertWasmProtocol(value.parser_protocol_version(), '解析器')
  return { module: value, memory }
}
export async function initializeWriterWasm(value: unknown): Promise<LoadedWasmModule<WriterWasmModule>> {
  if (!isWriterWasmModule(value)) throw new WasmProtocolError('WASM_API_MISSING', '写出器缺少协议 2 接口，请执行 build:writer')
  const memory = requireWasmMemory(await value.default())
  assertWasmProtocol(value.writer_protocol_version(), '写出器')
  return { module: value, memory }
}
export async function loadParserWasm(moduleUrl: string): Promise<LoadedWasmModule<ParserWasmModule>> {
  return initializeParserWasm(await importWasmModule(moduleUrl, '解析器'))
}
export async function loadWriterWasm(moduleUrl: string): Promise<LoadedWasmModule<WriterWasmModule>> {
  return initializeWriterWasm(await importWasmModule(moduleUrl, '写出器'))
}
