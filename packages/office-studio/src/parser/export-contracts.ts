import type { OfficeKind } from '../domain/model'
import type { WasmErrorData, WasmTaskControl, WasmTaskMetrics } from '../platform/wasm/task-protocol.ts'

export type ExportStage = 'loading' | 'writing'
export interface ExportProgress { stage: ExportStage; message: string }
export interface ExportWarning { code: string; message: string }
export interface ExportOutcome {
  bytes: Uint8Array; warnings: ExportWarning[]
  context: WasmTaskControl; metrics: WasmTaskMetrics
}
export interface ExportRequest {
  type: 'export'
  context: WasmTaskControl
  kind: OfficeKind
  data: string
  moduleUrl: string
}
export type ExportResponse =
  | { context: WasmTaskControl; type: 'progress'; progress: ExportProgress }
  | { context: WasmTaskControl; type: 'success'; bytes: Uint8Array; warnings: ExportWarning[]; metrics: WasmTaskMetrics }
  | { context: WasmTaskControl; type: 'error'; error: WasmErrorData }
  | { type: 'protocol-error'; error: WasmErrorData }

export function isExportWarnings(value: unknown): value is ExportWarning[] {
  return Array.isArray(value) && value.every((item) => typeof item === 'object' && item !== null
    && 'code' in item && typeof item.code === 'string' && 'message' in item && typeof item.message === 'string')
}
