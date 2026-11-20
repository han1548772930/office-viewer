import type { OfficeKind } from '../domain/model'

export type ExportStage = 'loading' | 'writing'
export interface ExportProgress { stage: ExportStage; message: string }
export interface ExportWarning { code: string; message: string }
export interface ExportOutcome { bytes: Uint8Array; warnings: ExportWarning[] }
export interface ExportRequest {
  requestId: string
  kind: OfficeKind
  data: string
  baseUrl: string
}
export type ExportResponse =
  | { requestId: string; type: 'progress'; progress: ExportProgress }
  | { requestId: string; type: 'success'; bytes: Uint8Array; warnings: string }
  | { requestId: string; type: 'error'; error: { code: string; message: string } }
