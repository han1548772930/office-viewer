import type { CellRange } from '../../domain/model.ts'
import { cloneSheetFormatPainterSource, type SheetFormatPainterSource } from './sheet-format-painter.ts'
import { sheetRangeCommandInputRanges, type SheetRangeCommandInput } from './sheet-range-command.ts'

export type SheetFormatPainterMode = 'off' | 'once' | 'continuous'
export const SHEET_FORMAT_PAINTER_COMMANDS = Object.freeze({
  activate: 'sheet.operation.format-painter', continuous: 'sheet.operation.continuous-format-painter',
  cancel: 'sheet.operation.cancel-format-painter', setStatus: 'sheet.operation.set-format-painter',
  once: 'sheet.command.set-once-format-painter', infinite: 'sheet.command.set-infinite-format-painter',
  apply: 'sheet.command.apply-format-painter', confirm: 'sheet.command.confirm-format-painter',
  cancelConfirmation: 'sheet.operation.cancel-format-painter-confirmation',
})
export const SHEET_FORMAT_PAINTER_ALIASES = Object.freeze({
  activate: 'ui.operation.activate-format-painter', continuous: 'ui.operation.continuous-format-painter',
  cancel: 'ui.operation.cancel-format-painter',
})
export const SHEET_FORMAT_PAINTER_COMMAND_IDS: readonly string[] = [...Object.values(SHEET_FORMAT_PAINTER_COMMANDS), ...Object.values(SHEET_FORMAT_PAINTER_ALIASES)]

export interface SheetFormatPainterCommandInput extends SheetRangeCommandInput {
  status?: SheetFormatPainterMode | 0 | 1 | 2
  sessionId?: string
  expectedRevision?: number
  requestId?: string
  /** An explicit API choice; UI confirmation uses the separately registered confirm command. */
  mergeConfirmed?: boolean
}
export interface SheetFormatPainterState {
  mode: SheetFormatPainterMode
  applying: boolean
  unitId?: string
  sourceSheetId?: string
  sourceRange?: CellRange
  sessionId?: string
  confirmationId?: string
}
export interface SheetFormatPainterConfirmationRequest {
  requestId: string
  sessionId: string
  workbookId: string
  sheetId: string
  sourceSheetId: string
  source: SheetFormatPainterSource
  baseRevision: number
  discardedContentCount: number
  input: SheetFormatPainterCommandInput
}
export type SheetFormatPainterAction = { kind: 'activate'; continuous: boolean; legacyToggle?: boolean }
  | { kind: 'cancel' } | { kind: 'apply' } | { kind: 'confirm' } | { kind: 'cancel-confirmation' }
  | { kind: 'set-mode'; mode: SheetFormatPainterMode }

export function sheetFormatPainterCommandInput(params?: unknown): SheetFormatPainterCommandInput {
  if (params === undefined) return {}
  if (!params || typeof params !== 'object' || Array.isArray(params)) throw new TypeError('格式刷参数必须是对象。')
  const input = params as SheetFormatPainterCommandInput
  sheetRangeCommandInputRanges(input)
  for (const id of [input.sessionId, input.requestId]) if (id !== undefined && (typeof id !== 'string' || !id.trim())) throw new TypeError('格式刷会话或请求标识无效。')
  if (input.expectedRevision !== undefined && (!Number.isSafeInteger(input.expectedRevision) || input.expectedRevision < 0)) throw new TypeError('工作簿版本无效。')
  if (input.mergeConfirmed !== undefined && typeof input.mergeConfirmed !== 'boolean') throw new TypeError('格式刷合并确认选项必须是布尔值。')
  if (input.status !== undefined) painterMode(input.status)
  return input
}
function painterMode(status: unknown): SheetFormatPainterMode {
  if (status === 0 || status === 'off') return 'off'
  if (status === 1 || status === 'once') return 'once'
  if (status === 2 || status === 'continuous') return 'continuous'
  throw new TypeError('格式刷状态必须是 off、once、continuous 或 0、1、2。')
}
export function formatPainterActionForCommand(id: string, input: SheetFormatPainterCommandInput = {}): SheetFormatPainterAction | undefined {
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.activate || id === SHEET_FORMAT_PAINTER_COMMANDS.once || id === SHEET_FORMAT_PAINTER_ALIASES.activate) return { kind: 'activate', continuous: false }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.continuous || id === SHEET_FORMAT_PAINTER_ALIASES.continuous) return { kind: 'activate', continuous: true }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.infinite) return { kind: 'activate', continuous: true, legacyToggle: true }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.cancel || id === SHEET_FORMAT_PAINTER_ALIASES.cancel) return { kind: 'cancel' }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.setStatus) return { kind: 'set-mode', mode: painterMode(input.status) }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.apply) return { kind: 'apply' }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.confirm) return { kind: 'confirm' }
  if (id === SHEET_FORMAT_PAINTER_COMMANDS.cancelConfirmation) return { kind: 'cancel-confirmation' }
  return undefined
}

interface CapturedPainterSession { id: string; unitId: string; sheetId: string; source: SheetFormatPainterSource }

/** Ephemeral interaction only. Persisted formatting remains owned by the domain commands and content history. */
export class SheetFormatPainterSession {
  private session: CapturedPainterSession | null = null
  private mode: SheetFormatPainterMode = 'off'
  private applying = false
  private confirmation: SheetFormatPainterConfirmationRequest | null = null
  constructor(private readonly changed: (state: SheetFormatPainterState) => void = () => {}) {}

  getState(): SheetFormatPainterState {
    return { mode: this.mode, applying: this.applying, unitId: this.session?.unitId,
      sourceSheetId: this.session?.sheetId, sourceRange: this.session ? { ...this.session.source.range } : undefined,
      sessionId: this.session?.id, confirmationId: this.confirmation?.requestId }
  }
  getSource(): SheetFormatPainterSource | null { return this.session ? structuredClone(this.session.source) : null }
  getConfirmation(): SheetFormatPainterConfirmationRequest | null { return this.confirmation ? structuredClone(this.confirmation) : null }
  private notify(): void { this.changed(this.getState()) }

  activate(unitId: string, sheetId: string, capture: () => SheetFormatPainterSource | null, continuous = false, legacyToggle = false): boolean {
    if (!unitId.trim() || !sheetId.trim()) throw new TypeError('格式刷工作簿和源工作表标识无效。')
    if (this.session) {
      if (unitId !== this.session.unitId) return false
      if (!continuous || legacyToggle && this.mode === 'continuous') this.cancel()
      else { this.mode = 'continuous'; this.notify() }
      return true
    }
    const source = capture()
    if (!source) return false
    this.session = { id: crypto.randomUUID(), unitId, sheetId, source: cloneSheetFormatPainterSource(source) }
    this.mode = continuous ? 'continuous' : 'once'
    this.notify()
    return true
  }
  setMode(unitId: string, sheetId: string, capture: () => SheetFormatPainterSource | null, mode: SheetFormatPainterMode): boolean {
    if (this.session && unitId !== this.session.unitId) return false
    if (mode === 'off') { this.cancel(); return true }
    if (!this.session && !this.activate(unitId, sheetId, capture, mode === 'continuous')) return false
    this.mode = mode
    this.notify()
    return true
  }
  cancel(): void { this.session = null; this.mode = 'off'; this.confirmation = null; this.notify() }
  cancelConfirmation(requestId?: string): boolean {
    if (!this.confirmation || requestId !== undefined && this.confirmation.requestId !== requestId) return false
    this.confirmation = null
    this.notify()
    return true
  }
  requestConfirmation(request: SheetFormatPainterConfirmationRequest): void {
    if (!this.session || request.sessionId !== this.session.id || request.workbookId !== this.session.unitId) throw new Error('格式刷会话已经失效。')
    this.confirmation = structuredClone(request)
    this.notify()
  }
  apply(unitId: string, perform: (source: SheetFormatPainterSource, state: SheetFormatPainterState) => boolean, sessionId?: string): boolean {
    const current = this.session
    if (!current || this.applying || unitId !== current.unitId || sessionId !== undefined && sessionId !== current.id) return false
    this.applying = true
    try {
      this.notify()
      // Observers may cancel or replace the session during the applying notification.
      if (this.session !== current) return false
      const operationState = this.getState()
      const accepted = perform(current.source, operationState)
      if (accepted && this.session === current) {
        if (this.mode === 'once') this.cancel()
        else { this.confirmation = null; this.notify() }
      }
      return accepted
    } finally { this.applying = false; this.notify() }
  }
}
