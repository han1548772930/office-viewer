import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import type { CellRange, SpreadsheetSheet } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { ApiSheetSelection } from '../../api/events.ts'
import type { HistorySavepoint } from '../../platform/history/history-types.ts'
import type { SheetAutoFillResult } from './sheet-auto-fill.ts'
import type { SheetFillApplyType } from './sheet-fill.ts'

export interface SheetFillOptionsView {
  id: string
  unitId: string
  sheetId: string
  range: CellRange
  applyType: SheetFillApplyType
  supportsSeries: boolean
  userId: string
  contentRevision: number
  savepoint: HistorySavepoint
  hasHistoryEntry: boolean
}
export interface SheetFillSession {
  id: string
  unitId: string
  sheetId: string
  userId: string
  sourceRange: CellRange
  targetRange: CellRange
  originalSheet: SpreadsheetSheet
  applyType: SheetFillApplyType
  supportsSeries: boolean
  savepoint: HistorySavepoint
  hasHistoryEntry: boolean
  contentRevision: number
  beforeSelection?: ApiSheetSelection
}
export interface SheetFillHistoryContext {
  captureSavepoint(unitId: string): HistorySavepoint
  canAmend(unitId: string, savepoint: HistorySavepoint): boolean
}
/** An accepted unchanged fill must never amend an unrelated earlier edit. */
export function sheetFillSessionIsCurrent(session: Pick<SheetFillSession, 'unitId' | 'savepoint' | 'hasHistoryEntry'>,
  history: SheetFillHistoryContext): boolean {
  if (session.hasHistoryEntry) return history.canAmend(session.unitId, session.savepoint)
  const current = history.captureSavepoint(session.unitId)
  return current.unitId === session.savepoint.unitId && current.revision === session.savepoint.revision
    && current.generation === session.savepoint.generation
}
/** Recent fill is transient application state; original contents never become a second authority. */
export const useSheetFillStore = defineStore('sheet-auto-fill', () => {
  const sessions = shallowRef<ReadonlyMap<string, SheetFillSession>>(new Map())
  const results = shallowRef<ReadonlyMap<string, SheetAutoFillResult>>(new Map())
  function getSession(unitId: string): SheetFillSession | undefined {
    const session = sessions.value.get(unitId)
    return session ? cloneModel(session) : undefined
  }
  function getView(unitId: string): SheetFillOptionsView | undefined {
    const session = sessions.value.get(unitId)
    return session ? { id: session.id, unitId: session.unitId, sheetId: session.sheetId, range: { ...session.targetRange },
      applyType: session.applyType, supportsSeries: session.supportsSeries, userId: session.userId,
      contentRevision: session.contentRevision, savepoint: { ...session.savepoint }, hasHistoryEntry: session.hasHistoryEntry } : undefined
  }
  function setSession(session: SheetFillSession): void {
    const next = new Map(sessions.value); next.set(session.unitId, cloneModel(session)); sessions.value = next
  }
  function setResult(unitId: string, result: SheetAutoFillResult): void {
    const next = new Map(results.value); next.set(unitId, cloneModel(result)); results.value = next
  }
  function getResult(unitId: string): SheetAutoFillResult | undefined {
    const result = results.value.get(unitId); return result ? cloneModel(result) : undefined
  }
  function clearSession(unitId: string): void {
    if (!sessions.value.has(unitId)) return
    const next = new Map(sessions.value); next.delete(unitId); sessions.value = next
  }
  function clear(unitId?: string): void {
    if (unitId === undefined) { sessions.value = new Map(); results.value = new Map(); return }
    const next = new Map(sessions.value); next.delete(unitId); sessions.value = next
    const attempts = new Map(results.value); attempts.delete(unitId); results.value = attempts
  }
  return { getSession, getView, setSession, getResult, setResult, clearSession, clear }
})
