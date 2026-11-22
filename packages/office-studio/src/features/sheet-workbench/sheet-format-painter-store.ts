import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import { SheetFormatPainterSession, type SheetFormatPainterConfirmationRequest, type SheetFormatPainterMode,
  type SheetFormatPainterState } from './sheet-format-painter-session.ts'
import type { SheetFormatPainterSource } from './sheet-format-painter.ts'

/** Pinia owns the observable mode and source mark; the immutable sample stays outside reactive document models. */
export const useSheetFormatPainterStore = defineStore('sheet-format-painter-session', () => {
  const state = shallowRef<SheetFormatPainterState>({ mode: 'off', applying: false })
  const session = new SheetFormatPainterSession(next => { state.value = next })
  const activate = (unitId: string, sheetId: string, capture: () => SheetFormatPainterSource | null, continuous = false, legacyToggle = false) => session.activate(unitId, sheetId, capture, continuous, legacyToggle)
  const setMode = (unitId: string, sheetId: string, capture: () => SheetFormatPainterSource | null, mode: SheetFormatPainterMode) => session.setMode(unitId, sheetId, capture, mode)
  const apply = (unitId: string, perform: (source: SheetFormatPainterSource, current: SheetFormatPainterState) => boolean, sessionId?: string) => session.apply(unitId, perform, sessionId)
  const cancel = () => session.cancel()
  const cancelConfirmation = (requestId?: string) => session.cancelConfirmation(requestId)
  const requestConfirmation = (request: SheetFormatPainterConfirmationRequest) => session.requestConfirmation(request)
  const getSource = () => session.getSource()
  const getConfirmation = () => session.getConfirmation()
  return { state, activate, setMode, apply, cancel, cancelConfirmation, requestConfirmation, getSource, getConfirmation }
})
