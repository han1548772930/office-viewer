import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import { cloneSheetSelection, type SheetSelectionState } from './sheet-selection.ts'

/** View state belongs to each workbook/sheet and is separate from content snapshots and undo stacks. */
export const useSheetSelectionStore = defineStore('sheet-selections', () => {
  const units = shallowRef<Record<string, Record<string, SheetSelectionState>>>({})
  function get(unitId: string, sheetId: string): SheetSelectionState | undefined {
    const sheets = Object.hasOwn(units.value, unitId) ? units.value[unitId] : undefined
    const state = sheets && Object.hasOwn(sheets, sheetId) ? sheets[sheetId] : undefined
    return state ? cloneSheetSelection(state) : undefined
  }
  function set(unitId: string, sheetId: string, state: SheetSelectionState): boolean {
    const sheets = Object.hasOwn(units.value, unitId) ? units.value[unitId] : undefined
    const previous = sheets && Object.hasOwn(sheets, sheetId) ? sheets[sheetId] : undefined
    if (previous && JSON.stringify(previous) === JSON.stringify(state)) return false
    units.value = { ...units.value, [unitId]: { ...sheets, [sheetId]: cloneSheetSelection(state) } }
    return true
  }
  function retainSheets(unitId: string, sheetIds: readonly string[]): void {
    const sheets = Object.hasOwn(units.value, unitId) ? units.value[unitId] : undefined
    if (!sheets) return
    const retained = new Set(sheetIds)
    if (Object.keys(sheets).every(id => retained.has(id))) return
    units.value = { ...units.value, [unitId]: Object.fromEntries(Object.entries(sheets).filter(([id]) => retained.has(id))) }
  }
  function removeUnit(unitId: string): void {
    if (!Object.hasOwn(units.value, unitId)) return
    const next = { ...units.value }; delete next[unitId]; units.value = next
  }
  function clear(): void { units.value = {} }
  return { units, get, set, retainSheets, removeUnit, clear }
})
