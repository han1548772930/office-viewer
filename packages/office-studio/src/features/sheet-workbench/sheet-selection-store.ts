import { defineStore } from 'pinia'
import { shallowRef } from 'vue'
import { cloneSheetSelection, type SheetSelectionState } from './sheet-selection.ts'

/** View state belongs to each workbook/sheet and is separate from content snapshots and undo stacks. */
export const useSheetSelectionStore = defineStore('sheet-selections', () => {
  const units = shallowRef<Record<string, Record<string, SheetSelectionState>>>({})
  const tabStarts = shallowRef<Record<string, Record<string, SheetSelectionState>>>({})
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
    setTabStart(unitId, sheetId)
    return true
  }
  function getTabStart(unitId: string, sheetId: string): SheetSelectionState | undefined {
    const sheets = Object.hasOwn(tabStarts.value, unitId) ? tabStarts.value[unitId] : undefined
    const state = sheets && Object.hasOwn(sheets, sheetId) ? sheets[sheetId] : undefined
    return state ? cloneSheetSelection(state) : undefined
  }
  function setTabStart(unitId: string, sheetId: string, state?: SheetSelectionState): void {
    const sheets = Object.hasOwn(tabStarts.value, unitId) ? tabStarts.value[unitId] : undefined
    if (state) tabStarts.value = { ...tabStarts.value, [unitId]: { ...sheets, [sheetId]: cloneSheetSelection(state) } }
    else if (sheets && Object.hasOwn(sheets, sheetId)) {
      const next = { ...sheets }; delete next[sheetId]
      tabStarts.value = { ...tabStarts.value, [unitId]: next }
    }
  }
  function retainSheets(unitId: string, sheetIds: readonly string[]): void {
    const retained = new Set(sheetIds)
    for (const source of [units, tabStarts]) {
      const sheets = Object.hasOwn(source.value, unitId) ? source.value[unitId] : undefined
      if (!sheets || Object.keys(sheets).every(id => retained.has(id))) continue
      source.value = { ...source.value, [unitId]: Object.fromEntries(Object.entries(sheets).filter(([id]) => retained.has(id))) }
    }
  }
  function removeUnit(unitId: string): void {
    for (const source of [units, tabStarts]) {
      if (!Object.hasOwn(source.value, unitId)) continue
      const next = { ...source.value }; delete next[unitId]; source.value = next
    }
  }
  function clear(): void { units.value = {}; tabStarts.value = {} }
  return { units, tabStarts, get, set, getTabStart, setTabStart, retainSheets, removeUnit, clear }
})
