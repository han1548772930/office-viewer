import { defineStore } from 'pinia'
import { computed, markRaw, reactive, ref } from 'vue'

export interface HistoryEntry {
  label: string
  undo: () => void
  redo: () => void
  before: number
  after: number
}
export interface UnitHistory { undo: HistoryEntry[]; redo: HistoryEntry[]; revision: number; savedRevision: number }
export const useHistoryStore = defineStore('content-history', () => {
  const units = reactive<Record<string, UnitHistory>>({})
  const activeUnit = ref('')
  let revisionCounter = 0
  const current = computed(() => units[activeUnit.value])
  const canUndo = computed(() => Boolean(current.value?.undo.length))
  const canRedo = computed(() => Boolean(current.value?.redo.length))
  function unit(id: string): UnitHistory {
    return units[id] ??= { undo: [], redo: [], revision: ++revisionCounter, savedRevision: revisionCounter }
  }
  function activate(id: string): void { unit(id); activeUnit.value = id }
  function execute(id: string, label: string, redo: () => void, undo: () => void): void {
    const target = unit(id)
    redo()
    const after = ++revisionCounter
    target.undo.push(markRaw({ label, undo, redo, before: target.revision, after }))
    if (target.undo.length > 100) target.undo.shift()
    target.redo = []
    target.revision = after
  }
  function undo(): string | undefined {
    const target = current.value
    const entry = target?.undo.at(-1)
    if (!target || !entry) return
    entry.undo()
    target.undo.pop()
    target.redo.push(entry)
    target.revision = entry.before
    return entry.label
  }
  function redo(): string | undefined {
    const target = current.value
    const entry = target?.redo.at(-1)
    if (!target || !entry) return
    entry.redo()
    target.redo.pop()
    target.undo.push(entry)
    target.revision = entry.after
    return entry.label
  }
  function clear(id: string): void { delete units[id]; unit(id) }
  function markSaved(id: string): void { const target = unit(id); target.savedRevision = target.revision }
  function isDirty(id: string): boolean { const target = units[id]; return Boolean(target && target.revision !== target.savedRevision) }
  return { units, activeUnit, canUndo, canRedo, activate, execute, undo, redo, clear, markSaved, isDirty }
})
