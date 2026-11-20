import { defineStore } from 'pinia'
import { computed, markRaw, reactive, ref } from 'vue'

export interface HistoryEntry {
  label: string
  undo: () => void
  redo: () => void
  before: number
  after: number
}
export interface HistorySavepoint {
  readonly unitId: string
  readonly revision: number
  /** Distinguishes a reopened document from an earlier history using the same ID. */
  readonly generation: number
}
export interface UnitHistory { undo: HistoryEntry[]; redo: HistoryEntry[]; revision: number; savedRevision: number; generation: number }
export const useHistoryStore = defineStore('content-history', () => {
  const units = reactive<Record<string, UnitHistory>>(Object.create(null) as Record<string, UnitHistory>)
  const activeUnit = ref('')
  let revisionCounter = 0
  let generationCounter = 0
  const current = computed(() => units[activeUnit.value])
  const canUndo = computed(() => Boolean(current.value?.undo.length))
  const canRedo = computed(() => Boolean(current.value?.redo.length))
  function unit(id: string): UnitHistory {
    return units[id] ??= { undo: [], redo: [], revision: ++revisionCounter, savedRevision: revisionCounter, generation: ++generationCounter }
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
  function captureSavepoint(id: string): HistorySavepoint {
    const target = unit(id)
    return { unitId: id, revision: target.revision, generation: target.generation }
  }
  /** A completed asynchronous save acknowledges its input revision, never later edits. */
  function markSaved(id: string, savepoint?: HistorySavepoint): boolean {
    const target = savepoint ? units[id] : unit(id)
    if (!target || (savepoint && (savepoint.unitId !== id || savepoint.generation !== target.generation
      || !Number.isSafeInteger(savepoint.revision) || savepoint.revision <= 0 || savepoint.revision > revisionCounter))) return false
    target.savedRevision = savepoint?.revision ?? target.revision
    return true
  }
  function isDirty(id: string): boolean { const target = units[id]; return Boolean(target && target.revision !== target.savedRevision) }
  return { units, activeUnit, canUndo, canRedo, activate, execute, undo, redo, clear, captureSavepoint, markSaved, isDirty }
})
