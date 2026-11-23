import { defineStore } from 'pinia'
import { computed, markRaw, reactive, ref } from 'vue'
import type { HistorySavepoint, UnitHistory } from '../platform/history/history-types.ts'

export type { HistoryEntry, HistorySavepoint, UnitHistory } from '../platform/history/history-types.ts'
export const useHistoryStore = defineStore('content-history', () => {
  const units = reactive<Record<string, UnitHistory>>(Object.create(null) as Record<string, UnitHistory>)
  const activeUnit = ref('')
  const applyingDepth = ref(0)
  const applyingUnits = new Set<string>()
  let revisionCounter = 0
  let generationCounter = 0
  const current = computed(() => units[activeUnit.value])
  const canUndo = computed(() => Boolean(current.value?.undo.length))
  const canRedo = computed(() => Boolean(current.value?.redo.length))
  const isApplying = computed(() => applyingDepth.value > 0)
  function unit(id: string): UnitHistory {
    return units[id] ??= { undo: [], redo: [], revision: ++revisionCounter, savedRevision: revisionCounter, generation: ++generationCounter }
  }
  function activate(id: string): void { unit(id); activeUnit.value = id }
  function begin(id: string): void {
    if (applyingUnits.has(id)) throw new Error('当前文档的历史事务正在提交，请在提交完成后编辑。')
    applyingUnits.add(id)
    applyingDepth.value++
  }
  function end(id: string): void {
    applyingUnits.delete(id)
    applyingDepth.value--
  }
  function execute(id: string, label: string, redo: () => void, undo: () => void): HistorySavepoint {
    const target = unit(id)
    begin(id)
    try {
      redo()
      const after = ++revisionCounter
      target.undo.push(markRaw({ label, undo, redo, before: target.revision, after }))
      if (target.undo.length > 100) target.undo.shift()
      target.redo = []
      target.revision = after
      return { unitId: id, revision: after, generation: target.generation }
    } finally { end(id) }
  }
  /** Only the exact, applied top entry can be changed by a recent-paste option. */
  function canAmend(id: string, savepoint: HistorySavepoint): boolean {
    const target = units[id], entry = target?.undo.at(-1)
    return Boolean(target && entry && !applyingUnits.has(id) && !target.redo.length
      && savepoint?.unitId === id && Number.isSafeInteger(savepoint.revision) && savepoint.revision > 0
      && Number.isSafeInteger(savepoint.generation) && savepoint.generation === target.generation
      && savepoint.revision === target.revision && entry.after === target.revision)
  }
  /** Replace the redo and after revision atomically, retaining the original undo and before revision. */
  function amendWithReceipt(id: string, savepoint: HistorySavepoint, label: string, redo: () => void): HistorySavepoint | undefined {
    if (!canAmend(id, savepoint)) return undefined
    const target = units[id]!, index = target.undo.length - 1, entry = target.undo[index]!
    begin(id)
    try {
      redo()
      const after = ++revisionCounter
      target.undo[index] = markRaw({ ...entry, label, redo, after })
      target.revision = after
      return { unitId: id, revision: after, generation: target.generation }
    } finally { end(id) }
  }
  function amend(id: string, savepoint: HistorySavepoint, label: string, redo: () => void): boolean {
    return amendWithReceipt(id, savepoint, label, redo) !== undefined
  }
  function undo(): string | undefined {
    const id = activeUnit.value
    const target = current.value
    const entry = target?.undo.at(-1)
    if (!target || !entry) return
    begin(id)
    try {
      entry.undo()
      target.undo.pop()
      target.redo.push(entry)
      target.revision = entry.before
    } finally { end(id) }
    return entry.label
  }
  function redo(): string | undefined {
    const id = activeUnit.value
    const target = current.value
    const entry = target?.redo.at(-1)
    if (!target || !entry) return
    begin(id)
    try {
      entry.redo()
      target.redo.pop()
      target.undo.push(entry)
      target.revision = entry.after
    } finally { end(id) }
    return entry.label
  }
  function clear(id: string): void {
    if (applyingUnits.has(id)) throw new Error('当前文档的历史事务正在提交，请在提交完成后编辑。')
    delete units[id]
    unit(id)
  }
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
  return { units, activeUnit, applyingDepth, canUndo, canRedo, isApplying, activate, execute, canAmend, amend, amendWithReceipt, undo, redo, clear, captureSavepoint, markSaved, isDirty }
})
