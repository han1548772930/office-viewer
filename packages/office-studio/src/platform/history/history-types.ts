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

export interface UnitHistory {
  undo: HistoryEntry[]
  redo: HistoryEntry[]
  revision: number
  savedRevision: number
  generation: number
}
