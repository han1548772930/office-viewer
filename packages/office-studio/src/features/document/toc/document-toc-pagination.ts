import type { DocumentModel } from '../../../domain/model'
import type { DocumentLayout } from '../layout/document-layout'

/** Published by the authoritative Canvas layout path, including pending font
 * work. Layout generations describe computations, not physical page numbers. */
export interface DocumentTocLayoutPublication {
  source: DocumentModel
  layoutGeneration: number
  fontVersion: number
  complete: boolean
  layout?: DocumentLayout
  engine: 'main-thread' | 'worker'
}

export interface DocumentTocLayoutSnapshot extends DocumentTocLayoutPublication {
  unitId: string
  modelRevision: number
  historyGeneration: number
}

export interface DocumentTocLayoutProvider {
  read(): DocumentTocLayoutSnapshot | undefined
  /** Notify completion, invalidation, font changes and host activation changes. */
  subscribe(listener: () => void): () => void
}

export interface DocumentTocPaginationResult {
  unitId: string
  tocId: string
  status: 'settled' | 'limit' | 'canceled' | 'unavailable' | 'rejected'
  iterations: number
  modelRevision: number
  error?: string
}
