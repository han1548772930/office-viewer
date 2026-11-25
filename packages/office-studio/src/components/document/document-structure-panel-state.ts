import type { DocumentCalloutConfig, DocumentColumnGroup, DocumentModel } from '../../domain/model'
import type { DocumentQuoteStyle } from '../../features/document/model/document-structured-appearance'

export type DocumentStructurePanelValue =
  | { kind: 'callout'; config: DocumentCalloutConfig; textColor?: string | null }
  | { kind: 'quote'; style: Partial<DocumentQuoteStyle> }
  | { kind: 'columns'; columnCount: number; gap: number; horizontalPadding: number; minHeight: number; layout: DocumentColumnGroup['layout']; responsive: DocumentColumnGroup['responsive']; widthRatios: number[]; minWidths: Array<number | null> }

export interface DocumentStandaloneStructureTarget {
  unitId: string
  source: DocumentModel
  standaloneSnapshot?: string
  standaloneEpoch?: number
  standaloneParentSource?: DocumentModel | string
}

export function documentStandaloneStructureTargetCurrent(target: DocumentStandaloneStructureTarget,
  current: { model: DocumentModel; epoch: number; parentSource?: DocumentModel | string }): boolean {
  return target.unitId === (current.model.id ?? '') && target.source === current.model
    && target.standaloneEpoch === current.epoch && target.standaloneParentSource === current.parentSource
    && target.standaloneSnapshot === JSON.stringify(current.model)
}

