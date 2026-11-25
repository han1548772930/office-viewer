import type { DocumentModel } from '../../../domain/model'
import type { DocumentSelection } from '../model/document-model'
import type { DocumentAuxiliaryBlockHost } from '../model/document-block-host'

export const DOCUMENT_LATEX_POPUP_OPERATIONS = {
  open: 'docs-latex.operation.open-formula-popup',
  close: 'docs-latex.operation.close-formula-popup',
  localOpen: 'doc.operation.open-latex',
  localClose: 'doc.operation.close-latex',
} as const

/** Fixed ITextRange plus the document selection manager's ITextRangeParam metadata. */
export interface DocumentLatexPopupTextRange {
  startOffset: number
  endOffset: number
  collapsed?: boolean
  direction?: 'none' | 'backward' | 'forward'
  segmentId?: string
  segmentPage?: number
  isActive?: boolean
  rangeType?: 'TEXT' | 'RECT'
}

export interface DocumentLatexPopupOperationRequest {
  unitId?: string
  rangeId?: string
  initialLatex?: string
  segmentId?: string
  segmentPage?: number
  /** The fixed upstream operation uses inclusive endIndex. */
  startIndex?: number
  endIndex?: number
  selection?: DocumentLatexPopupTextRange
  expectedRevision?: number
  expectedHistoryGeneration?: number
  expectedInputGeneration?: number
}

/** A runtime UI target; never persisted as part of the formula or document. */
export interface DocumentLatexPopupTarget {
  source: DocumentModel
  unitId: string
  selection: DocumentSelection
  host?: DocumentAuxiliaryBlockHost
  rangeId?: string
  initialLatex: string
  modelRevision: number
  historyGeneration: number
  inputGeneration: number
}

const requestFields = new Set(['unitId', 'rangeId', 'initialLatex', 'segmentId', 'segmentPage', 'startIndex', 'endIndex',
  'selection', 'expectedRevision', 'expectedHistoryGeneration', 'expectedInputGeneration'])
const natural = (value: unknown) => typeof value === 'number' && Number.isSafeInteger(value) && value >= 0

export function assertDocumentLatexPopupRequest(value: unknown): asserts value is DocumentLatexPopupOperationRequest {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some(key => !requestFields.has(key))) {
    throw new TypeError('公式编辑器参数无效。')
  }
  for (const key of ['unitId', 'rangeId', 'segmentId', 'initialLatex']) {
    const field: unknown = Reflect.get(value, key)
    if (field !== undefined && (typeof field !== 'string' || (key === 'unitId' || key === 'rangeId') && !field.trim())) {
      throw new TypeError('公式编辑器的标识或源文本无效。')
    }
    if (key === 'initialLatex' && typeof field === 'string' && field.length > 65_536) throw new RangeError('公式源文本超过编辑器预算。')
  }
  for (const key of ['segmentPage', 'startIndex', 'endIndex', 'expectedRevision', 'expectedHistoryGeneration', 'expectedInputGeneration']) {
    const field: unknown = Reflect.get(value, key)
    if (field !== undefined && !natural(field)) throw new TypeError('公式编辑器的位置或版本无效。')
  }
  const start: unknown = Reflect.get(value, 'startIndex'), end: unknown = Reflect.get(value, 'endIndex')
  if ((start === undefined) !== (end === undefined) || typeof start === 'number' && typeof end === 'number'
    && (end < start || !Number.isSafeInteger(end + 1))) throw new RangeError('公式编辑器需要完整的 inclusive 范围。')
  const selection: unknown = Reflect.get(value, 'selection')
  if (selection === undefined) return
  if (!selection || typeof selection !== 'object' || Array.isArray(selection)
    || Object.keys(selection).some(key => !['startOffset', 'endOffset', 'segmentId', 'collapsed', 'direction', 'segmentPage', 'isActive', 'rangeType'].includes(key))) throw new TypeError('公式编辑器的文字选区无效。')
  const from: unknown = Reflect.get(selection, 'startOffset'), to: unknown = Reflect.get(selection, 'endOffset')
  const segment: unknown = Reflect.get(selection, 'segmentId'), collapsed: unknown = Reflect.get(selection, 'collapsed')
  const direction: unknown = Reflect.get(selection, 'direction')
  const page: unknown = Reflect.get(selection, 'segmentPage'), active: unknown = Reflect.get(selection, 'isActive'), rangeType: unknown = Reflect.get(selection, 'rangeType')
  if (!natural(from) || !natural(to) || typeof from !== 'number' || typeof to !== 'number' || to < from
    || segment !== undefined && typeof segment !== 'string' || collapsed !== undefined && (typeof collapsed !== 'boolean' || collapsed !== (from === to))
    || direction !== undefined && direction !== 'none' && direction !== 'backward' && direction !== 'forward'
    || page !== undefined && !natural(page) || active !== undefined && typeof active !== 'boolean'
    || rangeType !== undefined && rangeType !== 'TEXT' && rangeType !== 'RECT') {
    throw new RangeError('公式编辑器的文字选区边界无效。')
  }
  if (typeof start === 'number' && typeof end === 'number' && (from !== start || to !== end + 1)
    || Reflect.get(value, 'segmentId') !== undefined && segment !== undefined && Reflect.get(value, 'segmentId') !== segment
    || Reflect.get(value, 'segmentPage') !== undefined && page !== undefined && Reflect.get(value, 'segmentPage') !== page) {
    throw new RangeError('公式编辑器的显式范围与文字选区不一致。')
  }
}
