import type { DocumentModel, DocumentColumnGroup } from '../../../domain/model'
import { cloneDocumentSelection } from '../model/document-table-selection'
import { documentInputLocation, sameDocumentInputLocation, cloneDocumentInputSession, type DocumentInputSession } from '../model/document-input-session'
import { sameDocumentBlockHost, type DocumentAuxiliaryBlockHost } from '../model/document-block-host'
import type { DocumentSelection } from '../model/document-model'
import type { DocumentStructureCommand } from '../commands/document-structure-commands'
import type { LayoutContainerDecoration } from '../layout/document-layout-types'
import { documentColumnWidths } from '../layout/document-structural-layout'

export interface DocumentStructuredCanvasTarget {
  source: DocumentModel
  selection: DocumentSelection
  host?: DocumentAuxiliaryBlockHost
  inputSession?: DocumentInputSession
  standaloneSnapshot?: string
}
export interface DocumentStructuredCanvasCommandRequest extends DocumentStructuredCanvasTarget { command: DocumentStructureCommand }
export interface DocumentStructuredCanvasOperationRequest extends DocumentStructuredCanvasTarget { id: string; params?: unknown }

export function captureDocumentStructuredCanvasTarget(source: DocumentModel, selection: DocumentSelection,
  host?: DocumentAuxiliaryBlockHost, inputSession?: DocumentInputSession): DocumentStructuredCanvasTarget {
  return { source, selection: cloneDocumentSelection(selection), ...(host ? { host: { ...host } } : {}),
    ...(inputSession ? { inputSession: cloneDocumentInputSession(inputSession) } : { standaloneSnapshot: JSON.stringify(source) }) }
}

export function documentStructuredCanvasTargetCurrent(target: DocumentStructuredCanvasTarget, current: {
  source: DocumentModel; selection: DocumentSelection; host?: DocumentAuxiliaryBlockHost; session?: DocumentInputSession; requireSession?: boolean; disabled?: boolean
}): boolean {
  if (current.disabled || target.source !== current.source || !sameDocumentBlockHost(target.host, current.host)
    || !sameDocumentInputLocation(documentInputLocation(target.selection, undefined, target.host), documentInputLocation(current.selection, undefined, current.host))) return false
  const captured = target.inputSession, session = current.session
  if (!captured) return !current.requireSession && !session && target.standaloneSnapshot === JSON.stringify(current.source)
  return !!captured.valid && !!session?.valid && captured.generation === session.generation && captured.modelRevision === session.modelRevision
    && sameDocumentInputLocation(captured.location, session.location)
}

export interface DocumentColumnResizeResult { widthRatios: number[]; widths: number[]; guideX: number }
/** Preserve adjacent ratio sum and all other weights; only pointer-up publishes the command. */
export function resizeDocumentColumnBoundary(group: DocumentColumnGroup, container: LayoutContainerDecoration,
  index: number, deltaX: number): DocumentColumnResizeResult | undefined {
  const geometry = container.columnGroup
  if (!geometry || geometry.stacked || !Number.isFinite(deltaX) || !Number.isSafeInteger(index) || index < 0 || index >= group.columns.length - 1
    || geometry.columnIds.length !== group.columns.length || geometry.columnIds.some((id, position) => id !== group.columns[position]?.id)) return undefined
  const widths = geometry.widths, pairWidth = widths[index]! + widths[index + 1]!, pairRatio = group.columns[index]!.widthRatio + group.columns[index + 1]!.widthRatio
  const minimum = (position: number) => Math.max(24, group.columns[position]!.minWidth ?? 0)
  if (pairWidth < minimum(index) + minimum(index + 1)) return undefined
  const desired = Math.max(minimum(index), Math.min(widths[index]! + deltaX, pairWidth - minimum(index + 1)))
  const ratios = group.columns.map(column => column.widthRatio), prefix = widths.slice(0, index).reduce((sum, width) => sum + width + geometry.gap, 0)
  if (Math.abs(deltaX) < 1e-6) return { widthRatios: ratios, widths: [...widths], guideX: container.x + prefix + widths[index]! + geometry.gap / 2 }
  const probe = (left: number) => documentColumnWidths({ ...group, columns: group.columns.map((column, position) =>
    position === index ? { ...column, widthRatio: left } : position === index + 1 ? { ...column, widthRatio: pairRatio - left } : column) }, container.width)
  const epsilon = Math.max(Number.MIN_VALUE, pairRatio * 1e-9)
  let low = epsilon, high = pairRatio - epsilon
  for (let iteration = 0; iteration < 28; iteration++) {
    const middle = (low + high) / 2
    if (probe(middle).widths[index]! < desired) low = middle
    else high = middle
  }
  ratios[index] = (low + high) / 2; ratios[index + 1] = pairRatio - ratios[index]!
  const result = probe(ratios[index]!), preceding = result.widths.slice(0, index).reduce((sum, width) => sum + width + result.gap, 0)
  if (result.widths[index]! + 1e-5 < minimum(index) || result.widths[index + 1]! + 1e-5 < minimum(index + 1)) return undefined
  return { widthRatios: ratios, widths: result.widths, guideX: container.x + preceding + result.widths[index]! + result.gap / 2 }
}
