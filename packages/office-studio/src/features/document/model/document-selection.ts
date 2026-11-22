import type { DocumentModel } from '../../../domain/model'
import {
  collapsedSelection, orderedSelection, runText, runsForTarget, sameTarget, textTargets,
  type DocumentPoint, type DocumentSelection,
} from './document-model'

export function documentSelectionContainsPoint(model: DocumentModel, selection: DocumentSelection, point: DocumentPoint): boolean {
  const targets = textTargets(model)
  const index = targets.findIndex(target => sameTarget(target, point))
  const target = targets[index]
  if (!target || !Number.isSafeInteger(point.offset) || point.offset < 0
    || point.offset > runText(runsForTarget(model, target)).length) return false

  const { start, end } = orderedSelection(model, selection)
  const first = targets.findIndex(target => sameTarget(target, start))
  const last = targets.findIndex(target => sameTarget(target, end))
  if (first < 0 || last < 0 || index < first || index > last) return false
  if (sameTarget(start, end) && start.offset === end.offset) return sameTarget(start, point) && start.offset === point.offset
  return (index > first || point.offset >= start.offset) && (index < last || point.offset < end.offset)
}

export function documentSelectionForContextMenu(model: DocumentModel, selection: DocumentSelection, point: DocumentPoint): DocumentSelection {
  return documentSelectionContainsPoint(model, selection, point) ? selection : collapsedSelection(point)
}
