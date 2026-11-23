import type { DocumentModel } from '../../../domain/model.ts'
import { cloneModel } from '../../../domain/serialization.ts'
import type { HistoryEntry } from '../../../platform/history/history-types.ts'
import { runText, runsForTarget, textTargets } from '../model/document-model.ts'
import type { DocumentStoryFieldRefreshPlan } from './document-story-fields.ts'

interface DocumentSnapshotPair { before: DocumentModel; after: DocumentModel }
const documentSnapshots = new WeakMap<() => void, DocumentSnapshotPair>()

/** Callbacks retain one pair of snapshots; removing an entry releases the pair. */
export function createDocumentSnapshotHistory(before: DocumentModel, after: DocumentModel, publish: (model: DocumentModel) => void): Pick<HistoryEntry, 'undo' | 'redo'> {
  const snapshots = { before, after }
  const undo = () => publish(snapshots.before)
  const redo = () => publish(snapshots.after)
  documentSnapshots.set(undo, snapshots)
  documentSnapshots.set(redo, snapshots)
  return { undo, redo }
}

export function applyDocumentFieldRefreshPlan(model: DocumentModel, plan: DocumentStoryFieldRefreshPlan): void {
  if (plan.sectionHeaderFooters !== undefined) model.sectionHeaderFooters = cloneModel(plan.sectionHeaderFooters)
  if (plan.blocks !== undefined) {
    model.blocks = cloneModel(plan.blocks)
    model.wordCount = textTargets(model).reduce((count, target) => count + runText(runsForTarget(model, target)).replace(/\s/gu, '').length, 0)
  }
  if (plan.contentControls !== undefined) model.contentControls = cloneModel(plan.contentControls)
  if (plan.comments !== undefined) model.comments = cloneModel(plan.comments)
  if (Object.hasOwn(plan, 'footnotes')) model.footnotes = plan.footnotes === undefined ? undefined : cloneModel(plan.footnotes)
  if (Object.hasOwn(plan, 'endnotes')) model.endnotes = plan.endnotes === undefined ? undefined : cloneModel(plan.endnotes)
}

/** Prepare every retained undo/redo cache before publishing the live update.
 * History labels, revisions, savepoints and stack lengths remain unchanged.
 */
export function prepareDocumentHistoryCacheRebase(entries: readonly Pick<HistoryEntry, 'undo' | 'redo'>[], refresh: (snapshot: DocumentModel) => DocumentModel): () => void {
  const pairs = new Set<DocumentSnapshotPair>()
  for (const entry of entries) {
    const snapshots = documentSnapshots.get(entry.undo)
    if (!snapshots || snapshots !== documentSnapshots.get(entry.redo)) {
      throw new Error('文档历史未提供可重基的快照，无法自动更新字段缓存。')
    }
    pairs.add(snapshots)
  }
  const prepared = [...pairs].map(snapshots => ({ snapshots,
    before: refresh(cloneModel(snapshots.before)), after: refresh(cloneModel(snapshots.after)),
  }))
  return () => {
    for (const { snapshots, before, after } of prepared) { snapshots.before = before; snapshots.after = after }
  }
}
