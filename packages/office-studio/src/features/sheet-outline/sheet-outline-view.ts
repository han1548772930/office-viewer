import type { OutlineAxis, SheetOutlineState } from '../../domain/model.ts'
import { buildSheetOutlineTree, getSheetOutlineHiddenSpans, validateSheetOutlineState } from './sheet-outline.ts'
import type { SheetOutlineTreeNode } from './sheet-outline-types.ts'

export interface SheetOutlineViewModel extends SheetOutlineTreeNode {
  /** Tree children carry structural metadata; the returned array carries each
   * visible node's viewport anchors and ancestor visibility. */
  children: SheetOutlineTreeNode[]
  /** Adjacent summary cell; null when the boundary has no summary cell. */
  anchor: number | null
  /** Detail boundary remains usable for a whole-axis collapsed group. */
  anchorBoundary: number
  summaryAtStart: boolean
  hiddenByAncestor: boolean
}
export function getSheetOutlineViewModels(state: SheetOutlineState, axis: OutlineAxis, count: number, includeHidden = false): SheetOutlineViewModel[] {
  const roots = buildSheetOutlineTree(state, axis), result: SheetOutlineViewModel[] = []
  getSheetOutlineHiddenSpans(state, axis, count) // Validate this axis against the actual viewport extent.
  const summaryAtStart = axis === 'row' ? state.summaryAbove : state.summaryLeft
  function visit(node: SheetOutlineTreeNode, collapsedAncestors: readonly SheetOutlineTreeNode[]): void {
    const ancestorCollapsed = collapsedAncestors.some(parent => parent.start !== node.start || parent.end !== node.end)
    const rawAnchor = summaryAtStart ? node.start - 1 : node.end + 1
    const view: SheetOutlineViewModel = { ...node, anchor: rawAnchor < 0 || rawAnchor >= count ? null : rawAnchor,
      anchorBoundary: summaryAtStart ? node.start : node.end + 1, summaryAtStart, hiddenByAncestor: ancestorCollapsed }
    if (includeHidden || !ancestorCollapsed) result.push(view)
    const ancestors = node.collapsed ? [...collapsedAncestors, node] : collapsedAncestors
    for (const child of node.children) visit(child, ancestors)
  }
  roots.forEach(root => visit(root, []))
  return result
}
/** Local Office behavior: level 1 folds all; maxDepth+1 shows all details.
 * The fixed plugin instead toggles its deepest button. Keep this difference
 * explicit while preserving the same model and public collapse commands.
 */
export function getSheetOutlineLevelControls(state: SheetOutlineState, axis: OutlineAxis): Array<{ axis: OutlineAxis; level: number; active: boolean }> {
  validateSheetOutlineState(state)
  const nodes: SheetOutlineTreeNode[] = []
  function collect(node: SheetOutlineTreeNode): void { nodes.push(node); node.children.forEach(collect) }
  buildSheetOutlineTree(state, axis).forEach(collect)
  const depth = nodes.reduce((max, node) => Math.max(max, node.depth), 0)
  if (!depth) return []
  const rangeDepths = new Map<string, number>()
  for (const node of nodes) {
    const key = `${node.start}:${node.end}`; rangeDepths.set(key, Math.max(rangeDepths.get(key) ?? 0, node.depth))
  }
  return Array.from({ length: depth + 1 }, (_, index) => ({ axis, level: index + 1,
    active: nodes.every(node => node.collapsed === (rangeDepths.get(`${node.start}:${node.end}`)! >= index + 1)) }))
}
