import type { SheetDrawingAffineGeometry, SheetDrawingLeafRef, SheetDrawingMatrix, SheetDrawingRef, SheetDrawingState } from '../../domain/sheet-drawings.ts'

/** Geometry is supplied by the leaf owner after projecting its real anchors to worksheet CSS pixels. */
export interface SheetDrawingObject {
  ref: SheetDrawingLeafRef
  sheetId: string
  geometry: SheetDrawingAffineGeometry
  revision?: number
  editable?: boolean
  groupable?: boolean
  transformable?: boolean
  /** Missing media is retained in state; an unavailable selected leaf cannot be transformed or grouped. */
  available?: boolean
}

export interface SheetDrawingSceneInput {
  sheetId: string
  state?: SheetDrawingState
  objects: readonly SheetDrawingObject[]
}
export interface SheetDrawingScene extends SheetDrawingSceneInput { state: SheetDrawingState }

export type SheetDrawingFailureReason =
  | 'invalid-input' | 'invalid-state' | 'invalid-geometry' | 'unknown-kind' | 'missing-object'
  | 'empty-selection' | 'too-few-objects' | 'not-found' | 'different-parent' | 'overlapping-selection'
  | 'read-only' | 'not-groupable' | 'not-transformable' | 'unavailable-object' | 'id-conflict' | 'at-boundary' | 'unchanged'
  | 'stale-state' | 'stale-object' | 'incomplete-copy' | 'removed-object-present'

export class SheetDrawingError extends Error {
  readonly code: SheetDrawingFailureReason
  readonly ref?: SheetDrawingRef
  constructor(code: SheetDrawingFailureReason, message: string, ref?: SheetDrawingRef) {
    super(message)
    this.name = 'SheetDrawingError'
    this.code = code
    this.ref = ref
  }
}

export interface SheetDrawingGeometryChange {
  ref: SheetDrawingLeafRef
  before: SheetDrawingAffineGeometry
  after: SheetDrawingAffineGeometry
}
export interface SheetDrawingObjectGuard {
  ref: SheetDrawingLeafRef
  geometry: SheetDrawingAffineGeometry
  revision?: number
}
export interface SheetDrawingTransaction {
  kind: 'group' | 'ungroup' | 'arrange' | 'transform'
  sheetId: string
  beforeState: SheetDrawingState
  afterState: SheetDrawingState
  geometryChanges: SheetDrawingGeometryChange[]
  expectedObjects: SheetDrawingObjectGuard[]
  selectionBefore: SheetDrawingRef[]
  selectionAfter: SheetDrawingRef[]
  /** Includes displaced siblings, affected groups, and all affected descendants for host permissions. */
  affectedRefs: SheetDrawingRef[]
}
export type SheetDrawingPlanResult =
  | { status: 'ready'; transaction: SheetDrawingTransaction }
  | { status: 'noop'; reason: SheetDrawingFailureReason }
  | { status: 'rejected'; reason: SheetDrawingFailureReason; message: string; ref?: SheetDrawingRef }

export interface SheetDrawingGroupRequest { selection: readonly SheetDrawingRef[]; groupId: string }
export interface SheetDrawingUngroupRequest { selection: readonly SheetDrawingRef[] }
export type SheetDrawingArrangeType = 'front' | 'forward' | 'back' | 'backward'
export type SheetDrawingArrangeRequest = { selection: readonly SheetDrawingRef[] } & (
  | { arrangeType: SheetDrawingArrangeType; zOrder?: never }
  | { arrangeType?: never; zOrder: number }
)
export type SheetDrawingGroupTransformRequest = { groupId: string } & (
  | { frame: SheetDrawingAffineGeometry; delta?: never }
  | { frame?: never; delta: SheetDrawingMatrix }
)
export type SheetDrawingSelectionTransformRequest = { selection: readonly SheetDrawingRef[] } & (
  | { frame: SheetDrawingAffineGeometry; delta?: never }
  | { frame?: never; delta: SheetDrawingMatrix }
)

export interface SheetDrawingCopyMapping {
  source: SheetDrawingRef
  target: SheetDrawingRef
  /** Optional explicit copied group world frame; otherwise it is fitted to the target children. */
  frame?: SheetDrawingAffineGeometry
}
export interface SheetDrawingReconcileChanges {
  /** Only explicit deletions may remove a stale reference; missing media must stay represented as an object. */
  removedRefs?: readonly SheetDrawingRef[]
  /** Copying a group requires mappings for the entire subtree, including all nested group IDs. */
  copies?: readonly SheetDrawingCopyMapping[]
}

export interface SheetDrawingApplyResult {
  scene: SheetDrawingScene
  selection: SheetDrawingRef[]
}
