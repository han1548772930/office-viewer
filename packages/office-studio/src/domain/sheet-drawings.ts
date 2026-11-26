import type { DrawingAffineGeometry, DrawingAffineMatrix } from './drawing-affine.ts'

export const SHEET_DRAWING_LEAF_KINDS = ['image', 'chart', 'shape', 'slicer', 'timeline'] as const
export type SheetDrawingLeafKind = typeof SHEET_DRAWING_LEAF_KINDS[number]
export type SheetDrawingKind = SheetDrawingLeafKind | 'group'
export interface SheetDrawingLeafRef { kind: SheetDrawingLeafKind; id: string }
export interface SheetDrawingGroupRef { kind: 'group'; id: string }
export type SheetDrawingRef = SheetDrawingLeafRef | SheetDrawingGroupRef
export type SheetDrawingMatrix = DrawingAffineMatrix
export type SheetDrawingAffineMatrix = DrawingAffineMatrix
export type SheetDrawingAffineGeometry = DrawingAffineGeometry
export type SheetDrawingGeometry = DrawingAffineGeometry
export const SHEET_DRAWING_ALIGN_TYPES = { default: '0', left: '1', center: '2', right: '3', top: '4', middle: '5', bottom: '6', horizon: '7', vertical: '8' } as const
export type SheetDrawingAlignType = typeof SHEET_DRAWING_ALIGN_TYPES[keyof typeof SHEET_DRAWING_ALIGN_TYPES]
/** Fixed Univer Direction values: UP, RIGHT, DOWN, LEFT. */
export type SheetDrawingMoveDirection = 0 | 1 | 2 | 3

/** Members and roots are ordered from back to front. Leaf coordinates remain owned by their models. */
export interface SheetDrawingGroup {
  id: string
  members: SheetDrawingRef[]
  /** World geometry for the group's frame, never a second cache of child geometry. */
  frame: SheetDrawingAffineGeometry
  metadata?: Record<string, unknown>
}

/** Persisted per worksheet. Every known drawing belongs to exactly one root or group container. */
export interface SheetDrawingState {
  version: 1
  revision: number
  rootOrder: SheetDrawingRef[]
  groups: SheetDrawingGroup[]
}

/** Shared shape input for planners; shape authoring and painting remain the owning feature's responsibility. */
export interface SheetDrawingShapeDTO {
  id: string
  sheetId: string
  geometry: SheetDrawingAffineGeometry
  source:
    | { kind: 'preset'; preset: string; adjustments?: Record<string, number> }
    | { kind: 'path'; path: string; viewBox: { left: number; top: number; width: number; height: number } }
  style?: Record<string, unknown>
  text?: { content: string; style?: Record<string, unknown> }
  metadata?: Record<string, unknown>
}
