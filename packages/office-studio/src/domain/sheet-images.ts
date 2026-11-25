import type { DrawingAffineGeometry } from './drawing-affine.ts'

/** Serializable media only. Decoded images, object URLs and Canvas handles belong to the render owner. */
export interface SheetImageResource {
  id: string
  name: string
  mimeType: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp' | 'image/webp'
  source: string
  width: number
  height: number
  byteLength: number
  /** Source resolution converted to 96-DPI worksheet CSS pixels. */
  scaleX?: number
  scaleY?: number
  metadata?: Record<string, unknown>
}

export interface SheetCellImage {
  id: string
  resource: SheetImageResource
  altText: string
  title: string
  metadata?: Record<string, unknown>
}

export const SheetDrawingAnchorTypes = { Position: '0', Both: '1', None: '2' } as const
export type SheetDrawingAnchorType = typeof SheetDrawingAnchorTypes[keyof typeof SheetDrawingAnchorTypes]
export interface SheetImageAnchor { row: number; column: number; offsetX: number; offsetY: number }
export interface SheetImageCellMarker { row: number; column: number; rowOffset: number; columnOffset: number }
export interface SheetImageBounds { left: number; top: number; width: number; height: number }
/** Public markers describe Excel's nearest major-axis rectangle, including rotated drawings. */
export type SheetImagePlacement =
  | { kind: '0'; from: SheetImageCellMarker; width: number; height: number }
  | { kind: '1'; from: SheetImageCellMarker; to: SheetImageCellMarker }
  | ({ kind: '2' } & SheetImageBounds)
export type SheetImagePlacementInput = SheetImagePlacement | { kind: SheetDrawingAnchorType; bounds: SheetImageBounds }
/** CSS-pixel offsets around the visible frame, matching Univer srcRect. Negative values outcrop. */
export interface SheetImageSourceRect { top: number; left: number; bottom: number; right: number }

export interface SheetFloatingImage extends SheetCellImage {
  anchor: SheetImageAnchor
  /** Position is the legacy default. Both uses anchor/endAnchor; None uses absolutePosition. */
  anchorType?: SheetDrawingAnchorType
  endAnchor?: SheetImageAnchor
  absolutePosition?: { left: number; top: number }
  /** Fixed unrotated extent for Position/None; last authored crop basis for Both, whose extent is derived from markers. */
  size: { width: number; height: number }
  keepRatio: boolean
  opacity: number
  crop?: SheetImageSourceRect
  rotation?: number
  flipX?: boolean
  flipY?: boolean
  /** Group transforms retain the complete affine geometry, including shear. */
  drawingGeometry?: DrawingAffineGeometry
}

export interface SheetEmbeddedBackgroundImage {
  resource: SheetImageResource
  scaleX: number
  scaleY: number
  metadata?: Record<string, unknown>
}
/** Source-only configuration preserves the synchronous public background API without inventing image dimensions. */
export interface SheetLinkedBackgroundImage {
  resource?: undefined
  source: string
  imageSourceType: 'URL' | 'UUID'
  scaleX: number
  scaleY: number
  metadata?: Record<string, unknown>
}
export type SheetBackgroundImage = SheetEmbeddedBackgroundImage | SheetLinkedBackgroundImage
