import type { DrawingAffineMatrix } from '../domain/drawing-affine.ts'
import type { SheetDrawingLeafRef } from '../domain/sheet-drawings.ts'

/** All scene coordinates are CSS pixels. The surface applies device pixel ratio. */
export interface RenderRect { x: number; y: number; width: number; height: number }
/** A drawing owner supplies its content; the worksheet owns the common stacking order. */
export interface DrawingCommandOwner {
  drawingRef?: SheetDrawingLeafRef
  drawingOverlay?: boolean
  /** Maps local command CSS pixels into the scene; clip is always in scene coordinates. */
  matrix?: DrawingAffineMatrix
  /** Applied in local coordinates before matrix. Useful for a rotated or sheared control's content. */
  localClip?: RenderRect
}
export interface RectCommand extends RenderRect, DrawingCommandOwner { kind: 'rect'; fill: string; stroke?: string; strokeWidth?: number; clip?: RenderRect }
export interface LineCommand extends DrawingCommandOwner { kind: 'line'; x1: number; y1: number; x2: number; y2: number; color: string; width?: number; clip?: RenderRect }
export interface PolygonCommand extends DrawingCommandOwner { kind: 'polygon'; points: Array<{ x: number; y: number }>; fill: string; clip?: RenderRect }
export interface TextCommand extends DrawingCommandOwner { kind: 'text'; x: number; y: number; text: string; font: string; color: string; align?: 'left' | 'center' | 'right'; baseline?: 'top' | 'middle' | 'alphabetic' | 'bottom'; maxWidth?: number; rotation?: number; clip?: RenderRect }
export interface ImageCommand extends RenderRect, DrawingCommandOwner {
  kind: 'image'; source: CanvasImageSource; sourceVersion?: number; opacity?: number; clip?: RenderRect
  /** Fractions of the complete source bitmap; out-of-bounds portions leave empty destination space. */
  sourceRect?: RenderRect
  /** Degrees around the center of the destination frame. */
  rotation?: number
  flipX?: boolean
  flipY?: boolean
  /** The inherited matrix is applied after the image's local rotation and flips. */
}
export type DrawCommand = RectCommand | LineCommand | PolygonCommand | TextCommand | ImageCommand
export interface RenderScene { width: number; height: number; background?: string; commands: DrawCommand[] }
export type RenderBackend = 'webgpu' | 'canvas2d'
export interface RenderSurfaceOptions { preferredBackend?: RenderBackend; onBackendChange?: (backend: RenderBackend, reason?: string) => void }
export interface RenderSurface {
  readonly backend: RenderBackend
  readonly canvas: HTMLCanvasElement
  resize(cssWidth: number, cssHeight: number, dpr?: number): void
  render(scene: RenderScene): void
  dispose(): void
}
export interface RenderDriver {
  readonly backend: RenderBackend
  resize(width: number, height: number, dpr: number): void
  render(scene: RenderScene): void
  dispose(): void
}
