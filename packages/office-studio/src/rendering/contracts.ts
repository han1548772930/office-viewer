/** All scene coordinates are CSS pixels. The surface applies device pixel ratio. */
export interface RenderRect { x: number; y: number; width: number; height: number }
export interface RectCommand extends RenderRect { kind: 'rect'; fill: string; stroke?: string; strokeWidth?: number; clip?: RenderRect }
export interface LineCommand { kind: 'line'; x1: number; y1: number; x2: number; y2: number; color: string; width?: number; clip?: RenderRect }
export interface PolygonCommand { kind: 'polygon'; points: Array<{ x: number; y: number }>; fill: string; clip?: RenderRect }
export interface TextCommand { kind: 'text'; x: number; y: number; text: string; font: string; color: string; align?: 'left' | 'center' | 'right'; baseline?: 'top' | 'middle' | 'alphabetic' | 'bottom'; maxWidth?: number; rotation?: number; clip?: RenderRect }
export interface ImageCommand extends RenderRect { kind: 'image'; source: CanvasImageSource; sourceVersion?: number; opacity?: number; clip?: RenderRect }
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
