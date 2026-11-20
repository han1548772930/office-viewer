import type { RenderDriver, RenderScene, RenderSurface, RenderSurfaceOptions } from './contracts'
import { createCanvas2DDriver } from './canvas2d'
import { createWebGPUDriver } from './webgpu'

export interface RenderDriverFactories {
  webgpu: typeof createWebGPUDriver
  canvas2d: typeof createCanvas2DDriver
}

/** Owns backend resources only. The authoritative scene survives every canvas replacement. */
export async function createRenderSurface(host: HTMLElement, options: RenderSurfaceOptions = {}, factories: RenderDriverFactories = { webgpu: createWebGPUDriver, canvas2d: createCanvas2DDriver }): Promise<RenderSurface> {
  let canvas: HTMLCanvasElement, driver: RenderDriver, disposed = false, fallingBack = false, initializing = true
  let pendingLoss: string | undefined, initialReason: string | undefined
  let width = 1, height = 1, dpr = Math.max(1, window.devicePixelRatio || 1), lastScene: RenderScene | undefined
  const makeCanvas = () => {
    const next = document.createElement('canvas'); next.dataset.renderSurface = 'true'; next.setAttribute('aria-hidden', 'true')
    Object.assign(next.style, { display: 'block', width: '100%', height: '100%', position: 'absolute', inset: '0', pointerEvents: 'none' })
    return next
  }
  const fallback = (reason: string) => {
    if (disposed || fallingBack || driver?.backend === 'canvas2d') return
    fallingBack = true
    try {
      driver.dispose()
      const replacement = makeCanvas(); canvas.replaceWith(replacement); canvas = replacement
      driver = factories.canvas2d(canvas); driver.resize(width, height, dpr)
      options.onBackendChange?.('canvas2d', reason)
      if (lastScene) driver.render(lastScene)
    } finally { fallingBack = false }
  }
  canvas = makeCanvas(); host.appendChild(canvas)
  try {
    if (options.preferredBackend === 'canvas2d') driver = factories.canvas2d(canvas)
    else driver = await factories.webgpu(canvas, reason => { if (initializing) pendingLoss = reason; else fallback(reason) })
  } catch (error) {
    const replacement = makeCanvas(); canvas.replaceWith(replacement); canvas = replacement
    try { driver = factories.canvas2d(canvas) } catch (fallbackError) { canvas.remove(); throw fallbackError }
    initialReason = error instanceof Error ? error.message : String(error)
  }
  initializing = false
  if (pendingLoss) fallback(pendingLoss)
  driver.resize(width, height, dpr)
  if (!pendingLoss) options.onBackendChange?.(driver.backend, initialReason)
  return {
    get backend() { return driver.backend }, get canvas() { return canvas },
    resize(cssWidth, cssHeight, pixelRatio = window.devicePixelRatio || 1) {
      if (disposed) return
      width = Math.max(0, cssWidth); height = Math.max(0, cssHeight); dpr = Math.max(1, Math.min(4, pixelRatio))
      if (!width || !height) return
      try { driver.resize(width, height, dpr) } catch (error) { if (driver.backend === 'canvas2d') throw error; fallback(String(error)) }
    },
    render(scene) {
      if (disposed) return
      lastScene = scene
      if (!width || !height) return
      try { driver.render(scene) } catch (error) { if (driver.backend === 'canvas2d') throw error; fallback(String(error)) }
    },
    dispose() { if (disposed) return; disposed = true; lastScene = undefined; driver.dispose(); canvas.remove() },
  }
}
