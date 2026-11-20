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
  let backendAnnounced = false, driverLive = false
  const normalizeDpr = (value: number) => Number.isFinite(value) && value > 0 ? Math.max(1, Math.min(4, value)) : 1
  const normalizeSize = (value: number) => Number.isFinite(value) ? Math.max(0, value) : 0
  let width = 1, height = 1, dpr = normalizeDpr(window.devicePixelRatio), lastScene: RenderScene | undefined
  const makeCanvas = () => {
    const next = document.createElement('canvas'); next.dataset.renderSurface = 'true'; next.setAttribute('aria-hidden', 'true')
    Object.assign(next.style, { display: 'block', width: '100%', height: '100%', position: 'absolute', inset: '0', pointerEvents: 'none' })
    return next
  }
  const setCanvasCssSize = (target: HTMLCanvasElement, cssWidth: number, cssHeight: number) => {
    target.style.width = `${cssWidth}px`
    target.style.height = `${cssHeight}px`
  }
  const disposeDriver = () => { if (driverLive) { driverLive = false; driver.dispose() } }
  const release = () => {
    disposed = true
    lastScene = undefined
    try { disposeDriver() } finally { canvas.remove() }
  }
  const fallback = (reason: string) => {
    if (disposed || fallingBack || driver?.backend === 'canvas2d') return
    fallingBack = true
    try {
      disposeDriver()
      const replacement = makeCanvas(); canvas.replaceWith(replacement); canvas = replacement
      setCanvasCssSize(canvas, width, height)
      driver = factories.canvas2d(canvas); driverLive = true
      if (width && height) driver.resize(width, height, dpr)
      canvas.dataset.renderBackend = driver.backend
      canvas.dataset.renderBackendReason = reason
      backendAnnounced = true
      options.onBackendChange?.('canvas2d', reason)
      if (lastScene && width && height) driver.render(lastScene)
    } catch (error) {
      release()
      throw error
    } finally { fallingBack = false }
  }
  canvas = makeCanvas(); host.appendChild(canvas)
  try {
    try {
      if (options.preferredBackend === 'canvas2d') driver = factories.canvas2d(canvas)
      else driver = await factories.webgpu(canvas, reason => { if (initializing) pendingLoss = reason; else fallback(reason) })
      driverLive = true
    } catch (error) {
      if (options.preferredBackend === 'canvas2d') throw error
      const replacement = makeCanvas(); canvas.replaceWith(replacement); canvas = replacement
      driver = factories.canvas2d(canvas); driverLive = true
      initialReason = error instanceof Error ? error.message : String(error)
    }
    setCanvasCssSize(canvas, width, height)
    try { driver.resize(width, height, dpr) } catch (error) {
      if (driver.backend === 'canvas2d') throw error
      fallback(error instanceof Error ? error.message : String(error))
    }
    initializing = false
    if (pendingLoss) fallback(pendingLoss)
    canvas.dataset.renderBackend = driver.backend
    if (initialReason) canvas.dataset.renderBackendReason = initialReason
    if (!backendAnnounced) options.onBackendChange?.(driver.backend, initialReason)
  } catch (error) {
    if (!disposed) release()
    throw error
  }
  return {
    get backend() { return driver.backend }, get canvas() { return canvas },
    resize(cssWidth, cssHeight, pixelRatio = window.devicePixelRatio || 1) {
      if (disposed) return
      width = normalizeSize(cssWidth); height = normalizeSize(cssHeight); dpr = normalizeDpr(pixelRatio)
      setCanvasCssSize(canvas, width, height)
      if (!width || !height) return
      try { driver.resize(width, height, dpr) } catch (error) { if (driver.backend === 'canvas2d') throw error; fallback(String(error)) }
    },
    render(scene) {
      if (disposed) return
      lastScene = scene
      if (!width || !height) return
      try { driver.render(scene) } catch (error) { if (driver.backend === 'canvas2d') throw error; fallback(String(error)) }
    },
    dispose() { if (!disposed) release() },
  }
}
