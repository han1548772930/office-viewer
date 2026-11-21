/// <reference types="@webgpu/types" />
import type { RenderDriver, RenderRect, RenderScene, TextCommand } from './contracts'
import { intersectRect, parseColor } from './geometry'

interface TextureEntry { texture: GPUTexture; bindGroup: GPUBindGroup; used: number; bytes: number }
interface TextMetricsEntry { advance: number; left: number; top: number; right: number; bottom: number }
interface Batch { texture: TextureEntry; clip: RenderRect; start: number; count: number }
const SHADER = `
struct VertexOut { @builtin(position) position: vec4f, @location(0) uv: vec2f, @location(1) color: vec4f }
@vertex fn vs(@location(0) position: vec2f, @location(1) uv: vec2f, @location(2) color: vec4f) -> VertexOut {
  var out: VertexOut; out.position = vec4f(position, 0.0, 1.0); out.uv = uv; out.color = color; return out;
}
@group(0) @binding(0) var image: texture_2d<f32>;
@group(0) @binding(1) var imageSampler: sampler;
@fragment fn fs(input: VertexOut) -> @location(0) vec4f { return textureSample(image, imageSampler, input.uv) * input.color; }
`

/** Text runs are shaped by the browser once, cached as small textures, and composited by GPU. */
export async function createWebGPUDriver(canvas: HTMLCanvasElement, onLost: (reason: string) => void): Promise<RenderDriver> {
  if (!navigator.gpu) throw new Error('浏览器未提供 WebGPU')
  const adapter = await navigator.gpu.requestAdapter({ powerPreference: 'low-power' })
  if (!adapter) throw new Error('没有可用的 WebGPU 适配器')
  const device = await adapter.requestDevice()
  const context = canvas.getContext('webgpu')
  if (!context) { device.destroy(); throw new Error('无法创建 WebGPU Canvas 上下文') }
  const format = navigator.gpu.getPreferredCanvasFormat()
  const module = device.createShaderModule({ label: 'office-scene-shader', code: SHADER })
  let pipeline: GPURenderPipeline
  try { pipeline = await device.createRenderPipelineAsync({
    label: 'office-scene-pipeline', layout: 'auto',
    vertex: { module, entryPoint: 'vs', buffers: [{ arrayStride: 32, attributes: [{ shaderLocation: 0, offset: 0, format: 'float32x2' }, { shaderLocation: 1, offset: 8, format: 'float32x2' }, { shaderLocation: 2, offset: 16, format: 'float32x4' }] }] },
    fragment: { module, entryPoint: 'fs', targets: [{ format, blend: { color: { srcFactor: 'src-alpha', dstFactor: 'one-minus-src-alpha', operation: 'add' }, alpha: { srcFactor: 'one', dstFactor: 'one-minus-src-alpha', operation: 'add' } } }] },
    primitive: { topology: 'triangle-list' },
  }) } catch (error) { context.unconfigure(); device.destroy(); throw error }
  const sampler = device.createSampler({ magFilter: 'linear', minFilter: 'linear' })
  let width = 1, height = 1, scaleX = 1, scaleY = 1, frame = 0, disposed = false
  let buffer: GPUBuffer | undefined, bufferSize = 0
  const textCache = new Map<string, TextureEntry>()
  const textMetrics = new Map<string, TextMetricsEntry>()
  const imageCache = new Map<CanvasImageSource, TextureEntry>()
  const imageVersions = new Map<CanvasImageSource, number>()
  const scratch = document.createElement('canvas')
  const measure = scratch.getContext('2d')!
  measure.textAlign = 'left'
  const makeEntry = (texture: GPUTexture, pixelWidth: number, pixelHeight: number): TextureEntry => ({ texture, used: frame, bytes: pixelWidth * pixelHeight * 4, bindGroup: device.createBindGroup({ layout: pipeline.getBindGroupLayout(0), entries: [{ binding: 0, resource: texture.createView() }, { binding: 1, resource: sampler }] }) })
  const textureFromCanvas = (source: HTMLCanvasElement): TextureEntry => {
    const texture = device.createTexture({ label: 'office-raster-texture', size: [source.width, source.height], format: 'rgba8unorm', usage: GPUTextureUsage.TEXTURE_BINDING | GPUTextureUsage.COPY_DST | GPUTextureUsage.RENDER_ATTACHMENT })
    device.queue.copyExternalImageToTexture({ source }, { texture, premultipliedAlpha: false }, [source.width, source.height])
    return makeEntry(texture, source.width, source.height)
  }
  const whiteTexture = device.createTexture({ size: [1, 1], format: 'rgba8unorm', usage: GPUTextureUsage.TEXTURE_BINDING | GPUTextureUsage.COPY_DST })
  device.queue.writeTexture({ texture: whiteTexture }, new Uint8Array([255, 255, 255, 255]), { bytesPerRow: 4 }, [1, 1])
  const white = makeEntry(whiteTexture, 1, 1)
  const clearTextTiles = () => { textCache.forEach(entry => entry.texture.destroy()); textCache.clear() }
  const measureText = (command: TextCommand): TextMetricsEntry => {
    const key = `${command.baseline ?? 'alphabetic'}\u0000${command.font}\u0000${command.text}`
    const cached = textMetrics.get(key)
    if (cached) return cached
    measure.font = command.font
    measure.textBaseline = command.baseline ?? 'alphabetic'
    const metrics = measure.measureText(command.text)
    const entry = { advance: metrics.width, left: Math.floor(-metrics.actualBoundingBoxLeft) - 2, top: Math.floor(-metrics.actualBoundingBoxAscent) - 2, right: Math.ceil(metrics.actualBoundingBoxRight) + 2, bottom: Math.ceil(metrics.actualBoundingBoxDescent) + 2 }
    if (textMetrics.size >= 4096) textMetrics.delete(textMetrics.keys().next().value!)
    textMetrics.set(key, entry); return entry
  }
  const getTextTile = (command: TextCommand, x: number, y: number, width: number, height: number, phaseX: number): TextureEntry => {
    const key = `${command.font}\u0000${command.color}\u0000${command.baseline ?? 'alphabetic'}\u0000${command.maxWidth ?? ''}\u0000${x},${y},${width},${height},${phaseX}\u0000${command.text}`
    const cached = textCache.get(key)
    if (cached) { cached.used = frame; return cached }
    const raster = document.createElement('canvas')
    raster.width = width
    raster.height = height
    const rasterContext = raster.getContext('2d')!
    rasterContext.setTransform(scaleX, 0, 0, scaleY, phaseX - x, -y)
    rasterContext.font = command.font
    rasterContext.fillStyle = command.color
    rasterContext.textAlign = 'left'
    rasterContext.textBaseline = command.baseline ?? 'alphabetic'
    rasterContext.fillText(command.text, 0, 0, command.maxWidth)
    const entry = textureFromCanvas(raster); textCache.set(key, entry); return entry
  }
  const getImage = (source: CanvasImageSource, sourceVersion = 0): TextureEntry => {
    const cached = imageCache.get(source)
    if (cached && imageVersions.get(source) === sourceVersion) { cached.used = frame; return cached }
    if (cached) { cached.texture.destroy(); imageCache.delete(source) }
    const raster = document.createElement('canvas')
    const sourceWidth = 'naturalWidth' in source ? source.naturalWidth : 'videoWidth' in source ? source.videoWidth : 'width' in source ? Number(source.width) : 'displayWidth' in source ? source.displayWidth : 1
    const sourceHeight = 'naturalHeight' in source ? source.naturalHeight : 'videoHeight' in source ? source.videoHeight : 'height' in source ? Number(source.height) : 'displayHeight' in source ? source.displayHeight : 1
    raster.width = Math.min(device.limits.maxTextureDimension2D, Math.max(1, sourceWidth)); raster.height = Math.min(device.limits.maxTextureDimension2D, Math.max(1, sourceHeight))
    raster.getContext('2d')!.drawImage(source, 0, 0, raster.width, raster.height)
    const entry = textureFromCanvas(raster); imageCache.set(source, entry); imageVersions.set(source, sourceVersion); return entry
  }
  const fontsChanged = () => { clearTextTiles(); textMetrics.clear() }
  document.fonts?.addEventListener('loadingdone', fontsChanged)
  void device.lost.then(info => { if (!disposed) onLost(`WebGPU 设备丢失：${info.message || info.reason}`) })
  device.addEventListener('uncapturederror', event => { if (!disposed) onLost(`WebGPU 绘制失败：${event.error.message}`) })

  return {
    backend: 'webgpu',
    resize(cssWidth, cssHeight, pixelRatio) {
      const backingWidth = Math.min(device.limits.maxTextureDimension2D, Math.max(1, Math.round(cssWidth * pixelRatio)))
      const backingHeight = Math.min(device.limits.maxTextureDimension2D, Math.max(1, Math.round(cssHeight * pixelRatio)))
      const nextScaleX = cssWidth > 0 ? backingWidth / cssWidth : 1
      const nextScaleY = cssHeight > 0 ? backingHeight / cssHeight : 1
      if (scaleX !== nextScaleX || scaleY !== nextScaleY) clearTextTiles()
      width = cssWidth; height = cssHeight; scaleX = nextScaleX; scaleY = nextScaleY
      canvas.width = backingWidth; canvas.height = backingHeight
      context.configure({ device, format, alphaMode: 'opaque' })
    },
    render(scene: RenderScene) {
      if (disposed || !width || !height) return
      frame += 1
      const vertices: number[] = [], batches: Batch[] = []
      const viewport: RenderRect = { x: 0, y: 0, width, height }
      const quad = (points: number[], color: string, texture: TextureEntry, clip: RenderRect, opacity = 1) => {
        if (clip.width <= 0 || clip.height <= 0) return
        const rgba = parseColor(color); rgba[3] *= opacity
        const start = vertices.length / 8
        const indices = [0, 1, 2, 0, 2, 3], uv = [[0, 0], [1, 0], [1, 1], [0, 1]]
        for (const index of indices) vertices.push(points[index * 2] / width * 2 - 1, 1 - points[index * 2 + 1] / height * 2, uv[index][0], uv[index][1], ...rgba)
        const previous = batches[batches.length - 1]
        if (previous && previous.texture === texture && previous.clip.x === clip.x && previous.clip.y === clip.y && previous.clip.width === clip.width && previous.clip.height === clip.height) previous.count += 6
        else batches.push({ texture, clip, start, count: 6 })
      }
      const triangle = (points: number[], color: string, clip: RenderRect) => {
        if (clip.width <= 0 || clip.height <= 0) return
        const rgba = parseColor(color)
        const start = vertices.length / 8
        for (const index of [0, 1, 2]) vertices.push(points[index * 2]! / width * 2 - 1, 1 - points[index * 2 + 1]! / height * 2, 0, 0, ...rgba)
        const previous = batches[batches.length - 1]
        if (previous && previous.texture === white && previous.clip.x === clip.x && previous.clip.y === clip.y && previous.clip.width === clip.width && previous.clip.height === clip.height) previous.count += 3
        else batches.push({ texture: white, clip, start, count: 3 })
      }
      const rect = (r: RenderRect, fill: string, clip: RenderRect, texture = white, opacity = 1) => quad([r.x, r.y, r.x + r.width, r.y, r.x + r.width, r.y + r.height, r.x, r.y + r.height], fill, texture, clip, opacity)
      for (const command of scene.commands) {
        let clip = intersectRect(viewport, command.clip ?? viewport)
        if (clip.width <= 0 || clip.height <= 0) continue
        switch (command.kind) {
          case 'rect': {
            rect(command, command.fill, clip)
            if (command.stroke) {
              const s = command.strokeWidth ?? 1, half = s / 2
              for (const side of [{ x: command.x - half, y: command.y - half, width: command.width + s, height: s }, { x: command.x - half, y: command.y + command.height - half, width: command.width + s, height: s }, { x: command.x - half, y: command.y + half, width: s, height: command.height - s }, { x: command.x + command.width - half, y: command.y + half, width: s, height: command.height - s }]) rect(side, command.stroke, clip)
            }
            break
          }
          case 'line': {
            const dx = command.x2 - command.x1, dy = command.y2 - command.y1, length = Math.hypot(dx, dy) || 1
            const ox = -dy / length * (command.width ?? 1) / 2, oy = dx / length * (command.width ?? 1) / 2
            quad([command.x1 + ox, command.y1 + oy, command.x2 + ox, command.y2 + oy, command.x2 - ox, command.y2 - oy, command.x1 - ox, command.y1 - oy], command.color, white, clip)
            break
          }
          case 'polygon': {
            if (command.points.length < 3) break
            const origin = command.points[0]!
            for (let index = 1; index < command.points.length - 1; index += 1) {
              const second = command.points[index]!, third = command.points[index + 1]!
              triangle([origin.x, origin.y, second.x, second.y, third.x, third.y], command.fill, clip)
            }
            break
          }
          case 'text': {
            if (!command.text) break
            const metrics = measureText(command)
            const textWidth = command.maxWidth === undefined ? metrics.advance : Math.min(metrics.advance, command.maxWidth)
            const left = command.x - (command.align === 'center' ? textWidth / 2 : command.align === 'right' ? textWidth : 0)
            if (command.maxWidth !== undefined && !command.rotation) {
              const x = command.x - (command.align === 'center' ? command.maxWidth / 2 : command.align === 'right' ? command.maxWidth : 0)
              clip = intersectRect(clip, { x, y: 0, width: command.maxWidth, height })
            }
            if (clip.width <= 0 || clip.height <= 0) break
            // Keep texture edges on device pixels so the GPU doesn't filter already
            // antialiased text. Quarter-pixel horizontal offsets preserve glyph spacing and cache reuse.
            const subpixelX = Math.round(left * scaleX * 4) / 4
            const originX = Math.floor(subpixelX), originY = Math.round(command.y * scaleY)
            const phaseX = subpixelX - originX
            const radians = (command.rotation ?? 0) * Math.PI / 180
            const cosine = Math.cos(radians), sine = Math.sin(radians)
            const rotate = (x: number, y: number): [number, number] => [
              command.x + (x - command.x) * cosine - (y - command.y) * sine,
              command.y + (x - command.x) * sine + (y - command.y) * cosine,
            ]
            const fitScale = metrics.advance > 0 ? textWidth / metrics.advance : 1
            const inkLeft = Math.floor(metrics.left * fitScale * scaleX + phaseX), inkRight = Math.ceil(metrics.right * fitScale * scaleX + phaseX)
            const inkTop = Math.floor(metrics.top * scaleY), inkBottom = Math.ceil(metrics.bottom * scaleY)
            const tileSize = Math.min(2048, device.limits.maxTextureDimension2D)
            const firstX = Math.max(0, Math.floor((Math.floor(clip.x * scaleX) - originX - inkLeft) / tileSize))
            const lastX = Math.ceil((Math.min(inkRight, Math.ceil((clip.x + clip.width) * scaleX) - originX) - inkLeft) / tileSize)
            const firstY = Math.max(0, Math.floor((Math.floor(clip.y * scaleY) - originY - inkTop) / tileSize))
            const lastY = Math.ceil((Math.min(inkBottom, Math.ceil((clip.y + clip.height) * scaleY) - originY) - inkTop) / tileSize)
            for (let ty = firstY; ty < lastY; ty += 1) for (let tx = firstX; tx < lastX; tx += 1) {
              const x = inkLeft + tx * tileSize, y = inkTop + ty * tileSize
              const tileWidth = Math.min(tileSize, inkRight - x), tileHeight = Math.min(tileSize, inkBottom - y)
              const texture = getTextTile(command, x, y, tileWidth, tileHeight, phaseX)
              const tileLeft = (originX + x) / scaleX, tileTop = (originY + y) / scaleY
              const tileRight = tileLeft + tileWidth / scaleX, tileBottom = tileTop + tileHeight / scaleY
              const points = [
                ...rotate(tileLeft, tileTop), ...rotate(tileRight, tileTop),
                ...rotate(tileRight, tileBottom), ...rotate(tileLeft, tileBottom),
              ]
              quad(points, '#ffffff', texture, clip)
            }
            break
          }
          case 'image': rect(command, '#ffffff', clip, getImage(command.source, command.sourceVersion), command.opacity ?? 1); break
        }
      }
      const data = new Float32Array(vertices)
      if (data.byteLength > bufferSize) { buffer?.destroy(); bufferSize = Math.max(4096, 2 ** Math.ceil(Math.log2(data.byteLength))); buffer = device.createBuffer({ size: bufferSize, usage: GPUBufferUsage.VERTEX | GPUBufferUsage.COPY_DST }) }
      if (data.byteLength) device.queue.writeBuffer(buffer!, 0, data)
      const encoder = device.createCommandEncoder({ label: 'office-scene-frame' })
      const bg = parseColor(scene.background ?? '#ffffff')
      const pass = encoder.beginRenderPass({ colorAttachments: [{ view: context.getCurrentTexture().createView(), clearValue: { r: bg[0], g: bg[1], b: bg[2], a: 1 }, loadOp: 'clear', storeOp: 'store' }] })
      if (data.byteLength) {
        pass.setPipeline(pipeline); pass.setVertexBuffer(0, buffer!)
        for (const batch of batches) {
          const x = Math.max(0, Math.floor(batch.clip.x * scaleX)), y = Math.max(0, Math.floor(batch.clip.y * scaleY))
          const right = Math.min(canvas.width, Math.ceil((batch.clip.x + batch.clip.width) * scaleX)), bottom = Math.min(canvas.height, Math.ceil((batch.clip.y + batch.clip.height) * scaleY))
          if (right <= x || bottom <= y) continue
          pass.setScissorRect(x, y, right - x, bottom - y); pass.setBindGroup(0, batch.texture.bindGroup); pass.draw(batch.count, 1, batch.start)
        }
      }
      pass.end(); device.queue.submit([encoder.finish()])
      const evict = <K>(cache: Map<K, TextureEntry>, onEvict?: (key: K) => void) => {
        let bytes = [...cache.values()].reduce((total, entry) => total + entry.bytes, 0)
        for (const [key, entry] of cache.entries()) if (entry.used < frame - 120 || cache.size > 1024 || bytes > 64 * 1024 * 1024) { entry.texture.destroy(); cache.delete(key); onEvict?.(key); bytes -= entry.bytes }
      }
      evict(textCache); evict(imageCache, source => imageVersions.delete(source))
    },
    dispose() { if (disposed) return; disposed = true; document.fonts?.removeEventListener('loadingdone', fontsChanged); fontsChanged(); imageCache.forEach(entry => entry.texture.destroy()); imageCache.clear(); imageVersions.clear(); whiteTexture.destroy(); buffer?.destroy(); context.unconfigure(); device.destroy() },
  }
}
