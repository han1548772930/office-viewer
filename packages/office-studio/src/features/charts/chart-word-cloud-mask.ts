import type { ChartWordCloudShape } from '../../domain/model'
import { ComputeError } from '../../compute/errors'
import { assertChartPreparationCurrent, assertChartPreparationMemory, type ChartComputePreparationContext } from './chart-compute-lifecycle'

export const CHART_WORD_CLOUD_MASK_CHOICES = [
  ['ArrowLeft', '左箭头'], ['ArrowRight', '右箭头'], ['Book', '书本'], ['Brain', '大脑'], ['ChatBubbles', '对话气泡'],
  ['Circle', '圆形'], ['Cloud', '云形'], ['Diamond', '菱形'], ['Ellipse', '椭圆'], ['Flower', '花朵'], ['Folder', '文件夹'],
  ['Head', '头像'], ['Heart', '心形'], ['Light', '灯泡'], ['Music', '音符'], ['Rectangle', '矩形'], ['Star', '星形'],
  ['ThumbDown', '向下拇指'], ['ThumbUp', '向上拇指'], ['Triangle', '三角形'],
] as const

type RasterContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D
export interface ChartWordCloudCanvas { canvas: OffscreenCanvas | HTMLCanvasElement; context: RasterContext; release(): void }

/** This is a browser platform adapter, not a second layout engine or UI tree. */
export function createChartWordCloudCanvas(width: number, height: number): ChartWordCloudCanvas {
  if (typeof OffscreenCanvas !== 'undefined') {
    const canvas = new OffscreenCanvas(width, height), context = canvas.getContext('2d', { willReadFrequently: true })
    if (!context) throw new ComputeError('COMPUTE_FAILED', '无法创建词云字形画布')
    return { canvas, context, release: () => { canvas.width = 0; canvas.height = 0 } }
  }
  if (typeof document === 'undefined') throw new ComputeError('COMPUTE_UNSUPPORTED_TYPE', '词云字形准备需要浏览器 Canvas')
  const canvas = document.createElement('canvas')
  canvas.width = width; canvas.height = height
  const context = canvas.getContext('2d', { willReadFrequently: true })
  if (!context) throw new ComputeError('COMPUTE_FAILED', '无法创建词云字形画布')
  return { canvas, context, release: () => { canvas.width = 0; canvas.height = 0 } }
}

function polygon(context: RasterContext, points: ReadonlyArray<readonly [number, number]>): void {
  context.beginPath()
  points.forEach(([x, y], index) => index ? context.lineTo(x, y) : context.moveTo(x, y))
  context.closePath(); context.fill()
}
function ellipse(context: RasterContext, x: number, y: number, rx: number, ry = rx): void {
  context.beginPath(); context.ellipse(x, y, rx, ry, 0, 0, Math.PI * 2); context.fill()
}
function star(context: RasterContext, count: number, inner = 1, start = -Math.PI / 2): void {
  polygon(context, Array.from({ length: count }, (_value, index) => {
    const radius = index % 2 ? inner : 1, angle = start + index * Math.PI * 2 / count
    return [Math.cos(angle) * radius, Math.sin(angle) * radius] as const
  }))
}
function heart(context: RasterContext): void {
  context.beginPath(); context.moveTo(0, 0.95)
  context.bezierCurveTo(-1.7, -0.15, -0.8, -1.5, 0, -0.45)
  context.bezierCurveTo(0.8, -1.5, 1.7, -0.15, 0, 0.95); context.fill()
}
function cloud(context: RasterContext): void {
  ellipse(context, 0, 0.3, 0.95, 0.5); ellipse(context, -0.45, -0.1, 0.46)
  ellipse(context, 0.08, -0.35, 0.57); ellipse(context, 0.58, -0.05, 0.37)
}

/** All twenty fixed IDs have independent Canvas contours. The commercial
 * bitmap assets are not copied; contour differences remain unaccepted. */
function builtinMask(context: RasterContext, id: string): boolean {
  switch (id) {
    case 'Circle': ellipse(context, 0, 0, 1); break
    case 'Ellipse': ellipse(context, 0, 0, 1, 0.7); break
    case 'Rectangle': context.fillRect(-1, -0.75, 2, 1.5); break
    case 'Cloud': cloud(context); break
    case 'Diamond': polygon(context, [[0, -1], [1, 0], [0, 1], [-1, 0]]); break
    case 'Triangle': star(context, 3); break
    case 'Star': star(context, 10, 0.42); break
    case 'Heart': heart(context); break
    case 'ArrowLeft': polygon(context, [[-1, 0], [-0.1, -0.85], [-0.1, -0.35], [1, -0.35], [1, 0.35], [-0.1, 0.35], [-0.1, 0.85]]); break
    case 'ArrowRight': context.scale(-1, 1); builtinMask(context, 'ArrowLeft'); break
    case 'Book': polygon(context, [[-1, -0.65], [-0.2, -0.85], [0, -0.65], [0.2, -0.85], [1, -0.65], [1, 0.85], [0.2, 0.65], [0, 0.85], [-0.2, 0.65], [-1, 0.85]]); break
    case 'Brain':
      ellipse(context, 0, 0, 0.85, 0.9)
      for (let index = 0; index < 8; index++) { const angle = index * Math.PI / 4; ellipse(context, Math.cos(angle) * 0.6, Math.sin(angle) * 0.65, 0.33) }
      break
    case 'ChatBubbles':
      ellipse(context, -0.2, -0.2, 0.75, 0.55); ellipse(context, 0.35, 0.35, 0.6, 0.45)
      polygon(context, [[-0.7, 0.05], [-0.85, 0.65], [-0.15, 0.25]])
      polygon(context, [[0.55, 0.5], [0.9, 0.95], [0.8, 0.25]])
      break
    case 'Flower':
      for (let index = 0; index < 8; index++) { const angle = index * Math.PI / 4; ellipse(context, Math.cos(angle) * 0.55, Math.sin(angle) * 0.55, 0.4) }
      ellipse(context, 0, 0, 0.6); break
    case 'Folder': polygon(context, [[-1, -0.7], [-0.15, -0.7], [0.05, -0.45], [1, -0.45], [0.85, 0.8], [-0.95, 0.8]]); break
    case 'Head':
      ellipse(context, 0, -0.35, 0.5, 0.6); ellipse(context, 0, 0.9, 0.92, 0.55)
      context.fillRect(-0.3, -0.1, 0.6, 0.85); break
    case 'Light':
      ellipse(context, 0, -0.25, 0.7, 0.75); polygon(context, [[-0.5, 0.1], [0.5, 0.1], [0.3, 0.7], [-0.3, 0.7]])
      context.fillRect(-0.3, 0.6, 0.6, 0.3); break
    case 'Music':
      ellipse(context, -0.5, 0.6, 0.4, 0.28); ellipse(context, 0.5, 0.4, 0.4, 0.28)
      context.fillRect(-0.3, -0.65, 0.16, 1.3); context.fillRect(0.7, -0.85, 0.16, 1.3)
      polygon(context, [[-0.3, -0.65], [0.86, -0.9], [0.86, -0.55], [-0.3, -0.3]]); break
    case 'ThumbDown': context.scale(1, -1); builtinMask(context, 'ThumbUp'); break
    case 'ThumbUp': polygon(context, [[-1, -0.1], [-0.55, -0.1], [-0.25, -0.45], [0, -1], [0.3, -0.85], [0.25, -0.1], [0.95, -0.1], [0.85, 0.75], [-0.45, 0.75], [-0.55, 0.5], [-1, 0.5]]); break
    default: return false
  }
  return true
}

export function drawChartWordCloudShape(context: RasterContext, shape: ChartWordCloudShape, width: number, height: number): void {
  context.save(); context.translate(width / 2, height / 2); context.scale(width * 0.48, height * 0.48); context.fillStyle = '#000'
  switch (shape) {
    case 'cloud': cloud(context); break
    case 'rect': context.fillRect(-1, -1, 2, 2); break
    case 'circle': ellipse(context, 0, 0, 1); break
    case 'heart': heart(context); break
    case 'diamond': builtinMask(context, 'Diamond'); break
    case 'triangle': star(context, 3); break
    case 'triangle-forward': star(context, 3, 1, 0); break
    case 'pentagon': star(context, 5); break
    case 'star': star(context, 10, 0.42); break
    case 'cardioid':
      polygon(context, Array.from({ length: 256 }, (_value, index) => {
        const angle = index * Math.PI * 2 / 256, radius = (1 - Math.sin(angle)) / 2
        return [Math.cos(angle) * radius, Math.sin(angle) * radius + 0.35] as const
      })); break
  }
  context.restore()
}

async function cancellable<T>(promise: Promise<T>, context: ChartComputePreparationContext, discard?: (value: T) => void): Promise<T> {
  assertChartPreparationCurrent(context)
  const signal = context.signal
  if (!signal) { const result = await promise; assertChartPreparationCurrent(context); return result }
  return new Promise<T>((resolve, reject) => {
    let cancelled = false
    const abort = () => { cancelled = true; reject(new ComputeError('COMPUTE_CANCELLED', '词云资源准备已取消')) }
    signal.addEventListener('abort', abort, { once: true })
    promise.then(value => {
      signal.removeEventListener('abort', abort)
      if (cancelled) { discard?.(value); return }
      try { assertChartPreparationCurrent(context); resolve(value) } catch (error) { discard?.(value); reject(error) }
    }, error => { signal.removeEventListener('abort', abort); if (!cancelled) reject(error) })
    if (signal.aborted) abort()
  })
}

export function waitChartWordCloudFont(font: string, text: string, context: ChartComputePreparationContext): Promise<unknown> {
  return typeof document !== 'undefined' && document.fonts ? cancellable(document.fonts.load(font, text), context) : Promise.resolve()
}

export async function drawChartWordCloudMaskImage(owner: ChartWordCloudCanvas, imageId: string, context: ChartComputePreparationContext): Promise<void> {
  const { canvas, context: drawing } = owner
  drawing.save(); drawing.translate(canvas.width / 2, canvas.height / 2); drawing.scale(canvas.width * 0.48, canvas.height * 0.48); drawing.fillStyle = '#000'
  const builtin = builtinMask(drawing, imageId)
  drawing.restore()
  if (builtin) return
  let bitmap: ImageBitmap | undefined
  let image: HTMLImageElement | undefined
  let objectUrl: string | undefined
  try {
    const response = await fetch(imageId, { signal: context.signal, credentials: 'omit' })
    if (!response.ok) throw new ComputeError('COMPUTE_INVALID_INPUT', `词云遮罩图片加载失败（${response.status}）`)
    const blob = await response.blob()
    assertChartPreparationCurrent(context); assertChartPreparationMemory(blob.size + canvas.width * canvas.height * 8, context)
    if (typeof createImageBitmap === 'function') {
      bitmap = await cancellable(createImageBitmap(blob), context, value => value.close())
      assertChartPreparationMemory(blob.size + bitmap.width * bitmap.height * 4 + canvas.width * canvas.height * 8, context)
      drawing.drawImage(bitmap, canvas.width * 0.02, canvas.height * 0.02, canvas.width * 0.96, canvas.height * 0.96)
    } else {
      if (typeof document === 'undefined') throw new ComputeError('COMPUTE_UNSUPPORTED_TYPE', '浏览器无法解码词云遮罩图片')
      objectUrl = URL.createObjectURL(blob); image = document.createElement('img'); image.src = objectUrl
      await cancellable(image.decode(), context)
      assertChartPreparationMemory(blob.size + image.naturalWidth * image.naturalHeight * 4 + canvas.width * canvas.height * 8, context)
      drawing.drawImage(image, canvas.width * 0.02, canvas.height * 0.02, canvas.width * 0.96, canvas.height * 0.96)
    }
  } catch (error) {
    assertChartPreparationCurrent(context)
    if (error instanceof ComputeError) throw error
    throw new ComputeError('COMPUTE_INVALID_INPUT', '词云遮罩图片无法加载或解码', error instanceof Error ? error.message : String(error))
  } finally { bitmap?.close(); if (image) image.src = ''; if (objectUrl) URL.revokeObjectURL(objectUrl) }
}
