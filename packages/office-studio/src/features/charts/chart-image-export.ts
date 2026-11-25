export const CHART_IMAGE_EXPORT_MAX_PIXELS = 16_777_216
export const CHART_IMAGE_EXPORT_MAX_BYTES = 64 * 1024 * 1024
export function chartImageExportFilename(chartId: string): string {
  if (typeof chartId !== 'string' || !chartId.trim() || chartId.length > 512) throw new TypeError('导出图表标识无效。')
  return `${chartId.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/, '_')}.png`
}
function assertNotAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw new DOMException('图表图片导出已取消。', 'AbortError')
}
export interface ChartImageExportEnvironment {
  createCanvas(): HTMLCanvasElement
  download(blob: Blob, filename: string, signal?: AbortSignal): void
}
export const chartImageExportEnvironment: ChartImageExportEnvironment = {
  createCanvas: () => document.createElement('canvas'),
  download(blob, filename, signal) {
    assertNotAborted(signal)
    const url = URL.createObjectURL(blob), anchor = document.createElement('a')
    let handedOff = false
    try {
      anchor.href = url; anchor.download = filename; anchor.style.display = 'none'
      document.body.append(anchor); assertNotAborted(signal); anchor.click(); handedOff = true
    } finally {
      anchor.remove()
      // Keep the owned blob URL alive until the browser consumes the click.
      if (handedOff) setTimeout(() => URL.revokeObjectURL(url), 1000)
      else URL.revokeObjectURL(url)
    }
  },
}

/** Snapshot the actual painted bitmap before awaiting encoding. Its owned
 * backing store is released on success, failure, abort and a late callback. */
export async function exportChartCanvasImage(source: HTMLCanvasElement, chartId: string,
  options: { signal?: AbortSignal; assertCurrent?(): void; environment?: ChartImageExportEnvironment } = {}): Promise<Blob> {
  assertNotAborted(options.signal); options.assertCurrent?.()
  const { width, height } = source
  if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width <= 0 || height <= 0 || width * height > CHART_IMAGE_EXPORT_MAX_PIXELS) throw new RangeError('图表位图无效或超出 1600 万像素导出预算。')
  const filename = chartImageExportFilename(chartId), environment = options.environment ?? chartImageExportEnvironment
  const owned = environment.createCanvas()
  if (owned === source) throw new Error('图片编码必须使用独立画布，不能释放图表的借用表面。')
  try {
    owned.width = width; owned.height = height
    const context = owned.getContext('2d')
    if (!context) throw new Error('无法创建图表图片编码画布。')
    context.drawImage(source, 0, 0)
    const blob = await new Promise<Blob>((resolve, reject) => {
      let settled = false
      const finish = (result?: Blob | null, failure?: unknown) => {
        if (settled) return
        settled = true; clearTimeout(timer); options.signal?.removeEventListener('abort', abort)
        if (failure) reject(failure)
        else if (!result || result.type !== 'image/png' || result.size <= 0 || result.size > CHART_IMAGE_EXPORT_MAX_BYTES) reject(new Error('图表 PNG 编码失败或超出 64 MiB 输出预算。'))
        else resolve(result)
      }
      const abort = () => finish(undefined, new DOMException('图表图片导出已取消。', 'AbortError'))
      const timer = setTimeout(() => finish(undefined, new Error('图表 PNG 编码超时。')), 15_000)
      options.signal?.addEventListener('abort', abort, { once: true })
      if (options.signal?.aborted) { abort(); return }
      try { owned.toBlob(result => finish(result), 'image/png') } catch (failure) { finish(undefined, failure) }
    })
    assertNotAborted(options.signal); options.assertCurrent?.()
    environment.download(blob, filename, options.signal)
    return blob
  } finally { owned.width = 0; owned.height = 0 }
}
