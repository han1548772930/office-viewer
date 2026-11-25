import type { ChartImageResource, ChartModel, WorkbookModel } from '../../domain/model'
import { ComputeError } from '../../compute/errors'

export const CHART_IMAGE_RESOURCE_KEY = 'chartImageResources' as const
export const CHART_IMAGE_MAX_BYTES = 8 * 1024 * 1024
export const CHART_IMAGE_REGISTRY_MAX_BYTES = 32 * 1024 * 1024
const IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/avif', 'image/bmp', 'image/svg+xml', 'image/x-icon', 'image/vnd.microsoft.icon'])

export function chartImageMimeType(value: unknown): string {
  if (typeof value !== 'string') throw new TypeError('图表图片需要明确的图片 MIME 类型。')
  const mimeType = value.split(';', 1)[0]!.trim().toLowerCase()
  if (!IMAGE_TYPES.has(mimeType)) throw new TypeError('图表遮罩须为 PNG、JPEG、GIF、WebP、AVIF、BMP、SVG 或图标图片。')
  return mimeType
}

/** A URL is a locator, never an unregistered resource ID or an ephemeral blob. */
export function chartImageUrl(value: unknown): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 16_384) throw new TypeError('图表图片 URL 无效。')
  if (!/^(?:https?:\/\/|\/|\.\.?\/)/i.test(value)) throw new TypeError('图片 URL 须为 HTTP(S) 或明确的相对路径；资源 ID 必须先注册。')
  const parsed = new URL(value, 'https://office-viewer.local/')
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') throw new TypeError('图表图片 URL 协议无效。')
  return value
}

export function normalizeChartImageResource(value: unknown): ChartImageResource {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('图表图片资源必须为对象。')
  const item = value as Partial<ChartImageResource>
  if (typeof item.id !== 'string' || !item.id.trim() || item.id.length > 512) throw new TypeError('图表图片资源 ID 无效。')
  if (item.name !== undefined && (typeof item.name !== 'string' || item.name.length > 1024)) throw new TypeError('图表图片资源名称无效。')
  const mimeType = chartImageMimeType(item.mimeType)
  let sizeBytes = item.sizeBytes, source: string
  if (sizeBytes !== undefined && (!Number.isSafeInteger(sizeBytes) || sizeBytes <= 0 || sizeBytes > CHART_IMAGE_MAX_BYTES)) throw new RangeError('图表图片须不超过 8 MiB。')
  if (item.sourceType === 'dataUrl') {
    if (typeof item.source !== 'string' || item.source.length > Math.ceil(CHART_IMAGE_MAX_BYTES / 3) * 4 + 128) throw new RangeError('图表图片数据超出预算。')
    const match = /^data:([^;,]+);base64,([A-Za-z0-9+/]*={0,2})$/.exec(item.source)
    if (!match || chartImageMimeType(match[1]) !== mimeType || match[2]!.length % 4 !== 0) throw new TypeError('图表图片 data URL 与 MIME 类型不匹配。')
    const payload = match[2]!, actual = payload.length / 4 * 3 - (payload.endsWith('==') ? 2 : payload.endsWith('=') ? 1 : 0)
    if (!Number.isSafeInteger(actual) || actual <= 0 || actual > CHART_IMAGE_MAX_BYTES || sizeBytes !== undefined && sizeBytes !== actual) throw new RangeError('图表图片数据长度与字节声明不匹配。')
    source = item.source; sizeBytes = actual
  } else if (item.sourceType === 'url') source = chartImageUrl(item.source)
  else throw new TypeError('图表图片资源须区分 URL 和 data URL。')
  return { id: item.id, sourceType: item.sourceType, source, mimeType,
    ...(sizeBytes === undefined ? {} : { sizeBytes }), ...(item.name === undefined ? {} : { name: item.name }) }
}

export function readChartImageResources(workbook: Pick<WorkbookModel, 'metadata'>): Record<string, ChartImageResource> {
  const value = workbook.metadata?.[CHART_IMAGE_RESOURCE_KEY]
  if (value === undefined) return Object.create(null) as Record<string, ChartImageResource>
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('图表图片资源注册表无效。')
  const result: Record<string, ChartImageResource> = Object.create(null) as Record<string, ChartImageResource>
  let bytes = 0, count = 0
  for (const [id, candidate] of Object.entries(value)) {
    const resource = normalizeChartImageResource(candidate)
    if (id !== resource.id || ++count > 4096) throw new TypeError('图表图片资源 ID 或注册表容量无效。')
    bytes += resource.source.length * 2 + id.length * 2 + 256
    if (!Number.isSafeInteger(bytes) || bytes > CHART_IMAGE_REGISTRY_MAX_BYTES) throw new RangeError('图表图片注册表超出 32 MiB 预算。')
    result[id] = resource
  }
  return result
}

export function resolveChartImageResource(workbook: Pick<WorkbookModel, 'metadata'>,
  chart: Pick<ChartModel, 'options'>, id: string): ChartImageResource | undefined {
  const pending = chart.options?.wordCloud?.maskImageResource
  if (pending) {
    const resource = normalizeChartImageResource(pending)
    if (resource.id !== (chart.options?.wordCloud?.maskImageId ?? chart.options?.wordCloud?.maskImage)) throw new TypeError('词云遮罩引用与候选图片资源不一致。')
    if (resource.id === id) return resource
  }
  return readChartImageResources(workbook)[id]
}

/** Called only inside the host's candidate content transaction. Validate the
 * entire registry before assigning; no failure publishes a partial resource. */
export function publishChartImageResourceCandidate(workbook: WorkbookModel, chart: ChartModel): void {
  const pending = chart.options?.wordCloud?.maskImageResource
  if (pending) {
    const resource = normalizeChartImageResource(pending)
    if (resource.id !== (chart.options?.wordCloud?.maskImageId ?? chart.options?.wordCloud?.maskImage)) throw new TypeError('词云遮罩引用与候选图片资源不一致。')
    const registry = readChartImageResources(workbook), existing = registry[resource.id]
    if (existing && (existing.source !== resource.source || existing.mimeType !== resource.mimeType || existing.sourceType !== resource.sourceType)) throw new Error(`图表图片资源 ID 已被其他内容使用：${resource.id}`)
    registry[resource.id] = resource
    readChartImageResources({ metadata: { [CHART_IMAGE_RESOURCE_KEY]: registry } })
    workbook.metadata = { ...workbook.metadata, [CHART_IMAGE_RESOURCE_KEY]: registry }
    delete chart.options!.wordCloud!.maskImageResource
  }
  pruneChartImageResources(workbook)
}

/** Reference ownership is derived from authored charts, never a second cache.
 * Snapshots/history retain their own registry and restore it atomically. */
export function pruneChartImageResources(workbook: WorkbookModel): void {
  if (workbook.metadata?.[CHART_IMAGE_RESOURCE_KEY] === undefined) return
  const registry = readChartImageResources(workbook), referenced = new Set<string>()
  for (const chart of workbook.charts ?? []) {
    const id = chart.options?.wordCloud?.maskImageId ?? chart.options?.wordCloud?.maskImage
    if (id) referenced.add(id)
  }
  const retained: Record<string, ChartImageResource> = Object.create(null) as Record<string, ChartImageResource>
  for (const [id, resource] of Object.entries(registry)) if (referenced.has(id)) retained[id] = resource
  const metadata = { ...workbook.metadata }
  if (Object.keys(retained).length) metadata[CHART_IMAGE_RESOURCE_KEY] = retained
  else delete metadata[CHART_IMAGE_RESOURCE_KEY]
  workbook.metadata = metadata
}

let nextImageId = 0
/** FileReader is only a browser input adapter. Its returned DTO is validated
 * and the host owns publication; abort/close releases the reader callbacks. */
export function readChartImageFile(file: File, signal?: AbortSignal): Promise<ChartImageResource> {
  const mimeType = chartImageMimeType(file.type)
  if (!Number.isSafeInteger(file.size) || file.size <= 0 || file.size > CHART_IMAGE_MAX_BYTES) return Promise.reject(new RangeError('词云遮罩图片须不超过 8 MiB。'))
  if (signal?.aborted) return Promise.reject(new ComputeError('COMPUTE_CANCELLED', '词云图片读取已取消'))
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    const cleanup = () => { reader.onload = null; reader.onerror = null; reader.onabort = null; signal?.removeEventListener('abort', abort) }
    const abort = () => { reader.abort(); cleanup(); reject(new ComputeError('COMPUTE_CANCELLED', '词云图片读取已取消')) }
    reader.onload = () => {
      cleanup()
      try {
        const id = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function' ? `chart-image:${crypto.randomUUID()}` : `chart-image:${Date.now()}:${++nextImageId}`
        resolve(normalizeChartImageResource({ id, sourceType: 'dataUrl', source: reader.result, mimeType, sizeBytes: file.size, name: file.name }))
      } catch (error) { reject(error) }
    }
    reader.onerror = () => { cleanup(); reject(reader.error ?? new Error('词云图片读取失败。')) }
    reader.onabort = () => { cleanup(); reject(new ComputeError('COMPUTE_CANCELLED', '词云图片读取已取消')) }
    signal?.addEventListener('abort', abort, { once: true })
    if (signal?.aborted) { abort(); return }
    try { reader.readAsDataURL(file) } catch (error) { cleanup(); reject(error) }
  })
}
