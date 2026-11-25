import type { SheetImageResource } from '../../domain/sheet-images.ts'
import { normalizeSheetImageResource, SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_FILE_BYTES, SHEET_IMAGE_MAX_PIXELS } from './sheet-image-model.ts'

import { inspectSheetImageBytes } from './sheet-image-bytes.ts'
export { inspectSheetImageBytes, decodeSheetImageDataUrl, type SheetImageHeader } from './sheet-image-bytes.ts'

function abortError(): DOMException { return new DOMException('图片读取已取消。', 'AbortError') }
function assertLive(signal?: AbortSignal, current?: () => boolean): void {
  if (signal?.aborted || current && !current()) throw abortError()
}
function readImageFile(file: Blob, signal?: AbortSignal): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    const dispose = () => { reader.onload = null; reader.onerror = null; reader.onabort = null; signal?.removeEventListener('abort', cancel) }
    const cancel = () => { reader.abort(); dispose(); reject(abortError()) }
    reader.onload = () => { const value = reader.result; dispose(); value instanceof ArrayBuffer ? resolve(value) : reject(new TypeError('无法读取图片数据。')) }
    reader.onerror = () => { const error = reader.error; dispose(); reject(error ?? new Error('图片读取失败。')) }
    reader.onabort = () => { dispose(); reject(abortError()) }
    signal?.addEventListener('abort', cancel, { once: true })
    if (signal?.aborted) { cancel(); return }
    reader.readAsArrayBuffer(file)
  })
}
function imageDataUrl(bytes: Uint8Array, mimeType: string): string {
  const chunks: string[] = []
  for (let offset = 0; offset < bytes.length; offset += 32_768) chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + 32_768)))
  return `data:${mimeType};base64,${btoa(chunks.join(''))}`
}
export function decodeSheetImageSource(source: string, signal?: AbortSignal): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image(), remote = /^https?:/i.test(source)
    const dispose = () => { image.onload = null; image.onerror = null; signal?.removeEventListener('abort', cancel) }
    const cancel = () => { dispose(); image.src = ''; reject(abortError()) }
    image.onload = () => {
      dispose()
      if (!image.naturalWidth || !image.naturalHeight || image.naturalWidth * image.naturalHeight > SHEET_IMAGE_MAX_PIXELS) {
        image.src = ''; reject(new RangeError('图片解码失败或超过内存预算。')); return
      }
      resolve(image)
    }
    image.onerror = () => { dispose(); image.src = ''; reject(new Error('图片无法解码或读取。')) }
    if (remote) image.crossOrigin = 'anonymous'
    signal?.addEventListener('abort', cancel, { once: true })
    if (signal?.aborted) { cancel(); return }
    image.src = source
  })
}
export async function prepareSheetImageFiles(files: readonly File[], options: {
  signal?: AbortSignal; current?: () => boolean; progress?: (completed: number, total: number) => void
} = {}): Promise<SheetImageResource[]> {
  if (files.length > 1000) throw new RangeError('一次最多插入 1000 张图片。')
  if (files.some(file => !(file instanceof File) || file.size <= 0 || file.size > SHEET_IMAGE_MAX_FILE_BYTES)) throw new RangeError('请选择非空且不超过 5 MiB 的图片。')
  if (files.reduce((sum, file) => sum + file.size, 0) > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('本次图片总大小超过 32 MiB。')
  const resources: SheetImageResource[] = []
  for (const file of files) {
    assertLive(options.signal, options.current)
    const bytes = new Uint8Array(await readImageFile(file, options.signal)), header = inspectSheetImageBytes(bytes)
    assertLive(options.signal, options.current)
    const source = imageDataUrl(bytes, header.mimeType), decoded = await decodeSheetImageSource(source, options.signal)
    try {
      assertLive(options.signal, options.current)
      resources.push(normalizeSheetImageResource({ id: crypto.randomUUID(), name: file.name, source, byteLength: file.size,
        ...header, width: decoded.naturalWidth, height: decoded.naturalHeight }))
    } finally { decoded.src = '' }
    options.progress?.(resources.length, files.length)
  }
  return resources
}

/** Used by programmatic URLs; fetched bytes are embedded so reopen does not depend on a temporary URL. */
export async function prepareSheetImageUrl(source: string, options: { signal?: AbortSignal; name?: string } = {}): Promise<SheetImageResource> {
  if (!/^(?:https?:\/\/|data:image\/(?:png|jpeg|jpg|gif|bmp|webp);base64,)/i.test(source)) throw new TypeError('图片来源无效。')
  const response = await fetch(source, { signal: options.signal })
  if (!response.ok) throw new Error('图片下载失败。')
  const reader = response.body?.getReader(), chunks: Uint8Array[] = []
  let length = 0
  if (!reader) throw new Error('无法读取图片响应。')
  try {
    for (;;) {
      const part = await reader.read()
      if (part.done) break
      length += part.value.byteLength
      if (length > SHEET_IMAGE_MAX_FILE_BYTES) throw new RangeError('图片超过 5 MiB 限制。')
      chunks.push(part.value)
    }
  } finally { await reader.cancel(); reader.releaseLock() }
  assertLive(options.signal)
  const resource = await prepareSheetImageFiles([new File(chunks, options.name ?? 'image')], options)
  return resource[0]!
}

export async function prepareSheetImageResources(resources: readonly SheetImageResource[], options: {
  signal?: AbortSignal; current?: () => boolean
} = {}): Promise<SheetImageResource[]> {
  const result: SheetImageResource[] = []
  let totalBytes = 0
  for (const value of resources) {
    assertLive(options.signal, options.current)
    const resource = normalizeSheetImageResource(value)
    const actual = /^https?:/i.test(resource.source)
      ? { ...resource, ...await prepareSheetImageUrl(resource.source, { signal: options.signal, name: resource.name }), id: resource.id, metadata: resource.metadata }
      : resource
    totalBytes += actual.byteLength
    if (totalBytes > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('本次图片总大小超过 32 MiB。')
    assertLive(options.signal, options.current)
    result.push(actual)
  }
  return result
}
