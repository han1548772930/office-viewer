import type { SheetImageResource } from '../../domain/sheet-images.ts'
import { assertSheetClipboardBytes } from '../sheet-workbench/sheet-clipboard-data.ts'
import { prepareSheetImageUrl } from './sheet-image-io.ts'
import { sheetImageResourceFromDataUrl } from './sheet-image-source.ts'
import { SHEET_IMAGE_MAX_BATCH_BYTES, SHEET_IMAGE_MAX_COUNT } from './sheet-image-model.ts'

export interface SheetClipboardImagePreparationOptions {
  signal?: AbortSignal
  current?: () => boolean
  progress?: (completed: number, total: number) => void
}

/** External HTML images are embedded before the existing synchronous paste transaction runs. */
export async function prepareSheetClipboardHtmlImages(html: string, options: SheetClipboardImagePreparationOptions = {}): Promise<string> {
  assertSheetClipboardBytes(html)
  const assertCurrent = () => {
    if (options.signal?.aborted || options.current && !options.current()) throw new DOMException('剪贴板图片读取已取消。', 'AbortError')
  }
  assertCurrent()
  if (!/<img\b/i.test(html)) return html
  if (typeof DOMParser === 'undefined') throw new Error('当前环境缺少 HTML 剪贴板解析器。')
  const parsed = new DOMParser().parseFromString(html, 'text/html'), images = Array.from(parsed.querySelectorAll('img'))
  if (images.length > SHEET_IMAGE_MAX_COUNT) throw new RangeError('剪贴板图片超过 1000 张。')
  const cache = new Map<string, SheetImageResource>()
  let bytes = 0, completed = 0
  for (const image of images) {
    assertCurrent()
    const source = image.getAttribute('src')?.trim() ?? ''
    let resource = cache.get(source)
    if (!resource) {
      resource = /^data:/i.test(source) ? sheetImageResourceFromDataUrl(source, image.getAttribute('title') || 'clipboard-image')
        : await prepareSheetImageUrl(source, { signal: options.signal, name: image.getAttribute('title') || 'clipboard-image' })
      cache.set(source, resource)
    }
    assertCurrent()
    bytes += resource.byteLength
    if (bytes > SHEET_IMAGE_MAX_BATCH_BYTES) throw new RangeError('剪贴板图片超过 32 MiB。')
    image.setAttribute('src', resource.source)
    // A browser must use the single prepared source instead of an unprepared srcset alternative.
    image.removeAttribute('srcset')
    options.progress?.(++completed, images.length)
  }
  assertCurrent()
  const result = parsed.documentElement.outerHTML
  assertSheetClipboardBytes(result)
  return result
}
