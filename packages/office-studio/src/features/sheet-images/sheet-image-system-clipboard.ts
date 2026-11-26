import type { SheetImageResource } from '../../domain/sheet-images.ts'
import { decodeSheetImageDataUrl } from './sheet-image-bytes.ts'
import { decodeSheetImageSource, prepareSheetImageResources } from './sheet-image-io.ts'
import { normalizeSheetImageResource } from './sheet-image-model.ts'

export interface SheetImageClipboardEncodingOptions { signal?: AbortSignal; current?: () => boolean }

/** Native image clipboard requires PNG; original media remains unchanged in the document and HTML flavour. */
export async function sheetImageClipboardPngBlob(resource: SheetImageResource, options: SheetImageClipboardEncodingOptions = {}): Promise<Blob> {
  const assertCurrent = () => {
    if (options.signal?.aborted || options.current && !options.current()) throw new DOMException('图片剪贴板准备已取消。', 'AbortError')
  }
  assertCurrent()
  const valid = normalizeSheetImageResource(resource)
  const actual = /^data:/i.test(valid.source) ? valid : (await prepareSheetImageResources([valid], options))[0]!
  assertCurrent()
  const data = decodeSheetImageDataUrl(actual.source), blob = new Blob([data.bytes], { type: actual.mimeType })
  if (actual.mimeType === 'image/png') return blob
  const decoded = await decodeSheetImageSource(actual.source, options.signal)
  try {
    assertCurrent()
    if (typeof OffscreenCanvas !== 'undefined') {
      const canvas = new OffscreenCanvas(actual.width, actual.height)
      try {
        const context = canvas.getContext('2d')
        if (!context) throw new Error('无法准备原生图片剪贴板。')
        context.drawImage(decoded, 0, 0)
        const converted = await canvas.convertToBlob({ type: 'image/png' })
        assertCurrent()
        return converted
      } finally { canvas.width = 0; canvas.height = 0 }
    }
    const canvas = document.createElement('canvas')
    canvas.width = actual.width; canvas.height = actual.height
    try {
      const context = canvas.getContext('2d')
      if (!context) throw new Error('无法准备原生图片剪贴板。')
      context.drawImage(decoded, 0, 0)
      const converted = await new Promise<Blob>((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error('图片剪贴板编码失败。')), 'image/png'))
      assertCurrent()
      return converted
    } finally { canvas.width = 0; canvas.height = 0 }
  } finally { decoded.src = '' }
}
