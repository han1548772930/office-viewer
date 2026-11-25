import type { SheetImageResource } from '../../domain/sheet-images.ts'
import { decodeSheetImageDataUrl } from './sheet-image-bytes.ts'
import { decodeSheetImageSource } from './sheet-image-io.ts'

/** Native image clipboard requires PNG; original media remains unchanged in the document and HTML flavour. */
export async function sheetImageClipboardPngBlob(resource: SheetImageResource): Promise<Blob> {
  const data = decodeSheetImageDataUrl(resource.source), blob = new Blob([data.bytes], { type: resource.mimeType })
  if (resource.mimeType === 'image/png') return blob
  const decoded = await decodeSheetImageSource(resource.source)
  try {
    if (typeof OffscreenCanvas !== 'undefined') {
      const canvas = new OffscreenCanvas(resource.width, resource.height), context = canvas.getContext('2d')
      if (!context) throw new Error('无法准备原生图片剪贴板。')
      context.drawImage(decoded, 0, 0)
      return await canvas.convertToBlob({ type: 'image/png' })
    }
    const canvas = document.createElement('canvas')
    canvas.width = resource.width; canvas.height = resource.height
    const context = canvas.getContext('2d')
    if (!context) throw new Error('无法准备原生图片剪贴板。')
    context.drawImage(decoded, 0, 0)
    try { return await new Promise<Blob>((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error('图片剪贴板编码失败。')), 'image/png')) }
    finally { canvas.width = 0; canvas.height = 0 }
  } finally { decoded.src = '' }
}
