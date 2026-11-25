import type { SheetImageResource } from '../../domain/sheet-images.ts'

export const SHEET_IMAGE_MAX_PIXELS = 16_777_216
export const SHEET_IMAGE_MAX_FILE_BYTES = 5 * 1024 * 1024
export interface SheetImageHeader { mimeType: SheetImageResource['mimeType']; width: number; height: number; scaleX?: number; scaleY?: number }
const matches = (bytes: Uint8Array, offset: number, values: readonly number[]) => values.every((value, index) => bytes[offset + index] === value)
const resolution = (x: number, y: number, units: 'meter' | 'inch' | 'cm'): Pick<SheetImageHeader, 'scaleX' | 'scaleY'> => {
  const factor = units === 'meter' ? .0254 : units === 'cm' ? 2.54 : 1
  if (x <= 0 || y <= 0) return {}
  const scaleX = 96 / (x * factor), scaleY = 96 / (y * factor)
  return Number.isFinite(scaleX) && Number.isFinite(scaleY) && scaleX >= .001 && scaleY >= .001 && scaleX <= 1000 && scaleY <= 1000
    ? { scaleX: Math.abs(scaleX - 1) <= .001 ? 1 : scaleX, scaleY: Math.abs(scaleY - 1) <= .001 ? 1 : scaleY } : {}
}

/** Inspect encoded dimensions before allocating a browser decode; shared by import, commands and rendering. */
export function inspectSheetImageBytes(bytes: Uint8Array): SheetImageHeader {
  if (!bytes.byteLength || bytes.byteLength > SHEET_IMAGE_MAX_FILE_BYTES) throw new RangeError('请选择非空且不超过 5 MiB 的图片。')
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let header: SheetImageHeader | undefined
  if (bytes.length >= 33 && matches(bytes, 0, [137, 80, 78, 71, 13, 10, 26, 10]) && matches(bytes, 12, [73, 72, 68, 82])) {
    header = { mimeType: 'image/png', width: view.getUint32(16), height: view.getUint32(20) }
    for (let offset = 8; offset + 12 <= bytes.length;) {
      const length = view.getUint32(offset), end = offset + length + 12
      if (end > bytes.length || end <= offset) throw new TypeError('PNG 图片数据不完整。')
      if (length === 9 && matches(bytes, offset + 4, [112, 72, 89, 115]) && bytes[offset + 16] === 1) {
        Object.assign(header, resolution(view.getUint32(offset + 8), view.getUint32(offset + 12), 'meter'))
      }
      if (matches(bytes, offset + 4, [73, 69, 78, 68])) break
      offset = end
    }
  } else if (bytes.length >= 20 && matches(bytes, 0, [82, 73, 70, 70]) && matches(bytes, 8, [87, 69, 66, 80])) {
    const end = view.getUint32(4, true) + 8, uint24 = (offset: number) => bytes[offset]! | bytes[offset + 1]! << 8 | bytes[offset + 2]! << 16
    if (end > bytes.length || end < 20) throw new TypeError('WebP 图片数据不完整。')
    for (let offset = 12; offset + 8 <= end;) {
      const length = view.getUint32(offset + 4, true), start = offset + 8, next = start + length + (length & 1)
      if (next > end || next <= offset) throw new TypeError('WebP 图片数据块无效。')
      if (!header && matches(bytes, offset, [86, 80, 56, 88]) && length >= 10) {
        header = { mimeType: 'image/webp', width: uint24(start + 4) + 1, height: uint24(start + 7) + 1 }
      } else if (!header && matches(bytes, offset, [86, 80, 56, 32]) && length >= 10 && matches(bytes, start + 3, [157, 1, 42])) {
        header = { mimeType: 'image/webp', width: view.getUint16(start + 6, true) & 0x3fff, height: view.getUint16(start + 8, true) & 0x3fff }
      } else if (!header && matches(bytes, offset, [86, 80, 56, 76]) && length >= 5 && bytes[start] === 0x2f) {
        header = { mimeType: 'image/webp', width: 1 + (bytes[start + 1]! | (bytes[start + 2]! & 0x3f) << 8),
          height: 1 + (bytes[start + 2]! >> 6 | bytes[start + 3]! << 2 | (bytes[start + 4]! & 0x0f) << 10) }
      }
      offset = next
    }
  } else if (bytes.length >= 10 && (matches(bytes, 0, [71, 73, 70, 56, 55, 97]) || matches(bytes, 0, [71, 73, 70, 56, 57, 97]))) {
    header = { mimeType: 'image/gif', width: view.getUint16(6, true), height: view.getUint16(8, true) }
  } else if (bytes.length >= 26 && matches(bytes, 0, [66, 77])) {
    const dib = view.getUint32(14, true)
    if (dib === 12) header = { mimeType: 'image/bmp', width: view.getUint16(18, true), height: view.getUint16(20, true) }
    else if (dib >= 40 && bytes.length >= 54) {
      header = { mimeType: 'image/bmp', width: view.getInt32(18, true), height: Math.abs(view.getInt32(22, true)),
        ...resolution(view.getInt32(38, true), view.getInt32(42, true), 'meter') }
    }
  } else if (bytes.length >= 4 && matches(bytes, 0, [255, 216])) {
    let density: Pick<SheetImageHeader, 'scaleX' | 'scaleY'> = {}
    for (let offset = 2; offset + 4 <= bytes.length;) {
      if (bytes[offset] !== 255) throw new TypeError('JPEG 图片标记无效。')
      while (bytes[offset] === 255) offset++
      const marker = bytes[offset++]!
      if (marker === 0xd9 || marker === 0xda) break
      if (marker === 1 || marker >= 0xd0 && marker <= 0xd7) continue
      if (offset + 2 > bytes.length) break
      const length = view.getUint16(offset), end = offset + length
      if (length < 2 || end > bytes.length) throw new TypeError('JPEG 图片数据不完整。')
      if (marker === 0xe0 && length >= 16 && matches(bytes, offset + 2, [74, 70, 73, 70, 0])) {
        const unit = bytes[offset + 9]
        if (unit === 1 || unit === 2) density = resolution(view.getUint16(offset + 10), view.getUint16(offset + 12), unit === 1 ? 'inch' : 'cm')
      }
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker) && length >= 8) {
        header = { mimeType: 'image/jpeg', width: view.getUint16(offset + 5), height: view.getUint16(offset + 3) }
      }
      offset = end
    }
    if (header) Object.assign(header, density)
  }
  if (!header || header.width < 1 || header.height < 1) throw new TypeError('无法识别图片，请选择有效的 PNG、JPEG、GIF、BMP 或 WebP 文件。')
  if (header.width * header.height > SHEET_IMAGE_MAX_PIXELS) throw new RangeError('图片解码尺寸超过内存预算。')
  return header
}

/** A data URL's byte size and dimensions are established from its payload, never from author-supplied metadata. */
export function decodeSheetImageDataUrl(source: string): { bytes: Uint8Array; header: SheetImageHeader } {
  const prefix = /^data:image\/(png|jpeg|jpg|gif|bmp|webp);base64,/i.exec(source)
  if (!prefix) throw new TypeError('图片 data URL 无效。')
  const encoded = source.slice(prefix[0].length)
  if (!encoded || encoded.length % 4 !== 0 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)) throw new TypeError('图片 Base64 数据无效。')
  const padding = encoded.endsWith('==') ? 2 : encoded.endsWith('=') ? 1 : 0, length = encoded.length / 4 * 3 - padding
  if (length > SHEET_IMAGE_MAX_FILE_BYTES) throw new RangeError('图片超过 5 MiB 限制。')
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/', bytes = new Uint8Array(length)
  let destination = 0
  for (let offset = 0; offset < encoded.length; offset += 4) {
    const a = alphabet.indexOf(encoded[offset]!), b = alphabet.indexOf(encoded[offset + 1]!),
      c = alphabet.indexOf(encoded[offset + 2]!), d = alphabet.indexOf(encoded[offset + 3]!)
    bytes[destination++] = a << 2 | b >> 4
    if (destination < length) bytes[destination++] = (b & 15) << 4 | c >> 2
    if (destination < length) bytes[destination++] = (c & 3) << 6 | d
  }
  const header = inspectSheetImageBytes(bytes), declared = prefix[1]!.toLowerCase().replace('jpg', 'jpeg')
  if (header.mimeType !== `image/${declared}`) throw new TypeError('图片 MIME 类型与内容不一致。')
  return { bytes, header }
}
