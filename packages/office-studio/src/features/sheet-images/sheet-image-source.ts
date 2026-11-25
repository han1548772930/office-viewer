import type { SheetImageResource } from '../../domain/sheet-images.ts'
import { decodeSheetImageDataUrl } from './sheet-image-bytes.ts'
import { normalizeSheetImageResource } from './sheet-image-model.ts'

/** Encoded sources can be checked synchronously without a DOM image or a fake dimension hint. */
export function sheetImageResourceFromDataUrl(source: string, name = 'image', id = crypto.randomUUID()): SheetImageResource {
  const { bytes, header } = decodeSheetImageDataUrl(source)
  return normalizeSheetImageResource({ id, name, source, byteLength: bytes.byteLength, ...header })
}
