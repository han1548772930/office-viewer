import { decodeSheetImageDataUrl } from './sheet-image-bytes.ts'
import { prepareSheetImageResources } from './sheet-image-io.ts'
import { sheetImageSaveNames, type SheetImageSaveConfig, type SheetImageSaveRequest } from './sheet-image-save.ts'
export interface SheetImageFileWriter { write(data: Blob): Promise<void>; close(): Promise<void>; abort?(): Promise<void> }
export interface SheetImageDirectoryHandle { getFileHandle(name: string, options: { create: true }): Promise<{ createWritable(): Promise<SheetImageFileWriter> }> }
export interface SheetImageSavePlatform {
  chooseDirectory?: () => Promise<SheetImageDirectoryHandle>
  download(blob: Blob, name: string): Promise<void>
}
export function browserSheetImageSavePlatform(): SheetImageSavePlatform {
  const browser = window as Window & { showDirectoryPicker?: (options: { mode: 'readwrite' }) => Promise<SheetImageDirectoryHandle> }
  return { chooseDirectory: browser.showDirectoryPicker ? () => browser.showDirectoryPicker!({ mode: 'readwrite' }) : undefined,
    download: async (blob, name) => {
      const source = URL.createObjectURL(blob), anchor = document.createElement('a')
      try { anchor.href = source; anchor.download = name; anchor.rel = 'noopener'; anchor.style.display = 'none'; document.body.append(anchor); anchor.click() }
      finally { anchor.remove(); setTimeout(() => URL.revokeObjectURL(source), 1000) }
    } }
}
/** Writes original bytes sequentially. Unsupported directory pickers fall back to individual browser downloads. */
export async function saveSheetImageFiles(request: SheetImageSaveRequest, config: SheetImageSaveConfig, options: {
  single?: boolean; signal?: AbortSignal; assertCurrent?: () => void; progress?: (completed: number, total: number) => void; platform?: SheetImageSavePlatform
} = {}): Promise<boolean> {
  if (!request.entries.length) return false
  const platform = options.platform ?? browserSheetImageSavePlatform(), names = sheetImageSaveNames(request, config)
  const assertCurrent = () => { if (options.signal?.aborted) throw new DOMException('图片保存已取消。', 'AbortError'); options.assertCurrent?.() }
  assertCurrent()
  const directory = !options.single && platform.chooseDirectory ? await platform.chooseDirectory() : undefined
  assertCurrent()
  for (let index = 0; index < request.entries.length; index++) {
    assertCurrent()
    const resource = (await prepareSheetImageResources([request.entries[index]!.image.resource], { signal: options.signal }))[0]!
    const data = decodeSheetImageDataUrl(resource.source), blob = new Blob([data.bytes], { type: resource.mimeType })
    assertCurrent()
    if (directory) {
      const handle = await directory.getFileHandle(names[index]!, { create: true }); assertCurrent()
      const writer = await handle.createWritable()
      try { assertCurrent(); await writer.write(blob); assertCurrent(); await writer.close() }
      catch (reason: unknown) { await writer.abort?.(); throw reason }
    } else await platform.download(blob, names[index]!)
    options.progress?.(index + 1, request.entries.length)
  }
  return true
}
