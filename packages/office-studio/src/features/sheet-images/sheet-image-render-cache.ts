import type { SheetBackgroundImage, SheetImageResource } from '../../domain/sheet-images.ts'
import type { GridViewport, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'
import { decodeSheetImageSource, prepareSheetImageResources, prepareSheetImageUrl } from './sheet-image-io.ts'

interface Entry { resource?: SheetImageResource; controller: AbortController; source?: HTMLImageElement; bytes: number }
const key = (resource: SheetImageResource) => `${resource.id}\0${resource.source}`
const backgroundId = (background: SheetBackgroundImage) => background.resource ? key(background.resource) : `linked\0${background.imageSourceType}\0${background.source}`
/** One editor owns all decodes and background rasterization; no decoded object enters the workbook. */
export class SheetImageRenderCache {
  private entries = new Map<string, Entry>()
  private disposed = false
  private background?: HTMLCanvasElement
  private backgroundKey = ''
  private backgroundGeneration = 0
  constructor(private changed: () => void, private failed: (message: string) => void = () => {}, private maxBytes = 64 * 1024 * 1024) {}
  get(resource: SheetImageResource): CanvasImageSource | undefined { return this.entries.get(key(resource))?.source }
  get backgroundVersion(): number { return this.backgroundGeneration }
  sync(resources: readonly SheetImageResource[], background?: SheetBackgroundImage | null): void {
    if (this.disposed) return
    const desired = new Map<string, { resource?: SheetImageResource; linked?: SheetBackgroundImage }>(resources.map(resource => [key(resource), { resource }]))
    if (background && !background.resource) desired.set(backgroundId(background), { linked: background })
    for (const [id, entry] of this.entries) if (!desired.has(id)) { this.release(entry); this.entries.delete(id) }
    let bytes = [...this.entries.values()].reduce((sum, entry) => sum + entry.bytes, 0)
    for (const [id, request] of desired) {
      if (this.entries.has(id)) continue
      const resource = request.resource, size = resource ? resource.width * resource.height * 4 : 0
      if (bytes + size > this.maxBytes) { this.failed('可见图片解码超过 64 MiB 预算，请缩小显示区域。'); continue }
      bytes += size
      const entry: Entry = { resource, controller: new AbortController(), bytes: size }
      this.entries.set(id, entry)
      const prepare = async () => {
        if (resource) return prepareSheetImageResources([resource], { signal: entry.controller.signal })
        const linked = request.linked
        if (!linked || linked.resource || linked.imageSourceType === 'UUID') throw new Error('UUID 背景图需要绑定实际媒体资源后再显示。')
        return [await prepareSheetImageUrl(linked.source, { signal: entry.controller.signal, name: 'background' })]
      }
      void prepare().then(async prepared => {
        const actual = prepared[0]!, decodedBytes = actual.width * actual.height * 4
        if (this.disposed || this.entries.get(id) !== entry) return undefined
        const otherBytes = [...this.entries.values()].reduce((sum, item) => sum + (item === entry ? 0 : item.bytes), 0)
        if (otherBytes + decodedBytes > this.maxBytes) throw new RangeError('可见图片解码超过 64 MiB 预算，请缩小显示区域。')
        entry.bytes = decodedBytes; entry.resource = actual
        return decodeSheetImageSource(actual.source, entry.controller.signal)
      }).then(source => {
        if (!source) return
        if (this.disposed || this.entries.get(id) !== entry) { source.src = ''; return }
        const decodedBytes = source.naturalWidth * source.naturalHeight * 4
        const otherBytes = [...this.entries.values()].reduce((sum, item) => sum + (item === entry ? 0 : item.bytes), 0)
        if (otherBytes + decodedBytes > this.maxBytes) { source.src = ''; throw new RangeError('图片实际解码尺寸超过内存预算。') }
        entry.bytes = decodedBytes
        entry.source = source; this.backgroundKey = ''; this.changed()
      }).catch((reason: unknown) => {
        if (this.disposed || this.entries.get(id) !== entry || entry.controller.signal.aborted) return
        this.failed(reason instanceof Error ? reason.message : '图片显示失败。'); this.changed()
      })
    }
  }
  /** A bounded Canvas pattern is composited by either backend; tiny tiles do not create millions of commands. */
  backgroundSource(geometry: SpreadsheetGeometry, viewport: GridViewport, config?: SheetBackgroundImage | null): HTMLCanvasElement | undefined {
    if (this.disposed || !config) { this.backgroundKey = ''; return undefined }
    const id = backgroundId(config), image = this.entries.get(id)?.source
    if (!image) return undefined
    const signature = `${id}\0${config.scaleX}/${config.scaleY}\0${JSON.stringify(viewport)}\0${geometry.frozenRows}/${geometry.frozenColumns}\0${geometry.panes(viewport).map(pane => `${pane.clip.x},${pane.clip.y},${pane.clip.width},${pane.clip.height}`).join(';')}`
    if (this.background && signature === this.backgroundKey) return this.background
    const width = Math.max(1, Math.ceil(viewport.width)), height = Math.max(1, Math.ceil(viewport.height))
    if (width * height * 4 > this.maxBytes) { this.failed('背景图绘制区域超过内存预算。'); return undefined }
    const canvas = this.background ??= document.createElement('canvas')
    canvas.width = width; canvas.height = height
    const context = canvas.getContext('2d')
    if (!context) { this.failed('无法准备工作表背景图。'); return undefined }
    for (const pane of geometry.panes(viewport)) {
      if (!pane.clip.width || !pane.clip.height) continue
      const pattern = context.createPattern(image, 'repeat')
      if (!pattern) continue
      const frozenX = geometry.frozenColumns > 0 && pane.clip.x === geometry.headerWidth
      const frozenY = geometry.frozenRows > 0 && pane.clip.y === geometry.headerHeight
      pattern.setTransform({ a: config.scaleX * viewport.zoom, d: config.scaleY * viewport.zoom,
        e: geometry.headerWidth - (frozenX ? 0 : viewport.scrollLeft), f: geometry.headerHeight - (frozenY ? 0 : viewport.scrollTop) })
      context.fillStyle = pattern; context.fillRect(pane.clip.x, pane.clip.y, pane.clip.width, pane.clip.height)
    }
    this.backgroundKey = signature
    this.backgroundGeneration++
    return canvas
  }
  private release(entry: Entry): void { entry.controller.abort(); if (entry.source) entry.source.src = '' }
  clear(): void {
    this.entries.forEach(entry => this.release(entry)); this.entries.clear(); this.backgroundKey = ''
    if (this.background) { this.background.width = 1; this.background.height = 1; this.background = undefined }
  }
  dispose(): void { this.disposed = true; this.clear() }
}
