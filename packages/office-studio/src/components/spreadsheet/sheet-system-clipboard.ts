export interface SheetSystemClipboardContent {
  text?: string
  html?: string
}

export interface SheetSystemClipboard {
  /** Optional adapter-specific probe. Merely having a clipboard object does not establish a reader. */
  canRead?(): boolean
  read(): Promise<SheetSystemClipboardContent | undefined>
  write(content: { plain: string; html: string }): Promise<boolean>
}

/** Browser I/O stays outside the model planner. Undefined means no reader, rather than an empty clipboard. */
export const sheetSystemClipboard: SheetSystemClipboard = {
  canRead() {
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard
    return typeof clipboard?.read === 'function' || typeof clipboard?.readText === 'function'
  },
  async read() {
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard
    if (typeof clipboard?.read === 'function') {
      try {
        const items = await clipboard.read()
        const item = items.find(candidate => candidate.types.includes('text/plain') || candidate.types.includes('text/html'))
        // An actual empty/media-only clipboard must not silently paste an older internal cache.
        if (!item) return { text: '', html: '' }
        const [text, html] = await Promise.all([
          item.types.includes('text/plain') ? item.getType('text/plain').then(blob => blob.text()) : undefined,
          item.types.includes('text/html') ? item.getType('text/html').then(blob => blob.text()) : undefined,
        ])
        return { ...(text === undefined ? {} : { text }), ...(html === undefined ? {} : { html }) }
      } catch (error) {
        if (typeof clipboard.readText !== 'function') throw error
      }
    }
    if (typeof clipboard?.readText === 'function') return { text: await clipboard.readText() }
    return undefined
  },
  async write(content) {
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard
    if (typeof clipboard?.write === 'function' && typeof ClipboardItem !== 'undefined') {
      try {
        await clipboard.write([new ClipboardItem({
          'text/plain': new Blob([content.plain], { type: 'text/plain' }),
          'text/html': new Blob([content.html], { type: 'text/html' }),
        })])
        return true
      } catch (error) {
        if (typeof clipboard.writeText !== 'function') throw error
      }
    }
    if (typeof clipboard?.writeText === 'function') { await clipboard.writeText(content.plain); return true }
    return false
  },
}

export function sheetSystemClipboardAvailable(adapter: SheetSystemClipboard = sheetSystemClipboard): boolean {
  return typeof adapter.read === 'function' && (adapter.canRead === undefined || typeof adapter.canRead === 'function' && adapter.canRead())
}
