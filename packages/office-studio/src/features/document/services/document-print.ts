import type { DocumentModel } from '../../../domain/model'
import { documentHtml } from '../model/document-model'

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character] ?? character)
}

function printableHtml(model: DocumentModel): string {
  const pageWidth = Math.max(1, model.page.width)
  const pageHeight = Math.max(1, model.page.height)
  const marginTop = Math.max(0, model.page.marginTop)
  const marginRight = Math.max(0, model.page.marginRight)
  const marginBottom = Math.max(0, model.page.marginBottom)
  const marginLeft = Math.max(0, model.page.marginLeft)
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>${escapeHtml(model.name)}</title>
    <style>
      @page { size: ${pageWidth}px ${pageHeight}px; margin: 0; }
      * { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; background: #fff; color: #1f2937; }
      body { font-family: Aptos, Arial, sans-serif; }
      .office-page { width: ${pageWidth}px; min-height: ${pageHeight}px; padding: ${marginTop}px ${marginRight}px ${marginBottom}px ${marginLeft}px; page-break-after: always; overflow: hidden; }
      .office-page:last-child { page-break-after: auto; }
      table { border-collapse: collapse; width: 100%; }
      td, th { border: 1px solid #aab3bf; padding: 6px; vertical-align: top; }
      img { max-width: 100%; height: auto; }
      p, h1, h2, h3, h4, h5, h6, blockquote { break-inside: avoid; }
    </style></head><body><section class="office-page">${documentHtml(model)}</section></body></html>`
}

/** Print the document layout in an isolated browsing context, excluding the app shell. */
export async function printDocumentModel(model: DocumentModel): Promise<void> {
  if (typeof window === 'undefined' || typeof document === 'undefined') throw new Error('打印需要浏览器环境')
  const frame = document.createElement('iframe')
  frame.setAttribute('aria-hidden', 'true')
  Object.assign(frame.style, { position: 'fixed', width: '1px', height: '1px', right: '0', bottom: '0', border: '0', opacity: '0' })
  document.body.appendChild(frame)
  const cleanup = () => { frame.remove() }
  try {
    const frameDocument = frame.contentDocument
    const frameWindow = frame.contentWindow
    if (!frameDocument || !frameWindow) throw new Error('无法创建打印文档')
    frameDocument.open()
    frameDocument.write(printableHtml(model))
    frameDocument.close()
    await new Promise<void>((resolve) => {
      const images = Array.from(frameDocument.images)
      if (!images.length) { resolve(); return }
      let remaining = images.length
      const done = () => { remaining -= 1; if (remaining <= 0) resolve() }
      images.forEach((image) => {
        if (image.complete) done()
        else { image.addEventListener('load', done, { once: true }); image.addEventListener('error', done, { once: true }) }
      })
      window.setTimeout(resolve, 4000)
    })
    await frameDocument.fonts?.ready
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
    frameWindow.addEventListener('afterprint', cleanup, { once: true })
    frameWindow.focus()
    frameWindow.print()
    window.setTimeout(cleanup, 1000)
  } catch (error) {
    cleanup()
    throw error
  }
}
