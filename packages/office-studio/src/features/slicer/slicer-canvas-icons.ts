import { h, render } from 'vue'
import { ChevronDown, ChevronUp, FilterX, ListChecks } from 'lucide-vue-next'

export type SlicerIconName = 'clear' | 'multi' | 'moreUp' | 'moreDown'
const icons = { clear: FilterX, multi: ListChecks, moreUp: ChevronUp, moreDown: ChevronDown }
const images = new Map<string, HTMLImageElement>()
const pending = new Map<string, Set<() => void>>()

/** Canvas actions use the public icon components used by the workbench. */
export function slicerCanvasIcon(name: SlicerIconName, color: string, onReady?: () => void): HTMLImageElement | undefined {
  const key = `${name}:${color}`
  const cached = images.get(key)
  if (cached || typeof document === 'undefined') return cached
  const waiting = pending.get(key)
  if (waiting) { if (onReady) waiting.add(onReady); return undefined }
  pending.set(key, new Set(onReady ? [onReady] : []))
  const host = document.createElement('div')
  render(h(icons[name], { size: 20, color, strokeWidth: 1.5 }), host)
  const svg = host.firstElementChild
  if (!svg) { pending.delete(key); render(null, host); return undefined }
  const image = new Image()
  image.onload = () => {
    images.set(key, image)
    const listeners = pending.get(key)
    pending.delete(key)
    listeners?.forEach(listener => listener())
  }
  image.onerror = () => { pending.delete(key) }
  image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(new XMLSerializer().serializeToString(svg))}`
  render(null, host)
  return undefined
}
