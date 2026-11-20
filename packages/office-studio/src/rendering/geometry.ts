import type { RenderRect } from './contracts'

export function intersectRect(a: RenderRect, b: RenderRect): RenderRect {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y)
  return { x, y, width: Math.max(0, Math.min(a.x + a.width, b.x + b.width) - x), height: Math.max(0, Math.min(a.y + a.height, b.y + b.height) - y) }
}

export function containsPoint(rect: RenderRect, x: number, y: number): boolean {
  return x >= rect.x && y >= rect.y && x < rect.x + rect.width && y < rect.y + rect.height
}

export function parseColor(value: string): [number, number, number, number] {
  const hex = /^#([0-9a-f]{3,8})$/i.exec(value.trim())?.[1]
  if (hex) {
    const full = hex.length === 3 || hex.length === 4 ? [...hex].map(x => x + x).join('') : hex
    return [parseInt(full.slice(0, 2), 16) / 255, parseInt(full.slice(2, 4), 16) / 255, parseInt(full.slice(4, 6), 16) / 255, full.length === 8 ? parseInt(full.slice(6, 8), 16) / 255 : 1]
  }
  const rgba = /^rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:\s*[,/]\s*([\d.]+))?\s*\)$/i.exec(value.trim())
  if (rgba) return [Number(rgba[1]) / 255, Number(rgba[2]) / 255, Number(rgba[3]) / 255, rgba[4] === undefined ? 1 : Number(rgba[4])]
  const names: Record<string, [number, number, number, number]> = { transparent: [0, 0, 0, 0], white: [1, 1, 1, 1], black: [0, 0, 0, 1], red: [1, 0, 0, 1], blue: [0, 0, 1, 1] }
  return names[value.toLowerCase()] ?? [0, 0, 0, 1]
}
