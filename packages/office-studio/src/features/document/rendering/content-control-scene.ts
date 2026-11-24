import type { RenderScene } from '../../../rendering/contracts'
import { contentControlFragmentVisibleRect, type DocumentContentControlGeometry } from '../layout/content-control-geometry'
import type { DocumentLayout } from '../layout/document-layout-types'
import { contentControlDisplayColor } from '../services/content-control-presentation'
import { sceneOffset, type DocumentViewport } from './document-scene'

export function appendDocumentContentControlScene(scene: RenderScene, geometries: readonly DocumentContentControlGeometry[],
  layout: DocumentLayout, viewport: DocumentViewport, options: { activeControlId?: string; primary: string; border: string }): void {
  const scale = viewport.zoom / 100, offset = sceneOffset(layout, viewport)
  const rect = (source: { x: number; y: number; width: number; height: number }) => ({ x: offset.x + source.x * scale,
    y: offset.y + source.y * scale, width: source.width * scale, height: source.height * scale })
  for (const geometry of geometries) {
    const control = geometry.control
    if (control.appearance === 'hidden') continue
    const active = control.id === options.activeControlId, color = contentControlDisplayColor(control, active ? options.primary : options.border)
    for (const fragment of geometry.fragments) {
      const visible = contentControlFragmentVisibleRect(fragment)
      if (!visible) continue
      const area = rect(visible)
      if (area.x + area.width < 0 || area.y + area.height < 0 || area.x > viewport.width || area.y > viewport.height) continue
      const bounds = rect(fragment), clip = fragment.clip ? rect(fragment.clip) : undefined, width = active ? 1.5 : 1
      if (control.appearance === 'tags') {
        for (const edge of [{ enabled: fragment.first, x: bounds.x, direction: 1 }, { enabled: fragment.last, x: bounds.x + bounds.width, direction: -1 }]) {
          if (!edge.enabled) continue
          scene.commands.push({ kind: 'line', x1: edge.x, x2: edge.x, y1: bounds.y, y2: bounds.y + bounds.height, color, width, clip })
          for (const y of [bounds.y, bounds.y + bounds.height]) scene.commands.push({ kind: 'line', x1: edge.x,
            x2: edge.x + 4 * edge.direction * scale, y1: y, y2: y, color, width, clip })
        }
      } else scene.commands.push({ kind: 'rect', ...bounds, fill: 'rgba(0,0,0,0)', stroke: color, strokeWidth: width, clip })
    }
  }
}
