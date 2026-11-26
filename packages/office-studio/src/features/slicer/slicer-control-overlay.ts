import { drawingAffinePoint } from '../../domain/drawing-affine.ts'
import type { SheetDrawingAffineGeometry, SheetDrawingLeafRef } from '../../domain/sheet-drawings.ts'
import { drawingPaintOrder, drawingRefKey, type SheetDrawingSceneInput } from '../drawing-groups/index.ts'
import { sheetDrawingPaneLayouts, type SheetDrawingPaneLayout } from '../sheet-drawings/sheet-drawing-viewport.ts'
import { clipRenderTexturePolygon, renderRectPoints, type RenderPoint } from '../../rendering/affine-command-geometry.ts'
import type { GridViewport, SpreadsheetGeometry } from '../../rendering/spreadsheet/layout.ts'
import type { RenderRect } from '../../rendering/contracts.ts'
import type { SlicerObjectLayout } from './slicer-renderer.ts'

export interface SlicerControlPaneLayout {
  key: string
  layout: SlicerObjectLayout
  pane: SheetDrawingPaneLayout
  /** Screen polygons retain the Canvas owner order, including overlapping occluders. */
  polygons: RenderPoint[][]
  clipPath: string
}

function signedArea(points: readonly RenderPoint[]): number {
  const origin = points[0]
  if (!origin) return 0
  return points.reduce((sum, point, index) => {
    const next = points[(index + 1) % points.length]!
    return sum + (point.x - origin.x) * (next.y - origin.y) - (next.x - origin.x) * (point.y - origin.y)
  }, 0) / 2
}
function useful(points: RenderPoint[]): boolean { return points.length >= 3 && Math.abs(signedArea(points)) > 1e-8 }
function halfPlane(points: readonly RenderPoint[], first: RenderPoint, last: RenderPoint, winding: number, inside: boolean): RenderPoint[] {
  if (!points.length) return []
  const result: RenderPoint[] = [], distance = (point: RenderPoint) =>
    winding * ((last.x - first.x) * (point.y - first.y) - (last.y - first.y) * (point.x - first.x))
  let previous = points.at(-1)!, before = distance(previous)
  for (const point of points) {
    const after = distance(point), beforeKept = inside ? before >= 0 : before <= 0, afterKept = inside ? after >= 0 : after <= 0
    if (beforeKept !== afterKept) {
      const t = before / (before - after)
      result.push({ x: previous.x + (point.x - previous.x) * t, y: previous.y + (point.y - previous.y) * t })
    }
    if (afterKept) result.push({ ...point })
    previous = point; before = after
  }
  return result
}
/** Subtract one convex drawing footprint, emitting disjoint convex pieces without an even/odd overlap hole. */
export function subtractSlicerControlPolygon(subject: readonly RenderPoint[], occluder: readonly RenderPoint[]): RenderPoint[][] {
  const area = signedArea(occluder)
  if (!Number.isFinite(area) || Math.abs(area) <= 1e-8) return useful([...subject]) ? [subject.map(point => ({ ...point }))] : []
  const result: RenderPoint[][] = [], winding = area > 0 ? 1 : -1
  let remaining = subject.map(point => ({ ...point }))
  for (let index = 0; index < occluder.length && remaining.length; index++) {
    const first = occluder[index]!, last = occluder[(index + 1) % occluder.length]!, outside = halfPlane(remaining, first, last, winding, false)
    if (useful(outside)) result.push(outside)
    remaining = halfPlane(remaining, first, last, winding, true)
  }
  return result
}
function clipped(points: readonly RenderPoint[], clip: readonly RenderPoint[]): RenderPoint[] {
  return clipRenderTexturePolygon(points.map(point => ({ ...point, u: 0, v: 0 })), clip).map(({ x, y }) => ({ x, y }))
}
function cssPath(polygons: readonly RenderPoint[][], clip: RenderRect): string {
  const number = (value: number) => String(Math.round(value * 1e6) / 1e6)
  const path = polygons.map(polygon => {
    const points = signedArea(polygon) < 0 ? [...polygon].reverse() : polygon
    return points.map((point, index) => `${index ? 'L' : 'M'}${number(point.x - clip.x)} ${number(point.y - clip.y)}`).join(' ') + ' Z'
  }).join(' ')
  return `path(nonzero, "${path}")`
}

/** Ant controls reuse the Canvas affine, pane clipping, preview and complete leaf paint order. */
export function slicerControlPaneLayouts(
  grid: SpreadsheetGeometry, viewport: GridViewport, layouts: readonly SlicerObjectLayout[], scene?: SheetDrawingSceneInput,
  previews?: ReadonlyMap<string, SheetDrawingAffineGeometry>, visible?: (ref: SheetDrawingLeafRef, geometry: SheetDrawingAffineGeometry) => boolean,
): SlicerControlPaneLayout[] {
  const owners = scene ? drawingPaintOrder(scene).flatMap(ref => {
    const object = scene.objects.find(object => drawingRefKey(object.ref) === drawingRefKey(ref)),
      geometry = previews?.get(drawingRefKey(ref)) ?? object?.geometry
    return object && object.available !== false && geometry && (!visible || visible(ref, geometry)) ? [{ ref, geometry }] : []
  }) : layouts.map(layout => ({ ref: { kind: layout.model.type, id: layout.id } as SheetDrawingLeafRef, geometry: layout.geometry }))
  const result: SlicerControlPaneLayout[] = []
  for (const layout of layouts) {
    const ref: SheetDrawingLeafRef = { kind: layout.model.type, id: layout.id }, order = owners.findIndex(owner => drawingRefKey(owner.ref) === drawingRefKey(ref))
    if (order < 0 || visible && !visible(ref, layout.geometry)) continue
    const later = owners.slice(order + 1).flatMap(owner => sheetDrawingPaneLayouts(grid, viewport, owner.geometry))
    for (const pane of sheetDrawingPaneLayouts(grid, viewport, layout.geometry)) {
      const polygon = clipped(renderRectPoints({ x: 0, y: 0, width: layout.geometry.width, height: layout.geometry.height }, pane.frame.matrix), renderRectPoints(pane.clip))
      let polygons = useful(polygon) ? [polygon] : []
      for (const higher of later) {
        const footprint = clipped(renderRectPoints({ x: 0, y: 0, width: higher.frame.width, height: higher.frame.height }, higher.frame.matrix), renderRectPoints(higher.clip))
        if (useful(footprint)) polygons = polygons.flatMap(subject => subtractSlicerControlPolygon(subject, footprint))
        if (!polygons.length) break
      }
      if (polygons.length) result.push({ key: `${drawingRefKey(ref)}:${pane.frozenX ? 1 : 0}:${pane.frozenY ? 1 : 0}`,
        layout, pane, polygons, clipPath: cssPath(polygons, pane.clip) })
    }
  }
  return result
}
export function slicerControlScreenPolygon(pane: SlicerControlPaneLayout, rect: RenderRect): RenderPoint[] {
  return renderRectPoints({ x: rect.x - pane.layout.rect.x, y: rect.y - pane.layout.rect.y, width: rect.width, height: rect.height }, pane.pane.frame.matrix)
}
export function slicerControlVisible(pane: SlicerControlPaneLayout, rect: RenderRect): boolean {
  const control = slicerControlScreenPolygon(pane, rect)
  return rect.width > 0 && rect.height > 0 && pane.polygons.some(polygon => useful(clipped(control, polygon)))
}
/** Only one frozen-pane copy of a logical control participates in Tab navigation. */
export function slicerControlTabIndex(panes: readonly SlicerControlPaneLayout[], pane: SlicerControlPaneLayout, rect: RenderRect): 0 | -1 {
  const candidates = panes.filter(candidate => candidate.layout.id === pane.layout.id && candidate.layout.model.type === pane.layout.model.type
    && slicerControlVisible(candidate, rect))
  const centerIn = (candidate: SlicerControlPaneLayout) => {
    const center = drawingAffinePoint(candidate.pane.frame.matrix, rect.x - pane.layout.rect.x + rect.width / 2, rect.y - pane.layout.rect.y + rect.height / 2)
    return candidate.polygons.some(polygon => {
      const winding = signedArea(polygon) >= 0 ? 1 : -1
      return polygon.every((first, index) => {
        const last = polygon[(index + 1) % polygon.length]!
        return winding * ((last.x - first.x) * (center.y - first.y) - (last.y - first.y) * (center.x - first.x)) >= -1e-8
      })
    })
  }
  return (candidates.find(centerIn) ?? candidates[0])?.key === pane.key ? 0 : -1
}
