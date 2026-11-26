import { drawingAffinePoint, invertDrawingAffine, type DrawingAffineMatrix } from '../domain/drawing-affine.ts'
import type { DrawingCommandOwner, RenderRect } from './contracts.ts'

export interface RenderPoint { x: number; y: number }
export interface RenderTextureVertex extends RenderPoint { u: number; v: number }

export function renderRectPoints(rect: RenderRect, matrix?: DrawingAffineMatrix): RenderPoint[] {
  return [[rect.x, rect.y], [rect.x + rect.width, rect.y], [rect.x + rect.width, rect.y + rect.height], [rect.x, rect.y + rect.height]]
    .map(([x, y]) => matrix ? drawingAffinePoint(matrix, x!, y!) : { x: x!, y: y! })
}

export function renderLocalClipPolygon(command: DrawingCommandOwner): RenderPoint[] | undefined {
  const clip = command.localClip
  if (!clip) return undefined
  if (clip.width <= 0 || clip.height <= 0) return []
  return renderRectPoints(clip, command.matrix)
}

/** A conservative local viewport for text-tile selection; exact clipping happens on emitted triangles. */
export function renderLocalViewportBounds(viewport: RenderRect, matrix: DrawingAffineMatrix): RenderRect {
  const points = renderRectPoints(viewport, invertDrawingAffine(matrix)), x = Math.min(...points.map(point => point.x)),
    y = Math.min(...points.map(point => point.y))
  return { x, y, width: Math.max(...points.map(point => point.x)) - x, height: Math.max(...points.map(point => point.y)) - y }
}

/** Clip a convex polygon and interpolate its texture coordinates; reflected clip winding is supported. */
export function clipRenderTexturePolygon(vertices: readonly RenderTextureVertex[], clip: readonly RenderPoint[]): RenderTextureVertex[] {
  if (clip.length < 3 || vertices.length < 3) return []
  const origin = clip[0]!
  const area = clip.reduce((sum, point, index) => {
    const next = clip[(index + 1) % clip.length]!
    return sum + (point.x - origin.x) * (next.y - origin.y) - (next.x - origin.x) * (point.y - origin.y)
  }, 0)
  if (!Number.isFinite(area) || area === 0) return []
  const winding = area > 0 ? 1 : -1
  let result = vertices.map(vertex => ({ ...vertex }))
  for (let edge = 0; edge < clip.length && result.length; edge++) {
    const first = clip[edge]!, last = clip[(edge + 1) % clip.length]!, distance = (point: RenderPoint) =>
      winding * ((last.x - first.x) * (point.y - first.y) - (last.y - first.y) * (point.x - first.x))
    const source = result
    result = []
    let previous = source[source.length - 1]!, previousDistance = distance(previous)
    for (const current of source) {
      const currentDistance = distance(current), previousInside = previousDistance >= 0, currentInside = currentDistance >= 0
      if (currentInside !== previousInside) {
        const t = previousDistance / (previousDistance - currentDistance)
        result.push({ x: previous.x + t * (current.x - previous.x), y: previous.y + t * (current.y - previous.y),
          u: previous.u + t * (current.u - previous.u), v: previous.v + t * (current.v - previous.v) })
      }
      if (currentInside) result.push(current)
      previous = current; previousDistance = currentDistance
    }
  }
  return result
}

/** Existing image geometry already applies its matrix, so callers omit it for those vertices. */
export function renderProjectedVertices(points: readonly number[], uv: readonly (readonly number[])[], matrix?: DrawingAffineMatrix): RenderTextureVertex[] {
  return uv.map((texture, index) => {
    const x = points[index * 2]!, y = points[index * 2 + 1]!, point = matrix ? drawingAffinePoint(matrix, x, y) : { x, y }
    return { ...point, u: texture[0]!, v: texture[1]! }
  })
}
