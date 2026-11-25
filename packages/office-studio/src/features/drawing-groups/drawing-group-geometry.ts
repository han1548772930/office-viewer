import type { SheetDrawingAffineGeometry, SheetDrawingMatrix } from '../../domain/sheet-drawings.ts'
import { drawingAffineBounds, drawingAffinePoint, invertDrawingAffine, multiplyDrawingAffine, normalizeDrawingAffineGeometry, normalizeDrawingAffineMatrix } from '../../domain/drawing-affine.ts'
import { SheetDrawingError } from './drawing-group-types.ts'

export function cloneDrawingGeometry(value: SheetDrawingAffineGeometry): SheetDrawingAffineGeometry {
  return { width: value.width, height: value.height, matrix: { ...value.matrix } }
}

export function normalizeDrawingGeometry(value: unknown): SheetDrawingAffineGeometry {
  try { return normalizeDrawingAffineGeometry(value) }
  catch (error) { throw new SheetDrawingError('invalid-geometry', error instanceof Error ? error.message : 'Invalid drawing geometry.') }
}

export function normalizeDrawingMatrix(value: unknown): SheetDrawingMatrix {
  try { return normalizeDrawingAffineMatrix(value) }
  catch (error) { throw new SheetDrawingError('invalid-geometry', error instanceof Error ? error.message : 'Invalid drawing matrix.') }
}

export function drawingGeometryEqual(left: SheetDrawingAffineGeometry, right: SheetDrawingAffineGeometry): boolean {
  const equal = (a: number, b: number) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b))
  return equal(left.width, right.width) && equal(left.height, right.height)
    && (['a', 'b', 'c', 'd', 'e', 'f'] as const).every(key => equal(left.matrix[key], right.matrix[key]))
}

export function drawingGeometryCorners(geometry: SheetDrawingAffineGeometry): { x: number; y: number }[] {
  return [[0, 0], [geometry.width, 0], [geometry.width, geometry.height], [0, geometry.height]]
    .map(([x, y]) => drawingAffinePoint(geometry.matrix, x!, y!))
}

/** Fit in the existing group's basis to retain its rotation, reflection, and shear during membership changes. */
export function drawingGroupFrame(geometries: readonly SheetDrawingAffineGeometry[], basis?: SheetDrawingAffineGeometry): SheetDrawingAffineGeometry {
  if (!geometries.length) throw new SheetDrawingError('invalid-geometry', 'A drawing group requires nonempty geometry.')
  const inverse = basis ? invertDrawingAffine(basis.matrix) : undefined
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity
  for (const geometry of geometries) for (const world of drawingGeometryCorners(geometry)) {
    const point = inverse ? drawingAffinePoint(inverse, world.x, world.y) : world
    left = Math.min(left, point.x); top = Math.min(top, point.y)
    right = Math.max(right, point.x); bottom = Math.max(bottom, point.y)
  }
  const spanX = right - left, spanY = bottom - top
  // Large worksheet groups use a bounded local frame; the matrix retains the full world extent.
  const width = Math.min(spanX, 100_000), height = Math.min(spanY, 100_000)
  const local = { a: spanX / width, b: 0, c: 0, d: spanY / height, e: left, f: top }
  return normalizeDrawingGeometry({ width, height, matrix: basis ? multiplyDrawingAffine(basis.matrix, local) : local })
}

export function transformDrawingGeometry(geometry: SheetDrawingAffineGeometry, delta: SheetDrawingMatrix): SheetDrawingAffineGeometry {
  return normalizeDrawingGeometry({ width: geometry.width, height: geometry.height, matrix: multiplyDrawingAffine(delta, geometry.matrix) })
}

/** A resized target frame includes its local-size ratio as well as its world matrix. */
export function drawingGroupFrameDelta(before: SheetDrawingAffineGeometry, after: SheetDrawingAffineGeometry): SheetDrawingMatrix {
  const scale = { a: after.width / before.width, b: 0, c: 0, d: after.height / before.height, e: 0, f: 0 }
  return normalizeDrawingMatrix(multiplyDrawingAffine(multiplyDrawingAffine(after.matrix, scale), invertDrawingAffine(before.matrix)))
}

export function drawingGeometryContains(geometry: SheetDrawingAffineGeometry, x: number, y: number): boolean {
  if (!Number.isFinite(x) || !Number.isFinite(y)) return false
  const point = drawingAffinePoint(invertDrawingAffine(geometry.matrix), x, y)
  return point.x >= 0 && point.y >= 0 && point.x <= geometry.width && point.y <= geometry.height
}

export { drawingAffineBounds as drawingGeometryBounds }
