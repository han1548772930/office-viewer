/** Serializable drawing geometry. Matrices map local CSS pixels to host CSS pixels. */
export interface DrawingAffineMatrix { a: number; b: number; c: number; d: number; e: number; f: number }
export interface DrawingAffineGeometry { width: number; height: number; matrix: DrawingAffineMatrix }
export interface DrawingAffineBounds { left: number; top: number; width: number; height: number }
export const DRAWING_AFFINE_IDENTITY: Readonly<DrawingAffineMatrix> = Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 })

export function normalizeDrawingAffineMatrix(value: unknown): DrawingAffineMatrix {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('绘图变换矩阵必须是对象。')
  const raw = value as Record<string, unknown>, matrix = {} as DrawingAffineMatrix
  for (const key of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
    const component = raw[key]
    if (typeof component !== 'number' || !Number.isFinite(component) || Math.abs(component) > Number.MAX_SAFE_INTEGER) throw new RangeError('绘图变换矩阵无效。')
    matrix[key] = component
  }
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c
  if (!Number.isFinite(determinant) || determinant === 0) throw new RangeError('绘图变换矩阵不可逆。')
  return matrix
}
export function normalizeDrawingAffineGeometry(value: unknown): DrawingAffineGeometry {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('绘图几何必须是对象。')
  const raw = value as Record<string, unknown>
  if (typeof raw.width !== 'number' || typeof raw.height !== 'number' || !Number.isFinite(raw.width) || !Number.isFinite(raw.height)
    || raw.width <= 0 || raw.height <= 0 || raw.width > 100_000 || raw.height > 100_000) throw new RangeError('绘图局部尺寸无效。')
  const result = { width: raw.width, height: raw.height, matrix: normalizeDrawingAffineMatrix(raw.matrix) }
  const bounds = drawingAffineBounds(result)
  if (![bounds.left, bounds.top, bounds.width, bounds.height].every(value => Number.isFinite(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER)) throw new RangeError('绘图变换后的范围无效。')
  return result
}
/** Composition is left * right: right is applied first. */
export function multiplyDrawingAffine(left: DrawingAffineMatrix, right: DrawingAffineMatrix): DrawingAffineMatrix {
  return { a: left.a * right.a + left.c * right.b, b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d, d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e, f: left.b * right.e + left.d * right.f + left.f }
}
export function invertDrawingAffine(matrix: DrawingAffineMatrix): DrawingAffineMatrix {
  const valid = normalizeDrawingAffineMatrix(matrix), det = valid.a * valid.d - valid.b * valid.c
  const inverse = { a: valid.d / det, b: -valid.b / det, c: -valid.c / det, d: valid.a / det,
    e: (valid.c * valid.f - valid.d * valid.e) / det, f: (valid.b * valid.e - valid.a * valid.f) / det }
  return normalizeDrawingAffineMatrix(inverse)
}
export function drawingAffinePoint(matrix: DrawingAffineMatrix, x: number, y: number): { x: number; y: number } {
  return { x: matrix.a * x + matrix.c * y + matrix.e, y: matrix.b * x + matrix.d * y + matrix.f }
}
export function drawingAffineBounds(geometry: DrawingAffineGeometry): DrawingAffineBounds {
  const points = [[0, 0], [geometry.width, 0], [geometry.width, geometry.height], [0, geometry.height]]
    .map(([x, y]) => drawingAffinePoint(geometry.matrix, x!, y!)),
    left = Math.min(...points.map(point => point.x)), top = Math.min(...points.map(point => point.y))
  return { left, top, width: Math.max(...points.map(point => point.x)) - left, height: Math.max(...points.map(point => point.y)) - top }
}
/** The two displayed edge lengths and center; shear remains in the matrix. */
export function drawingAffineFrame(geometry: DrawingAffineGeometry): DrawingAffineBounds {
  const m = geometry.matrix, center = drawingAffinePoint(m, geometry.width / 2, geometry.height / 2),
    width = Math.hypot(m.a, m.b) * geometry.width, height = Math.hypot(m.c, m.d) * geometry.height
  return { left: center.x - width / 2, top: center.y - height / 2, width, height }
}
export function drawingAffineWithFrame(geometry: DrawingAffineGeometry, frame: DrawingAffineBounds): DrawingAffineGeometry {
  const before = drawingAffineFrame(geometry), width = geometry.width * frame.width / before.width,
    height = geometry.height * frame.height / before.height, matrix = { ...geometry.matrix },
    center = drawingAffinePoint(matrix, width / 2, height / 2)
  matrix.e += frame.left + frame.width / 2 - center.x; matrix.f += frame.top + frame.height / 2 - center.y
  return normalizeDrawingAffineGeometry({ width, height, matrix })
}
export function drawingAffineFromFrame(frame: DrawingAffineBounds, rotation = 0): DrawingAffineGeometry {
  if (!Number.isFinite(rotation)) throw new RangeError('绘图旋转角度无效。')
  const angle = rotation % 360 * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle),
    cx = frame.left + frame.width / 2, cy = frame.top + frame.height / 2
  return normalizeDrawingAffineGeometry({ width: frame.width, height: frame.height,
    matrix: { a: cos, b: sin, c: -sin, d: cos, e: cx - cos * frame.width / 2 + sin * frame.height / 2,
      f: cy - sin * frame.width / 2 - cos * frame.height / 2 } })
}
