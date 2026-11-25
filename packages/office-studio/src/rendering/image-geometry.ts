import type { ImageCommand, RenderRect } from './contracts.ts'
import { drawingAffinePoint } from '../domain/drawing-affine.ts'

export function imageSourceDimensions(source: CanvasImageSource): { width: number; height: number } {
  return { width: 'naturalWidth' in source ? source.naturalWidth : 'videoWidth' in source ? source.videoWidth : 'width' in source ? Number(source.width) : 'displayWidth' in source ? source.displayWidth : 1,
    height: 'naturalHeight' in source ? source.naturalHeight : 'videoHeight' in source ? source.videoHeight : 'height' in source ? Number(source.height) : 'displayHeight' in source ? source.displayHeight : 1 }
}
/** Shared clipping and transform math for both render backends. Source coordinates are normalized. */
export function imageCommandGeometry(command: ImageCommand): { rect: RenderRect; sourceRect: RenderRect; points: number[]; uv: number[][] } | undefined {
  const crop = command.sourceRect ?? { x: 0, y: 0, width: 1, height: 1 }
  if (![command.x, command.y, command.width, command.height, crop.x, crop.y, crop.width, crop.height, command.rotation ?? 0].every(Number.isFinite)
    || command.width <= 0 || command.height <= 0 || crop.width <= 0 || crop.height <= 0) return undefined
  const matrix = command.matrix
  if (matrix && (![matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f].every(Number.isFinite)
    || matrix.a * matrix.d - matrix.b * matrix.c === 0)) return undefined
  const x = Math.max(0, crop.x), y = Math.max(0, crop.y), right = Math.min(1, crop.x + crop.width), bottom = Math.min(1, crop.y + crop.height)
  if (right <= x || bottom <= y) return undefined
  const rect = { x: command.x + (x - crop.x) / crop.width * command.width, y: command.y + (y - crop.y) / crop.height * command.height,
    width: (right - x) / crop.width * command.width, height: (bottom - y) / crop.height * command.height }
  const centerX = command.x + command.width / 2, centerY = command.y + command.height / 2, radians = ((command.rotation ?? 0) % 360) * Math.PI / 180
  const cos = Math.cos(radians), sin = Math.sin(radians)
  const point = (x: number, y: number) => {
    const dx = (x - centerX) * (command.flipX ? -1 : 1), dy = (y - centerY) * (command.flipY ? -1 : 1)
    const px = centerX + dx * cos - dy * sin, py = centerY + dx * sin + dy * cos,
      transformed = matrix ? drawingAffinePoint(matrix, px, py) : { x: px, y: py }
    return [transformed.x, transformed.y]
  }
  return { rect, sourceRect: { x, y, width: right - x, height: bottom - y },
    points: [...point(rect.x, rect.y), ...point(rect.x + rect.width, rect.y), ...point(rect.x + rect.width, rect.y + rect.height), ...point(rect.x, rect.y + rect.height)],
    uv: [[x, y], [right, y], [right, bottom], [x, bottom]] }
}
