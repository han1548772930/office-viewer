import type { ChartPointShape, ChartRelationForceOptions, ChartRelationLayout, ChartRelationOptions, ChartType } from '../../domain/model'
import { isChartPointShape } from './chart-markers'

export const ChartRelationLayouts = Object.freeze({ None: 'none', Circular: 'circular', Force: 'force' } as const)
export const CHART_RELATION_DEFAULTS = Object.freeze({ layout: 'none' as const, useValueAsSymbolSize: true,
  emphasisEnabled: true, circularLabelRotation: false, nodeShape: 'circle' as const, gravity: 0.1, repulsion: 5000 })

export function assertChartRelationType(type: ChartType): void {
  if (type !== 'relation') throw new TypeError('此操作仅适用于关系图。')
}
export function chartRelationLayoutValue(value: unknown): ChartRelationLayout {
  if (value !== 'none' && value !== 'circular' && value !== 'force') throw new TypeError('关系图布局无效。')
  return value
}
export function chartRelationBoolean(value: unknown): boolean {
  if (typeof value !== 'boolean') throw new TypeError('关系图开关必须为布尔值。')
  return value
}
export function chartRelationNodeShape(value: unknown): ChartPointShape {
  if (!isChartPointShape(value)) throw new TypeError('关系图节点形状无效。')
  return value
}
/** Omitted force fields preserve author overrides; reset removes the group. */
export function chartRelationForce(value: ChartRelationForceOptions): ChartRelationForceOptions {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError('关系图力导向参数必须为对象。')
  const limits = { repulsion: 100_000, gravity: 10, edgeLength: 10_000 }
  const result: ChartRelationForceOptions = {}
  for (const key of ['repulsion', 'gravity', 'edgeLength'] as const) {
    const current = value[key]
    if (current !== undefined && (typeof current !== 'number' || !Number.isFinite(current) || current < 0 || current > limits[key])) throw new RangeError(`关系图 ${key} 参数无效。`)
    if (current !== undefined) result[key] = current
  }
  return result
}
export function resolvedChartRelationForce(options: ChartRelationOptions | undefined, width: number, height: number): Required<ChartRelationForceOptions> {
  return { repulsion: options?.force?.repulsion ?? CHART_RELATION_DEFAULTS.repulsion,
    gravity: options?.force?.gravity ?? CHART_RELATION_DEFAULTS.gravity,
    edgeLength: options?.force?.edgeLength ?? Math.min(width, height) / 5 }
}

/** The fixed None layout seeds non-central nodes from UTF-16 names. */
export function chartRelationNameSeed(name: string): number {
  let hash = 0
  for (let index = 0; index < name.length; index++) hash = (Math.imul(hash, 31) + name.charCodeAt(index)) | 0
  return Math.abs(hash)
}
