import { isUnsignedCount } from '../platform/wasm/task-protocol.ts'
import { isRecord } from './protocol.ts'
import type { SlicerMembersOutput, SlicerMembersParams } from './slicer-member-protocol.ts'

function owned(view: Uint8Array | Uint32Array): boolean {
  if (!(view.buffer instanceof ArrayBuffer) || view.byteOffset !== 0 || view.byteLength !== view.buffer.byteLength) return false
  // Detached empty views also report zero extents. A zero-length constructor
  // probe distinguishes them from real owned empty window buffers.
  try { new Uint8Array(view.buffer, 0, 0); return true } catch { return false }
}

/** Dimensions/ownership plus a bounded window scan. Never scan the full source
 * on the host when accepting a result from the single Rust member authority.
 */
export function readSlicerMembersOutput(value: unknown, params: SlicerMembersParams): SlicerMembersOutput | null {
  if (!isRecord(value) || value.kind !== 'slicerMembers' || !(value.indices instanceof Uint32Array)
    || !(value.hasData instanceof Uint8Array) || !owned(value.indices) || !owned(value.hasData)
    || value.indices.buffer === value.hasData.buffer || value.indices.length !== value.hasData.length
    || !isUnsignedCount(value.total) || value.total > params.rowCount
    || !isUnsignedCount(value.dateCount) || value.dateCount > params.rowCount) return null
  const length = params.mode === 'metadata' ? 0 : Math.min(params.count, Math.max(0, value.total - params.start))
  if (value.indices.length !== length || params.mode === 'metadata' && value.total !== 0) return null
  if (value.dateCount === 0 ? value.minDate !== null || value.maxDate !== null
    : typeof value.minDate !== 'number' || !Number.isFinite(value.minDate)
      || typeof value.maxDate !== 'number' || !Number.isFinite(value.maxDate) || value.minDate > value.maxDate) return null
  const seen = new Set<number>()
  for (let index = 0; index < length; index++) {
    const row = value.indices[index], available = value.hasData[index]
    if (row >= params.rowCount || seen.has(row) || (params.mode === 'canvas' ? available > 5 || (available & 6) === 6 : available > 1)
      || (params.crossFilter === 'none' || params.hideItemsWithNoData) && (available & 1) !== 1) return null
    seen.add(row)
  }
  return { kind: 'slicerMembers', indices: value.indices, hasData: value.hasData,
    total: value.total, dateCount: value.dateCount, minDate: value.minDate as number | null, maxDate: value.maxDate as number | null }
}
