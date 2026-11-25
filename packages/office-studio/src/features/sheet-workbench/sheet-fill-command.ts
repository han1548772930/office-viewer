import type { CellRange } from '../../domain/model.ts'
import { assertSheetCommandRanges } from './sheet-range-command.ts'
import type { SheetFillApplyType } from './sheet-fill.ts'

export interface SheetFillCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  sourceRange?: CellRange
  targetRange?: CellRange
  range?: CellRange
  clearRange?: CellRange
  selectionRange?: CellRange
  applyType?: SheetFillApplyType
  /** Univer's RefillCommand uses `type`; normalize it to applyType. */
  type?: SheetFillApplyType
  sessionId?: string
  expectedRevision?: number
  expectedHistoryGeneration?: number
  host?: { unitId: string; sheetId: string }
}
function record(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}必须是对象。`)
  return value as Record<string, unknown>
}
function identifier(value: unknown, label: string): string | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${label}无效。`)
  return value
}
function coordinate(value: unknown): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value)) throw new RangeError('填充范围坐标必须是整数。')
  return value
}
function range(value: unknown, input: SheetFillCommandInput): CellRange {
  const data = record(value, '填充范围')
  if (data.unitId !== undefined && input.unitId !== undefined && data.unitId !== input.unitId
    || data.sheetId !== undefined && (input.sheetId ?? input.subUnitId) !== undefined && data.sheetId !== (input.sheetId ?? input.subUnitId)) {
    throw new TypeError('填充范围必须属于目标工作簿和工作表。')
  }
  if (data.startCol !== undefined && data.startColumn !== undefined && data.startCol !== data.startColumn
    || data.endCol !== undefined && data.endColumn !== undefined && data.endCol !== data.endColumn) throw new TypeError('填充列坐标不一致。')
  const result = { startRow: coordinate(data.startRow), endRow: coordinate(data.endRow),
    startCol: coordinate(data.startCol ?? data.startColumn), endCol: coordinate(data.endCol ?? data.endColumn) }
  assertSheetCommandRanges([result])
  return result
}
export function sheetFillApplyType(value: unknown): SheetFillApplyType | undefined {
  if (value === undefined) return undefined
  if (value === 'COPY' || value === 'SERIES' || value === 'ONLY_FORMAT' || value === 'NO_FORMAT') return value
  throw new TypeError('自动填充模式无效。')
}
export function parseSheetFillCommandInput(value: unknown): SheetFillCommandInput {
  if (value === undefined) return {}
  const data = record(value, '填充参数'), result: SheetFillCommandInput = {
    unitId: identifier(data.unitId, '工作簿标识'), sheetId: identifier(data.sheetId, '工作表标识'),
    subUnitId: identifier(data.subUnitId, '工作表标识'), sessionId: identifier(data.sessionId, '填充会话标识'),
    applyType: sheetFillApplyType(data.applyType ?? data.type),
  }
  if (data.applyType !== undefined && data.type !== undefined && sheetFillApplyType(data.type) !== result.applyType) {
    throw new TypeError('填充模式参数不一致。')
  }
  if (result.sheetId !== undefined && result.subUnitId !== undefined && result.sheetId !== result.subUnitId) throw new TypeError('工作表标识不一致。')
  for (const key of ['sourceRange', 'targetRange', 'range', 'clearRange', 'selectionRange'] as const) {
    if (data[key] !== undefined) result[key] = range(data[key], result)
  }
  for (const key of ['expectedRevision', 'expectedHistoryGeneration'] as const) if (data[key] !== undefined) {
    const current = coordinate(data[key])
    if (current < (key === 'expectedRevision' ? 0 : 1)) throw new RangeError('填充内容或历史版本无效。')
    result[key] = current
  }
  if (data.host !== undefined) {
    const host = record(data.host, '填充宿主'), unitId = identifier(host.unitId, '宿主工作簿标识'), sheetId = identifier(host.sheetId, '宿主工作表标识')
    if (!unitId || !sheetId) throw new TypeError('填充宿主标识无效。')
    result.host = { unitId, sheetId }
  }
  return result
}
