import type { ApiRange } from './events.ts'
import type { SheetSelectionCell, SheetSelectionOptions } from '../features/sheet-workbench/sheet-selection.ts'
import type { SheetSelectionDirection, SheetSelectionShortcutKey } from '../features/sheet-workbench/sheet-selection-navigation.ts'

export const SET_SHEET_SELECTIONS_OPERATION = 'sheet.operation.set-selections'
export const SELECT_SHEET_RANGE_COMMAND = 'sheet.command.select-range'
export const MOVE_SHEET_SELECTION_ENTER_TAB_COMMAND = 'sheet.command.move-selection-enter-tab'

export interface SheetSelectionOperationRequest {
  unitId: string
  sheetId: string
  ranges: ApiRange[]
  options: SheetSelectionOptions
}
export interface SheetSelectionShortcutRequest { unitId: string; sheetId: string; direction: SheetSelectionDirection; key: SheetSelectionShortcutKey }

function record(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${label}必须是对象。`)
  return value as Record<string, unknown>
}
function identifier(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim()) throw new TypeError(`${label}无效。`)
  return value
}
function coordinate(value: unknown): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value)) throw new RangeError('选区坐标必须是整数。')
  return value
}
function primaryCell(value: unknown): SheetSelectionCell {
  const primary = record(value, '活动单元格')
  return { row: coordinate(primary.row ?? primary.actualRow ?? primary.startRow),
    column: coordinate(primary.column ?? primary.actualColumn ?? primary.startColumn) }
}
function range(value: unknown, unitId: string, sheetId: string): ApiRange {
  const source = record(value, '选区')
  if (source.unitId !== undefined && source.unitId !== unitId || source.sheetId !== undefined && source.sheetId !== sheetId) {
    throw new Error('选区必须属于指定的工作簿和工作表。')
  }
  const rangeType = source.rangeType
  if (rangeType !== undefined && rangeType !== 0 && rangeType !== 1 && rangeType !== 2 && rangeType !== 3) throw new TypeError('选区类型无效。')
  return { startRow: coordinate(source.startRow), endRow: coordinate(source.endRow),
    startColumn: coordinate(source.startColumn ?? source.startCol), endColumn: coordinate(source.endColumn ?? source.endCol),
    unitId, sheetId, ...(rangeType === undefined ? {} : { rangeType }) }
}

/** Accept the project's range list and the fixed-version Univer selection operation envelope. */
export function parseSheetSelectionOperation(value: unknown, single = false): SheetSelectionOperationRequest {
  const input = record(value, '选区命令参数'), unitId = identifier(input.unitId, '工作簿标识')
  const sheetId = identifier(input.sheetId ?? input.subUnitId ?? input.subUnit, '工作表标识')
  for (const id of [input.sheetId, input.subUnitId, input.subUnit]) if (id !== undefined && id !== sheetId) throw new TypeError('工作表标识不一致。')
  const parameters = input.options === undefined ? input : record(input.options, '选区选项')
  const options: SheetSelectionOptions = {}
  if (parameters.activeRangeIndex !== undefined) options.activeRangeIndex = coordinate(parameters.activeRangeIndex)
  if (parameters.primary !== undefined) options.primary = primaryCell(parameters.primary)
  let values: unknown[]
  if (single) values = [input.range]
  else {
    const entries = input.ranges ?? input.selections
    if (!Array.isArray(entries)) throw new TypeError('选区必须是范围数组。')
    values = entries
  }
  let descriptorPrimary = false
  const ranges = values.map((value, index) => {
    const source = record(value, '选区')
    if (Object.hasOwn(source, 'range')) {
      if (!descriptorPrimary && source.primary !== undefined && source.primary !== null && options.primary === undefined) {
        options.primary = primaryCell(source.primary)
        options.activeRangeIndex ??= index
        descriptorPrimary = true
      }
      return range(source.range, unitId, sheetId)
    }
    return range(source, unitId, sheetId)
  })
  return { unitId, sheetId, ranges, options }
}

export function parseSheetSelectionShortcut(value: unknown, target: { unitId: string; sheetId: string },
  resolveActiveSheet?: (unitId: string) => string | undefined): SheetSelectionShortcutRequest {
  const input = record(value, '选区移动命令参数')
  const unitId = identifier(input.unitId ?? target.unitId, '工作簿标识')
  const explicitSheet = input.sheetId ?? input.subUnitId ?? input.subUnit
  const sheetId = identifier(explicitSheet ?? (resolveActiveSheet ? resolveActiveSheet(unitId)
    : unitId === target.unitId ? target.sheetId : undefined), '工作表标识')
  for (const id of [input.sheetId, input.subUnitId, input.subUnit]) if (id !== undefined && id !== sheetId) throw new TypeError('工作表标识不一致。')
  const direction = input.direction
  if (direction !== 0 && direction !== 1 && direction !== 2 && direction !== 3) throw new TypeError('选区移动方向无效。')
  const key = input.key ?? (input.keycode === 9 ? 'Tab' : input.keycode === 13 ? 'Enter' : undefined)
  if (key !== 'Tab' && key !== 'Enter' || key === 'Tab' && direction !== 1 && direction !== 3 || key === 'Enter' && direction !== 0 && direction !== 2) {
    throw new TypeError('选区移动按键与方向不一致。')
  }
  if (input.keycode !== undefined && input.keycode !== (key === 'Tab' ? 9 : 13)) throw new TypeError('选区移动按键参数不一致。')
  return { unitId, sheetId, direction, key }
}
