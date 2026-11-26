import type { OutlineAxis } from '../../domain/model.ts'
import type { SheetSelectionRange } from '../sheet-workbench/sheet-selection.ts'
import { assertOutlineAxis, getSheetOutlineUpdate, type ChangeSheetOutlineInput, type OutlineRange, type OutlineWorksheet } from './sheet-outline.ts'
import { DimensionOutlineErrorReason, SheetOutlineError, sheetOutlineErrorMessage } from './sheet-outline-types.ts'

export const SHEET_OUTLINE_CORE_COMMAND_IDS = [
  'sheet.command.add-dimension-outline', 'sheet.command.remove-dimension-outline',
  'sheet.command.set-dimension-outline-collapsed', 'sheet.command.clear-dimension-outlines',
] as const
export const SHEET_OUTLINE_UI_COMMAND_IDS = [
  'sheet.command.outline-selected-dimension', 'sheet.command.outline-selected-rows', 'sheet.command.outline-selected-columns',
  'sheet.command.clear-selected-dimension-outlines', 'sheet.command.clear-selected-rows-outlines', 'sheet.command.clear-selected-columns-outlines',
  'sheet.command.outline-selected-dimension-shortcut', 'sheet.command.clear-selected-dimension-outlines-shortcut',
] as const
export const SHEET_OUTLINE_LOCAL_COMMAND_IDS = [
  'sheet.command.set-outline-level', 'sheet.command.set-outline-summary-position', 'sheet.command.set-outline-summary-settings',
] as const
export const SHEET_OUTLINE_COMMAND_IDS = [...SHEET_OUTLINE_CORE_COMMAND_IDS, ...SHEET_OUTLINE_UI_COMMAND_IDS, ...SHEET_OUTLINE_LOCAL_COMMAND_IDS] as const

export interface SheetOutlineCommandInput {
  unitId?: string; subUnitId?: string; sheetId?: string; axis?: OutlineAxis; start?: number; end?: number; outlineId?: string;
  ranges?: readonly SheetSelectionRange[];
  collapsed?: boolean; value?: OutlineAxis; level?: number; atStart?: boolean; summaryAbove?: boolean; summaryLeft?: boolean;
  expectedRevision?: number; expectedHistoryGeneration?: number; expectedUserId?: string; host?: 'api' | 'workbench'
}
export interface SheetOutlineCommandContext {
  unitId: string; sheetId: string; sheet: OutlineWorksheet; userId: string; selections?: readonly SheetSelectionRange[]
}
export interface SheetOutlinePromptChoice extends OutlineRange {
  label: string; disabled: boolean; reason?: string; action: { kind: 'group' | 'remove'; axis: OutlineAxis; start: number; end: number }
}
export type SheetOutlineCommandResolution = { kind: 'change'; input: ChangeSheetOutlineInput }
  | { kind: 'prompt'; operation: 'group' | 'remove'; choices: SheetOutlinePromptChoice[] }

export function sheetOutlineCommandInput(value: unknown): SheetOutlineCommandInput {
  if (value === undefined) return {}
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw invalid('大纲命令参数必须是对象。')
  const input = value as SheetOutlineCommandInput
  for (const key of ['unitId', 'subUnitId', 'sheetId', 'outlineId', 'expectedUserId'] as const) {
    if (input[key] !== undefined && (typeof input[key] !== 'string' || !input[key]!.trim())) throw invalid(`大纲参数 ${key} 无效。`)
  }
  for (const key of ['start', 'end', 'level', 'expectedRevision', 'expectedHistoryGeneration'] as const) {
    if (input[key] !== undefined && (!Number.isSafeInteger(input[key]) || input[key]! < 0)) throw invalid(`大纲参数 ${key} 无效。`)
  }
  if (input.expectedHistoryGeneration === 0) throw invalid('大纲历史代次必须为正整数。')
  for (const key of ['collapsed', 'atStart', 'summaryAbove', 'summaryLeft'] as const) {
    if (input[key] !== undefined && typeof input[key] !== 'boolean') throw invalid(`大纲参数 ${key} 必须是布尔值。`)
  }
  if (input.axis !== undefined) assertOutlineAxis(input.axis)
  if (input.value !== undefined) assertOutlineAxis(input.value)
  if (input.host !== undefined && input.host !== 'api' && input.host !== 'workbench') throw invalid('大纲调用宿主无效。')
  if (input.sheetId !== undefined && input.subUnitId !== undefined && input.sheetId !== input.subUnitId) throw invalid('大纲工作表 ID 不一致。')
  let ranges: SheetSelectionRange[] | undefined
  if (input.ranges !== undefined) {
    if (!Array.isArray(input.ranges)) throw invalid('大纲选区必须是数组。')
    ranges = input.ranges.map(range => {
      if (!range || typeof range !== 'object' || ![range.startRow, range.endRow, range.startCol, range.endCol].every(Number.isSafeInteger)
        || range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol
        || range.rangeType !== undefined && ![0, 1, 2, 3].includes(range.rangeType)) throw invalid('大纲选区无效。')
      return { startRow: range.startRow, endRow: range.endRow, startCol: range.startCol, endCol: range.endCol,
        ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }) }
    })
  }
  return { ...input, ...(ranges === undefined ? {} : { ranges }) }
}
export function getSheetOutlinePromptChoices(context: SheetOutlineCommandContext, operation: 'group' | 'remove'): SheetOutlinePromptChoice[] {
  const selected = singleSelection(context)
  if (!selected) return []
  return (['row', 'column'] as const).map(axis => {
    const range = selectionRange(selected, axis), action = { ...range, kind: operation }, label = outlineSelectionLabel(range, operation)
    let disabled = false, reason: string | undefined
    try {
      let previewId = '__outline_preview__'
      while (context.sheet.outline?.groups.some(group => group.id === previewId)) previewId += '_'
      if (!getSheetOutlineUpdate(context.sheet, operation === 'group' ? { ...action, groupId: previewId } : action, context.userId)) {
        disabled = true; reason = '选区内没有可以取消的大纲组。'
      }
    } catch (error) { disabled = true; reason = sheetOutlineErrorMessage(error) }
    return { ...range, action, label, disabled, ...(reason ? { reason } : {}) }
  })
}
export function resolveSheetOutlineCommand(id: string, value: unknown, context: SheetOutlineCommandContext): SheetOutlineCommandResolution | undefined {
  if (!(SHEET_OUTLINE_COMMAND_IDS as readonly string[]).includes(id)) return undefined
  const params = sheetOutlineCommandInput(value)
  if (params.ranges !== undefined) context = { ...context, selections: params.ranges }
  if (params.unitId !== undefined && params.unitId !== context.unitId
    || (params.subUnitId ?? params.sheetId) !== undefined && (params.subUnitId ?? params.sheetId) !== context.sheetId) throw invalid('大纲命令上下文不匹配。')
  const change = (action: ChangeSheetOutlineInput['action']): SheetOutlineCommandResolution => ({ kind: 'change', input: {
    unitId: context.unitId, sheetId: context.sheetId, action,
    ...(params.expectedRevision !== undefined ? { expectedRevision: params.expectedRevision } : {}),
    ...(params.expectedHistoryGeneration !== undefined ? { expectedHistoryGeneration: params.expectedHistoryGeneration } : {}),
    ...(params.expectedUserId !== undefined ? { expectedUserId: params.expectedUserId } : {}), ...(params.host ? { host: params.host } : {}),
  } })
  if (id === 'sheet.command.add-dimension-outline' || id === 'sheet.command.clear-dimension-outlines') {
    assertOutlineAxis(params.axis)
    if (params.start === undefined || params.end === undefined) throw invalid('大纲命令需要 inclusive 起止范围。')
    return change({ axis: params.axis, start: params.start, end: params.end, kind: id === SHEET_OUTLINE_CORE_COMMAND_IDS[0] ? 'group' : 'clear' })
  }
  if (id === 'sheet.command.remove-dimension-outline' || id === 'sheet.command.set-dimension-outline-collapsed') {
    if (!params.outlineId) throw invalid('大纲命令需要 outlineId。')
    const group = context.sheet.outline?.groups.find(group => group.id === params.outlineId)
    if (!group) return undefined
    if (id === 'sheet.command.remove-dimension-outline') return change({ kind: 'remove-group', axis: group.axis, groupId: group.id })
    if (params.collapsed === undefined) throw invalid('大纲折叠命令需要 collapsed。')
    return change({ kind: 'set-collapsed', axis: group.axis, groupId: group.id, collapsed: params.collapsed })
  }
  if (id === 'sheet.command.set-outline-summary-settings') {
    if (params.summaryAbove === undefined || params.summaryLeft === undefined) throw invalid('大纲汇总设置需要行、列两个方向。')
    return change({ kind: 'summary-settings', summaryAbove: params.summaryAbove, summaryLeft: params.summaryLeft })
  }
  if (id === 'sheet.command.set-outline-level' || id === 'sheet.command.set-outline-summary-position') {
    assertOutlineAxis(params.axis)
    if (id === 'sheet.command.set-outline-level') {
      if (params.level === undefined) throw invalid('大纲层级命令需要 level。')
      return change({ kind: 'set-level', axis: params.axis, level: params.level })
    }
    if (params.atStart === undefined) throw invalid('大纲汇总位置命令需要 atStart。')
    return change({ kind: 'summary-position', axis: params.axis, atStart: params.atStart })
  }
  const selection = singleSelection(context)
  if (!selection) return undefined
  const operation = id.includes('clear-selected') ? 'remove' : 'group'
  const explicitAxis = id.includes('selected-rows') ? 'row' : id.includes('selected-columns') ? 'column' : params.value ?? params.axis
  const axis = explicitAxis ?? (selection.rangeType === 1 ? 'row' : selection.rangeType === 2 ? 'column' : undefined)
  if (axis) return change({ ...selectionRange(selection, axis), kind: operation })
  if (!id.endsWith('-shortcut')) return undefined
  const choices = getSheetOutlinePromptChoices(context, operation)
  return choices.some(choice => !choice.disabled) ? { kind: 'prompt', operation, choices } : undefined
}
function singleSelection(context: SheetOutlineCommandContext): SheetSelectionRange | undefined {
  const ranges = context.selections
  if (!ranges || ranges.length !== 1) return undefined
  const range = ranges[0]!
  if (![range.startRow, range.endRow, range.startCol, range.endCol].every(Number.isSafeInteger)
    || range.startRow < 0 || range.startCol < 0 || range.endRow < range.startRow || range.endCol < range.startCol
    || range.endRow >= context.sheet.rowCount || range.endCol >= context.sheet.columnCount
    || range.rangeType !== undefined && ![0, 1, 2, 3].includes(range.rangeType)) throw invalid('大纲选区无效。')
  return range
}
function selectionRange(range: SheetSelectionRange, axis: OutlineAxis): OutlineRange {
  return { axis, start: axis === 'row' ? range.startRow : range.startCol, end: axis === 'row' ? range.endRow : range.endCol }
}
export function outlineSelectionLabel(range: OutlineRange, operation: 'group' | 'remove' = 'group'): string {
  const verb = operation === 'group' ? '组合' : '取消分组'
  return range.axis === 'row' ? `${verb}第 ${range.start + 1}–${range.end + 1} 行` : `${verb} ${columnLabel(range.start)}–${columnLabel(range.end)} 列`
}
function columnLabel(index: number): string {
  let current = index + 1, result = ''
  while (current > 0) { result = String.fromCharCode(65 + (current - 1) % 26) + result; current = Math.floor((current - 1) / 26) }
  return result
}
function invalid(message: string): SheetOutlineError { return new SheetOutlineError(DimensionOutlineErrorReason.INVALID_RANGE, message) }
