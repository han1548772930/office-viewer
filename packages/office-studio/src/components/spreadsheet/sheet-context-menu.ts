import type { GridHit, GridRange } from '../../rendering/spreadsheet/layout.ts'
import { mergeSheetAxisIntervals } from '../../domain/sheet-axis-structure.ts'
import { sheetBorderLabels } from '../../features/sheet-workbench/sheet-cell-border.ts'
import { SHEET_BORDER_FORMAT_COMMANDS } from '../../features/sheet-workbench/sheet-border-format-command.ts'

export interface SheetContextMenuItem {
  id: string
  label: string
  translationKey?: string
  count?: number
  commandId: string
  type?: 'divider'
  params?: unknown
  disabled?: boolean
  separatorBefore?: boolean
  quick?: boolean
  icon?: string
  children?: SheetContextMenuItem[]
}

export interface SheetContextMenuContext {
  sheetId: string
  canDeleteRow: boolean
  canDeleteColumn: boolean
  frozen: boolean
  hasHiddenRows: boolean
  hasHiddenColumns: boolean
  canFormat?: boolean
  canConvertText?: boolean
  hasHyperlink?: boolean
  canShiftHorizontally?: boolean
  canShiftVertically?: boolean
  selectedRanges?: readonly GridRange[]
  rowCount?: number
  columnCount?: number
  canStructureRows?: boolean
  canStructureColumns?: boolean
  canVisibilityRows?: boolean
  canVisibilityColumns?: boolean
}

function targetParams(hit: GridHit, sheetId: string, axis: 'row' | 'column', offset = 0) {
  const index = axis === 'row'
    ? hit.kind === 'row' || hit.kind === 'cell' ? hit.row : undefined
    : hit.kind === 'column' || hit.kind === 'cell' ? hit.column : undefined
  return index === undefined ? undefined : { sheetId, index: index + offset }
}

export function sheetContextStructureTarget(hit: GridHit, options: Pick<SheetContextMenuContext, 'sheetId' | 'selectedRanges' | 'rowCount' | 'columnCount'>, axis: 'row' | 'column') {
  const rows = options.rowCount ?? 1_048_576, columns = options.columnCount ?? 16_384
  const selected = options.selectedRanges ?? []
  const inside = selected.some(range => hit.kind === 'cell'
    ? hit.row >= range.startRow && hit.row <= range.endRow && hit.column >= range.startColumn && hit.column <= range.endColumn
    : hit.kind === 'row' ? range.startColumn === 0 && range.endColumn === columns - 1 && hit.row >= range.startRow && hit.row <= range.endRow
      : hit.kind === 'column' && range.startRow === 0 && range.endRow === rows - 1 && hit.column >= range.startColumn && hit.column <= range.endColumn)
  const index = axis === 'row' ? hit.kind === 'row' || hit.kind === 'cell' ? hit.row : 0 : hit.kind === 'column' || hit.kind === 'cell' ? hit.column : 0
  const intervals = mergeSheetAxisIntervals(inside ? selected.map(range => axis === 'row'
    ? { start: range.startRow, end: range.endRow } : { start: range.startColumn, end: range.endColumn }) : [{ start: index, end: index }])
  const ranges = intervals.map(interval => axis === 'row' ? { startRow: interval.start, endRow: interval.end, startCol: 0, endCol: columns - 1 }
    : { startRow: 0, endRow: rows - 1, startCol: interval.start, endCol: interval.end })
  return { params: { sheetId: options.sheetId, ranges }, count: intervals.reduce((sum, interval) => sum + interval.end - interval.start + 1, 0), continuous: intervals.length === 1 }
}

export function createSheetContextMenuItems(hit: GridHit, options: SheetContextMenuContext): SheetContextMenuItem[] {
  if (hit.kind !== 'cell' && hit.kind !== 'row' && hit.kind !== 'column' && hit.kind !== 'all') return []

  const clipboard: SheetContextMenuItem[] = [
    { id: 'copy', label: '复制', commandId: 'sheet.command.copy', quick: true, icon: 'CopyDoubleIcon' },
    { id: 'cut', label: '剪切', commandId: 'sheet.command.cut', quick: true, icon: 'CutIcon' },
    { id: 'paste', label: '粘贴', commandId: 'sheet.command.paste', quick: true, icon: 'PasteSpecialDoubleIcon' },
    { id: 'copy-special', label: '选择性复制', commandId: '', icon: 'CopyDoubleIcon', separatorBefore: true, children: [
      { id: 'copy-formulas', label: '仅复制公式', commandId: 'sheet.command.copy-formulas' },
    ] },
    {
      id: 'paste-special', label: '选择性粘贴', commandId: '', icon: 'PasteSpecialDoubleIcon', children: [
        { id: 'paste-values', label: '仅粘贴值', commandId: 'sheet.command.paste-values' },
        { id: 'paste-formats', label: '仅粘贴格式', commandId: 'sheet.command.paste-formats' },
        { id: 'paste-column-widths', label: '仅粘贴列宽', commandId: 'sheet.command.paste-column-widths' },
        { id: 'paste-except-borders', label: '仅粘贴边框以外内容', commandId: 'sheet.command.paste-except-borders' },
        { id: 'paste-formulas', label: '仅粘贴公式', commandId: 'sheet.command.paste-formulas' },
      ],
    },
    { id: 'clear-selection', label: '清除', commandId: '', icon: 'ClearFormatDoubleIcon', children: [
      { id: 'clear-content', label: '清除内容', commandId: 'sheet.command.clear-content' },
      { id: 'clear-format', label: '清除格式', commandId: 'sheet.command.clear-format' },
      { id: 'clear-all', label: '清除全部', commandId: 'sheet.command.clear-all' },
    ] },
    { id: 'borders', label: '边框', translationKey: 'sheet.items.sheet.menu.borders', commandId: '', icon: 'AllBorderIcon', disabled: options.canFormat === false,
      children: Object.entries(sheetBorderLabels).map(([preset, label]) => ({
        id: `border-${preset}`, label, translationKey: `sheet.items.sheet.menu.border.${preset}`,
        commandId: SHEET_BORDER_FORMAT_COMMANDS.position, params: preset, disabled: options.canFormat === false,
      })),
    },
    ...(options.canConvertText ? [{ id: 'text-to-number', label: '文本转数字', commandId: 'sheet.command.text-to-number', icon: 'ConvertToNumberIcon' }] : []),
  ]
  const linkActions: SheetContextMenuItem[] = hit.kind === 'cell' && options.hasHyperlink
    ? [{ id: 'open-hyperlink', label: '打开超链接', commandId: 'sheet.operation.open-cell-hyperlink', params: { key: hit.key }, icon: 'LinkIcon', separatorBefore: true }]
    : []
  const freeze: SheetContextMenuItem[] = [
    { id: 'freeze', label: '冻结', commandId: '', icon: 'FreezeToSelectedIcon', children: [
      { id: 'freeze-selection', label: hit.kind === 'row' ? `冻结至第 ${hit.row + 1} 行` : hit.kind === 'column' ? `冻结至第 ${hit.column + 1} 列` : hit.kind === 'cell' ? `冻结至活动单元格（${hit.row + 1}行${hit.column + 1}列）` : '冻结至活动单元格', commandId: 'sheet.command.freeze-selection', icon: hit.kind === 'row' ? 'FreezeRowIcon' : hit.kind === 'column' ? 'FreezeColumnIcon' : 'FreezeToSelectedIcon', params: hit.kind === 'row' ? { sheetId: options.sheetId, axis: 'row', index: hit.row } : hit.kind === 'column' ? { sheetId: options.sheetId, axis: 'column', index: hit.column } : undefined },
      ...(hit.kind === 'column' ? [] : [{ id: 'freeze-first-row', label: '冻结首行', commandId: 'sheet.command.freeze-first-row', icon: 'FreezeRowIcon' }]),
      ...(hit.kind === 'row' ? [] : [{ id: 'freeze-first-column', label: '冻结首列', commandId: 'sheet.command.freeze-first-column', icon: 'FreezeColumnIcon' }]),
      { id: 'unfreeze', label: '取消冻结', commandId: 'sheet.command.unfreeze', icon: 'CancelFreezeIcon', disabled: !options.frozen },
    ] },
  ]
  const range = hit.kind === 'row' ? `A${hit.row + 1}:${columnLabel(16383)}${hit.row + 1}`
    : hit.kind === 'column' ? `${columnLabel(hit.column)}1:${columnLabel(hit.column)}1048576`
      : hit.kind === 'cell' ? hit.key : undefined
  const protect: SheetContextMenuItem = {
    id: 'protect-range', label: '保护行列', commandId: '', icon: 'ProtectIcon', children: [
      { id: 'protect-range-add', label: '新增保护范围', commandId: 'sheet.operation.open-range-protection', icon: 'LockIcon', params: { sheetId: options.sheetId, ...(range ? { range } : {}), mode: 'add' } },
      { id: 'protect-range-view', label: '查看所有保护范围', commandId: 'sheet.operation.open-range-protection', icon: 'EyeIcon', params: { sheetId: options.sheetId, ...(range ? { range } : {}) } },
    ],
  }
  const rowTarget = targetParams(hit, options.sheetId, 'row')
  const columnTarget = targetParams(hit, options.sheetId, 'column')
  const rowStructure = sheetContextStructureTarget(hit, options, 'row'), columnStructure = sheetContextStructureTarget(hit, options, 'column')

  if (hit.kind === 'row') return [
    ...clipboard,
    { id: 'insert-row-above', label: `在上方插入 ${rowStructure.count} 行`, count: rowStructure.count, commandId: 'sheet.command.insert-row-before', icon: 'InsertRowAboveDoubleIcon', params: rowStructure.params, separatorBefore: true, disabled: !rowStructure.continuous || options.canStructureRows === false },
    { id: 'insert-row-below', label: `在下方插入 ${rowStructure.count} 行`, count: rowStructure.count, commandId: 'sheet.command.insert-row-after', icon: 'InsertRowBelowDoubleIcon', params: rowStructure.params, disabled: !rowStructure.continuous || options.canStructureRows === false },
    { id: 'delete-row', label: '删除选中行', commandId: 'sheet.command.delete-row', icon: 'DeleteRowDoubleIcon', params: rowStructure.params, disabled: !options.canDeleteRow || options.canStructureRows === false },
    { id: 'hide-row', label: '隐藏选中行', commandId: 'sheet.command.hide-row-confirm', icon: 'HideDoubleIcon', params: rowStructure.params, disabled: options.canVisibilityRows === false },
    { id: 'show-hidden-rows', label: '显示隐藏行', commandId: 'sheet.command.set-specific-rows-visible', icon: 'EyeIcon', params: rowStructure.params, disabled: !options.hasHiddenRows || options.canVisibilityRows === false },
    { id: 'row-height', label: '行高', commandId: 'sheet.command.set-row-height', icon: 'AdjustHeightDoubleIcon', params: rowTarget },
    { id: 'auto-fit-row', label: '适合数据', commandId: 'sheet.command.auto-fit-row', icon: 'AutoHeightDoubleIcon', params: rowTarget },
    ...freeze,
    protect,
  ]
  if (hit.kind === 'column') return [
    ...clipboard,
    { id: 'insert-column-left', label: `在左侧插入 ${columnStructure.count} 列`, count: columnStructure.count, commandId: 'sheet.command.insert-col-before', icon: 'LeftInsertColumnDoubleIcon', params: columnStructure.params, separatorBefore: true, disabled: !columnStructure.continuous || options.canStructureColumns === false },
    { id: 'insert-column-right', label: `在右侧插入 ${columnStructure.count} 列`, count: columnStructure.count, commandId: 'sheet.command.insert-col-after', icon: 'RightInsertColumnDoubleIcon', params: columnStructure.params, disabled: !columnStructure.continuous || options.canStructureColumns === false },
    { id: 'hide-column', label: '隐藏选中列', commandId: 'sheet.command.hide-col-confirm', icon: 'HideDoubleIcon', params: columnStructure.params, disabled: options.canVisibilityColumns === false },
    { id: 'show-hidden-columns', label: '显示隐藏列', commandId: 'sheet.command.set-col-visible-on-cols', icon: 'EyeIcon', params: columnStructure.params, disabled: !options.hasHiddenColumns || options.canVisibilityColumns === false },
    { id: 'delete-column', label: '删除选中列', commandId: 'sheet.command.delete-column', icon: 'DeleteColumnDoubleIcon', params: columnStructure.params, disabled: !options.canDeleteColumn || options.canStructureColumns === false },
    { id: 'column-width', label: '列宽', commandId: 'sheet.command.set-worksheet-col-width', icon: 'AdjustWidthDoubleIcon', params: columnTarget },
    { id: 'auto-fit-column', label: '适合数据', commandId: 'sheet.command.auto-fit-column', icon: 'AutoWidthDoubleIcon', params: columnTarget },
    ...freeze,
    protect,
  ]
  if (hit.kind === 'all') return [
    ...clipboard,
    ...freeze,
    protect,
  ]
  return [
    ...clipboard,
    ...linkActions,
    {
      id: 'insert-cells', label: '插入', commandId: '', icon: 'InsertDoubleIcon', separatorBefore: true, children: [
        { id: 'insert-range-right', label: '右移单元格', commandId: 'sheet.command.insert-range-move-right', icon: 'InsertCellShiftRightDoubleIcon', disabled: options.canShiftHorizontally === false },
        { id: 'insert-range-down', label: '下移单元格', commandId: 'sheet.command.insert-range-move-down', icon: 'InsertCellDownDoubleIcon', disabled: options.canShiftVertically === false },
        { id: 'insert-row', label: '在上方插入行', commandId: 'sheet.command.insert-row-before', icon: 'InsertRowAboveDoubleIcon', params: rowStructure.params, disabled: !rowStructure.continuous || options.canStructureRows === false },
        { id: 'insert-column', label: '在左侧插入列', commandId: 'sheet.command.insert-col-before', icon: 'LeftInsertColumnDoubleIcon', params: columnStructure.params, disabled: !columnStructure.continuous || options.canStructureColumns === false },
      ],
    },
    {
      id: 'delete-cells', label: '删除', commandId: '', icon: 'ReduceDoubleIcon', children: [
        { id: 'delete-range-left', label: '左移单元格', commandId: 'sheet.command.delete-range-move-left', icon: 'DeleteCellShiftLeftDoubleIcon', disabled: options.canShiftHorizontally === false },
        { id: 'delete-range-up', label: '上移单元格', commandId: 'sheet.command.delete-range-move-up', icon: 'DeleteCellShiftUpDoubleIcon', disabled: options.canShiftVertically === false },
        { id: 'delete-column', label: '删除选中列', commandId: 'sheet.command.delete-column', icon: 'DeleteColumnDoubleIcon', params: columnStructure.params, disabled: !options.canDeleteColumn || options.canStructureColumns === false },
        { id: 'delete-row', label: '删除选中行', commandId: 'sheet.command.delete-row', icon: 'DeleteRowDoubleIcon', params: rowStructure.params, disabled: !options.canDeleteRow || options.canStructureRows === false },
      ],
    },
    ...freeze,
    protect,
  ]
}

function columnLabel(index: number) {
  let column = index + 1
  let label = ''
  while (column > 0) {
    const remainder = (column - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    column = Math.floor((column - 1) / 26)
  }
  return label
}
