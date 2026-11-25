import type { OfficeStudio } from './office-studio.ts'
import type { Disposable } from './events.ts'
import type { SheetStyleSource } from '../domain/model.ts'
import { assignSheetAxisSpan } from '../domain/sheet-axis.ts'
import { normalizeSheetStyleInput, resolveSheetStyleSource, type SheetStyleInput } from '../features/sheet-workbench/sheet-style-inheritance.ts'

export const SHEET_DEFAULT_STYLE_COMMANDS = Object.freeze({
  worksheet: 'sheet.command.set-worksheet-default-style',
  workbook: 'sheet.command.set-workbook-default-style',
  row: 'sheet.command.set-row-default-style',
  column: 'sheet.command.set-column-default-style',
  rowPriority: 'sheet.command.set-row-style-priority',
})
export interface SheetDefaultStyleCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  defaultStyle?: SheetStyleInput
  style?: SheetStyleInput
  index?: number
  rowPriority?: boolean
}

export function registerSheetDefaultStyleCommands(api: OfficeStudio): Disposable {
  const registrations = Object.entries(SHEET_DEFAULT_STYLE_COMMANDS).map(([kind, id]) => api.registerCommand<SheetDefaultStyleCommandInput, boolean>(id, input => {
    if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('默认样式命令参数必须是对象。')
    const unitId = input.unitId ?? api.activeWorkbookId, workbook = api.getWorkbookModel(unitId)
    if (!workbook) throw new Error(`Workbook not found: ${unitId}`)
    const sheetId = input.sheetId ?? input.subUnitId ?? workbook.activeSheetId
    const sheet = workbook.sheets.find(item => item.id === sheetId)
    if (!sheet) throw new Error(`Worksheet not found: ${sheetId}`)
    let style: SheetStyleSource | undefined
    if (kind === 'rowPriority') {
      if (typeof input.rowPriority !== 'boolean') throw new TypeError('行样式优先级必须是布尔值。')
    } else {
      const value = input.defaultStyle === undefined ? input.style : input.defaultStyle
      if (value === undefined) throw new TypeError('请提供默认样式或 null。')
      style = normalizeSheetStyleInput(value)
      resolveSheetStyleSource(style, workbook)
    }
    const axis = kind === 'row' || kind === 'column'
    if (axis && (!Number.isSafeInteger(input.index) || input.index! < 0 || input.index! >= (kind === 'row' ? 1_048_576 : 16_384))) throw new RangeError('默认样式的行列索引越界。')
    const affected = kind === 'workbook' ? workbook.sheets : [sheet]
    for (const target of affected) {
      api.assertCanEditWorksheet(unitId, target.id)
      const range = kind === 'row' ? { startRow: input.index!, endRow: input.index!, startCol: 0, endCol: 16_383 }
        : kind === 'column' ? { startRow: 0, endRow: 1_048_575, startCol: input.index!, endCol: input.index! }
          : { startRow: 0, endRow: 1_048_575, startCol: 0, endCol: 16_383 }
      api.assertCanEditRange(unitId, target.id, range)
    }
    return api.commitWorkbook(unitId, kind === 'rowPriority' ? '设置行列样式优先级' : '设置默认样式', candidate => {
      if (kind === 'workbook') { candidate.defaultStyle = style; return }
      const target = candidate.sheets.find(item => item.id === sheetId)!
      if (kind === 'worksheet') target.defaultStyle = style
      else if (kind === 'rowPriority') target.rowStylePriority = input.rowPriority
      else {
        const key = kind === 'row' ? 'rowStyles' : 'columnStyles'
        const next = { ...target[key] }
        if (style === null) {
          delete next[input.index!]
          const spanKey = kind === 'row' ? 'rowStyleSpans' : 'columnStyleSpans'
          target[spanKey] = assignSheetAxisSpan(target[spanKey], input.index!, input.index!)
        }
        else next[input.index!] = style!
        target[key] = Object.keys(next).length ? next : undefined
      }
    }, { recalculate: false })
  }))
  return { dispose: () => registrations.forEach(item => item.dispose()) }
}
