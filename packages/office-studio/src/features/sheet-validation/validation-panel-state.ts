import type { CellRange, SheetDataValidationOperator, SheetDataValidationRule, SheetDataValidationType } from '../../domain/model.ts'
import { parseSheetAddress } from '../sheet-workbench/sheet-workbench.ts'
import { createSheetDataValidationRule, parseValidationListItems, serializeValidationListItems } from './validation-engine.ts'

export type ValidationRangeTarget = 'range' | 'list' | 'formula1' | 'formula2'
export interface ValidationListItem { id: string; label: string; color: string }
export interface ValidationPanelSettings {
  type: SheetDataValidationType
  operator: SheetDataValidationOperator | ''
  formula1: string
  formula2: string
  checkboxCustomValues: boolean
  listMode: 'custom' | 'reference'
  listItems: ValidationListItem[]
  listReference: string
}

const COMPARISON_TYPES: readonly SheetDataValidationType[] = ['whole', 'decimal', 'textLength', 'date', 'time']
export const VALIDATION_OPERATORS: readonly SheetDataValidationOperator[] = [
  'between', 'equal', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'notBetween', 'notEqual',
]

export function validationUsesOperator(type: SheetDataValidationType): boolean { return COMPARISON_TYPES.includes(type) }
export function validationUsesTwoFormulas(type: SheetDataValidationType, operator: ValidationPanelSettings['operator']): boolean {
  return validationUsesOperator(type) && (operator === 'between' || operator === 'notBetween')
}

export function validationPanelSettings(rule?: SheetDataValidationRule | null, initialType?: SheetDataValidationType): ValidationPanelSettings {
  const type = rule?.type ?? initialType ?? 'whole'
  const reference = rule?.formula1?.trim().startsWith('=') ?? false
  const labels = !reference && (type === 'list' || type === 'listMultiple') ? parseValidationListItems(rule?.formula1 ?? '') : []
  const colors = rule?.formula2?.split(',') ?? []
  return {
    type,
    operator: rule ? rule.operator ?? '' : validationUsesOperator(type) ? type === 'date' ? '' : 'between' : '',
    formula1: rule?.formula1 ?? '',
    formula2: rule?.formula2 ?? '',
    checkboxCustomValues: type === 'checkbox' && (rule?.formula1 !== undefined || rule?.formula2 !== undefined),
    listMode: reference ? 'reference' : 'custom',
    listItems: labels.map((label, index) => ({ id: crypto.randomUUID(), label, color: colors[index] ?? '' })),
    listReference: reference ? rule?.formula1 ?? '=' : '=',
  }
}

export function changeValidationType(settings: ValidationPanelSettings, type: SheetDataValidationType): ValidationPanelSettings {
  if (settings.type === type) return settings
  const next = validationPanelSettings(null, type)
  if (settings.type.startsWith('list') && type.startsWith('list')) {
    return { ...settings, type, operator: '' }
  }
  return next
}

export function validationRuleFromPanel(
  settings: ValidationPanelSettings,
  ranges: readonly CellRange[],
  options: Partial<SheetDataValidationRule>,
  previous?: SheetDataValidationRule | null,
): SheetDataValidationRule {
  const draft: Omit<SheetDataValidationRule, 'id' | 'ranges'> & { id?: string } = {
    ...previous, ...options, type: settings.type, operator: undefined, formula1: undefined, formula2: undefined,
  }
  if (settings.type === 'list' || settings.type === 'listMultiple') {
    if (settings.listMode === 'reference') {
      draft.formula1 = settings.listReference.trim()
    } else {
      const unique = new Map<string, ValidationListItem>()
      for (const item of settings.listItems) {
        if (item.label.trim() && !unique.has(item.label)) unique.set(item.label, item)
      }
      draft.formula1 = serializeValidationListItems([...unique.keys()])
      draft.formula2 = [...unique.values()].map(item => item.color).join(',')
    }
  } else if (settings.type === 'checkbox') {
    if (settings.checkboxCustomValues) {
      draft.formula1 = settings.formula1
      draft.formula2 = settings.formula2
    }
  } else if (settings.type === 'custom') {
    draft.formula1 = settings.formula1.trim()
  } else if (settings.operator) {
    draft.operator = settings.operator
    draft.formula1 = settings.formula1.trim()
    if (validationUsesTwoFormulas(settings.type, settings.operator)) draft.formula2 = settings.formula2.trim()
  }
  return createSheetDataValidationRule(draft, ranges)
}

export function parseValidationRanges(value: string, sheetName?: string): CellRange[] {
  const tokens = value.match(/(?:'[^']*(?:''[^']*)*'|[^,;])+/g) ?? []
  const ranges = tokens.map(token => {
    const text = token.trim()
    const qualified = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(text)
    const qualifier = qualified?.[1]?.replaceAll("''", "'") ?? qualified?.[2]
    if (qualifier && qualifier !== sheetName) throw new RangeError('应用区域必须属于当前工作表。')
    const parsed = parseSheetAddress(qualified?.[3] ?? text, { rowCount: 1_048_576, columnCount: 16_384 })
    if (!parsed) throw new TypeError('请输入有效的单元格区域，例如 A1:B4、A:A 或 1:3。')
    return { startRow: parsed.range.startRow, startCol: parsed.range.startColumn, endRow: parsed.range.endRow, endCol: parsed.range.endColumn }
  })
  if (!ranges.length) throw new TypeError('请选择至少一个数据验证区域。')
  return ranges
}

export function formatValidationRange(range: CellRange, sheetName?: string): string {
  const first = `${columnLabel(range.startCol)}${range.startRow + 1}`
  const last = `${columnLabel(range.endCol)}${range.endRow + 1}`
  const address = first === last ? first : `${first}:${last}`
  return sheetName ? `'${sheetName.replaceAll("'", "''")}'!${address}` : address
}

function columnLabel(index: number): string {
  let number = index + 1
  let label = ''
  while (number > 0) {
    number -= 1
    label = String.fromCharCode(number % 26 + 65) + label
    number = Math.floor(number / 26)
  }
  return label
}
