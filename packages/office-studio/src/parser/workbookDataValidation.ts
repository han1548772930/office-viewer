import type { SheetDataValidationRule } from '../domain/model.ts'
import { cloneModel } from '../domain/serialization.ts'
import { array, record } from './validation.ts'

const types: Readonly<Record<string, SheetDataValidationRule['type']>> = {
  any: 'any', none: 'any', whole: 'whole', integer: 'whole', decimal: 'decimal', textlength: 'textLength',
  date: 'date', time: 'time', checkbox: 'checkbox', list: 'list', listmultiple: 'listMultiple', custom: 'custom',
}
const operators: Readonly<Record<string, NonNullable<SheetDataValidationRule['operator']>>> = {
  between: 'between', equal: 'equal', greaterthan: 'greaterThan', greaterthanorequal: 'greaterThanOrEqual',
  lessthan: 'lessThan', lessthanorequal: 'lessThanOrEqual', notbetween: 'notBetween', notequal: 'notEqual',
}
const knownFields = new Set(['uid', 'id', 'type', 'ranges', 'formulaAnchor', 'formula1', 'formula2', 'operator', 'allowBlank',
  'showDropDown', 'showErrorMessage', 'error', 'errorTitle', 'errorStyle', 'showInputMessage', 'prompt', 'promptTitle', 'renderMode', 'metadata'])

function coordinate(value: unknown, maximum: number): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 && value <= maximum
}
function sourceRule(value: unknown): Record<string, unknown> {
  const outer = record(value)
  if (outer.rule === undefined) return outer
  if (!outer.rule || typeof outer.rule !== 'object' || Array.isArray(outer.rule)) throw new TypeError('数据验证包装中的规则无效。')
  return record(outer.rule)
}

/** Keep range order and the original formula origin; parser wrappers are explicit. */
export function normalizeWorkbookDataValidations(value: unknown): SheetDataValidationRule[] {
  if (value !== null && value !== undefined && !Array.isArray(value)) throw new TypeError('解析结果的数据验证列表无效。')
  const ids = new Set<string>()
  return array(value).map(sourceRule).map((item): SheetDataValidationRule => {
    const type = types[typeof item.type === 'string' ? item.type.toLowerCase() : '']
    const id = typeof item.uid === 'string' ? item.uid : typeof item.id === 'string' ? item.id : ''
    if (!type || !id.trim()) throw new TypeError('解析结果包含无效的数据验证类型或标识。')
    if (ids.has(id)) throw new TypeError('解析结果包含重复的数据验证标识。')
    ids.add(id)
    if (!Array.isArray(item.ranges) || !item.ranges.length) throw new TypeError('数据验证缺少应用区域。')
    const ranges = item.ranges.map((value) => {
      const range = record(value)
      const startRow = range.startRow ?? range.firstRow, startCol = range.startColumn ?? range.startCol ?? range.firstColumn ?? range.firstCol
      const endRow = range.endRow ?? range.lastRow ?? startRow, endCol = range.endColumn ?? range.endCol ?? range.lastColumn ?? range.lastCol ?? startCol
      if (!coordinate(startRow, 1_048_575) || !coordinate(endRow, 1_048_575) || !coordinate(startCol, 16_383) || !coordinate(endCol, 16_383)
        || endRow < startRow || endCol < startCol) throw new TypeError('数据验证应用区域超出工作表边界。')
      return { startRow, startCol, endRow, endCol }
    })
    const anchor = record(item.formulaAnchor)
    if (item.formulaAnchor !== undefined && item.formulaAnchor !== null && (!coordinate(anchor.row, 1_048_575) || !coordinate(anchor.col, 16_383))) {
      throw new TypeError('数据验证公式起点无效。')
    }
    const result: SheetDataValidationRule = {
      ...cloneModel(Object.fromEntries(Object.entries(item).filter(([key]) => !knownFields.has(key)))),
      id, type, ranges,
      ...(coordinate(anchor.row, 1_048_575) && coordinate(anchor.col, 16_383) ? { formulaAnchor: { row: anchor.row, col: anchor.col } } : {}),
    }
    for (const field of ['formula1', 'formula2', 'error', 'errorTitle', 'prompt', 'promptTitle'] as const) {
      const value = item[field]
      if (value === null || value === undefined) continue
      if (typeof value !== 'string') throw new TypeError(`数据验证 ${field} 必须是文本。`)
      result[field] = value
    }
    for (const field of ['allowBlank', 'showDropDown', 'showErrorMessage', 'showInputMessage'] as const) {
      const value = item[field]
      if (value === null || value === undefined) continue
      if (typeof value !== 'boolean') throw new TypeError(`数据验证 ${field} 必须是布尔值。`)
      result[field] = value
    }
    if (item.operator !== null && item.operator !== undefined) {
      const operator = typeof item.operator === 'string' ? operators[item.operator.toLowerCase()] : undefined
      if (!operator) throw new TypeError('数据验证比较条件无效。')
      result.operator = operator
    }
    if (item.errorStyle !== null && item.errorStyle !== undefined) {
      const style = typeof item.errorStyle === 'string' ? item.errorStyle.toLowerCase() : ''
      if (style !== 'stop' && style !== 'warning' && style !== 'info' && style !== 'information') throw new TypeError('数据验证错误样式无效。')
      result.errorStyle = style === 'information' ? 'info' : style
    }
    if (item.renderMode !== null && item.renderMode !== undefined) {
      const mode = typeof item.renderMode === 'string' ? item.renderMode.toLowerCase() : ''
      if (mode !== 'text' && mode !== 'arrow' && mode !== 'custom') throw new TypeError('数据验证显示方式无效。')
      result.renderMode = mode
    }
    if (item.metadata !== null && item.metadata !== undefined) {
      if (typeof item.metadata !== 'object' || Array.isArray(item.metadata)) throw new TypeError('数据验证扩展数据无效。')
      result.metadata = cloneModel(record(item.metadata))
    }
    return result
  })
}
