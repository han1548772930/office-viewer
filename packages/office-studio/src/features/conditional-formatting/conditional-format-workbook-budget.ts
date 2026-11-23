import { conditionalFormatModelByteLimit, conditionalFormatRuleArraySerializedSize } from './conditional-format-source-budget.ts'

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

/** Parser sheets and facade sheet maps share one logical array without changing each sheet's JSON keys. */
export function assertWorkbookConditionalFormatBudget(sheets: unknown, limit = conditionalFormatModelByteLimit): void {
  let bytes = 2, populated = false
  if (bytes > limit) throw new RangeError('CF_SOURCE_TOTAL_LIMIT: 条件格式模型超过累计字节预算。')
  const entries = Array.isArray(sheets) ? sheets : isRecord(sheets) ? Object.values(sheets) : []
  for (const value of entries) {
    if (!isRecord(value) || value.conditionalFormats === undefined) continue
    if (!Array.isArray(value.conditionalFormats)) throw new TypeError('条件格式规则必须是数组。')
    // A serialized array is nonempty exactly when it uses more than its two brackets.
    // Keep each local index and the conditionalFormats property context for pure toJSON conversions.
    const size = conditionalFormatRuleArraySerializedSize(value.conditionalFormats, limit - bytes + 2, 'conditionalFormats')
    if (size === 2) continue
    bytes += size - 2 + (populated ? 1 : 0)
    if (bytes > limit) throw new RangeError('CF_SOURCE_TOTAL_LIMIT: 条件格式模型超过累计字节预算。')
    populated = true
  }
}
