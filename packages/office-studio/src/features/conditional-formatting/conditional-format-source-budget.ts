/** Shared with office-io-common; byte budgets describe UTF-8, not UTF-16 units. */
export const conditionalFormatOfficeFragmentByteLimit = 16 * 1024 * 1024
export const conditionalFormatOfficeSourceXmlByteLimit = 5 * conditionalFormatOfficeFragmentByteLimit
export const conditionalFormatModelByteLimit = 256 * 1024 * 1024

/** Returns at most limit + 1, avoiding an additional large TextEncoder buffer. */
export function conditionalFormatUtf8Size(value: string, limit = Number.MAX_SAFE_INTEGER): number {
  let bytes = 0
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    if (code < 0x80) bytes++
    else if (code < 0x800) bytes += 2
    else if (code >= 0xd800 && code <= 0xdbff && index + 1 < value.length
      && value.charCodeAt(index + 1) >= 0xdc00 && value.charCodeAt(index + 1) <= 0xdfff) { bytes += 4; index++ }
    else bytes += 3
    if (bytes > limit) return limit + 1
  }
  return bytes
}

function jsonStringSize(value: string, limit: number): number {
  let bytes = 2
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    if (code === 0x22 || code === 0x5c || [8, 9, 10, 12, 13].includes(code)) bytes += 2
    else if (code < 0x20) bytes += 6
    else if (code < 0x80) bytes++
    else if (code < 0x800) bytes += 2
    else if (code >= 0xd800 && code <= 0xdbff && index + 1 < value.length
      && value.charCodeAt(index + 1) >= 0xdc00 && value.charCodeAt(index + 1) <= 0xdfff) { bytes += 4; index++ }
    else if (code >= 0xd800 && code <= 0xdfff) bytes += 6
    else bytes += 3
    if (bytes > limit) return limit + 1
  }
  return bytes
}

function totalLimit(): RangeError { return new RangeError('CF_SOURCE_TOTAL_LIMIT: 条件格式模型超过累计字节预算。') }
function invalidJson(): TypeError { return new TypeError('CF_SOURCE_XML_INVALID: 条件格式来源必须是可序列化的 JSON 数据。') }

function serializedSize(value: unknown, limit: number, rootKey: string, requireRootArray = false): number {
  let bytes = 0
  const ancestors = new Set<object>()
  const add = (count: number) => { bytes += count; if (bytes > limit) throw totalLimit() }
  const visit = (input: unknown, key: string, depth: number, inArray: boolean): boolean => {
    if (depth > 128) throw invalidJson()
    let item = input
    if (item && (typeof item === 'object' || typeof item === 'function')
      && 'toJSON' in item && typeof item.toJSON === 'function') item = item.toJSON(key)
    if (requireRootArray && depth === 0 && !Array.isArray(item)) throw new TypeError('条件格式规则的 JSON 转换结果必须是数组。')
    if (item === null) { add(4); return true }
    if (typeof item === 'string') { add(jsonStringSize(item, limit - bytes)); return true }
    if (typeof item === 'number') { add(Number.isFinite(item) ? String(item).length : 4); return true }
    if (typeof item === 'boolean') { add(item ? 4 : 5); return true }
    if (typeof item === 'bigint') throw invalidJson()
    if (typeof item !== 'object') { if (inArray) { add(4); return true } return false }
    if (ancestors.has(item)) throw invalidJson()
    ancestors.add(item)
    if (Array.isArray(item)) {
      add(2)
      for (let index = 0; index < item.length; index++) {
        if (index) add(1)
        visit(item[index], String(index), depth + 1, true)
      }
    } else {
      const prototype = Object.getPrototypeOf(item)
      if (prototype !== Object.prototype && prototype !== null) throw invalidJson()
      add(2)
      let entries = 0
      for (const key of Object.keys(item)) {
        const child = (item as Record<string, unknown>)[key]
        // Functions with toJSON can retain a property; visit them before deciding to omit it.
        if (child === undefined || typeof child === 'symbol') continue
        // A toJSON result may also omit an object property. Count its key only if retained.
        if (!visit(child, key, depth + 1, false)) continue
        if (entries++) add(1)
        add(jsonStringSize(key, limit - bytes) + 1)
      }
    }
    ancestors.delete(item)
    return true
  }
  if (!visit(value, rootKey, 0, false)) throw invalidJson()
  return bytes
}

/** rootKey preserves toJSON's property context when counting a value inside a snapshot. */
export function conditionalFormatSerializedSize(value: unknown, limit = conditionalFormatModelByteLimit, rootKey = ''): number {
  return serializedSize(value, limit, rootKey)
}

/** Count a sheet's actual rule-array representation, applying its root conversion exactly once. */
export function conditionalFormatRuleArraySerializedSize(values: readonly unknown[], limit = conditionalFormatModelByteLimit, rootKey = 'conditionalFormats'): number {
  return serializedSize(values, limit, rootKey, true)
}

/** Include array brackets/separators and every retained rule/source extension. */
export function assertConditionalFormatModelBudget(values: Iterable<unknown>, limit = conditionalFormatModelByteLimit): void {
  let bytes = 2, entries = 0
  if (bytes > limit) throw totalLimit()
  for (const value of values) {
    if (entries++) bytes++
    if (bytes > limit) throw totalLimit()
    bytes += serializedSize(value, limit - bytes, String(entries - 1))
  }
}
