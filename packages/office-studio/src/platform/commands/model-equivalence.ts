/** Compare JSON model data independently of object-key order and omitted optional fields. */
export function equivalentModelData(left: unknown, right: unknown, ignoredRootKeys: readonly string[] = []): boolean {
  const ignored = new Set(ignoredRootKeys)
  function equal(first: unknown, second: unknown, root: boolean): boolean {
    if (Object.is(first, second)) return true
    if (Array.isArray(first)) {
      return Array.isArray(second) && first.length === second.length
        && first.every((value: unknown, index: number) => equal(value, second[index], false))
    }
    if (Array.isArray(second) || !isRecord(first) || !isRecord(second)) return false
    const keys = (record: Record<string, unknown>) => Object.keys(record)
      .filter((key) => record[key] !== undefined && (!root || !ignored.has(key)))
    const firstKeys = keys(first)
    const secondKeys = keys(second)
    return firstKeys.length === secondKeys.length
      && firstKeys.every((key) => Object.hasOwn(second, key) && equal(first[key], second[key], false))
  }
  return equal(left, right, true)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object'
}
