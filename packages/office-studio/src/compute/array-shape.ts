/** Control arrays must be bounded and dense before they cross the WASM DTO
 * boundary. Array.every skips holes, while serde and subsequent indexed reads
 * require a value at every position. Inherited array entries are not DTO data. */
export function isBoundedDenseArray<T>(value: unknown, maximum: number,
  valid: (item: unknown) => item is T, minimum = 0): value is T[] {
  if (!Array.isArray(value) || value.length < minimum || value.length > maximum) return false
  for (let position = 0; position < value.length; position++) {
    if (!Object.prototype.hasOwnProperty.call(value, position) || !valid(value[position])) return false
  }
  return true
}
