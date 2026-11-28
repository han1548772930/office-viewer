/** v1.6.0 text filters use * and ?; all other regexp syntax (including ~) is literal. */
export function matchesSheetFilterWildcard(value: string, pattern: string, caseSensitive = false): boolean {
  const source = pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\\\*/g, '.*').replace(/\\\?/g, '.')
  return new RegExp(`^${source}$`, caseSensitive ? '' : 'i').test(value)
}
