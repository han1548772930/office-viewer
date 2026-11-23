/** A host's regional settings identify its stable lifetime across Facade wrappers. */
const refreshed = new WeakMap<object, Map<string, number>>()

/** Reserve before publishing so synchronous observers cannot refresh the same open twice. */
export function reserveDocumentFieldOpenRefresh(host: object, unitId: string, historyGeneration: number): boolean {
  if (!host || typeof host !== 'object' || !unitId.trim() || !Number.isSafeInteger(historyGeneration) || historyGeneration < 1) {
    throw new TypeError('文档打开字段更新的宿主、标识或历史代数无效。')
  }
  let units = refreshed.get(host)
  if (!units) { units = new Map(); refreshed.set(host, units) }
  if (units.get(unitId) === historyGeneration) return false
  units.set(unitId, historyGeneration)
  return true
}
