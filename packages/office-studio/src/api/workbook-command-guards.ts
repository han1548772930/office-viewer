/** Local execution context is attached by identity, keeping command DTOs and
 * observer snapshots serializable. Weak ownership ends with the request. */
const guards = new WeakMap<object, () => void>()

export function bindWorkbookCommandGuard<T extends object>(params: T, assertCurrent: () => void): T {
  if (typeof assertCurrent !== 'function') throw new TypeError('工作簿命令需要有效的上下文检查。')
  guards.set(params, assertCurrent)
  return params
}

export function workbookCommandGuard(params: unknown): (() => void) | undefined {
  return params !== null && typeof params === 'object' ? guards.get(params) : undefined
}
