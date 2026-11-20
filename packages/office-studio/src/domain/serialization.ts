/** Domain snapshots contain JSON data only; this also unwraps nested reactive proxies. */
export function cloneModel<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}
