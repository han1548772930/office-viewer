import type { DocumentContentControl } from '../../../domain/model'

/** The original public customRanges index is provenance, never a local ID sort. */
export function documentContentControlSourceOrder(control: DocumentContentControl): number | undefined {
  return readDocumentContentControlSourceOrder(control.metadata?.univerBodySdtOrder)
}

/** Public snapshot adapters validate the persisted provenance before using it. */
export function readDocumentContentControlSourceOrder(source: unknown): number | undefined {
  if (!source || typeof source !== 'object' || Array.isArray(source) || Reflect.get(source, 'version') !== 1) return undefined
  const order = Reflect.get(source, 'order')
  return typeof order === 'number' && Number.isSafeInteger(order) && order >= 0 ? order : undefined
}

/** Copied or mixed imports can duplicate provenance; ambiguous orders cannot imply parenthood. */
export function documentContentControlSourceOrders(controls: readonly DocumentContentControl[]): ReadonlyMap<string, number> {
  const orders = new Map<string, number>(), counts = new Map<number, number>()
  for (const control of controls) {
    const order = documentContentControlSourceOrder(control)
    if (order === undefined) continue
    orders.set(control.id, order); counts.set(order, (counts.get(order) ?? 0) + 1)
  }
  for (const [id, order] of orders) if (counts.get(order) !== 1) orders.delete(id)
  return orders
}
