/** Labels/query are lowercased by the same JavaScript locale boundary as the
 * existing list UI, then encoded as exact UTF-16LE. Lone surrogates are retained.
 * Rust owns substring matching/windowing; offsets are even byte offsets.
 */
export interface ValidationListQueryParams {
  rowCount: number
  textByteLength: number
  queryByteLength: number
  start: number
  count: number
}
export interface ValidationListQueryTask {
  kind: 'validationListQuery'
  offsets: Uint32Array
  textBytes: Uint8Array
  queryBytes: Uint8Array
  params: ValidationListQueryParams
}
export interface ValidationListQueryOutput {
  kind: 'validationListQuery'
  indices: Uint32Array
  total: number
}
