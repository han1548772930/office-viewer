/** Typed source rows, separate from numeric CELL_STATE. Values are never captions. */
export const SLICER_MEMBER_TYPE = { blank: 0, number: 1, date: 2, string: 3, boolean: 4, error: 5 } as const
export const SLICER_NO_TEXT = 0xffff_ffff

export interface SlicerMembersParams {
  mode: 'items' | 'canvas' | 'metadata'
  rowCount: number
  dictionaryCount: number
  textByteLength: number
  start: number
  count: number
  locale: string
  sortOrder: 'ascending' | 'descending'
  customListSort: boolean
  crossFilter: 'none' | 'showItemsWithDataAtTop' | 'showItemsWithNoData'
  hideItemsWithNoData: boolean
}

/** String/Error use a shared, exact UTF-16LE dictionary. It preserves empty
 * strings and lone surrogates. Nontext rows use SLICER_NO_TEXT. The Worker
 * fills textRanks using the same Intl comparator as the synchronous adapter;
 * equal collation/calendar ranks do not collapse distinct raw identities.
 * Canvas keeps two rowCount segments in textIds: raw text IDs followed by
 * exact source-caption IDs for every row. Its availability bit 0 is hasData,
 * bit 1 is selected; items/metadata retain the original 0/1 availability.
 */
export interface SlicerMembersTask {
  kind: 'slicerMembers'
  values: Float64Array
  states: Uint8Array
  textIds: Uint32Array
  availability: Uint8Array
  textOffsets: Uint32Array
  textBytes: Uint8Array
  textRanks: Uint32Array
  params: SlicerMembersParams
}

/** Indices are first-occurrence source-row offsets. Only the requested window
 * crosses back to the host; total counts the complete display projection.
 * Bounds/count refer to all typed Date rows before availability projection.
 * Canvas first projects sorted typed members, globally folds equal source
 * captions, then windows the resulting caption total. Canvas hasData packs
 * bit 0=hasData, bit 1=fully selected, bit 2=partial (bits 1/2 exclusive).
 * Metadata mode returns no members (total=0) and skips deduplication/sorting.
 */
export interface SlicerMembersOutput {
  kind: 'slicerMembers'
  indices: Uint32Array
  hasData: Uint8Array
  total: number
  dateCount: number
  minDate: number | null
  maxDate: number | null
}
