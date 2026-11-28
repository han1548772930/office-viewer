/** Closed rectangle order: startRow,endRow,startColumn,endColumn. Inputs are disjoint and clipped to worksheet dimensions. */
export interface CheckboxCellsParams {
  rowCount: number; columnCount: number; rangeCount: number; mergeCount: number; existingCount: number
  coveredCells: number; maxOutputRuns: number
}
/** Only existing cells need explicit actions: 0=preserve, 1=unchecked, 2=checked. Missing cells imply action 1. */
export interface CheckboxCellsTask {
  kind: 'checkboxCells'; ranges: Uint32Array; merges: Uint32Array; positions: Uint32Array; actions: Uint8Array; params: CheckboxCellsParams
}
/** Each run is startRow,endRow,startColumn,endColumn,action (1 or 2); no partial results are published. */
export interface CheckboxCellsOutput { kind: 'checkboxCells'; runs: Uint32Array; changedCells: number }
