import { isRecord } from './protocol.ts'
import { isUnsignedCount } from '../platform/wasm/task-protocol.ts'
import type { CheckboxCellsOutput, CheckboxCellsParams } from './checkbox-cell-protocol.ts'

/** Output extents and ownership are checked at the host; the Rust kernel owns range/merge membership and non-overlap correctness. */
export function readCheckboxCellsOutput(value: unknown, params: CheckboxCellsParams): CheckboxCellsOutput | null {
  if (!isRecord(value) || value.kind !== 'checkboxCells' || !(value.runs instanceof Uint32Array)
    || !(value.runs.buffer instanceof ArrayBuffer) || value.runs.byteOffset !== 0 || value.runs.byteLength !== value.runs.buffer.byteLength
    || value.runs.length % 5 || value.runs.length > params.maxOutputRuns * 5
    || !isUnsignedCount(value.changedCells) || value.changedCells > params.coveredCells) return null
  try { new Uint8Array(value.runs.buffer, 0, 0) } catch { return null }
  let cells = 0
  for (let i = 0; i < value.runs.length; i += 5) {
    const [sr, er, sc, ec, action] = value.runs.subarray(i, i + 5)
    if (sr > er || sc > ec || er >= params.rowCount || ec >= params.columnCount || action !== 1 && action !== 2) return null
    cells += (er - sr + 1) * (ec - sc + 1)
  }
  return cells === value.changedCells ? { kind: 'checkboxCells', runs: value.runs, changedCells: value.changedCells } : null
}
