/** The fixed public IRange sentinel is provenance, not a substitute data range. */
export interface SparklineBrokenSourceRange {
  startRow: -9999
  endRow: -9999
  startColumn: -9999
  endColumn: -9999
  sheetId?: string
  unitId?: string
  rangeType?: 0 | 1 | 2 | 3
}

export interface SparklineSourceReference {
  version: 1
  invalidRange: SparklineBrokenSourceRange
}
