/** Inert Office source metadata; the editable per-cell SparklineSpec stays authoritative. */
export interface OfficeSparklineSource {
  version: 1
  /** A detached x14:sparklineGroup template, excluding its member sparklines. */
  groupXml: string
  /** Resolved colors at import, used to detect authored changes to native color bindings. */
  resolvedColors: Record<string, string>
  themeXml?: string
  indexedColors?: string[]
}
