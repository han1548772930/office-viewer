export interface DocumentParagraphFragmentPlan {
  lineCount: number
  /** Required at a new fragment when the current remaining space cannot obey the rule. */
  requiredHeight: number
}

const EPSILON = 1e-6

/** Greedy whole-line fragments with two-line first/last pieces when feasible.
 * Impossible one-line frames and oversized lines still make finite progress. */
export function planDocumentParagraphFragment(heights: readonly number[], start: number, room: number, fullRoom: number,
  widowControl: boolean, trailingHeight = 0): DocumentParagraphFragmentPlan {
  let used = 0, count = 0
  for (let index = start; index < heights.length; index++) {
    const next = used + heights[index]!, tail = index === heights.length - 1 ? trailingHeight : 0
    if (next + tail > room + EPSILON) break
    used = next; count++
  }
  const remaining = heights.length - start, first = heights[start] ?? 0
  if (remaining <= 0) return { lineCount: 0, requiredHeight: 0 }
  if (!count) return { lineCount: 0, requiredHeight: first + (remaining === 1 ? trailingHeight : 0) }
  if (!widowControl || count === remaining || remaining < 2) return { lineCount: count, requiredHeight: first }
  const firstTwo = first + heights[start + 1]! + (remaining === 2 ? trailingHeight : 0)
  if (count === 1 && firstTwo <= fullRoom + EPSILON) return { lineCount: 0, requiredHeight: firstTwo }
  if (remaining - count === 1) {
    if (count >= 3) return { lineCount: count - 1, requiredHeight: first }
    const whole = heights.slice(start).reduce((sum, value) => sum + value, trailingHeight)
    if (whole <= fullRoom + EPSILON) return { lineCount: 0, requiredHeight: whole }
  }
  return { lineCount: count, requiredHeight: first }
}
