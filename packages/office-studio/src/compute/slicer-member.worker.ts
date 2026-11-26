import { ComputeError } from './errors.ts'
import type { SlicerMembersTask } from './slicer-member-protocol.ts'
import { createSlicerTextComparator } from '../features/slicer/slicer-member-order.ts'

/** Only the locale boundary stays in JS. Exact identity/deduplication,
 * availability OR and final typed ordering/windowing belong to Rust.
 * Decoding/rank merge writes are bounded and yield in the existing Worker.
 */
export async function prepareSlicerTextRanks(task: SlicerMembersTask, yieldAndCheck: () => Promise<void>): Promise<void> {
  if (!task.params.dictionaryCount) return
  let compare: (left: string, right: string) => number
  try { compare = createSlicerTextComparator(task.params.customListSort, task.params.locale) }
  catch (error) { throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器 locale 无法建立文本比较规则', String(error)) }
  if (task.textOffsets[0] !== 0 || task.textOffsets[task.params.dictionaryCount] !== task.textBytes.length) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器 UTF-16 字典边界无效')
  }
  const texts: string[] = [], count = task.params.dictionaryCount
  let decodedBytes = 0
  for (let index = 0; index < count; index++) {
    const start = task.textOffsets[index], end = task.textOffsets[index + 1]
    if (start > end || end > task.textBytes.length || start % 2 !== 0 || end % 2 !== 0) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器 UTF-16 字典偏移无效')
    }
    const parts: string[] = []
    for (let position = start; position < end;) {
      const units: number[] = [], limit = Math.min(end, position + 8192)
      for (; position < limit; position += 2) units.push(task.textBytes[position] | task.textBytes[position + 1] << 8)
      parts.push(String.fromCharCode(...units))
      decodedBytes += units.length * 2
      if (decodedBytes >= 65_536) { decodedBytes = 0; await yieldAndCheck() }
    }
    texts.push(parts.join(''))
    if ((index + 1) % 2048 === 0) await yieldAndCheck()
  }
  let order = new Uint32Array(count), scratch = new Uint32Array(count)
  for (let index = 0; index < count; index++) {
    order[index] = index
    if ((index + 1) % 16_384 === 0) await yieldAndCheck()
  }
  let writes = 0
  for (let width = 1; width < count; width *= 2) {
    for (let left = 0; left < count; left += width * 2) {
      const middle = Math.min(count, left + width), right = Math.min(count, left + width * 2)
      let first = left, second = middle
      for (let out = left; out < right; out++) {
        const takeLeft = first < middle && (second >= right || compare(texts[order[first]], texts[order[second]]) <= 0)
        scratch[out] = takeLeft ? order[first++] : order[second++]
        if (++writes >= 16_384) { writes = 0; await yieldAndCheck() }
      }
    }
    const previous = order; order = scratch; scratch = previous
  }
  let rank = 0
  for (let index = 0; index < count; index++) {
    if (index && compare(texts[order[index - 1]], texts[order[index]]) !== 0) rank++
    task.textRanks[order[index]] = rank
    if ((index + 1) % 16_384 === 0) await yieldAndCheck()
  }
  await yieldAndCheck()
}
