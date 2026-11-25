import type { DocumentTextFieldRange } from './document-text-fields'
import { compareDocumentNoteBoundaries } from './document-note-boundary-edit'

const MAX_FIELD_DEPTH = 512
const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)

/** Source parent identity orders empty/equal result ranges; other fields infer a
 * parent from the same offset + cumulative-reference geometry used by DOCX. */
export function documentTextFieldTopologyParents(ranges: readonly DocumentTextFieldRange[]): readonly (number | undefined)[] {
  const byId = new Map(ranges.map((range, index) => [range.field.id, index]))
  const parents: (number | undefined)[] = ranges.map(() => undefined), infer = ranges.map(() => true)
  const sequence = ranges.map((range, index) => {
    const source = range.field.metadata?.source
    if (source !== undefined && source !== null && !record(source)) throw new TypeError('正文字段 source 必须为对象。')
    const value = record(source) ? source.beginSequence : undefined
    if (value !== undefined && (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)) throw new RangeError('正文字段 beginSequence 必须为非负安全整数。')
    return typeof value === 'number' ? value : index
  })
  const start = (left: number, right: number) => compareDocumentNoteBoundaries(ranges[left]!.absoluteStart, ranges[left]!.startNoteReferenceCount, ranges[right]!.absoluteStart, ranges[right]!.startNoteReferenceCount)
  const end = (left: number, right: number) => compareDocumentNoteBoundaries(ranges[left]!.absoluteEnd, ranges[left]!.endNoteReferenceCount, ranges[right]!.absoluteEnd, ranges[right]!.endNoteReferenceCount)
  const endStart = (left: number, right: number) => compareDocumentNoteBoundaries(ranges[left]!.absoluteEnd, ranges[left]!.endNoteReferenceCount, ranges[right]!.absoluteStart, ranges[right]!.startNoteReferenceCount)
  const nonempty = (index: number) => compareDocumentNoteBoundaries(ranges[index]!.absoluteStart, ranges[index]!.startNoteReferenceCount, ranges[index]!.absoluteEnd, ranges[index]!.endNoteReferenceCount) < 0
  for (const [index, range] of ranges.entries()) {
    const source = range.field.metadata?.source, parentId = record(source) ? source.parentFieldId : undefined
    if (parentId === undefined) continue
    infer[index] = false
    if (parentId === null) continue
    if (typeof parentId !== 'string' || !parentId.trim() || parentId.length > 4096) throw new TypeError('正文字段 parentFieldId 必须为有效标识或 null。')
    const parent = byId.get(parentId)
    if (parent === undefined || parent === index) throw new RangeError('正文字段父身份不存在或指向自身。')
    if (start(parent, index) > 0 || end(parent, index) < 0) throw new RangeError('正文字段结果范围超出显式父字段。')
    parents[index] = parent
  }
  const depths = () => {
    const state = ranges.map(() => 0), result = ranges.map(() => 0)
    for (let first = 0; first < ranges.length; first++) {
      const path: number[] = []
      let current: number | undefined = first
      while (current !== undefined && state[current] !== 2) {
        if (state[current] === 1) throw new RangeError('正文字段父身份形成循环。')
        state[current] = 1; path.push(current)
        if (path.length > MAX_FIELD_DEPTH + 1) throw new RangeError('正文字段嵌套层级超出 512。')
        current = parents[current]
      }
      for (const index of path.reverse()) {
        const parent = parents[index]
        result[index] = parent === undefined ? 0 : result[parent]! + 1
        if (result[index]! > MAX_FIELD_DEPTH) throw new RangeError('正文字段嵌套层级超出 512。')
        state[index] = 2
      }
    }
    return result
  }
  const explicitDepths = depths(), enclosing: number[] = []
  const order = ranges.map((_, index) => index).sort((left, right) => start(left, right) || end(right, left)
    || explicitDepths[left]! - explicitDepths[right]! || sequence[left]! - sequence[right]! || left - right)
  for (const index of order) {
    while (enclosing.length && endStart(enclosing.at(-1)!, index) <= 0) enclosing.pop()
    const parent = enclosing.at(-1)
    if (infer[index] && parent !== undefined) {
      if (end(parent, index) < 0) throw new RangeError('正文字段结果范围部分交叉。')
      parents[index] = parent
    }
    if (nonempty(index)) enclosing.push(index)
  }
  depths()
  const siblings = new Map<number | undefined, number[]>()
  for (const [index, parent] of parents.entries()) {
    const group = siblings.get(parent) ?? []
    group.push(index); siblings.set(parent, group)
  }
  for (const group of siblings.values()) {
    group.sort((left, right) => start(left, right) || Number(nonempty(left)) - Number(nonempty(right)) || sequence[left]! - sequence[right]! || left - right)
    for (let index = 1; index < group.length; index++) if (endStart(group[index - 1]!, group[index]!) > 0) {
      throw new RangeError('正文字段兄弟结果范围交叉或声明了错误的父身份。')
    }
  }
  return parents
}
