import { ComputeError } from '../../compute/errors.ts'
import { COMPUTE_LIMITS } from '../../compute/protocol.ts'
import { estimateSlicerMembersMemory } from '../../compute/slicer-member-tasks.ts'
import { SLICER_MEMBER_TYPE, SLICER_NO_TEXT } from '../../compute/slicer-member-protocol.ts'
import type { SlicerMembersParams, SlicerMembersTask } from '../../compute/slicer-member-protocol.ts'
import { normalizeRecordValue, type RecordValue } from '../../domain/record-value.ts'

export type SlicerMemberInputOptions = Omit<SlicerMembersParams, 'rowCount' | 'dictionaryCount' | 'textByteLength'>
export interface SlicerCanvasSourceRow { caption: string; selected: boolean }

/** Source adapter only interns literal dictionary strings and packs rows. It
 * never deduplicates typed members or chooses their order/availability/window.
 * Metadata encodes non-Date rows as blank; no arbitrary serial/string fallback.
 */
export class SlicerMemberInputBuilder {
  private readonly values: Float64Array
  private readonly states: Uint8Array
  private readonly textIds: Uint32Array
  private readonly availability: Uint8Array
  private readonly dictionary = new Map<string, number>()
  private texts: string[] = []
  private textByteLength = 0
  private position = 0
  private finished = false

  constructor(readonly rowCount: number, private readonly options: SlicerMemberInputOptions,
    private readonly memoryBudgetBytes = COMPUTE_LIMITS.defaultMemoryBudget,
    private readonly reserveMemory?: (estimatedPeakBytes: number) => void) {
    if (!Number.isSafeInteger(rowCount) || rowCount < 0 || rowCount > COMPUTE_LIMITS.maxRows
      || !Number.isSafeInteger(memoryBudgetBytes) || memoryBudgetBytes <= 0 || memoryBudgetBytes > 0xffff_ffff) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器来源行数或内存预算无效')
    }
    this.checkBudget(0, 0)
    try {
      this.values = new Float64Array(rowCount)
      this.states = new Uint8Array(rowCount)
      this.textIds = new Uint32Array(rowCount * (options.mode === 'canvas' ? 2 : 1))
      this.textIds.fill(SLICER_NO_TEXT)
      this.availability = new Uint8Array(rowCount)
    } catch (error) { throw new ComputeError('COMPUTE_MEMORY_LIMIT', '无法分配切片器 typed 来源列', String(error)) }
  }

  private checkBudget(dictionaryCount = this.texts.length, textByteLength = this.textByteLength): void {
    const memory = estimateSlicerMembersMemory({ ...this.options, rowCount: this.rowCount, dictionaryCount, textByteLength })
    if (!Number.isSafeInteger(memory.estimatedPeakBytes) || memory.estimatedPeakBytes > this.memoryBudgetBytes || textByteLength > 0xffff_ffff) {
      throw new ComputeError('COMPUTE_MEMORY_LIMIT', '切片器成员和字典超过查询内存预算')
    }
    this.reserveMemory?.(memory.estimatedPeakBytes)
  }

  private internText(text: string): number {
    let id = this.dictionary.get(text)
    if (id === undefined) {
      const nextBytes = this.textByteLength + text.length * 2
      this.checkBudget(this.texts.length + 1, nextBytes)
      id = this.texts.length
      this.dictionary.set(text, id); this.texts.push(text); this.textByteLength = nextBytes
    }
    return id
  }

  append(input: RecordValue, hasData = true, canvas?: SlicerCanvasSourceRow): void {
    if (this.finished || this.position >= this.rowCount) throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器来源列已结束或超界')
    if (typeof hasData !== 'boolean') throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器来源可用性必须为 boolean')
    if (this.options.mode === 'canvas' && (!canvas || typeof canvas.caption !== 'string' || typeof canvas.selected !== 'boolean')) {
      throw new ComputeError('COMPUTE_INVALID_INPUT', 'Canvas 成员来源必须保留原始 caption 和选中位')
    }
    const raw = normalizeRecordValue(input), value = this.options.mode === 'metadata' && raw.type !== 'date' ? { type: 'blank' } as const : raw
    const row = this.position
    this.states[row] = SLICER_MEMBER_TYPE[value.type]
    this.availability[row] = (hasData ? 1 : 0) | (this.options.mode === 'canvas' && canvas?.selected ? 2 : 0)
    if (value.type === 'number' || value.type === 'date') this.values[row] = value.value
    else if (value.type === 'boolean') this.values[row] = value.value ? 1 : 0
    else if (value.type === 'string' || value.type === 'error') {
      this.textIds[row] = this.internText(value.value)
    }
    if (this.options.mode === 'canvas') this.textIds[this.rowCount + row] = this.internText(canvas!.caption)
    this.position++
  }

  async finish(yieldAndCheck: () => Promise<void>): Promise<SlicerMembersTask> {
    if (this.finished || this.position !== this.rowCount) throw new ComputeError('COMPUTE_INVALID_INPUT', '切片器来源列未完成或已被消费')
    this.finished = true
    this.checkBudget()
    const texts = this.texts, dictionaryCount = texts.length
    this.dictionary.clear()
    this.texts = []
    let textOffsets: Uint32Array, textBytes: Uint8Array, textRanks: Uint32Array
    try {
      textOffsets = new Uint32Array(dictionaryCount + 1)
      textBytes = new Uint8Array(this.textByteLength)
      textRanks = new Uint32Array(dictionaryCount)
    } catch (error) { throw new ComputeError('COMPUTE_MEMORY_LIMIT', '无法分配切片器 UTF-16 来源字典', String(error)) }
    let offset = 0, written = 0
    for (let index = 0; index < dictionaryCount; index++) {
      textOffsets[index] = offset
      const text = texts[index]
      for (let position = 0; position < text.length; position++) {
        const unit = text.charCodeAt(position)
        textBytes[offset++] = unit & 255
        textBytes[offset++] = unit >>> 8
        if (++written >= 32_768) { written = 0; await yieldAndCheck() }
      }
      if ((index + 1) % 2048 === 0) await yieldAndCheck()
    }
    textOffsets[dictionaryCount] = offset
    await yieldAndCheck()
    return { kind: 'slicerMembers', values: this.values, states: this.states, textIds: this.textIds,
      availability: this.availability, textOffsets, textBytes, textRanks,
      params: { ...this.options, rowCount: this.rowCount, dictionaryCount, textByteLength: this.textByteLength } }
  }
}
