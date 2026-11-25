import type { FontStyle } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'
import type { ExcelCellView } from '../../services/excelViewAdapter.ts'
import { isSheetTextNumberFormat } from './sheet-number-format.ts'
import { sheetCellRichTextMatches, spliceSheetCellRichText } from './sheet-cell-rich-text-edit.ts'
import { setSheetCellRichTextFont, sheetCellFontValues, type SheetRichTextFontField, type SheetRichTextFontValue } from './sheet-cell-rich-text-style.ts'
import { applySheetCellTextEdit, isSheetCellTextBoundary, projectSheetCellInputEditTrace,
  type SheetCellInputEditTrace, type SheetCellTextEditOperation } from './sheet-cell-input-trace.ts'
import { nextSheetFontFormatSize, type SheetFontFormatAction } from './sheet-font-format-command.ts'

export interface SheetCellInlineFormatEdit {
  /** Number of text edits already applied when the font command ran. */
  at: number
  input: string
  start: number
  end: number
  field: SheetRichTextFontField
  value: SheetRichTextFontValue
}
export interface SheetCellInlineFormatReplay {
  cell: ExcelCellView
  typing?: { position: number; font: Partial<FontStyle> }
}
const fontFields: readonly SheetRichTextFontField[] = ['family', 'size', 'bold', 'italic', 'underline', 'strike', 'color']

export function readSheetCellInlineFormats(value: unknown, trace: SheetCellInputEditTrace): SheetCellInlineFormatEdit[] {
  if (!Array.isArray(value)) throw new TypeError('局部文字格式轨迹必须是数组。')
  const inputs = [trace.source]
  for (const operation of trace.operations) {
    const next = applySheetCellTextEdit(inputs.at(-1)!, operation)
    if (next === undefined) throw new TypeError('文字格式对应的输入轨迹无效。')
    inputs.push(next)
  }
  let previous = 0
  return value.map(item => {
    if (!item || typeof item !== 'object' || Array.isArray(item)) throw new TypeError('局部文字格式操作必须是对象。')
    const edit = item as Record<string, unknown>
    if (typeof edit.at !== 'number' || !Number.isSafeInteger(edit.at) || edit.at < previous || edit.at > trace.operations.length
      || typeof edit.input !== 'string' || edit.input !== inputs[edit.at]
      || typeof edit.start !== 'number' || typeof edit.end !== 'number' || edit.start > edit.end
      || !isSheetCellTextBoundary(edit.input, edit.start) || !isSheetCellTextBoundary(edit.input, edit.end)
      || typeof edit.field !== 'string' || !fontFields.includes(edit.field as SheetRichTextFontField)) throw new TypeError('局部文字格式范围或输入版本无效。')
    const field = edit.field as SheetRichTextFontField, current = edit.value
    if (current !== null && (field === 'family' ? typeof current !== 'string' || !current.trim()
      : field === 'size' ? typeof current !== 'number' || !Number.isFinite(current) || current <= 0
        : field === 'color' ? typeof current !== 'string' : typeof current !== 'boolean')) throw new TypeError('局部文字格式值无效。')
    previous = edit.at
    return { at: edit.at, input: edit.input, start: edit.start, end: edit.end, field, value: current as SheetRichTextFontValue }
  })
}

/** Split source entities only where a selected interval changes; an original ID is retained once. */
export function formatSheetCellRichTextRange(runs: ReadonlyArray<Record<string, unknown>>, text: string,
  start: number, end: number, field: SheetRichTextFontField, value: SheetRichTextFontValue): Record<string, unknown>[] | undefined {
  if (!sheetCellRichTextMatches(runs, text) || !isSheetCellTextBoundary(text, start)
    || !isSheetCellTextBoundary(text, end) || start > end) return undefined
  let offset = 0
  const result: Record<string, unknown>[] = []
  for (const run of runs) {
    const source = run.text as string, first = offset; offset += source.length
    const selectedStart = Math.max(first, start), selectedEnd = Math.min(offset, end)
    const styled = selectedStart < selectedEnd ? setSheetCellRichTextFont([run], field, value) : undefined
    if (!styled?.changed) { result.push(cloneModel(run)); continue }
    const parts = [
      { ...cloneModel(run), text: source.slice(0, selectedStart - first) },
      { ...cloneModel(styled.runs![0]!), text: source.slice(selectedStart - first, selectedEnd - first) },
      { ...cloneModel(run), text: source.slice(selectedEnd - first) },
    ].filter(part => part.text.length > 0)
    parts.forEach((part, index) => { if (index > 0) delete part.id })
    result.push(...parts)
  }
  return result
}

/** Interleave style operations with the existing precise input trace, including apostrophe and typing affinity. */
export function replaySheetCellInlineFormats(cell: ExcelCellView | undefined, trace: SheetCellInputEditTrace,
  formats: readonly SheetCellInlineFormatEdit[], options: { numberFormat: string; overwrite?: boolean }): SheetCellInlineFormatReplay | undefined {
  const textFormat = isSheetTextNumberFormat(options.numberFormat)
  let editor = trace.source
  const rawInput = (input: string) => !textFormat && input.startsWith("'") ? input.slice(1) : input
  let raw = rawInput(editor)
  if (!options.overwrite && cell?.richText?.length && !sheetCellRichTextMatches(cell.richText, raw)) return undefined
  let runs = !options.overwrite && cell?.richText?.length ? cell.richText.map(run => cloneModel(run)) : [{ text: raw } as Record<string, unknown>]
  let typing: SheetCellInlineFormatReplay['typing'], formatIndex = 0
  let emptyDonor: Record<string, unknown> | undefined, emptyDonorPosition: number | undefined
  for (let at = 0; at <= trace.operations.length; at++) {
    while (formats[formatIndex]?.at === at) {
      const format = formats[formatIndex++]!
      if (format.input !== editor) return undefined
      const prefix = editor.length - raw.length, start = Math.max(0, format.start - prefix), end = Math.max(0, format.end - prefix)
      if (start === end) {
        if (!typing || typing.position !== start) typing = { position: start, font: {} }
        const view = { ...(cell ?? { value: '' }), value: raw, rawValue: raw, formula: undefined, richText: runs }
        const inherited = sheetCellFontValues({ ...view, richText: undefined }, format.field)[0]
        const effective = format.value ?? inherited
        // Values remain typed by readSheetCellInlineFormats(); assign without widening FontStyle's fields.
        typing.font = { ...typing.font, [format.field]: effective }
      } else {
        const next = formatSheetCellRichTextRange(runs, raw, start, end, format.field, format.value)
        if (!next) return undefined
        runs = next; typing = undefined
      }
    }
    const source = trace.operations[at]
    if (!source) break
    const nextEditor = applySheetCellTextEdit(editor, source)
    if (nextEditor === undefined) return undefined
    const nextRaw = rawInput(nextEditor), projected = projectSheetCellInputEditTrace(
      { ...trace, source: editor, operations: [source] }, raw, nextRaw, textFormat)
    if (!projected) return undefined
    for (const operation of projected) {
      if (operation.end > operation.start) {
        let offset = 0
        for (let index = 0; index < runs.length; index++) {
          const run = runs[index]!, end = offset + (run.text as string).length
          if (offset <= operation.start && end > operation.start) { emptyDonor = cloneModel(run); emptyDonorPosition = index; break }
          offset = end
        }
      }
      const next = applySheetCellTextEdit(raw, operation)
      if (next === undefined) return undefined
      const edited = spliceSheetCellRichText(runs, raw, operation, emptyDonor, emptyDonorPosition)
      if (!edited) return undefined
      runs = edited
      if (typing && operation.start === typing.position && operation.text) {
        for (const field of fontFields) {
          const value = typing.font[field]
          if (value === undefined) continue
          const styled = formatSheetCellRichTextRange(runs, next, operation.start, operation.start + operation.text.length, field, value)
          if (!styled) return undefined
          runs = styled
        }
      }
      if (typing) {
        if (operation.end <= typing.position) typing.position += operation.text.length - (operation.end - operation.start)
        else if (operation.start <= typing.position) typing.position = operation.start + operation.text.length
      }
      raw = next
    }
    editor = nextEditor
  }
  if (formatIndex !== formats.length || raw !== rawInput(editor)) return undefined
  const richText = cell?.richText?.length && !options.overwrite || runs.some(run => Object.keys(run).some(key => key !== 'text')) ? runs : undefined
  return { cell: { ...(cell ?? { value: '' }), formula: undefined, value: raw, rawValue: raw, richText }, ...(typing ? { typing } : {}) }
}

export function sheetInlineFontChangeForAction(action: SheetFontFormatAction, cell: ExcelCellView | undefined,
  selection: { start: number; end: number }, typing?: SheetCellInlineFormatReplay['typing']): Pick<SheetCellInlineFormatEdit, 'field' | 'value'> | undefined {
  if (action.kind === 'color') return action.field === 'color' ? { field: 'color', value: action.value } : undefined
  if (action.kind === 'family' || action.kind === 'size') return { field: action.kind, value: action.value }
  const field = action.kind === 'mark' ? action.field : 'size'
  const caretValue = selection.start === selection.end && typing?.position === selection.start ? typing.font[field] : undefined
  const values = caretValue === undefined ? sheetCellFontValues(cell, field, selection) : [caretValue]
  return action.kind === 'mark' ? { field, value: action.value ?? !values.every(value => value === true) }
    : { field: 'size', value: nextSheetFontFormatSize(typeof values[0] === 'number' ? values[0] : cell?.fontSize ?? 11, action.direction) }
}

interface AnchoredFormat { edit: SheetCellInlineFormatEdit; prefix: SheetCellTextEditOperation[] }
/** Transient formatting follows native text undo/redo and drops entries belonging to an abandoned text branch. */
export class SheetCellInlineFormatDraft {
  private entries: AnchoredFormat[] = []
  private undone: AnchoredFormat[] = []
  canUndo(trace: SheetCellInputEditTrace): boolean { return this.formats(trace).some(edit => edit.at === trace.operations.length) }
  canRedo(trace: SheetCellInputEditTrace): boolean {
    const entry = this.undone.at(-1)
    return !!entry && entry.edit.at === trace.operations.length && entry.prefix.every((operation, index) => JSON.stringify(operation) === JSON.stringify(trace.operations[index]))
  }
  formats(trace: SheetCellInputEditTrace): SheetCellInlineFormatEdit[] {
    return this.entries.filter(entry => entry.edit.at <= trace.operations.length
      && entry.prefix.every((operation, index) => JSON.stringify(operation) === JSON.stringify(trace.operations[index])))
      .map(entry => ({ ...entry.edit })).sort((left, right) => left.at - right.at)
  }
  reconcile(trace: SheetCellInputEditTrace): void {
    const matches = (entry: AnchoredFormat) => entry.prefix.slice(0, trace.operations.length)
      .every((operation, index) => JSON.stringify(operation) === JSON.stringify(trace.operations[index]))
    this.entries = this.entries.filter(matches)
    this.undone = this.undone.filter(matches)
  }
  append(trace: SheetCellInputEditTrace, edit: SheetCellInlineFormatEdit): void {
    this.reconcile(trace)
    this.entries = this.entries.filter(entry => entry.edit.at <= trace.operations.length)
    this.entries.push({ edit: { ...edit }, prefix: trace.operations.map(operation => ({ ...operation })) })
    this.undone = []
  }
  undo(trace: SheetCellInputEditTrace): boolean {
    this.reconcile(trace)
    let index = this.entries.length - 1
    while (index >= 0 && this.entries[index]!.edit.at !== trace.operations.length) index--
    if (index < 0) return false
    this.undone.push(this.entries.splice(index, 1)[0]!)
    return true
  }
  redo(trace: SheetCellInputEditTrace): boolean {
    this.reconcile(trace)
    const entry = this.undone.at(-1)
    if (!entry || entry.edit.at !== trace.operations.length) return false
    this.undone.pop(); this.entries.push(entry)
    return true
  }
}
