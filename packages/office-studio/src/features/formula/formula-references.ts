import type { FormulaReference } from './formula-types.ts'

const MAX_COLUMN = 16_383
const MAX_ROW = 1_048_575

interface ReferenceSpan { start: number; end: number; replacementStart: number; reference: FormulaReference; raw: string }
interface TextPatch { start: number; end: number; replacement: string }
export interface FormulaStructureEdit {
  kind: 'insert' | 'delete'
  axis: 'row' | 'column'
  start: number
  count: number
  bandStart: number
  bandEnd: number
}

/**
 * Translate relative references as a formula is copied/moved. Absolute row and
 * column markers remain fixed; references moved outside the worksheet become
 * #REF!. String literals and sheet names are never scanned as cell references.
 */
export function translateFormulaReferences(formula: string, rowDelta: number, columnDelta: number): string {
  if (typeof formula !== 'string' || !formula.length) throw new TypeError('Formula must be a non-empty string')
  if (!Number.isSafeInteger(rowDelta) || !Number.isSafeInteger(columnDelta)) throw new TypeError('Formula reference offsets must be safe integers')
  const references = findReferenceSpans(formula)
  const patches: TextPatch[] = []
  for (let index = 0; index < references.length; index += 1) {
    const current = references[index]!
    const next = references[index + 1]
    if (next && formula.slice(current.end, next.start).match(/^\s*:\s*$/)) {
      const first = shiftReference(current.reference, rowDelta, columnDelta)
      const second = shiftReference(next.reference, rowDelta, columnDelta)
      if (!first || !second) {
        patches.push({ start: current.replacementStart, end: next.end, replacement: '#REF!' })
        index += 1
        continue
      }
      patches.push({ start: current.replacementStart, end: current.end, replacement: `${sheetPrefix(formula, current)}${formatReference(first)}` })
      patches.push({ start: next.replacementStart, end: next.end, replacement: `${sheetPrefix(formula, next)}${formatReference(second)}` })
      index += 1
      continue
    }
    const moved = shiftReference(current.reference, rowDelta, columnDelta)
    patches.push({ start: current.replacementStart, end: current.end, replacement: moved ? `${sheetPrefix(formula, current)}${formatReference(moved)}` : '#REF!' })
  }
  let result = formula
  for (const patch of patches.sort((left, right) => right.start - left.start)) {
    result = result.slice(0, patch.start) + patch.replacement + result.slice(patch.end)
  }
  return result
}

/** Update references affected by inserting or deleting cells in one worksheet band. */
export function transformFormulaReferencesForStructureEdit(
  formula: string,
  edit: FormulaStructureEdit,
  formulaSheetName: string | undefined,
  targetSheetName: string,
): string {
  if (typeof formula !== 'string' || !formula.length) throw new TypeError('Formula must be a non-empty string')
  if (!Number.isSafeInteger(edit.start) || edit.start < 0 || !Number.isSafeInteger(edit.count) || edit.count < 1
    || !Number.isSafeInteger(edit.bandStart) || edit.bandStart < 0 || !Number.isSafeInteger(edit.bandEnd) || edit.bandEnd < edit.bandStart) {
    throw new RangeError('Invalid worksheet structure edit')
  }
  const references = findReferenceSpans(formula)
  const patches: TextPatch[] = []
  for (let index = 0; index < references.length; index += 1) {
    const first = references[index]!
    const second = references[index + 1]
    const isRange = Boolean(second && formula.slice(first.end, second.start).match(/^\s*:\s*$/))
    const qualifiedSheet = referenceSheetName(first.raw)
    const rangeSheet = qualifiedSheet ?? (isRange && second ? referenceSheetName(second.raw) : undefined)
    const applies = rangeSheet === undefined
      ? formulaSheetName?.toLowerCase() === targetSheetName.toLowerCase()
      : rangeSheet.toLowerCase() === targetSheetName.toLowerCase()
    if (!applies) {
      if (isRange) index += 1
      continue
    }

    if (isRange && second) {
      // Univer's formula ref-range path moves a rectangular reference only when
      // its entire perpendicular span is contained in the shifted band.
      const firstBand = edit.axis === 'row' ? first.reference.column : first.reference.row
      const secondBand = edit.axis === 'row' ? second.reference.column : second.reference.row
      if (Math.min(firstBand, secondBand) < edit.bandStart || Math.max(firstBand, secondBand) > edit.bandEnd) {
        index += 1
        continue
      }
      const firstResult = transformStructureReference(first.reference, edit)
      const secondResult = transformStructureReference(second.reference, edit)
      const firstAxisIndex = edit.axis === 'row' ? first.reference.row : first.reference.column
      const secondAxisIndex = edit.axis === 'row' ? second.reference.row : second.reference.column
      const replacement = (edit.kind === 'insert' && (firstResult.deleted || secondResult.deleted)) || (firstResult.deleted && secondResult.deleted)
        ? '#REF!'
        : `${firstResult.deleted ? formatReference(editBoundaryReference(first.reference, edit, firstAxisIndex <= secondAxisIndex ? 0 : -1)) : formatReference(firstResult.reference)}:${sheetPrefix(formula, second)}${secondResult.deleted ? formatReference(editBoundaryReference(second.reference, edit, secondAxisIndex <= firstAxisIndex ? 0 : -1)) : formatReference(secondResult.reference)}`
      patches.push({ start: first.replacementStart, end: second.end, replacement: `${sheetPrefix(formula, first)}${replacement}` })
      index += 1
      continue
    }

    const result = transformStructureReference(first.reference, edit)
    patches.push({
      start: first.replacementStart,
      end: first.end,
      replacement: result.deleted ? '#REF!' : `${sheetPrefix(formula, first)}${formatReference(result.reference)}`,
    })
  }

  let result = formula
  for (const patch of patches.sort((left, right) => right.start - left.start)) {
    result = result.slice(0, patch.start) + patch.replacement + result.slice(patch.end)
  }
  return result
}

function transformStructureReference(reference: FormulaReference, edit: FormulaStructureEdit): { reference: FormulaReference; deleted: boolean } {
  const bandIndex = edit.axis === 'row' ? reference.column : reference.row
  const axisIndex = edit.axis === 'row' ? reference.row : reference.column
  if (bandIndex < edit.bandStart || bandIndex > edit.bandEnd || axisIndex < edit.start) return { reference, deleted: false }
  if (edit.kind === 'insert') {
    const moved = setAxisIndex(reference, edit.axis, axisIndex + edit.count)
    return { reference: moved, deleted: moved.row > MAX_ROW || moved.column > MAX_COLUMN }
  }
  if (axisIndex < edit.start + edit.count) return { reference, deleted: true }
  return { reference: setAxisIndex(reference, edit.axis, axisIndex - edit.count), deleted: false }
}

function editBoundaryReference(reference: FormulaReference, edit: FormulaStructureEdit, offset: 0 | -1): FormulaReference {
  return setAxisIndex(reference, edit.axis, Math.max(0, edit.start + offset))
}

function setAxisIndex(reference: FormulaReference, axis: FormulaStructureEdit['axis'], value: number): FormulaReference {
  return axis === 'row' ? { ...reference, row: value } : { ...reference, column: value }
}

function referenceSheetName(raw: string): string | undefined {
  const separator = raw.lastIndexOf('!')
  if (separator < 0) return undefined
  const source = raw.slice(0, separator)
  return source.startsWith("'") && source.endsWith("'")
    ? source.slice(1, -1).replace(/''/g, "'")
    : source
}

function findReferenceSpans(source: string): ReferenceSpan[] {
  const references: ReferenceSpan[] = []
  let index = 0
  while (index < source.length) {
    const character = source[index]!
    if (character === '"') {
      index += 1
      while (index < source.length) {
        if (source[index] === '"') {
          if (source[index + 1] === '"') { index += 2; continue }
          index += 1; break
        }
        index += 1
      }
      continue
    }
    if (character === "'") {
      const start = index++
      while (index < source.length) {
        if (source[index] === "'") {
          if (source[index + 1] === "'") { index += 2; continue }
          index += 1
          if (source[index] === '!') {
            const reference = readAddress(source, index + 1)
            if (reference) {
              references.push({ start, end: reference.end, replacementStart: start, reference: reference.reference, raw: source.slice(start, reference.end) })
              index = reference.end
              continue
            }
          }
          break
        }
        index += 1
      }
      continue
    }

    if (isIdentifierPart(source[index - 1])) { index += 1; continue }
    const unquotedQualifier = /^([\p{L}_\\][\p{L}\p{N}_.\\]*)!/u.exec(source.slice(index))
    if (unquotedQualifier) {
      const qualifiedAddress = readAddress(source, index + unquotedQualifier[0].length)
      if (qualifiedAddress) {
        references.push({
          start: index,
          end: qualifiedAddress.end,
          replacementStart: index,
          reference: { ...qualifiedAddress.reference, sheetName: unquotedQualifier[1] },
          raw: source.slice(index, qualifiedAddress.end),
        })
        index = qualifiedAddress.end
        continue
      }
    }

    const address = readAddress(source, index)
    if (!address) { index += 1; continue }

    // A sheet called "Q1" is lexically indistinguishable from a cell reference
    // until the following !. Do not shift the qualifier itself.
    if (source[address.end] === '!') {
      const sheetReference = readAddress(source, address.end + 1)
      if (sheetReference) {
        references.push({ start: index, end: sheetReference.end, replacementStart: index, reference: sheetReference.reference, raw: source.slice(index, sheetReference.end) })
        index = sheetReference.end
        continue
      }
    }
    references.push({ start: index, end: address.end, replacementStart: index, reference: address.reference, raw: source.slice(index, address.end) })
    index = address.end
  }
  return references
}

function readAddress(source: string, start: number): { reference: FormulaReference; end: number } | undefined {
  let cursor = start
  const absoluteColumn = source[cursor] === '$'
  if (absoluteColumn) cursor += 1
  const lettersStart = cursor
  while (/[A-Za-z]/.test(source[cursor] ?? '')) cursor += 1
  const letters = source.slice(lettersStart, cursor)
  if (!letters || letters.length > 3) return undefined
  const absoluteRow = source[cursor] === '$'
  if (absoluteRow) cursor += 1
  const digitsStart = cursor
  while (/\d/.test(source[cursor] ?? '')) cursor += 1
  const digits = source.slice(digitsStart, cursor)
  if (!digits || /^0+$/.test(digits) || isIdentifierPart(source[cursor])) return undefined
  const column = letters.toUpperCase().split('').reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1
  const row = Number(digits) - 1
  if (column < 0 || column > MAX_COLUMN || row < 0 || row > MAX_ROW) return undefined
  return { reference: { kind: 'cell', row, column, absoluteRow, absoluteColumn }, end: cursor }
}

function shiftReference(reference: FormulaReference, rowDelta: number, columnDelta: number): FormulaReference | undefined {
  const row = reference.absoluteRow ? reference.row : reference.row + rowDelta
  const column = reference.absoluteColumn ? reference.column : reference.column + columnDelta
  if (row < 0 || row > MAX_ROW || column < 0 || column > MAX_COLUMN) return undefined
  return { ...reference, row, column }
}

function formatReference(reference: FormulaReference): string {
  let column = reference.column + 1
  let label = ''
  while (column > 0) {
    const remainder = (column - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    column = Math.floor((column - 1) / 26)
  }
  return `${reference.absoluteColumn ? '$' : ''}${label}${reference.absoluteRow ? '$' : ''}${reference.row + 1}`
}

function sheetPrefix(_source: string, span: ReferenceSpan): string {
  const raw = span.raw
  const bang = raw.lastIndexOf('!')
  return bang < 0 ? '' : raw.slice(0, bang + 1)
}

function isIdentifierPart(value: string | undefined): boolean { return value !== undefined && /[\p{L}\p{N}_.\\]/u.test(value) }

