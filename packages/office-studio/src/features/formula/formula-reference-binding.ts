import type { CellRange, SpreadsheetSheet, WorkbookModel } from '../../domain/model'
import { parseFormula } from './formula-parser'
import { resolveFormulaDefinedName, type FormulaDefinedName } from './formula-defined-names'
import { findFormulaStructuredTable, resolveFormulaStructuredRange } from './formula-structured-references'
import type { FormulaCellAddress, FormulaNode, FormulaRangeReference } from './formula-types'

export function formulaRangeReference(range: CellRange, sheetName: string): FormulaRangeReference {
  return { kind: 'range', sheetName,
    start: { kind: 'cell', sheetName, row: range.startRow, column: range.startCol, absoluteRow: true, absoluteColumn: true },
    end: { kind: 'cell', sheetName, row: range.endRow, column: range.endCol, absoluteRow: true, absoluteColumn: true } }
}
const rangeNode = (range: CellRange, sheet: SpreadsheetSheet): FormulaNode => {
  const reference = formulaRangeReference(range, sheet.name)
  return range.startRow === range.endRow && range.startCol === range.endCol ? { kind: 'reference', reference: reference.start } : { kind: 'range', reference }
}
interface BindingTask {
  node: FormulaNode
  sheet: SpreadsheetSheet
  qualify: boolean
  assign(node: FormulaNode): void
  leaveName?: string
}

/** Bind names and current table footprints to the existing AST. This compiler
 * never evaluates values and never maintains another calculation engine.
 * An explicit work stack handles long name chains and reports name cycles. */
export function bindFormulaReferences(node: FormulaNode, workbook: WorkbookModel, currentSheet: SpreadsheetSheet,
  names: ReadonlyMap<string, FormulaDefinedName>, cell?: Pick<FormulaCellAddress, 'row' | 'column'>): FormulaNode {
  let result: FormulaNode = node, work = 0
  const active = new Set<string>(), bound = new Map<string, FormulaNode>()
  const sheets = new Map(workbook.sheets.map(sheet => [sheet.name.toLowerCase(), sheet]))
  const tasks: BindingTask[] = [{ node, sheet: currentSheet, qualify: false, assign(value) { result = value } }]
  while (tasks.length) {
    const task = tasks.pop()!
    if (task.leaveName) { active.delete(task.leaveName); continue }
    if (++work > 65_536) throw new RangeError('公式名称与结构化引用展开超出计算预算。')
    const node = task.node
    if (node.kind === 'name') {
      const scope = node.sheetName ? sheets.get(node.sheetName.toLowerCase()) : task.sheet
      if (!scope) { task.assign({ kind: 'error', code: '#REF!' }); continue }
      const name = resolveFormulaDefinedName(names, node.name, scope.id)
      if (!name) {
        const table = workbook.tables?.find(table => table.name.toLowerCase() === node.name.toLowerCase())
        if (!table || node.sheetName && table.sheetId !== scope.id) { task.assign({ kind: 'error', code: '#NAME?' }); continue }
        const sheet = workbook.sheets.find(sheet => sheet.id === table.sheetId), range = resolveFormulaStructuredRange(workbook, table, '[#Data]', cell)
        task.assign(sheet && range ? rangeNode(range, sheet) : { kind: 'error', code: '#REF!' }); continue
      }
      const owner = name.owner ? workbook.sheets.find(sheet => sheet.id === name.owner) : scope
      if (!owner || name.expression === undefined) { task.assign({ kind: 'error', code: '#REF!' }); continue }
      const identity = `${name.identity}:${owner.id}`
      if (active.has(identity)) { task.assign({ kind: 'error', code: '#CIRCULAR!' }); continue }
      const cached = bound.get(identity)
      if (cached) { task.assign(cached); continue }
      active.add(identity)
      const parsed = parseFormula(name.expression)
      tasks.push({ ...task, leaveName: identity })
      tasks.push({ node: parsed, sheet: owner, qualify: true,
        assign(value) { bound.set(identity, value); task.assign(value) } })
      continue
    }
    if (node.kind === 'structured') {
      const table = findFormulaStructuredTable(workbook, node, task.sheet.id, cell)
      const sheet = table && workbook.sheets.find(sheet => sheet.id === table.sheetId)
      const explicit = node.sheetName ? sheets.get(node.sheetName.toLowerCase()) : undefined
      const range = table && resolveFormulaStructuredRange(workbook, table, node.specifier, cell)
      task.assign(sheet && range && (!node.sheetName || explicit?.id === sheet.id) ? rangeNode(range, sheet) : { kind: 'error', code: '#REF!' })
      continue
    }
    if (node.kind === 'reference' || node.kind === 'range') {
      if (!task.qualify || node.reference.sheetName) { task.assign(node); continue }
      task.assign(node.kind === 'reference' ? { ...node, reference: { ...node.reference, sheetName: task.sheet.name } }
        : { ...node, reference: { ...node.reference, sheetName: task.sheet.name,
          start: { ...node.reference.start, sheetName: task.sheet.name }, end: { ...node.reference.end, sheetName: task.sheet.name } } })
      continue
    }
    if (node.kind === 'unary') {
      const next = { ...node }; task.assign(next)
      tasks.push({ ...task, node: node.operand, assign(value) { next.operand = value } }); continue
    }
    if (node.kind === 'binary') {
      const next = { ...node }; task.assign(next)
      tasks.push({ ...task, node: node.right, assign(value) { next.right = value } })
      tasks.push({ ...task, node: node.left, assign(value) { next.left = value } }); continue
    }
    if (node.kind === 'function') {
      const next = { ...node, arguments: [...node.arguments] }; task.assign(next)
      for (let index = node.arguments.length - 1; index >= 0; index--) tasks.push({ ...task, node: node.arguments[index]!,
        assign(value) { next.arguments[index] = value } })
      continue
    }
    task.assign(node)
  }
  return result
}
