import type { CellScalar } from '../../domain/model.ts'

export interface SheetFilterCalendarDay { year: number; month: number; day: number }
export interface SheetFilterValueCandidate {
  key: string
  label: string
  value: CellScalar
  count: number
  calendar?: SheetFilterCalendarDay
}
export interface SheetFilterValueTreeNode {
  key: string
  title: string
  count: number
  valueKeys: string[]
  children?: SheetFilterValueTreeNode[]
}

/** Dates are view groups. Every leaf retains the original typed values, including distinct times on the same day. */
export function buildSheetFilterValueTree(values: readonly SheetFilterValueCandidate[], locale = 'zh-CN'): SheetFilterValueTreeNode[] {
  const roots: SheetFilterValueTreeNode[] = [], nodes = new Map<string, SheetFilterValueTreeNode>()
  const monthTitle = new Intl.DateTimeFormat(locale, { month: 'long', timeZone: 'UTC' })
  const node = (key: string, title: string, parent?: SheetFilterValueTreeNode): SheetFilterValueTreeNode => {
    const existing = nodes.get(key)
    if (existing) return existing
    const created: SheetFilterValueTreeNode = { key, title, count: 0, valueKeys: [] }
    nodes.set(key, created)
    if (parent) (parent.children ??= []).push(created)
    else roots.push(created)
    return created
  }
  const add = (target: SheetFilterValueTreeNode, item: SheetFilterValueCandidate) => {
    target.count += item.count
    if (!target.valueKeys.includes(item.key)) target.valueKeys.push(item.key)
  }
  for (const item of values) {
    const date = item.calendar
    if (!date) { add(node(`value:${item.key}`, item.label), item); continue }
    const year = node(`date:${date.year}`, String(date.year))
    const month = node(`date:${date.year}-${date.month}`, monthTitle.format(new Date(Date.UTC(2000, date.month - 1, 1))), year)
    const day = node(`date:${date.year}-${date.month}-${date.day}`, String(date.day), month)
    add(year, item); add(month, item); add(day, item)
  }
  const compare = new Intl.Collator(locale, { numeric: true, sensitivity: 'base' })
  roots.sort((left, right) => Number(!!right.children) - Number(!!left.children) || compare.compare(left.title, right.title))
  for (const year of roots) if (year.children) {
    year.children.sort((left, right) => Number(left.key.split('-')[1]) - Number(right.key.split('-')[1]))
    for (const month of year.children) month.children?.sort((left, right) => Number(left.key.split('-')[2]) - Number(right.key.split('-')[2]))
  }
  return roots
}

/** Source search is case-insensitive OR across whitespace-separated keywords; a matched parent keeps its descendants. */
export function searchSheetFilterValueTree(nodes: readonly SheetFilterValueTreeNode[], search: string,
  values: readonly SheetFilterValueCandidate[]): SheetFilterValueTreeNode[] {
  const keywords = search.toLocaleLowerCase().split(/\s+/u).filter(Boolean)
  if (!keywords.length) return [...nodes]
  const labels = new Map(values.map(item => [item.key, item.label.toLocaleLowerCase()]))
  const visit = (items: readonly SheetFilterValueTreeNode[]): SheetFilterValueTreeNode[] => items.flatMap(item => {
    if (keywords.some(keyword => item.title.toLocaleLowerCase().includes(keyword)
      || !item.children && item.valueKeys.some(key => labels.get(key)?.includes(keyword)))) return [item]
    const children = item.children && visit(item.children)
    return children?.length ? [{ ...item, children, count: children.reduce((sum, child) => sum + child.count, 0),
      valueKeys: [...new Set(children.flatMap(child => child.valueKeys))] }] : []
  })
  return visit(nodes)
}

export function sheetFilterValueTreeSelection(nodes: readonly SheetFilterValueTreeNode[], selectedKeys: readonly string[]): {
  checked: string[]; halfChecked: string[]
} {
  const selected = new Set(selectedKeys), checked: string[] = [], halfChecked: string[] = []
  const visit = (items: readonly SheetFilterValueTreeNode[]) => { for (const item of items) {
    const count = item.valueKeys.filter(key => selected.has(key)).length
    if (count && count === item.valueKeys.length) checked.push(item.key)
    else if (count) halfChecked.push(item.key)
    if (item.children) visit(item.children)
  } }
  visit(nodes)
  return { checked, halfChecked }
}

export function toggleSheetFilterValueTreeNode(nodes: readonly SheetFilterValueTreeNode[], selectedKeys: readonly string[], key: string): string[] {
  const find = (items: readonly SheetFilterValueTreeNode[]): SheetFilterValueTreeNode | undefined => {
    for (const item of items) { if (item.key === key) return item; const nested = item.children && find(item.children); if (nested) return nested }
    return undefined
  }
  const node = find(nodes)
  if (!node) throw new RangeError('筛选值树节点不存在。')
  const selected = new Set(selectedKeys), remove = node.valueKeys.every(valueKey => selected.has(valueKey))
  for (const valueKey of node.valueKeys) { if (remove) selected.delete(valueKey); else selected.add(valueKey) }
  return [...selected]
}
