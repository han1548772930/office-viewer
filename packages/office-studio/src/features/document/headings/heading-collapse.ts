import type { DocumentBlock, DocumentModel } from '../../../domain/model'
import { blockAtPoint, collapsedSelection, runText, sameTarget, selectionCollapsed, tableForBlock, textTargets, type DocumentPoint, type DocumentSelection, type TextTarget } from '../model/document-model'
import { documentBlockChildContainers } from '../model/document-structural-blocks'
import { assertDocumentPermission, permissionTargetsForSelection, type DocumentPermissionContext } from '../permissions/document-permissions'
import { assertDocumentTextSelectionEditable } from '../services/document-text-formatting'

export const SET_DOCUMENT_HEADING_COLLAPSED_COMMAND = 'doc.command.set-heading-collapsed'
export interface DocumentHeadingSection {
  paragraphId: string
  target: TextTarget
  block: DocumentBlock
  level: number
  collapsed: boolean
  siblings: readonly DocumentBlock[]
  index: number
  endIndex: number
}
export interface DocumentHeadingVisibility {
  sections: DocumentHeadingSection[]
  sectionById: Map<string, DocumentHeadingSection>
  identityCounts: Map<string, number>
  /** Hidden object/paragraph ID -> its outermost visible collapsed heading. */
  hiddenOwners: Map<string, string>
  blocks: DocumentBlock[]
}

export function supportsDocumentHeadingCollapse(model: DocumentModel): boolean {
  const metadata = model.metadata ?? {}
  const style = metadata.documentStyle
  const source = metadata.univerDocumentSource
  const snapshot = source && typeof source === 'object' && !Array.isArray(source) ? Reflect.get(source, 'snapshot') : undefined
  const sourceStyle = snapshot && typeof snapshot === 'object' && !Array.isArray(snapshot) ? Reflect.get(snapshot, 'documentStyle') : undefined
  const flavor = metadata.documentFlavor ?? (style && typeof style === 'object' ? Reflect.get(style, 'documentFlavor') : undefined)
    ?? (sourceStyle && typeof sourceStyle === 'object' ? Reflect.get(sourceStyle, 'documentFlavor') : undefined)
  return flavor === 2 || typeof flavor === 'string' && flavor.toLowerCase() === 'modern'
}

/** Derived structural sharing: hiding content never deletes authored nodes,
 * clears nested fold flags, or clones every text run/media resource. */
export function documentHeadingVisibility(model: DocumentModel): DocumentHeadingVisibility {
  const result: DocumentHeadingVisibility = { sections: [], sectionById: new Map(), identityCounts: new Map(), hiddenOwners: new Map(), blocks: model.blocks }
  if (!supportsDocumentHeadingCollapse(model)) return result
  const targets = textTargets(model), byId = new Map(targets.map(target => [target.cellBlockId ?? target.blockId, target]))
  const scopes = new Map<readonly DocumentBlock[], Map<number, DocumentHeadingSection>>()
  const collect = (blocks: readonly DocumentBlock[], depth = 0) => {
    if (depth > 64) throw new RangeError('标题折叠结构超过 64 层。')
    const open: DocumentHeadingSection[] = [], indexed = new Map<number, DocumentHeadingSection>()
    scopes.set(blocks, indexed)
    for (const [index, block] of blocks.entries()) {
      result.identityCounts.set(block.id, (result.identityCounts.get(block.id) ?? 0) + 1)
      const target = byId.get(block.id)
      if (block.type === 'heading' && Number.isInteger(block.level) && block.level >= 1 && block.level <= 5 && target) {
        while (open.length && open.at(-1)!.level >= block.level) open.pop()!.endIndex = index
        const section: DocumentHeadingSection = { paragraphId: block.id, target, block, level: block.level,
          collapsed: block.metadata.headingCollapsed === true, siblings: blocks, index, endIndex: blocks.length }
        result.sections.push(section); result.sectionById.set(block.id, section); indexed.set(index, section); open.push(section)
      }
      for (const children of documentBlockChildContainers(block)) collect(children, depth + 1)
    }
  }
  collect(model.blocks)
  const hide = (block: DocumentBlock, owner: string, depth = 0) => {
    if (depth > 64) throw new RangeError('标题折叠结构超过 64 层。')
    result.hiddenOwners.set(block.id, owner)
    for (const children of documentBlockChildContainers(block)) for (const child of children) hide(child, owner, depth + 1)
  }
  const project = (blocks: DocumentBlock[]): DocumentBlock[] => {
    const output: DocumentBlock[] = [], indexed = scopes.get(blocks)
    let hidden: DocumentHeadingSection | undefined, changed = false
    for (const [index, block] of blocks.entries()) {
      if (hidden && index >= hidden.endIndex) hidden = undefined
      if (hidden) { hide(block, hidden.paragraphId); changed = true; continue }
      let projected = block
      if (block.children) {
        const children = project(block.children)
        if (children !== block.children) projected = { ...projected, children }
      }
      if (block.columns) {
        const columns = block.columns.columns.map(column => {
          const children = project(column.blocks)
          return children === column.blocks ? column : { ...column, blocks: children }
        })
        if (columns.some((column, i) => column !== block.columns!.columns[i])) projected = { ...projected, columns: { ...block.columns, columns } }
      }
      const table = block.type === 'table' ? tableForBlock(block) : undefined
      if (table) {
        const rows = table.rows.map(row => {
          const cells = row.cells.map(cell => {
            const children = project(cell.blocks), source = cell.style.nativeContinuationContent
            const continuations = source && Object.fromEntries(Object.entries(source).map(([key, content]) => {
              const contentBlocks = project(content.blocks)
              return [key, contentBlocks === content.blocks ? content : { ...content, blocks: contentBlocks }]
            }))
            const continued = source && continuations && Object.keys(source).some(key => source[key] !== continuations[key])
            return children === cell.blocks && !continued ? cell : { ...cell, blocks: children,
              ...(continued ? { style: { ...cell.style, nativeContinuationContent: continuations } } : {}) }
          })
          return cells.every((cell, i) => cell === row.cells[i]) ? row : { ...row, cells }
        })
        if (rows.some((row, i) => row !== table.rows[i])) projected = { ...projected, metadata: { ...projected.metadata, table: { ...table, rows } } }
      }
      changed ||= projected !== block; output.push(projected)
      const heading = indexed?.get(index)
      if (heading?.collapsed && heading.endIndex > index + 1) hidden = heading
    }
    return changed ? output : blocks
  }
  result.blocks = project(model.blocks)
  return result
}

export function visibleDocumentTextTargets(model: DocumentModel, visibility = documentHeadingVisibility(model)): TextTarget[] {
  return textTargets(model).filter(target => !visibility.hiddenOwners.has(target.cellBlockId ?? target.blockId) && !visibility.hiddenOwners.has(target.blockId))
}

/** View/API selections cannot retain a caret inside invisible authored content. */
export function normalizeDocumentHeadingSelection(model: DocumentModel, selection: DocumentSelection,
  visibility = documentHeadingVisibility(model)): DocumentSelection {
  if (!visibility.hiddenOwners.size) return selection
  const point = (value: DocumentPoint): DocumentPoint => {
    const id = visibility.hiddenOwners.get(value.cellBlockId ?? value.blockId) ?? visibility.hiddenOwners.get(value.blockId)
    const section = id && visibility.sectionById.get(id)
    return section ? { ...section.target, offset: runText(section.block.runs).length } : value
  }
  const anchor = point(selection.anchor), focus = point(selection.focus)
  return anchor === selection.anchor && focus === selection.focus ? selection : { anchor, focus }
}

function assertHeading(model: DocumentModel, paragraphId: string, collapsed: boolean, context: DocumentPermissionContext,
  visibility: DocumentHeadingVisibility): DocumentHeadingSection {
  if (typeof paragraphId !== 'string' || !paragraphId.trim() || typeof collapsed !== 'boolean') throw new TypeError('标题折叠参数无效。')
  if (!supportsDocumentHeadingCollapse(model)) throw new RangeError('标题折叠适用于现代文档。')
  const source = model.metadata?.univerDocumentSource
  const snapshot = source && typeof source === 'object' && !Array.isArray(source) ? Reflect.get(source, 'snapshot') : undefined
  if (model.metadata?.disabled === true || snapshot && typeof snapshot === 'object' && Reflect.get(snapshot, 'disabled') === true) throw new RangeError('文档当前不可编辑。')
  const section = visibility.sectionById.get(paragraphId)
  if (visibility.identityCounts.get(paragraphId) !== 1 || !section) throw new RangeError('标题不存在或段落标识不唯一。')
  if (visibility.hiddenOwners.has(paragraphId)) throw new RangeError('请先展开上层标题。')
  if (collapsed && section.endIndex <= section.index + 1) throw new RangeError('该标题后没有可折叠内容。')
  const selection = { anchor: { ...section.target, offset: 0 }, focus: { ...section.target, offset: runText(section.block.runs).length } }
  assertDocumentPermission(model.permissionRules ?? [], context, 'edit', permissionTargetsForSelection(model, selection))
  assertDocumentTextSelectionEditable(model, selection)
  return section
}

export function documentHeadingCollapseState(model: DocumentModel, paragraphId: string, collapsed: boolean,
  context: DocumentPermissionContext, visibility = documentHeadingVisibility(model)): { available: boolean; reason?: string } {
  try { assertHeading(model, paragraphId, collapsed, context, visibility); return { available: true } }
  catch (error) { return { available: false, reason: error instanceof Error ? error.message : String(error) } }
}

export function setDocumentHeadingCollapsed(model: DocumentModel, selection: DocumentSelection, paragraphId: string, collapsed: boolean,
  context: DocumentPermissionContext): DocumentSelection {
  const visibility = documentHeadingVisibility(model), section = assertHeading(model, paragraphId, collapsed, context, visibility)
  if (section.collapsed === collapsed) return normalizeDocumentHeadingSelection(model, selection, visibility)
  section.block.metadata.headingCollapsed = collapsed
  if (collapsed) {
    const hidden = new Set<string>()
    const visit = (block: DocumentBlock) => { hidden.add(block.id); for (const children of documentBlockChildContainers(block)) children.forEach(visit) }
    section.siblings.slice(section.index + 1, section.endIndex).forEach(visit)
    const targets = textTargets(model), first = targets.findIndex(target => sameTarget(target, selection.anchor)), last = targets.findIndex(target => sameTarget(target, selection.focus))
    if (first >= 0 && last >= 0 && targets.slice(Math.min(first, last), Math.max(first, last) + 1)
      .some(target => hidden.has(target.cellBlockId ?? target.blockId) || hidden.has(target.blockId))) {
      return collapsedSelection({ ...section.target, offset: runText(section.block.runs).length })
    }
  }
  return normalizeDocumentHeadingSelection(model, selection)
}

/** Enter opens the edited heading before splitting. Boundary Backspace/Delete
 * opens hidden content instead of deleting an invisible neighboring paragraph. */
export function documentHeadingExpansionForEdit(model: DocumentModel, selection: DocumentSelection,
  command: { id: string; text?: string; direction?: number; host?: unknown }): string | undefined {
  if (command.id !== 'doc.command.delete' && (command.id !== 'doc.command.insert-text' || !/[\r\n]/.test(command.text ?? ''))) return undefined
  if (command.host || !selectionCollapsed(selection) || !supportsDocumentHeadingCollapse(model)) return undefined
  const visibility = documentHeadingVisibility(model), block = blockAtPoint(model, selection.focus)
  if (!block) return undefined
  const heading = visibility.sectionById.get(block.id)
  if (heading?.collapsed && (command.id === 'doc.command.insert-text' && /[\r\n]/.test(command.text ?? '')
    || command.id === 'doc.command.delete' && command.direction === 1 && selection.focus.offset === runText(block.runs).length)) return heading.paragraphId
  if (command.id === 'doc.command.delete' && command.direction === -1 && selection.focus.offset === 0) {
    const targets = textTargets(model), index = targets.findIndex(target => sameTarget(target, selection.focus)), previous = targets[index - 1]
    if (previous) return visibility.hiddenOwners.get(previous.cellBlockId ?? previous.blockId) ?? visibility.hiddenOwners.get(previous.blockId)
  }
  return undefined
}
