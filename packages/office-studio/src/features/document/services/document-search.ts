import type { DocumentModel } from '../../../domain/model'
import { executeDocumentCommand } from '../commands/document-commands.ts'
import type { DocumentPermissionContext } from '../permissions/document-permissions'
import { orderedSelection, runText, runsForTarget, sameTarget, targetKey, textTargets, type DocumentPoint, type DocumentSelection, type TextTarget } from '../model/document-model.ts'
import { documentTableSelectedTextTargets } from '../model/document-table-selection'

export interface DocumentSearchOptions {
  caseSensitive?: boolean
  wholeWord?: boolean
  locale?: string
  selection?: DocumentSelection
}

export interface DocumentSearchMatch {
  target: TextTarget
  start: number
  end: number
}

export function findDocumentMatches(
  model: DocumentModel,
  query: string,
  options: DocumentSearchOptions = {},
): DocumentSearchMatch[] {
  if (!query) return []
  const expression = new RegExp(escapeRegExp(query), options.caseSensitive ? 'gu' : 'giu')
  const segmenter = options.wholeWord ? new Intl.Segmenter(options.locale, { granularity: 'word' }) : undefined
  const matches: DocumentSearchMatch[] = []
  const targets = textTargets(model)
  let selectedTableTargets: Set<string> | undefined
  if (options.selection?.tableSelection) {
    try { selectedTableTargets = new Set(documentTableSelectedTextTargets(model, options.selection.tableSelection).map(targetKey)) }
    catch (error) { if (error instanceof RangeError || error instanceof TypeError) return []; throw error }
  }
  const scope = options.selection && !options.selection.tableSelection ? orderedSelection(model, options.selection) : undefined
  const firstTarget = scope ? targets.findIndex((target) => sameTarget(target, scope.start)) : 0
  const lastTarget = scope ? targets.findIndex((target) => sameTarget(target, scope.end)) : targets.length - 1
  for (const [targetIndex, target] of targets.entries()) {
    if (targetIndex < firstTarget || targetIndex > lastTarget) continue
    if (selectedTableTargets && !selectedTableTargets.has(targetKey(target))) continue
    const text = runText(runsForTarget(model, target))
    const from = scope && targetIndex === firstTarget ? scope.start.offset : 0
    const to = scope && targetIndex === lastTarget ? scope.end.offset : text.length
    const wordStarts = new Set<number>()
    const wordEnds = new Set<number>()
    if (segmenter) {
      for (const segment of segmenter.segment(text)) {
        if (!segment.isWordLike) continue
        wordStarts.add(segment.index)
        wordEnds.add(segment.index + segment.segment.length)
      }
    }
    expression.lastIndex = from
    for (const match of text.matchAll(expression)) {
      const end = match.index + match[0].length
      if (end > to) break
      if (segmenter && (!wordStarts.has(match.index) || !wordEnds.has(end))) continue
      matches.push({ target: { ...target }, start: match.index, end })
    }
  }
  return matches
}

export function pointForMatch(match: DocumentSearchMatch): { anchor: DocumentPoint; focus: DocumentPoint } {
  return {
    anchor: { ...match.target, offset: match.start },
    focus: { ...match.target, offset: match.end },
  }
}

export function replaceDocumentMatches(
  model: DocumentModel,
  matches: readonly DocumentSearchMatch[],
  replacement: string,
  permissionContext?: DocumentPermissionContext,
): DocumentSelection | undefined {
  const targets = textTargets(model)
  const targetOrder = new Map(targets.map((target, index) => [targetKey(target), index]))
  const ordered = [...matches].sort((a, b) => {
    const targetDelta = (targetOrder.get(targetKey(b.target)) ?? -1) - (targetOrder.get(targetKey(a.target)) ?? -1)
    return targetDelta || b.start - a.start
  })
  let selection: DocumentSelection | undefined
  for (const match of ordered) {
    const target = targets.find((candidate) => sameTarget(candidate, match.target))
    if (!target) continue
    selection = executeDocumentCommand(model, pointForMatch({ ...match, target }), {
      id: 'doc.command.insert-text',
      text: replacement,
    }, permissionContext)
  }
  return selection
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
