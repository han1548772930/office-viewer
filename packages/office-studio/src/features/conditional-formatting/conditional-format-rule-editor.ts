import type { SheetConditionalFormatRule, SheetConditionalFormatValueConfig, SheetConditionalFormatVisualConfig } from '../../domain/model.ts'
import { cloneModel } from '../../domain/serialization.ts'

const editorSubtypes: Partial<Record<SheetConditionalFormatRule['type'], SheetConditionalFormatRule['subType']>> = {
  cellIs: 'number', highlightCell: 'number', expression: 'formula', top10: 'rank', average: 'average',
  containsText: 'text', notContainsText: 'text', beginsWith: 'text', endsWith: 'text', timePeriod: 'timePeriod',
  duplicateValues: 'duplicateValues', uniqueValues: 'uniqueValues',
}

/** Preserve the exact imported formula when the editor has not changed its expression. */
export function conditionalFormatEditorFormula(value: string, previous?: string): string {
  const normalized = value.trim().replace(/^=/, '').trim()
  return previous !== undefined && previous.trim().replace(/^=/, '').trim() === normalized ? previous : normalized
}

/** Office stores the matching literal separately from its generated predicate. */
export function conditionalFormatEditorText(value: string, previous?: SheetConditionalFormatRule): Pick<SheetConditionalFormatRule, 'formula1' | 'value'> {
  return previous && String(previous.value ?? previous.formula1) === value
    ? { formula1: previous.formula1, value: previous.value }
    : { formula1: value, value }
}

/** Stable ties place a new rule first without renumbering unedited rules. */
export function prependConditionalFormatRuleDraft(rules: readonly SheetConditionalFormatRule[], rule: SheetConditionalFormatRule): SheetConditionalFormatRule[] {
  const priority = rules.reduce((minimum, item) => Math.min(minimum, item.priority), rules[0]?.priority ?? 1)
  return cloneModel([{ ...rule, priority }, ...rules])
}

export function removeConditionalFormatRuleDraft(rules: readonly SheetConditionalFormatRule[], id: string): SheetConditionalFormatRule[] {
  return cloneModel(rules.filter(rule => rule.id !== id))
}

/** An adjacent move exchanges only the two affected priority slots; equal priorities retain the new array order. */
export function moveConditionalFormatRuleDraft(rules: readonly SheetConditionalFormatRule[], id: string, offset: -1 | 1): SheetConditionalFormatRule[] {
  const result = cloneModel([...rules])
  const index = result.findIndex(rule => rule.id === id), target = index + offset
  if (index < 0 || target < 0 || target >= result.length) return result
  const rule = result[index]!, neighbour = result[target]!
  const priority = rule.priority
  rule.priority = neighbour.priority
  neighbour.priority = priority
  result[index] = neighbour
  result[target] = rule
  return result
}

function mergeThreshold(previous: SheetConditionalFormatValueConfig | undefined, draft: SheetConditionalFormatValueConfig): SheetConditionalFormatValueConfig {
  return { ...previous, ...draft, value: draft.value }
}

function mergeVisualConfig(previous: SheetConditionalFormatVisualConfig | undefined, draft: SheetConditionalFormatVisualConfig | undefined): SheetConditionalFormatVisualConfig | undefined {
  if (!draft) return previous
  const result = { ...previous, ...draft }
  if (draft.min) result.min = mergeThreshold(previous?.min, draft.min)
  if (draft.max) result.max = mergeThreshold(previous?.max, draft.max)
  if (draft.thresholds) result.thresholds = draft.thresholds.map((threshold, index, entries) => {
    // A two/three-stop color scale keeps extension fields attached to the same endpoint.
    const before = index === entries.length - 1 ? previous?.thresholds?.at(-1)
      : index === 0 ? previous?.thresholds?.[0] : previous?.thresholds?.length === entries.length ? previous?.thresholds?.[index] : undefined
    return { ...before, ...threshold, value: mergeThreshold(before?.value, threshold.value) }
  })
  if (draft.icons) result.icons = draft.icons.map((icon, index) => ({
    ...previous?.icons?.[index], ...icon, value: mergeThreshold(previous?.icons?.[index]?.value, icon.value),
  }))
  return result
}

/** Replace edited fields while retaining unedited JSON extensions and isolating the resulting snapshot. */
export function mergeConditionalFormatRuleDraft(previous: SheetConditionalFormatRule | undefined, draft: SheetConditionalFormatRule): SheetConditionalFormatRule {
  const preserveType = previous && (previous.type === draft.type || previous.type === 'highlightCell')
    && (previous.subType ?? editorSubtypes[previous.type]) === (draft.subType ?? editorSubtypes[draft.type])
  return cloneModel({
    ...previous, ...draft,
    ...(preserveType ? { type: previous.type, subType: previous.subType } : {}),
    style: { ...previous?.style, ...draft.style },
    config: mergeVisualConfig(previous?.config, draft.config),
    metadata: previous?.metadata || draft.metadata ? { ...previous?.metadata, ...draft.metadata } : undefined,
  })
}
