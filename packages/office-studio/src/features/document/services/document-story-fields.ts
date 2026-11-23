import type { DocumentPageField, DocumentSectionHeaderFooter } from '../../../domain/model'
import { cloneModel } from '../../../domain/serialization'
import { sectionStoryLinkedToPrevious } from '../model/document-section-stories'
import { documentStoryDefaultStyle, documentStoryFieldCache, documentStoryFieldDisplayRuns, type DocumentStoryKey, type StoryDefaultsSource } from '../model/document-story'
import { resolveDocumentDateTimeFieldResult } from './document-date-time-field'

/** Keep cached run boundaries when a refreshed result changes within several formats. */
export function refreshDocumentStoryFieldResult(field: DocumentPageField, result: string, model?: StoryDefaultsSource): DocumentPageField {
  if (result === field.cachedResult) return field
  const cache = documentStoryFieldCache(field)
  if (!cache) return cloneModel({ ...field, cachedResult: result })
  const defaults = documentStoryDefaultStyle(model)
  const runs = documentStoryFieldDisplayRuns(field, result, defaults.text)
  const next = cloneModel({ ...field, cachedResult: result, metadata: { ...field.metadata,
    ...(field.textStyle ? { officeStoryStyleBaseline: { ...field.textStyle } } : {}), officeStoryFieldCache: { version: 1, runs } } })
  documentStoryFieldCache(next)
  return next
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function sourceLanguage(style: unknown): string | undefined {
  return isRecord(style) && typeof style.language === 'string' ? style.language : undefined
}

/** Parsed cache styles retain duke-pages' resolved w:lang in their provenance. */
function storyDateLanguage(field: DocumentPageField, fallback: string): string {
  const runs = documentStoryFieldCache(field)?.runs
  const first = runs?.find(run => run.text.length > 0)
  const runSource = first?.metadata?.source, fieldSource = field.metadata?.source
  const emptySource = isRecord(runSource) && runSource.text === ''
  if (field.metadata?.officeStoryTextFormatCleared === true || !emptySource && first?.metadata?.officeStoryTextFormatCleared === true) return fallback
  // Expanding an empty cache retains its source style. That empty source must
  // not become the language authority on the next refresh.
  return sourceLanguage(isRecord(runSource) && !emptySource ? runSource.style : undefined)
    ?? sourceLanguage(isRecord(fieldSource) ? fieldSource.textStyle : undefined)
    ?? fallback
}

/** Explicit date-time pictures are refreshed only in the selected, unlinked story. */
export function refreshDocumentStoryDateFields(section: DocumentSectionHeaderFooter, story: DocumentStoryKey, locale: string, date: Date, model?: StoryDefaultsSource): DocumentPageField[] | undefined {
  if (sectionStoryLinkedToPrevious(section, story)) throw new Error('请先取消此页眉页脚与前节的链接，再更新字段。')
  return section.pageFields?.map(field => {
    if (field.story !== story || field.field !== 'DATE' && field.field !== 'TIME') return field
    const result = resolveDocumentDateTimeFieldResult(field, storyDateLanguage(field, locale), date)
    return result === undefined ? field : refreshDocumentStoryFieldResult(field, result, model)
  })
}
