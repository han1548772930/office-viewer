import {
  DOCUMENT_EMOJI_CATEGORIES, DOCUMENT_EMOJI_SKIN_TONES, type DocumentCharacterItem, type DocumentCharacterLocale,
  type DocumentCharacterRow, type DocumentCharacterSection, type DocumentCharacterSectionPosition, type DocumentEmojiFamily,
  type DocumentEmojiLocaleData, type DocumentEmojiSkinTone, type DocumentEmojiVariant, type DocumentSymbolCategory,
} from '../model/document-character-catalog'
import { DOCUMENT_DEFAULT_RECENT_EMOJIS, DOCUMENT_EMOJI_GROUPS } from '../model/document-emojis.generated'
import { DOCUMENT_CHARACTER_PICKER_LABELS } from '../model/document-emoji-ui.generated'
import { DOCUMENT_SYMBOL_GROUPS } from '../model/document-symbols.generated'

export { DOCUMENT_EMOJI_CATEGORIES, DOCUMENT_EMOJI_SKIN_TONES, DOCUMENT_SYMBOL_GROUPS }
export const DOCUMENT_EMOJI_RECENT_LIMIT = 11
export const DOCUMENT_SYMBOL_RECENT_LIMIT = 32
export const DOCUMENT_CHARACTER_STORED_VALUE_LIMIT = 8192
export const DOCUMENT_CHARACTER_COLUMNS = 10
export const DOCUMENT_CHARACTER_ROW_HEIGHT = 32
export const DOCUMENT_CHARACTER_HEADER_HEIGHT = 24
export const DOCUMENT_CHARACTER_SECTION_SCROLL_OFFSET = 28

const emojiFamilies = DOCUMENT_EMOJI_CATEGORIES.flatMap(key => DOCUMENT_EMOJI_GROUPS[key])
const emojiVariants = emojiFamilies.flatMap(family => [family, ...(family.variants ?? [])])
const familyByEmoji = new Map<string, DocumentEmojiFamily>()
const variantByEmoji = new Map<string, DocumentEmojiVariant>()
for (const family of emojiFamilies) for (const item of [family, ...(family.variants ?? [])]) {
  familyByEmoji.set(item.emoji, family)
  variantByEmoji.set(item.emoji, item)
}
const tonePattern = /[\u{1F3FB}-\u{1F3FF}]/gu
const localeLoaders = import.meta.glob<DocumentEmojiLocaleData>('../model/emoji-locales/*.generated.ts', { import: 'default' })
const localePromises = new Map<DocumentCharacterLocale, Promise<DocumentEmojiLocaleData>>()
const localizedSearch = new WeakMap<DocumentEmojiLocaleData, ReadonlyMap<string, string>>()
const englishSearch = new Map(emojiFamilies.map(family => [family.emoji,
  [family, ...(family.variants ?? [])].flatMap(item => [item.emoji, item.title]).join(' ').toLowerCase()]))

export function documentEmojiFamilies(): readonly DocumentEmojiFamily[] { return emojiFamilies }
export function documentEmojiVariants(): readonly DocumentEmojiVariant[] { return emojiVariants }
export function documentCharacterPickerLabels(locale: DocumentCharacterLocale) { return DOCUMENT_CHARACTER_PICKER_LABELS[locale] }
export function documentEmojiGroup(category: (typeof DOCUMENT_EMOJI_CATEGORIES)[number]): readonly DocumentEmojiFamily[] { return DOCUMENT_EMOJI_GROUPS[category] }

/** Each locale is a separate build module. A rejected import can be retried; UI owns stale-result guards. */
export function loadDocumentEmojiLocale(locale: DocumentCharacterLocale): Promise<DocumentEmojiLocaleData> {
  const existing = localePromises.get(locale)
  if (existing) return existing
  const loader = localeLoaders[`../model/emoji-locales/${locale}.generated.ts`]
  if (!loader) return Promise.reject(new RangeError(`Unknown document emoji locale ${locale}`))
  const pending = loader().then(data => {
    if (data.locale !== locale) throw new TypeError('Document emoji locale module has a mismatched locale')
    return data
  }).catch(error => { localePromises.delete(locale); throw error })
  localePromises.set(locale, pending)
  return pending
}

export function documentEmojiFamilyKey(value: string): string { return familyByEmoji.get(value)?.emoji ?? value }
export function documentEmojiFamilyVariants(value: string): readonly DocumentEmojiVariant[] {
  const family = familyByEmoji.get(value)
  return family ? [family, ...(family.variants ?? [])] : []
}
export function documentEmojiVariant(value: string): DocumentEmojiVariant { return variantByEmoji.get(value) ?? { emoji: value, title: value } }
export function documentLocalizedEmojiTitle(item: DocumentEmojiVariant, locale?: DocumentEmojiLocaleData): string {
  return locale?.emojiTitles[item.emoji] ?? item.title
}
export function documentEmojiHasMixedTones(value: string): boolean {
  return familyByEmoji.get(value)?.variants?.some(item => (item.emoji.match(tonePattern) ?? []).length > 1) ?? false
}
export function documentEmojiForSkinTone(value: DocumentEmojiVariant, tone: DocumentEmojiSkinTone): DocumentEmojiVariant {
  const family = familyByEmoji.get(value.emoji)
  if (!family) return value
  if (!tone) return family
  return family.variants?.find(item => {
    const modifiers = item.emoji.match(tonePattern) ?? []
    return modifiers.length > 0 && modifiers.every(modifier => modifier === tone)
  }) ?? family
}
export function parseDocumentEmojiSkinTone(value: unknown): DocumentEmojiSkinTone { return DOCUMENT_EMOJI_SKIN_TONES.find(tone => tone === value) ?? '' }

export function searchDocumentEmojis(keyword: string, locale?: DocumentEmojiLocaleData): readonly DocumentEmojiFamily[] {
  const query = keyword.trim().toLowerCase()
  if (!query) return []
  let index: ReadonlyMap<string, string> = englishSearch
  if (locale) {
    let localized = localizedSearch.get(locale)
    if (!localized) {
      localized = new Map(emojiFamilies.map(family => [family.emoji, [englishSearch.get(family.emoji),
        ...[family, ...(family.variants ?? [])].map(item => locale.emojiTitles[item.emoji]), locale.emojiSearchIndex[family.emoji]]
        .filter(value => typeof value === 'string').join(' ').toLowerCase()]))
      localizedSearch.set(locale, localized)
    }
    index = localized
  }
  return emojiFamilies.filter(family => index.get(family.emoji)?.includes(query))
}
export function randomDocumentEmoji(random = Math.random, tone: DocumentEmojiSkinTone = ''): DocumentEmojiVariant {
  const sample = random()
  const index = Number.isFinite(sample) && sample >= 0 && sample < 1 ? Math.floor(sample * emojiFamilies.length) : 0
  return documentEmojiForSkinTone(emojiFamilies[index]!, tone)
}

function storedCharacter(value: unknown): value is string { return typeof value === 'string' && value.length > 0 && value.length <= DOCUMENT_CHARACTER_STORED_VALUE_LIMIT }
export function normalizeDocumentRecentEmojis(value: unknown): string[] {
  if (!Array.isArray(value)) return DOCUMENT_DEFAULT_RECENT_EMOJIS.map(item => item.emoji)
  const seen = new Set<string>(), values: string[] = []
  for (const raw of value) {
    const item: unknown = typeof raw === 'object' && raw !== null && 'emoji' in raw ? raw.emoji : raw
    if (!storedCharacter(item)) continue
    const family = documentEmojiFamilyKey(item)
    if (seen.has(family)) continue
    seen.add(family); values.push(item)
    if (values.length === DOCUMENT_EMOJI_RECENT_LIMIT) break
  }
  return values.length ? values : DOCUMENT_DEFAULT_RECENT_EMOJIS.map(item => item.emoji)
}
export function normalizeDocumentRecentSymbols(value: unknown): string[] {
  return Array.isArray(value) ? [...new Set(value.filter(storedCharacter))].slice(0, DOCUMENT_SYMBOL_RECENT_LIMIT) : []
}
export function promoteDocumentRecentEmoji(recents: readonly string[], value: string): string[] {
  return storedCharacter(value) ? [value, ...recents.filter(item => documentEmojiFamilyKey(item) !== documentEmojiFamilyKey(value))].slice(0, DOCUMENT_EMOJI_RECENT_LIMIT) : [...recents]
}
export function promoteDocumentRecentSymbol(recents: readonly string[], value: string): string[] {
  return storedCharacter(value) ? [value, ...recents.filter(item => item !== value)].slice(0, DOCUMENT_SYMBOL_RECENT_LIMIT) : [...recents]
}

export function documentEmojiItem(value: DocumentEmojiVariant, locale?: DocumentEmojiLocaleData, keyPrefix = 'emoji'): DocumentCharacterItem {
  const variants = documentEmojiFamilyVariants(value.emoji)
  return { key: `${keyPrefix}-${value.emoji}`, value: value.emoji, title: documentLocalizedEmojiTitle(value, locale),
    familyKey: documentEmojiFamilyKey(value.emoji), ...(variants.length > 1 ? { variants } : {}) }
}
export function documentEmojiSections(locale: DocumentCharacterLocale, recents: readonly string[], tone: DocumentEmojiSkinTone,
  query = '', localeData?: DocumentEmojiLocaleData): DocumentCharacterSection[] {
  const labels = documentCharacterPickerLabels(locale)
  if (query.trim()) return [{ key: 'search', title: labels.searchResults,
    items: searchDocumentEmojis(query, localeData).map(family => documentEmojiItem(documentEmojiForSkinTone(family, tone), localeData, 'search')) }]
  return [{ key: 'recent', title: labels.recents, items: recents.map(value => documentEmojiItem(documentEmojiVariant(value), localeData, 'recent')) },
    ...DOCUMENT_EMOJI_CATEGORIES.map(key => ({ key, title: labels[key], items: DOCUMENT_EMOJI_GROUPS[key].map(family =>
      documentEmojiItem(documentEmojiForSkinTone(family, tone), localeData, key)) }))]
}
export function documentSymbolSections(labels: Readonly<Record<DocumentSymbolCategory, string>>, recentLabel: string, recents: readonly string[], query = ''): DocumentCharacterSection[] {
  const groups: DocumentCharacterSection[] = [...(recents.length ? [{ key: 'recent', title: recentLabel, items: recents.map((value, index) => ({ key: `symbol-recent-${index}`, value, title: value })) }] : []),
    ...DOCUMENT_SYMBOL_GROUPS.map(group => ({ key: group.key, title: labels[group.key], items: group.values.map((value, index) =>
      ({ key: `symbol-${group.key}-${index}`, value, title: value, source: group.source })) }))]
  const keyword = query.trim().toLowerCase()
  return keyword ? groups.filter(group => group.key !== 'recent').map(group => ({ ...group, items: group.items.filter(item =>
    [item.value, group.title, group.key].join(' ').toLowerCase().includes(keyword)) })) : groups
}

export function createDocumentCharacterRows(sections: readonly DocumentCharacterSection[], keyPrefix: string): {
  rows: DocumentCharacterRow[]; positions: DocumentCharacterSectionPosition[]; height: number
} {
  const rows: DocumentCharacterRow[] = [], positions: DocumentCharacterSectionPosition[] = []
  let top = 0
  for (const section of sections) {
    if (!section.items.length) continue
    const rowKey = `${keyPrefix}-${section.key}-header`
    positions.push({ key: section.key, rowKey, rowIndex: rows.length, top })
    rows.push({ key: rowKey, type: 'header', sectionKey: section.key, title: section.title, top, height: DOCUMENT_CHARACTER_HEADER_HEIGHT })
    top += DOCUMENT_CHARACTER_HEADER_HEIGHT
    for (let start = 0; start < section.items.length; start += DOCUMENT_CHARACTER_COLUMNS) {
      rows.push({ key: `${keyPrefix}-${section.key}-${start}`, type: 'characters', sectionKey: section.key,
        items: section.items.slice(start, start + DOCUMENT_CHARACTER_COLUMNS), top, height: DOCUMENT_CHARACTER_ROW_HEIGHT })
      top += DOCUMENT_CHARACTER_ROW_HEIGHT
    }
  }
  return { rows, positions, height: top }
}
export function activeDocumentCharacterSection(positions: readonly DocumentCharacterSectionPosition[], scrollTop: number): string | undefined {
  const top = Math.max(0, Number.isFinite(scrollTop) ? scrollTop : 0) + DOCUMENT_CHARACTER_SECTION_SCROLL_OFFSET
  let key = positions[0]?.key
  for (const position of positions) if (position.top <= top) key = position.key
  return key
}

/** Arrow navigation follows actual rows, including short rows and section headers. */
export function documentCharacterNeighbor(rows: readonly DocumentCharacterRow[], key: string, direction: 'left' | 'right' | 'up' | 'down' | 'first' | 'last', rtl = false): { rowKey: string; item: DocumentCharacterItem } | undefined {
  const characterRows = rows.filter((row): row is Extract<DocumentCharacterRow, { type: 'characters' }> => row.type === 'characters')
  const rowIndex = characterRows.findIndex(row => row.items.some(item => item.key === key))
  if (rowIndex < 0) return undefined
  const row = characterRows[rowIndex]!, index = row.items.findIndex(item => item.key === key)
  if (direction === 'first' || direction === 'last') {
    const target = direction === 'first' ? characterRows[0] : characterRows.at(-1)
    const item = direction === 'first' ? target?.items[0] : target?.items.at(-1)
    return target && item ? { rowKey: target.key, item } : undefined
  }
  if (direction === 'up' || direction === 'down') {
    const target = characterRows[rowIndex + (direction === 'up' ? -1 : 1)]
    const item = target?.items[Math.min(index, target.items.length - 1)]
    return target && item ? { rowKey: target.key, item } : undefined
  }
  const delta = (direction === 'left' ? -1 : 1) * (rtl ? -1 : 1)
  const target = row.items[index + delta]
  if (target) return { rowKey: row.key, item: target }
  const nextRow = characterRows[rowIndex + delta]
  const item = delta < 0 ? nextRow?.items.at(-1) : nextRow?.items[0]
  return nextRow && item ? { rowKey: nextRow.key, item } : undefined
}
