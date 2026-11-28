export const DOCUMENT_EMOJI_CATEGORIES = ['people', 'nature', 'foods', 'activity', 'places', 'objects', 'symbols'] as const
export type DocumentEmojiCategory = (typeof DOCUMENT_EMOJI_CATEGORIES)[number]
export const DOCUMENT_EMOJI_SKIN_TONES = ['', '🏻', '🏼', '🏽', '🏾', '🏿'] as const
export type DocumentEmojiSkinTone = (typeof DOCUMENT_EMOJI_SKIN_TONES)[number]
export type DocumentCharacterMode = 'symbol' | 'emoji'

export interface DocumentEmojiVariant {
  readonly emoji: string
  readonly title: string
}
export interface DocumentEmojiFamily extends DocumentEmojiVariant {
  readonly variants?: readonly DocumentEmojiVariant[]
}
export const DOCUMENT_CHARACTER_LOCALES = ['ar-SA', 'ca-ES', 'de-DE', 'en-US', 'es-ES', 'fa-IR', 'fr-FR', 'id-ID', 'it-IT', 'ja-JP',
  'ko-KR', 'pl-PL', 'pt-BR', 'ru-RU', 'sk-SK', 'vi-VN', 'zh-CN', 'zh-HK', 'zh-TW'] as const
export type DocumentCharacterLocale = (typeof DOCUMENT_CHARACTER_LOCALES)[number]
export interface DocumentEmojiLocaleData {
  readonly locale: DocumentCharacterLocale
  readonly emojiTitles: Readonly<Record<string, string>>
  readonly emojiSearchIndex: Readonly<Record<string, string>>
}
export interface DocumentCharacterPickerLabels {
  readonly search: string
  readonly random: string
  readonly recents: string
  readonly people: string
  readonly nature: string
  readonly foods: string
  readonly activity: string
  readonly places: string
  readonly objects: string
  readonly symbols: string
  readonly searchResults: string
  readonly noResults: string
  readonly mathematics: string
  readonly greek: string
  readonly common: string
  readonly more: string
}
export type DocumentSymbolCategory = 'mathematics' | 'greek' | 'common' | 'arrows' | 'punctuation' | 'currency' | 'additional'
export interface DocumentSymbolGroup {
  readonly key: DocumentSymbolCategory
  readonly source: 'univer' | 'extension'
  readonly values: readonly string[]
}

/** Layout rows contain complete Unicode strings; offsets always refer to rows, never UTF-16 code units. */
export interface DocumentCharacterItem {
  readonly key: string
  readonly value: string
  readonly title: string
  readonly familyKey?: string
  readonly variants?: readonly DocumentEmojiVariant[]
  readonly source?: 'univer' | 'extension' | 'custom'
}
export interface DocumentCharacterSection {
  readonly key: string
  readonly title: string
  readonly items: readonly DocumentCharacterItem[]
}
export type DocumentCharacterRow =
  | { readonly key: string; readonly type: 'header'; readonly sectionKey: string; readonly title: string; readonly top: number; readonly height: number }
  | { readonly key: string; readonly type: 'characters'; readonly sectionKey: string; readonly items: readonly DocumentCharacterItem[]; readonly top: number; readonly height: number }
export interface DocumentCharacterSectionPosition {
  readonly key: string
  readonly rowKey: string
  readonly rowIndex: number
  readonly top: number
}
