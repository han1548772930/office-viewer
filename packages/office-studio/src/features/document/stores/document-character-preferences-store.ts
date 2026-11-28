import { defineStore } from 'pinia'
import { computed, ref, shallowRef } from 'vue'
import { DOCUMENT_EMOJI_SKIN_TONES, type DocumentCharacterMode, type DocumentEmojiSkinTone } from '../model/document-character-catalog'
import { DOCUMENT_CHARACTER_STORED_VALUE_LIMIT, normalizeDocumentRecentEmojis, normalizeDocumentRecentSymbols, parseDocumentEmojiSkinTone, promoteDocumentRecentEmoji, promoteDocumentRecentSymbol } from '../services/document-special-characters'

export const DOCUMENT_CHARACTER_PREFERENCES_KEY = 'office-studio.document.character-preferences.v1'
export const DOCUMENT_CHARACTER_PREFERENCES_VALUE_BUDGET = 32768
export const DOCUMENT_CHARACTER_PREFERENCES_STORAGE_BUDGET = 131072
export interface DocumentCharacterPreferencesSnapshot {
  readonly version: 1
  readonly recentEmojis: readonly string[]
  readonly recentSymbols: readonly string[]
  readonly skinTone: DocumentEmojiSkinTone
  readonly activeSymbol: string
  readonly activeEmoji: string
}
export interface DocumentCharacterPreferenceStorage {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
}
export function normalizeDocumentCharacterPreferences(value: unknown): DocumentCharacterPreferencesSnapshot {
  const record = typeof value === 'object' && value !== null && !Array.isArray(value) ? value as Record<string, unknown> : {}
  const activeSymbol = typeof record.activeSymbol === 'string' && record.activeSymbol.length <= DOCUMENT_CHARACTER_STORED_VALUE_LIMIT ? record.activeSymbol : ''
  const activeEmoji = typeof record.activeEmoji === 'string' && record.activeEmoji.length <= DOCUMENT_CHARACTER_STORED_VALUE_LIMIT ? record.activeEmoji : ''
  const recentEmojis: string[] = [], recentSymbols: string[] = []
  const next: DocumentCharacterPreferencesSnapshot = { version: 1, recentEmojis, recentSymbols,
    skinTone: parseDocumentEmojiSkinTone(record.skinTone), activeSymbol, activeEmoji }
  let valueUnits = activeSymbol.length + activeEmoji.length, storageUnits = JSON.stringify(next).length
  const append = (target: string[], item: string): void => {
    const encodedUnits = JSON.stringify(item).length + (target.length ? 1 : 0)
    if (valueUnits + item.length > DOCUMENT_CHARACTER_PREFERENCES_VALUE_BUDGET
      || storageUnits + encodedUnits > DOCUMENT_CHARACTER_PREFERENCES_STORAGE_BUDGET) return
    target.push(item); valueUnits += item.length; storageUnits += encodedUnits
  }
  const emojis = normalizeDocumentRecentEmojis(record.recentEmojis), symbols = normalizeDocumentRecentSymbols(record.recentSymbols)
  // Keep complete values and newer entries from both modes. JSON escaping is
  // budgeted too, so every valid write can be read by the same storage boundary.
  for (let index = 0; index < Math.max(emojis.length, symbols.length); index++) {
    if (emojis[index]) append(recentEmojis, emojis[index]!)
    if (symbols[index]) append(recentSymbols, symbols[index]!)
  }
  return next
}

/** Application preference state is shared by both pickers and is never copied into a document snapshot or content history. */
export const useDocumentCharacterPreferencesStore = defineStore('document-character-preferences', () => {
  const preferences = shallowRef<DocumentCharacterPreferencesSnapshot>(normalizeDocumentCharacterPreferences(undefined))
  const loaded = ref(false), storageError = ref<string>()
  let storage: DocumentCharacterPreferenceStorage | undefined
  const recentEmojis = computed(() => preferences.value.recentEmojis), recentSymbols = computed(() => preferences.value.recentSymbols)
  const skinTone = computed(() => preferences.value.skinTone), activeSymbol = computed(() => preferences.value.activeSymbol), activeEmoji = computed(() => preferences.value.activeEmoji)

  function ensureLoaded(target?: DocumentCharacterPreferenceStorage): void {
    if (loaded.value) return
    loaded.value = true
    try {
      storage = target ?? (typeof window !== 'undefined' ? window.localStorage : undefined)
      const saved = storage?.getItem(DOCUMENT_CHARACTER_PREFERENCES_KEY)
      if (!saved) return
      if (saved.length > DOCUMENT_CHARACTER_PREFERENCES_STORAGE_BUDGET) throw new RangeError('Character preference data exceeds its storage budget')
      const value: unknown = JSON.parse(saved)
      if (!value || typeof value !== 'object' || !('version' in value) || value.version !== 1) throw new TypeError('Unsupported character preference snapshot version')
      preferences.value = normalizeDocumentCharacterPreferences(value)
    } catch (error) { storageError.value = error instanceof Error ? error.message : String(error) }
  }
  function save(next: DocumentCharacterPreferencesSnapshot): void {
    ensureLoaded()
    preferences.value = normalizeDocumentCharacterPreferences(next)
    try { if (storage) { storage.setItem(DOCUMENT_CHARACTER_PREFERENCES_KEY, JSON.stringify(preferences.value)); storageError.value = undefined } }
    catch (error) { storageError.value = error instanceof Error ? error.message : String(error) }
  }
  function select(mode: DocumentCharacterMode, value: string): boolean {
    if (typeof value !== 'string' || !value.length) return false
    ensureLoaded()
    const current = preferences.value
    // Arbitrarily long custom text still inserts intact; preferences keep a bounded recent-value budget.
    save(mode === 'emoji' ? { ...current, activeEmoji: value.length <= DOCUMENT_CHARACTER_STORED_VALUE_LIMIT ? value : '', recentEmojis: promoteDocumentRecentEmoji(current.recentEmojis, value) }
      : { ...current, activeSymbol: value.length <= DOCUMENT_CHARACTER_STORED_VALUE_LIMIT ? value : '', recentSymbols: promoteDocumentRecentSymbol(current.recentSymbols, value) })
    return true
  }
  function setSkinTone(value: DocumentEmojiSkinTone): void {
    if (!DOCUMENT_EMOJI_SKIN_TONES.some(tone => tone === value)) throw new TypeError('Unsupported emoji skin tone')
    ensureLoaded()
    if (value !== preferences.value.skinTone) save({ ...preferences.value, skinTone: value })
  }
  function snapshot(): DocumentCharacterPreferencesSnapshot {
    return { ...preferences.value, recentEmojis: [...preferences.value.recentEmojis], recentSymbols: [...preferences.value.recentSymbols] }
  }
  /** No listeners, timers or asynchronous storage tasks are retained. Disposal releases only the storage adapter. */
  function releaseStorage(): void { storage = undefined }
  return { preferences, loaded, storageError, recentEmojis, recentSymbols, skinTone, activeSymbol, activeEmoji, ensureLoaded, select, setSkinTone, snapshot, releaseStorage }
})
export type DocumentCharacterPreferencesStore = ReturnType<typeof useDocumentCharacterPreferencesStore>
