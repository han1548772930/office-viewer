import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { afterAll, afterEach, beforeAll, test } from 'vitest'
import { createPinia } from 'pinia'
import { createServer } from 'vite'
import { buildDocumentEmojiGroups, buildDocumentEmojiLocale, documentEmojiLocaleTitleFallbacks, extractUniverSymbolGroups } from '../../../scripts/generate-document-emojis.mjs'

// Deferred real catalog/generator/Pinia/public-API contracts, 2026-11-28.
// The generator's pure functions are imported; this suite never downloads or generates files.
let server, catalog, characters, preferences, metadata, OfficeStudio, FDocument, domain, documentModel
const stores = new Set(), studios = new Set()
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  catalog = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-character-catalog.ts')
  characters = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-special-characters.ts')
  preferences = await server.ssrLoadModule('/packages/office-studio/src/features/document/stores/document-character-preferences-store.ts')
  ;({ OfficeStudio } = await server.ssrLoadModule('/packages/office-studio/src/api/office-studio.ts'))
  ;({ FDocument } = await server.ssrLoadModule('/packages/office-studio/src/api/facade.ts'))
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  documentModel = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  metadata = JSON.parse(await readFile('packages/office-studio/src/features/document/model/document-character-source.generated.json', 'utf8'))
})
afterEach(() => { for (const store of stores) { store.releaseStorage(); store.$dispose() } stores.clear(); for (const studio of studios) studio.dispose(); studios.clear() })
afterAll(async () => { await server?.close() })
const preferenceStore = () => { const store = preferences.useDocumentCharacterPreferencesStore(createPinia()); stores.add(store); return store }
function memoryStorage(initial = null, rejectWrites = false) {
  let value = initial, reads = 0
  const writes = []
  return { getItem(key) { assert.equal(key, preferences.DOCUMENT_CHARACTER_PREFERENCES_KEY); reads++; return value },
    setItem(key, next) { assert.equal(key, preferences.DOCUMENT_CHARACTER_PREFERENCES_KEY); if (rejectWrites) throw new Error('Storage unavailable'); value = next; writes.push(next) },
    get value() { return value }, get reads() { return reads }, writes }
}

test('generated symbols preserve the fixed three source arrays, duplicates and complete multi-character values before local extensions', async () => {
  const source = await readFile(path.resolve(process.cwd(), '..', 'univer/packages/ui/src/views/symbol-picker/symbols.ts'), 'utf8')
  const fixed = extractUniverSymbolGroups(source), groups = characters.DOCUMENT_SYMBOL_GROUPS
  assert.deepEqual(groups.slice(0, 3), fixed)
  assert.deepEqual(fixed.map(group => [group.key, group.values.length]), [['mathematics', 70], ['greek', 48], ['common', 32]])
  assert.equal(fixed.flatMap(group => group.values).filter(value => value === '%').length, 2)
  for (const value of ['|a|', 'kg', 'mm', 'cm', 'm²', '°C']) assert.ok(fixed.some(group => group.values.includes(value)))
  assert.equal(createHash('sha256').update(source).digest('hex'), metadata.sources.find(source => source.path?.endsWith('symbols.ts')).sha256)
  assert.deepEqual(groups.slice(3).map(group => group.key), ['arrows', 'punctuation', 'currency', 'additional'])
  assert.ok(groups.slice(3).every(group => group.source === 'extension'))
  assert.equal(groups.slice(3).reduce((count, group) => count + group.values.length, 0), 78)
  assert.equal(new Set(groups.slice(3).flatMap(group => group.values).filter(value => !new Set(fixed.flatMap(group => group.values)).has(value))).size, 71)
})

test('Unicode parsing retains fully qualified whole sequences, folds actual tone variants and excludes the fixed Flags group', () => {
  const input = [
    '# group: People & Body',
    '1F44D ; fully-qualified # 👍 E0.6 thumbs up',
    '1F44D 1F3FF ; fully-qualified # 👍🏿 E1.0 thumbs up: dark skin tone',
    '1F91D ; fully-qualified # 🤝 E3.0 handshake',
    '1FAF1 1F3FB 200D 1FAF2 1F3FF ; fully-qualified # 🫱🏻‍🫲🏿 E14.0 handshake: light skin tone, dark skin tone',
    '1F44D ; minimally-qualified # ✋ E0.6 wrong form',
    '# group: Animals & Nature', '1F98A ; fully-qualified # 🦊 E3.0 fox',
    '# group: Flags', '1F1FA 1F1F8 ; fully-qualified # 🇺🇸 E0.6 flag: United States',
  ].join('\n')
  const groups = buildDocumentEmojiGroups(input)
  assert.deepEqual(Object.keys(groups), ['people', 'nature', 'foods', 'activity', 'places', 'objects', 'symbols'])
  assert.deepEqual(groups.people.map(family => [family.emoji, family.variants.map(item => item.emoji)]), [['👍', ['👍🏿']], ['🤝', ['🫱🏻‍🫲🏿']]])
  assert.deepEqual(groups.nature, [{ emoji: '🦊', title: 'Fox' }])
  assert.ok(!Object.values(groups).flat().some(item => item.emoji === '🇺🇸' || item.emoji === '✋'))
})

test('CLDR annotation and derived titles are joined per actual sequence; matching English text is not a missing-title fallback', () => {
  const families = [{ emoji: '✋', title: 'Raised Hand', variants: [{ emoji: '✋🏻', title: 'Raised Hand: Light Skin Tone' }] },
    { emoji: '🤝', title: 'Handshake', variants: [{ emoji: '🫱🏻‍🫲🏿', title: 'Handshake: Light Skin Tone, Dark Skin Tone' }] },
    { emoji: '☀️', title: 'Sun' }]
  const base = { annotations: { annotations: { '✋': { tts: ['رفع اليد'], default: ['توقف', 'يد'] },
    '🤝': { tts: ['Handshake'], default: ['agreement'] }, '☀': { tts: ['الشمس'], default: ['ساطع'] } } } }
  const derived = { annotationsDerived: { annotations: { '✋🏻': { tts: ['يد فاتحة'], default: ['فاتح'] } } } }
  const locale = buildDocumentEmojiLocale('ar-SA', families, base, derived)
  assert.equal(locale.emojiTitles['✋🏻'], 'يد فاتحة'); assert.equal(locale.emojiTitles['☀️'], 'الشمس')
  assert.equal(locale.emojiTitles['🫱🏻‍🫲🏿'], families[1].variants[0].title)
  assert.match(locale.emojiSearchIndex['✋'], /توقف/)
  assert.deepEqual(documentEmojiLocaleTitleFallbacks(families.flatMap(item => [item, ...(item.variants ?? [])]), base, derived), ['🫱🏻‍🫲🏿'])
})

test('the seven catalog groups cover every recorded family and actual sequence with unique family membership', () => {
  assert.equal(metadata.unicodeVersion, '17.0.0'); assert.equal(metadata.cldrVersion, '48.2.0')
  assert.equal(metadata.univerCommit, 'eaf4cfa04fabdbcfc922c692f3450459f8a29546')
  assert.deepEqual(metadata.excludedUnicodeGroups, ['Flags'])
  assert.deepEqual(catalog.DOCUMENT_EMOJI_CATEGORIES, metadata.categories)
  const families = characters.documentEmojiFamilies(), variants = characters.documentEmojiVariants()
  assert.equal(families.length, 1644); assert.equal(variants.length, 3674); assert.equal(variants.length - families.length, 2030)
  assert.equal(new Set(variants.map(item => item.emoji)).size, variants.length)
  for (const category of catalog.DOCUMENT_EMOJI_CATEGORIES) {
    const group = characters.documentEmojiGroup(category)
    assert.equal(group.length, metadata.categoryCounts[category].families)
    assert.equal(group.reduce((count, family) => count + 1 + (family.variants?.length ?? 0), 0), metadata.categoryCounts[category].sequences)
    for (const family of group) for (const item of [family, ...(family.variants ?? [])]) assert.equal(characters.documentEmojiFamilyKey(item.emoji), family.emoji)
  }
})

test('uniform tones always select an existing family sequence and the full mixed-tone list stays available', () => {
  const known = new Set(characters.documentEmojiVariants().map(item => item.emoji))
  for (const family of characters.documentEmojiFamilies()) for (const tone of catalog.DOCUMENT_EMOJI_SKIN_TONES) {
    const item = characters.documentEmojiForSkinTone(family, tone)
    assert.ok(known.has(item.emoji)); assert.equal(characters.documentEmojiFamilyKey(item.emoji), family.emoji)
    if (item.emoji !== family.emoji) assert.ok((item.emoji.match(/[\u{1F3FB}-\u{1F3FF}]/gu) ?? []).every(modifier => modifier === tone))
  }
  assert.equal(characters.documentEmojiForSkinTone(characters.documentEmojiVariant('👍'), '🏽').emoji, '👍🏽')
  assert.equal(characters.documentEmojiForSkinTone(characters.documentEmojiVariant('🦊'), '🏽').emoji, '🦊')
  assert.ok(characters.documentEmojiFamilyVariants('🤝').some(item => item.emoji === '🫱🏻‍🫲🏿'))
  assert.equal(characters.documentEmojiHasMixedTones('🤝'), true)
  assert.deepEqual(characters.documentEmojiFamilyVariants('custom👨‍👩‍👧‍👦'), [])
  assert.equal(characters.documentEmojiVariant('kg').emoji, 'kg')
  for (const sample of [0, 0.5, 0.999999, 1, -1, Number.NaN]) assert.ok(known.has(characters.randomDocumentEmoji(() => sample, '🏿').emoji))
})

test('all 19 lazy locale modules cover every sequence, retain recorded missing titles and search localized, English and actual characters', async () => {
  assert.deepEqual(Object.keys(metadata.localeMapping), catalog.DOCUMENT_CHARACTER_LOCALES)
  const variants = characters.documentEmojiVariants(), titles = new Map(variants.map(item => [item.emoji, item.title]))
  const locales = await Promise.all(catalog.DOCUMENT_CHARACTER_LOCALES.map(locale => characters.loadDocumentEmojiLocale(locale)))
  for (const data of locales) {
    assert.equal(Object.keys(data.emojiTitles).length, 3674)
    assert.ok(variants.every(item => typeof data.emojiTitles[item.emoji] === 'string' && data.emojiTitles[item.emoji].length))
    const source = metadata.locales[data.locale]
    assert.equal(source.titleCount, variants.length); assert.equal(source.englishFallbackCount, source.englishFallbackSequences.length)
    for (const emoji of source.englishFallbackSequences) assert.equal(data.emojiTitles[emoji], titles.get(emoji))
    assert.ok(characters.searchDocumentEmojis(data.emojiTitles['🦊'], data).some(item => item.emoji === '🦊'))
    assert.ok(characters.searchDocumentEmojis('fox', data).some(item => item.emoji === '🦊'))
    assert.ok(characters.searchDocumentEmojis('👍🏿', data).some(item => item.emoji === '👍'))
    assert.ok(characters.documentCharacterPickerLabels(data.locale).search.length)
  }
  assert.deepEqual(characters.searchDocumentEmojis('no such emoji qz8920'), [])
  await assert.rejects(characters.loadDocumentEmojiLocale('invalid-locale'), /Unknown document emoji locale/)
})

test('10-column rows and keyboard neighbors handle headers, short rows, offsets and RTL without indexing Unicode code units', () => {
  const items = (prefix, count) => Array.from({ length: count }, (_, index) => ({ key: prefix + index, value: '👨‍👩‍👧‍👦' + index, title: 'Item ' + index }))
  const model = characters.createDocumentCharacterRows([{ key: 'a', title: 'A', items: items('a', 11) },
    { key: 'b', title: 'B', items: items('b', 3) }, { key: 'empty', title: 'Empty', items: [] }], 'fixture')
  assert.equal(model.height, 144)
  assert.deepEqual(model.rows.map(row => [row.type, row.top, row.height]), [['header', 0, 24], ['characters', 24, 32], ['characters', 56, 32], ['header', 88, 24], ['characters', 112, 32]])
  const neighbor = (key, movement, rtl = false) => characters.documentCharacterNeighbor(model.rows, key, movement, rtl)?.item.key
  assert.equal(neighbor('a9', 'right'), 'a10'); assert.equal(neighbor('a9', 'down'), 'a10')
  assert.equal(neighbor('a10', 'down'), 'b0'); assert.equal(neighbor('b2', 'up'), 'a10')
  assert.equal(neighbor('a0', 'left', true), 'a1'); assert.equal(neighbor('a0', 'right', true), undefined)
  assert.equal(neighbor('b2', 'first'), 'a0'); assert.equal(neighbor('a0', 'last'), 'b2')
  assert.equal(characters.activeDocumentCharacterSection(model.positions, 59), 'a')
  assert.equal(characters.activeDocumentCharacterSection(model.positions, 60), 'b')
  assert.equal(characters.activeDocumentCharacterSection(model.positions, Number.NaN), 'a')
})

test('preferences load once, persist whole strings, deduplicate recents by family and return detached snapshots', () => {
  const storage = memoryStorage(), store = preferenceStore()
  store.ensureLoaded(storage)
  assert.deepEqual(store.recentEmojis, ['👍', '😀', '😘', '😍', '😆', '😜', '😅', '😂', '😱', '😞', '😒'])
  store.select('emoji', '👍🏻'); store.select('emoji', '👍🏿'); store.setSkinTone('🏽')
  store.select('symbol', '|a|e\u0301'); store.select('symbol', 'kg')
  assert.equal(store.recentEmojis[0], '👍🏿'); assert.equal(store.recentEmojis.filter(value => characters.documentEmojiFamilyKey(value) === '👍').length, 1)
  assert.equal(store.recentEmojis.length, 11); assert.deepEqual(store.recentSymbols, ['kg', '|a|e\u0301'])
  const snapshot = store.snapshot(), restored = preferenceStore()
  restored.ensureLoaded(memoryStorage(storage.value)); assert.deepEqual(restored.snapshot(), snapshot)
  snapshot.recentSymbols.push('mutated'); assert.deepEqual(store.recentSymbols, ['kg', '|a|e\u0301'])
  store.ensureLoaded(memoryStorage('invalid JSON')); assert.equal(storage.reads, 1)
  const writes = storage.writes.length; store.releaseStorage(); store.select('symbol', 'm²')
  assert.equal(storage.writes.length, writes); assert.equal(store.activeSymbol, 'm²')
})

test('aggregate and encoded storage budgets discard complete older preferences while arbitrary insertion strings stay intact', () => {
  const storage = memoryStorage(), store = preferenceStore(), selected = new Set(characters.normalizeDocumentRecentEmojis(undefined))
  store.ensureLoaded(storage)
  for (let index = 0; index < 32; index++) { const value = String(index) + '\u0000'.repeat(4094); selected.add(value); store.select('symbol', value) }
  for (let index = 0; index < 11; index++) { const value = 'emoji' + index + '\u0000'.repeat(4090); selected.add(value); store.select('emoji', value) }
  const snapshot = store.snapshot(), all = [snapshot.activeSymbol, snapshot.activeEmoji, ...snapshot.recentSymbols, ...snapshot.recentEmojis]
  assert.ok(all.reduce((sum, item) => sum + item.length, 0) <= preferences.DOCUMENT_CHARACTER_PREFERENCES_VALUE_BUDGET)
  assert.ok(storage.writes.every(value => value.length <= preferences.DOCUMENT_CHARACTER_PREFERENCES_STORAGE_BUDGET))
  assert.ok(all.every(value => selected.has(value)))
  const restored = preferenceStore(); restored.ensureLoaded(memoryStorage(storage.value)); assert.deepEqual(restored.snapshot(), snapshot)
  const longValue = 'e\u0301👨‍👩‍👧‍👦'.repeat(1000), before = [...store.recentSymbols]
  assert.equal(store.select('symbol', longValue), true); assert.equal(store.activeSymbol, '')
  assert.deepEqual(store.recentSymbols, before)
  assert.deepEqual(characters.documentEmojiVariant(longValue), { emoji: longValue, title: longValue })
})

test('malformed, unsupported and oversized storage fail safely; unavailable storage preserves in-memory preferences', () => {
  for (const saved of ['{', '{"version":2}', 'x'.repeat(131073)]) {
    const store = preferenceStore(); store.ensureLoaded(memoryStorage(saved))
    assert.ok(store.storageError); assert.equal(store.skinTone, ''); assert.equal(store.recentEmojis.length, 11)
  }
  const store = preferenceStore(); store.ensureLoaded(memoryStorage(null, true)); store.select('symbol', '°C')
  assert.match(store.storageError, /Storage unavailable/); assert.equal(store.activeSymbol, '°C')
  assert.throws(() => store.setSkinTone('invalid'), /Unsupported emoji skin tone/)
  assert.deepEqual(characters.normalizeDocumentRecentSymbols(['kg', 'kg', null, { value: 'mm' }, '']), ['kg'])
  assert.equal(characters.normalizeDocumentRecentSymbols(Array.from({ length: 40 }, (_, index) => String(index))).length, 32)
})

test('one application preference store is shared across documents and remains isolated from snapshots, input, content history and undo/reopen', () => {
  const model = domain.createBlankDocument(); model.id = 'character-api'; model.blocks = [documentModel.createParagraph('ab')]
  const studio = new OfficeStudio({ document: model, currentUserId: 'editor' }); studios.add(studio); studio.activateDocument(model.id)
  const document = new FDocument(studio, model.id), store = document.getCharacterPreferences(), storage = memoryStorage()
  assert.equal(store, studio.documentCharacterPreferences); store.ensureLoaded(storage)
  const caret = documentModel.collapsedSelection({ blockId: model.blocks[0].id, offset: 1 }); document.setInputSelection(caret)
  const snapshot = document.getSnapshot(), input = document.getInputState()
  store.select('emoji', '🫱🏻‍🫲🏿'); store.select('symbol', '|a|'); store.setSkinTone('🏽')
  assert.deepEqual(document.getSnapshot(), snapshot); assert.deepEqual(document.getInputState(), input)
  assert.equal(studio.editor.history.units[model.id]?.undo.length ?? 0, 0); assert.equal(studio.editor.history.isDirty(model.id), false)
  const otherModel = domain.createBlankDocument(); otherModel.id = 'character-api-other'; studio.createDocument(otherModel)
  assert.equal(new FDocument(studio, otherModel.id).getCharacterPreferences(), store)
  const result = document.executeCommand({ id: 'doc.command.insert-text', text: '|a|' }, caret, 'doc.command.insert-special-character')
  assert.equal(result.changed, true); const chosen = store.snapshot()
  document.undo(); assert.deepEqual(store.snapshot(), chosen); assert.deepEqual(document.getSnapshot().body, snapshot.body)
  document.redo(); const saved = document.save()
  const reopened = new OfficeStudio({ document: saved, currentUserId: 'editor' }); studios.add(reopened)
  const next = new FDocument(reopened, saved.id); next.getCharacterPreferences().ensureLoaded(memoryStorage(storage.value))
  assert.deepEqual(next.getCharacterPreferences().snapshot(), chosen)
  assert.equal(JSON.stringify(saved).includes('character-preferences'), false)
})

