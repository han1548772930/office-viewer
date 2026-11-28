/**
 * Character data generator for the local Vue workbench.
 * Emoji grouping and picker labels are adapted from DreamNum Co., Ltd.,
 * Univer v1.6.0 (Apache-2.0), commit eaf4cfa04fabdbcfc922c692f3450459f8a29546.
 * https://www.apache.org/licenses/LICENSE-2.0
 * Unicode/CLDR data are covered by the Unicode Data Files and Software License.
 * The generated files retain attribution and a copy of that license in docs/third-party.
 */
import { execFileSync, spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const UNIVER_CHARACTER_COMMIT = 'eaf4cfa04fabdbcfc922c692f3450459f8a29546'
export const UNICODE_EMOJI_VERSION = '17.0.0'
export const CLDR_JSON_VERSION = '48.2.0'
export const EMOJI_CATEGORY_KEYS = ['people', 'nature', 'foods', 'activity', 'places', 'objects', 'symbols']
export const CLDR_LOCALE_MAP = {
  'ar-SA': 'ar', 'ca-ES': 'ca', 'de-DE': 'de', 'en-US': 'en', 'es-ES': 'es', 'fa-IR': 'fa', 'fr-FR': 'fr',
  'id-ID': 'id', 'it-IT': 'it', 'ja-JP': 'ja', 'ko-KR': 'ko', 'pl-PL': 'pl', 'pt-BR': 'pt', 'ru-RU': 'ru',
  'sk-SK': 'sk', 'vi-VN': 'vi', 'zh-CN': 'zh', 'zh-HK': 'zh-Hant-HK', 'zh-TW': 'zh-Hant',
}
const CATEGORY_MAP = { 'Smileys & Emotion': 'people', 'People & Body': 'people', 'Animals & Nature': 'nature',
  'Food & Drink': 'foods', Activities: 'activity', 'Travel & Places': 'places', Objects: 'objects', Symbols: 'symbols' }
const TONE_NAMES = new Set(['Light Skin Tone', 'Medium-light Skin Tone', 'Medium Skin Tone', 'Medium-dark Skin Tone', 'Dark Skin Tone'])
const TONE_PATTERN = /[\u{1F3FB}-\u{1F3FF}]/u
export const DEFAULT_RECENT_EMOJIS = ['👍', '😀', '😘', '😍', '😆', '😜', '😅', '😂', '😱', '😞', '😒']
const scriptFile = fileURLToPath(import.meta.url)
const workspace = path.resolve(path.dirname(scriptFile), '..')
const fixedSource = path.resolve(workspace, '..', 'univer')
const modelDirectory = path.join(workspace, 'packages/office-studio/src/features/document/model')
const cacheDirectory = path.join(workspace, '.research.local/document-character-sources')
const emojiTestUrl = `https://www.unicode.org/Public/${UNICODE_EMOJI_VERSION}/emoji/emoji-test.txt`
const cldrBaseUrl = `https://cdn.jsdelivr.net/gh/unicode-org/cldr-json@${CLDR_JSON_VERSION}/cldr-json`
const unicodeLicenseUrl = 'https://www.unicode.org/license.txt'

const titleCase = text => text.split(' ').map(word => word ? word[0].toUpperCase() + word.slice(1) : word).join(' ')
const sha256 = text => createHash('sha256').update(text).digest('hex')

export function skinToneFamilyTitle(title, baseItemsByTitle) {
  const separator = title.indexOf(': ')
  if (separator < 0) return title
  const baseTitle = title.slice(0, separator)
  const descriptors = title.slice(separator + 2).split(', ').filter(value => !TONE_NAMES.has(value))
  const familyTitle = descriptors.length ? `${baseTitle}: ${descriptors.join(', ')}` : baseTitle
  return familyTitle.endsWith(': Person, Person') && baseItemsByTitle.has(baseTitle) ? baseTitle : familyTitle
}

export function groupEmojiSkinToneFamilies(items) {
  const bases = new Map(items.filter(item => !TONE_PATTERN.test(item.emoji)).map(item => [item.title, item]))
  const variants = new Map(), grouped = new Set()
  for (const item of items) {
    if (!TONE_PATTERN.test(item.emoji)) continue
    const base = bases.get(skinToneFamilyTitle(item.title, bases))
    if (!base) continue
    const family = variants.get(base.emoji) ?? []
    family.push(item); variants.set(base.emoji, family); grouped.add(item.emoji)
  }
  return items.filter(item => !grouped.has(item.emoji)).map(item => variants.has(item.emoji) ? { ...item, variants: variants.get(item.emoji) } : item)
}

export function buildDocumentEmojiGroups(emojiTest) {
  const groups = Object.fromEntries(EMOJI_CATEGORY_KEYS.map(key => [key, []]))
  let category
  for (const line of emojiTest.split(/\r?\n/)) {
    const group = line.match(/^# group: (.+)$/)
    if (group) { category = CATEGORY_MAP[group[1]]; continue }
    if (!category) continue
    const item = line.match(/^\S(?:.+?)\s+;\s+fully-qualified\s+#\s+(\S+)\s+E[\d.]+\s+(.+)$/)
    if (item) groups[category].push({ emoji: item[1], title: titleCase(item[2]) })
  }
  return Object.fromEntries(EMOJI_CATEGORY_KEYS.map(key => [key, groupEmojiSkinToneFamilies(groups[key])]))
}

function annotation(source, emoji) {
  const entries = source.annotations?.annotations ?? source.annotationsDerived?.annotations
  return entries?.[emoji] ?? entries?.[emoji.replace(/\uFE0F/g, '')]
}
export function documentEmojiLocaleTitleFallbacks(items, ...sources) {
  return items.filter(item => !sources.some(source => annotation(source, item.emoji)?.tts?.[0])).map(item => item.emoji)
}
export function buildDocumentEmojiLocale(locale, families, ...sources) {
  const emojiTitles = {}, emojiSearchIndex = {}
  for (const family of families) {
    const titles = [], keywords = new Set()
    for (const item of [family, ...(family.variants ?? [])]) {
      const annotations = sources.map(source => annotation(source, item.emoji)).filter(Boolean)
      const title = annotations.find(entry => entry.tts?.[0])?.tts?.[0] ?? item.title
      emojiTitles[item.emoji] = title
      titles.push(item.title.toLowerCase(), title.toLowerCase())
      for (const entry of annotations) for (const word of [...(entry.tts ?? []), ...(entry.default ?? [])]) if (word) keywords.add(word.toLowerCase())
    }
    const aliases = [...keywords].filter(word => !titles.some(title => title.includes(word)))
    const searchText = aliases.filter(word => !aliases.some(alias => alias !== word && alias.includes(word))).join(' ')
    if (searchText) emojiSearchIndex[family.emoji] = searchText
  }
  return { locale, emojiTitles, emojiSearchIndex }
}

function quotedText(literal) {
  return literal.slice(1, -1).replace(/\\(u\{[\da-fA-F]+\}|u[\da-fA-F]{4}|x[\da-fA-F]{2}|.)/g, (_whole, escape) => {
    if (escape.startsWith('u{')) return String.fromCodePoint(Number.parseInt(escape.slice(2, -1), 16))
    if (escape[0] === 'u' || escape[0] === 'x') return String.fromCodePoint(Number.parseInt(escape.slice(1), 16))
    return ({ n: '\n', r: '\r', t: '\t', b: '\b', f: '\f' })[escape] ?? escape
  })
}
const literalPattern = /([a-zA-Z][\w]*):\s*('(?:\\.|[^'\\])*'|"(?:\\.|[^"\\])*")/g
function literalSection(source, section) {
  const body = source.match(new RegExp(`\\b${section}:\\s*\\{([\\s\\S]*?)\\n\\s*\\},`))?.[1]
  if (!body) throw new Error(`Fixed Univer source has no ${section} literal section`)
  return Object.fromEntries([...body.matchAll(literalPattern)].map(match => [match[1], quotedText(match[2])]))
}
export function extractUniverCharacterLabels(source) {
  const emoji = literalSection(source, 'emojiPicker'), symbol = literalSection(source, 'symbolPicker')
  const fields = { search: emoji.search, random: emoji.random, recents: emoji.recents, people: emoji.emojis, nature: emoji.animals,
    foods: emoji.food, activity: emoji.activities, places: emoji.places, objects: emoji.objects, symbols: emoji.symbols,
    searchResults: emoji.searchResults, noResults: emoji.noResults, mathematics: symbol.mathematics, greek: symbol.greek, common: symbol.common,
    more: literalSection(source, 'ribbon').more }
  for (const [key, value] of Object.entries(fields)) if (typeof value !== 'string' || !value.length) throw new Error(`Missing fixed picker label ${key}`)
  return fields
}
export function extractUniverSymbolGroups(source) {
  const groups = [...source.matchAll(/key:\s*'(mathematics|greek|common)'[\s\S]*?symbols:\s*\[([\s\S]*?)\]/g)]
    .map(match => ({ key: match[1], source: 'univer', values: [...match[2].matchAll(/'(?:\\.|[^'\\])*'/g)].map(item => quotedText(item[0])) }))
  if (groups.length !== 3 || groups.reduce((count, group) => count + group.values.length, 0) !== 150) throw new Error('Fixed symbol catalog must contain three groups and 150 array entries')
  return groups
}

const localSymbolGroups = {
  common: '© ® ™ § ¶ † ‡ • … ‰ ° ℃ ℉ № ✓ ✗ ★ ☆ ♥ ♦ ♣ ♠',
  mathematics: '± × ÷ ≠ ≈ ≤ ≥ ∞ √ ∑ ∏ ∫ ∂ ∆ ∇ ∈ ∉ ⊂ ⊃ ∩ ∪ ∅ ∀ ∃',
  arrows: '← ↑ → ↓ ↔ ↕ ↖ ↗ ↘ ↙ ↩ ↪ ⇐ ⇑ ⇒ ⇓ ⇔ ⇕ ➜ ➤ ➔ ➟',
  punctuation: '“ ” ‘ ’ « » ‹ › 「 」 『 』 【 】 〔 〕 〈 〉 《 》 — – · …',
  currency: '¥ $ € £ ₩ ₽ ₹ ₫ ฿ ₱ ₺ ₴ ₦ ¢ ¤',
}
export function appendLocalSymbolExtensions(groups) {
  const upstream = new Set(groups.flatMap(group => group.values))
  const additional = [...new Set([localSymbolGroups.common, localSymbolGroups.mathematics].flatMap(group => group.split(' ')))].filter(value => !upstream.has(value))
  return [...groups, ...['arrows', 'punctuation', 'currency'].map(key => ({ key, source: 'extension', values: localSymbolGroups[key].split(' ') })),
    { key: 'additional', source: 'extension', values: additional }]
}

const sourceHeader = `/**
 * Generated by scripts/generate-document-emojis.mjs; do not edit manually.
 * Unicode Emoji ${UNICODE_EMOJI_VERSION}, CLDR JSON ${CLDR_JSON_VERSION}.
 * Copyright © Unicode, Inc. Unicode Data Files and Software License:
 * https://www.unicode.org/license.txt; local copy: docs/third-party/UNICODE_DATA_LICENSE.txt.
 * Picker labels/symbol data: Copyright 2023-present DreamNum Co., Ltd. (Apache-2.0).
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy at https://www.apache.org/licenses/LICENSE-2.0.
 * Unless required by applicable law or agreed to in writing, distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * Fixed source: dream-num/univer@${UNIVER_CHARACTER_COMMIT}.
 */\n`
const typedModule = (imports, declaration, value) => `${sourceHeader}${imports}\n\n${declaration} = ${JSON.stringify(value, null, 2)}\n`

function getCldrUrl(locale, derived) {
  return derived ? `${cldrBaseUrl}/cldr-annotations-derived-full/annotationsDerived/${locale}/annotations.json`
    : `${cldrBaseUrl}/cldr-annotations-full/annotations/${locale}/annotations.json`
}
async function downloadSource(url, cacheName) {
  const cachePath = path.join(cacheDirectory, cacheName)
  try { return await readFile(cachePath, 'utf8') } catch (error) { if (error?.code !== 'ENOENT') throw error }
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(30_000) })
      if (!response.ok) throw new Error(`HTTP ${response.status} loading ${url}`)
      const text = await response.text()
      await mkdir(cacheDirectory, { recursive: true }); await writeFile(cachePath, text, 'utf8')
      return text
    } catch (error) {
      if (attempt === 2) throw error
      await new Promise(resolve => setTimeout(resolve, 500 * 2 ** attempt))
    }
  }
  throw new Error(`Could not read ${url}`)
}
async function mapBounded(items, mapper) {
  const output = new Array(items.length)
  let next = 0
  await Promise.all(Array.from({ length: Math.min(4, items.length) }, async () => {
    while (next < items.length) { const index = next++; output[index] = await mapper(items[index]) }
  }))
  return output
}

export async function generateDocumentCharacters() {
  const commit = execFileSync('git', ['-C', fixedSource, 'rev-parse', 'HEAD'], { encoding: 'utf8' }).trim()
  if (commit !== UNIVER_CHARACTER_COMMIT) throw new Error(`Expected fixed Univer ${UNIVER_CHARACTER_COMMIT}, found ${commit}`)
  const emojiTest = await downloadSource(emojiTestUrl, `emoji-test-${UNICODE_EMOJI_VERSION}.txt`)
  const groups = buildDocumentEmojiGroups(emojiTest)
  const families = EMOJI_CATEGORY_KEYS.flatMap(key => groups[key])
  if (!families.length || EMOJI_CATEGORY_KEYS.some(key => !groups[key].length)) throw new Error('Unicode source produced an empty catalog category')
  const all = families.flatMap(family => [family, ...(family.variants ?? [])])
  if (new Set(all.map(item => item.emoji)).size !== all.length) throw new Error('Generated catalog contains duplicate Unicode sequences')
  const titles = new Map(all.map(item => [item.emoji, item.title]))
  const recents = DEFAULT_RECENT_EMOJIS.map(emoji => { const title = titles.get(emoji); if (!title) throw new Error(`Default recent emoji ${emoji} is missing`); return { emoji, title } })
  const symbolSource = await readFile(path.join(fixedSource, 'packages/ui/src/views/symbol-picker/symbols.ts'), 'utf8')
  const symbols = appendLocalSymbolExtensions(extractUniverSymbolGroups(symbolSource))
  const uiLabels = {}, localeResults = await mapBounded(Object.entries(CLDR_LOCALE_MAP), async ([locale, cldrLocale]) => {
    const baseUrl = getCldrUrl(cldrLocale, false), derivedUrl = getCldrUrl(cldrLocale, true)
    const base = await downloadSource(baseUrl, `cldr-${CLDR_JSON_VERSION}-${cldrLocale}-annotations.json`)
    const derived = await downloadSource(derivedUrl, `cldr-${CLDR_JSON_VERSION}-${cldrLocale}-derived.json`)
    const localeSource = await readFile(path.join(fixedSource, `packages/ui/src/locale/${locale}.ts`), 'utf8')
    uiLabels[locale] = extractUniverCharacterLabels(localeSource)
    const annotationSources = [JSON.parse(base), JSON.parse(derived)]
    const data = buildDocumentEmojiLocale(locale, families, ...annotationSources)
    // A localized title may intentionally equal the English title. Only a
    // missing CLDR title is a fallback; string equality is not evidence.
    const fallbackTitles = documentEmojiLocaleTitleFallbacks(all, ...annotationSources)
    return { locale, cldrLocale, data, fallbackTitles, sources: [{ url: baseUrl, sha256: sha256(base) }, { url: derivedUrl, sha256: sha256(derived) }], localeSourceSha256: sha256(localeSource) }
  })
  const license = await downloadSource(unicodeLicenseUrl, 'unicode-license.txt')
  const categoryCounts = Object.fromEntries(EMOJI_CATEGORY_KEYS.map(key => [key, { families: groups[key].length,
    sequences: groups[key].reduce((count, family) => count + 1 + (family.variants?.length ?? 0), 0) }]))
  const upstreamValues = new Set(symbols.filter(group => group.source === 'univer').flatMap(group => group.values))
  const extensionValues = symbols.filter(group => group.source === 'extension').flatMap(group => group.values)
  const metadata = { unicodeVersion: UNICODE_EMOJI_VERSION, cldrVersion: CLDR_JSON_VERSION, univerCommit: commit,
    categories: EMOJI_CATEGORY_KEYS, excludedUnicodeGroups: ['Flags'], families: families.length, sequences: all.length,
    skinToneVariants: all.length - families.length, defaultRecentCount: recents.length, categoryCounts,
    symbolCounts: symbols.map(group => ({ key: group.key, source: group.source, entries: group.values.length })),
    symbolUpstreamEntries: 150, symbolExtensionEntries: extensionValues.length,
    symbolExtensionUniqueAdditions: new Set(extensionValues.filter(value => !upstreamValues.has(value))).size,
    sources: [{ url: emojiTestUrl, sha256: sha256(emojiTest) }, { path: 'packages/ui/src/views/symbol-picker/symbols.ts', sha256: sha256(symbolSource) }],
    localeMapping: CLDR_LOCALE_MAP, locales: Object.fromEntries(localeResults.map(result => [result.locale, { cldrLocale: result.cldrLocale,
      titleCount: Object.keys(result.data.emojiTitles).length, searchFamilyCount: Object.keys(result.data.emojiSearchIndex).length,
      englishFallbackCount: result.fallbackTitles.length, englishFallbackSequences: result.fallbackTitles, sources: result.sources,
      localeSourceSha256: result.localeSourceSha256 }])), unicodeLicense: { url: unicodeLicenseUrl, sha256: sha256(license) } }
  const generated = new Map([
    ['document-emojis.generated.ts', typedModule("import type { DocumentEmojiCategory, DocumentEmojiFamily } from './document-character-catalog'", 'export const DOCUMENT_EMOJI_GROUPS: Readonly<Record<DocumentEmojiCategory, readonly DocumentEmojiFamily[]>>', groups) + `\nexport const DOCUMENT_DEFAULT_RECENT_EMOJIS = ${JSON.stringify(recents, null, 2)}\n`],
    ['document-symbols.generated.ts', typedModule("import type { DocumentSymbolGroup } from './document-character-catalog'", 'export const DOCUMENT_SYMBOL_GROUPS: readonly DocumentSymbolGroup[]', symbols)],
    ['document-emoji-ui.generated.ts', typedModule("import type { DocumentCharacterLocale, DocumentCharacterPickerLabels } from './document-character-catalog'", 'export const DOCUMENT_CHARACTER_PICKER_LABELS: Readonly<Record<DocumentCharacterLocale, DocumentCharacterPickerLabels>>', Object.fromEntries(Object.keys(CLDR_LOCALE_MAP).map(locale => [locale, uiLabels[locale]])))],
    ['document-character-source.generated.json', JSON.stringify(metadata, null, 2) + '\n'],
  ])
  for (const result of localeResults) generated.set(`emoji-locales/${result.locale}.generated.ts`,
    typedModule("import type { DocumentEmojiLocaleData } from '../document-character-catalog'", 'const data: DocumentEmojiLocaleData', result.data) + '\nexport default data\n')
  for (const [relative, text] of generated) { const file = path.join(modelDirectory, relative); await mkdir(path.dirname(file), { recursive: true }); await writeFile(file, text, 'utf8') }
  const licensePath = path.join(workspace, 'docs/third-party/UNICODE_DATA_LICENSE.txt')
  await mkdir(path.dirname(licensePath), { recursive: true }); await writeFile(licensePath, license, 'utf8')
  const bytes = [...generated.values()].reduce((count, text) => count + Buffer.byteLength(text, 'utf8'), 0)
  console.log(JSON.stringify({ files: generated.size, generatedBytes: bytes, families: families.length, sequences: all.length, categoryCounts,
    symbolUpstreamEntries: 150, symbolExtensionEntries: extensionValues.length, locales: localeResults.length }, null, 2))
}

function relaunchForConfiguredProxy() {
  const proxyConfigured = ['HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY', 'http_proxy', 'https_proxy', 'all_proxy'].some(key => !!process.env[key])
  if (!proxyConfigured || process.env.OFFICE_DOCUMENT_EMOJI_PROXY_REEXEC || !process.allowedNodeEnvironmentFlags.has('--use-env-proxy')
    || process.execArgv.includes('--use-env-proxy') || process.env.NODE_OPTIONS?.split(/\s+/).includes('--use-env-proxy')) return false
  const result = spawnSync(process.execPath, ['--use-env-proxy', ...process.execArgv, ...process.argv.slice(1)],
    { env: { ...process.env, OFFICE_DOCUMENT_EMOJI_PROXY_REEXEC: '1' }, stdio: 'inherit' })
  if (result.error) throw result.error
  process.exitCode = result.status ?? 1
  return true
}
if (process.argv[1] && path.resolve(process.argv[1]) === scriptFile && !relaunchForConfiguredProxy()) {
  generateDocumentCharacters().catch(error => { console.error(error); process.exitCode = 1 })
}
