import { writeFile } from 'node:fs/promises'
import { coreEnUS } from '../packages/office-studio/src/i18n/core/locales/en-US.ts'
import { docsEnUS } from '../packages/office-studio/src/i18n/docs/locales/en-US.ts'
import { sheetsEnUS } from '../packages/office-studio/src/i18n/sheets/locales/en-US.ts'
import { supportedOfficeLocales } from '../packages/office-studio/src/i18n/office-i18n.ts'

const targets = {
  'ar-SA': 'ar', 'ca-ES': 'ca', 'de-DE': 'de', 'es-ES': 'es', 'fa-IR': 'fa', 'fr-FR': 'fr',
  'id-ID': 'id', 'it-IT': 'it', 'ja-JP': 'ja', 'ko-KR': 'ko', 'pl-PL': 'pl', 'pt-BR': 'pt',
  'ru-RU': 'ru', 'sk-SK': 'sk', 'vi-VN': 'vi', 'zh-CN': 'zh-CN', 'zh-HK': 'zh-TW', 'zh-TW': 'zh-TW',
}

function protectPlaceholders(value) {
  const placeholders = []
  const protectedValue = value.replace(/\{[^{}]+\}/g, placeholder => {
    const token = `ZXQPLACEHOLDER${placeholders.length}QXZ`
    placeholders.push([token, placeholder])
    return token
  })
  return {
    value: protectedValue,
    restore(translated) {
      return placeholders.reduce((result, [token, placeholder]) => result.replaceAll(token, placeholder), translated)
    },
  }
}

async function translateBatch(values, target) {
  const protectedValues = values.map(protectPlaceholders)
  const query = protectedValues.map(item => item.value).join('\n')
  const url = new URL('https://translate.googleapis.com/translate_a/single')
  url.searchParams.set('client', 'gtx')
  url.searchParams.set('sl', 'en')
  url.searchParams.set('tl', target)
  url.searchParams.set('dt', 't')
  url.searchParams.set('q', query)

  for (let attempt = 0; attempt < 4; attempt += 1) {
    const response = await fetch(url)
    if (response.ok) {
      const data = await response.json()
      const translated = data[0].map(part => part[0]).join('').split('\n')
      if (translated.length === values.length) return translated.map((value, index) => protectedValues[index].restore(value))
    }
    await new Promise(resolve => setTimeout(resolve, 500 * (attempt + 1)))
  }
  throw new Error(`Could not translate batch to ${target}`)
}

async function translateLocale(locale, source) {
  const target = targets[locale]
  const translated = {}
  const entries = Object.entries(source)
  for (let offset = 0; offset < entries.length; offset += 18) {
    const batch = entries.slice(offset, offset + 18)
    const values = await translateBatch(batch.map(([, value]) => value), target)
    batch.forEach(([key], index) => { translated[key] = values[index] })
  }
  return translated
}

async function generate(feature, source, folder, prefix) {
  const locales = supportedOfficeLocales.filter(locale => locale !== 'en-US')
  const resources = {}
  for (let offset = 0; offset < locales.length; offset += 2) {
    const batch = locales.slice(offset, offset + 2)
    const results = await Promise.all(batch.map(async locale => [locale, await translateLocale(locale, source)]))
    results.forEach(([locale, translated]) => { resources[locale] = translated })
    process.stdout.write(`${feature}: ${Math.min(offset + batch.length, locales.length)}/${locales.length} locales\n`)
  }

  const lines = [
    '/* Machine-translated UI labels. Pending human review. */',
    `export const ${prefix}Translations: Partial<Record<string, Record<string, string>>> = {`,
  ]
  for (const locale of locales) {
    lines.push(`  ${JSON.stringify(locale)}: {`)
    for (const [key, value] of Object.entries(resources[locale])) lines.push(`    ${JSON.stringify(key)}: ${JSON.stringify(value)},`)
    lines.push('  },')
  }
  lines.push('}')
  await writeFile(folder, `${lines.join('\n')}\n`, 'utf8')
}

const requestedFeature = process.argv[2]
if (!requestedFeature || requestedFeature === 'core') await generate('core', coreEnUS, 'packages/office-studio/src/i18n/core/locales/generated.ts', 'core')
if (!requestedFeature || requestedFeature === 'sheets') await generate('sheets', sheetsEnUS, 'packages/office-studio/src/i18n/sheets/locales/generated.ts', 'sheets')
if (!requestedFeature || requestedFeature === 'docs') await generate('docs', docsEnUS, 'packages/office-studio/src/i18n/docs/locales/generated.ts', 'docs')
