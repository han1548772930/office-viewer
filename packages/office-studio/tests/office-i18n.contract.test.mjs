import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createOfficeMessages, createOfficeTranslator, supportedOfficeLocales } from '../src/i18n/office-i18n.ts'
import { sheetsEnUS } from '../src/i18n/sheets/locales/en-US.ts'
import { docsEnUS } from '../src/i18n/docs/locales/en-US.ts'
import { coreEnUS } from '../src/i18n/core/locales/en-US.ts'
import { createSheetRibbon } from '../src/components/spreadsheet/sheet-ribbon.ts'
import { createDocumentRibbon } from '../src/components/document/document-ribbon.ts'
import { localizeRibbon } from '../src/i18n/office-i18n.ts'

test('all fixed Univer locales have complete non-empty fallback catalogs', () => {
  const messages = createOfficeMessages()
  assert.deepEqual(supportedOfficeLocales, [
    'ar-SA', 'ca-ES', 'de-DE', 'en-US', 'es-ES', 'fa-IR', 'fr-FR', 'id-ID', 'it-IT', 'ja-JP',
    'ko-KR', 'pl-PL', 'pt-BR', 'ru-RU', 'sk-SK', 'vi-VN', 'zh-CN', 'zh-HK', 'zh-TW',
  ])

  for (const locale of supportedOfficeLocales) {
    assert.ok(Object.keys(messages[locale]).length >= 300, `${locale} catalog is incomplete`)
    assert.ok(Object.values(messages[locale]).every(value => typeof value === 'string' && value.length > 0), `${locale} has an empty translation`)
  }
})

test('each supported locale translates every Ribbon and registered panel key', () => {
  const messages = createOfficeMessages()
  const productionKeys = [...new Set([
    ...Object.keys(sheetsEnUS),
    ...Object.keys(docsEnUS),
    ...Object.keys(coreEnUS),
  ])]
  const expected = supportedOfficeLocales.filter(locale => locale !== 'en-US')
  const identicalTerms = {
    'ca-ES': ['groups.font', 'items.sheet.command.set-font-family', 'items.sheet.operation.set-zoom', 'items.sheet.operation.open-function.category-3', 'options.sheet.command.set-number-format.general', 'options.sheet.command.set-number-format.text', 'options.sheet.command.set-calculation-mode.manual', 'items.doc-font-family', 'items.doc-emoji', 'items.doc-zoom', 'common.general', 'common.text', 'common.no', 'workspace.document', 'workspace.pending-document', 'workspace.severity-error', 'number-format.text', 'validation.type.decimal'],
    'de-DE': ['items.sheet.operation.open-function.category-3', 'items.sheet.menu.text-rotation.0', 'options.sheet.command.set-number-format.text', 'items.doc-link', 'items.doc-symbol', 'items.doc-emoji', 'document.latex.inline', 'common.standard', 'common.text', 'number-format.text', 'page-settings.margin-top'],
    'es-ES': ['common.no'],
    'fr-FR': ['options.sheet.command.set-number-format.date', 'items.doc-orientation', 'document.code-block.source', 'common.date', 'common.portrait', 'sparkline.minimum', 'number-format.date', 'validation.type.date', 'content-control.type.date'],
    'id-ID': ['tabs.data'],
    'it-IT': ['items.sheet.operation.set-zoom', 'items.doc-emoji', 'items.doc-zoom', 'common.no'],
    'pl-PL': ['items.doc-link', 'items.doc-symbol', 'items.doc-emoji'],
    'pt-BR': ['items.sheet.operation.set-zoom', 'items.doc-emoji', 'items.doc-zoom'],
    'sk-SK': ['items.sheet.data.filter', 'items.sheet.operation.set-zoom', 'items.sheet.operation.open-function.category-3', 'options.sheet.command.set-number-format.text', 'items.doc-symbol', 'items.doc-emoji', 'items.doc-zoom', 'document.latex.inline', 'common.text', 'sparkline.minimum', 'number-format.text'],
  }

  for (const locale of expected) {
    const allowedIdentical = new Set(identicalTerms[locale] ?? [])
    const untranslated = productionKeys.filter(key => !messages[locale][key] || (messages[locale][key] === messages['en-US'][key] && !allowedIdentical.has(key)))
    assert.deepEqual(untranslated, [], `${locale} has English fallbacks: ${untranslated.join(', ')}`)
  }
})

test('localized production Ribbons use a non-English catalog value for every visible display string', () => {
  const sheetTabs = createSheetRibbon({
    canUndo: true, canRedo: true, hasSheet: true, canMerge: true, canUnmerge: true,
    canDeleteRow: true, canDeleteColumn: true, chartAvailable: true, dataValidationAvailable: true, canSort: true,
    fontFamily: 'Arial', fontSize: 11, numberFormat: 'general', bold: false, italic: false, underline: false,
    wrap: true, align: 'left', gridlines: true, frozen: false, zoom: 100,
  })
  const docState = { marks: {}, mixedMarks: [], paragraph: { align: 'left' }, hasSelection: false }
  const docTabs = createDocumentRibbon({
    state: docState, canUndo: true, canRedo: true, outline: false, focus: false, fullscreen: false,
    paper: 'A4', orientation: 'portrait', zoom: 100, contentControlsAvailable: true,
  })
  const messages = createOfficeMessages()

  for (const locale of supportedOfficeLocales.filter(value => value !== 'en-US')) {
    for (const [namespace, source] of [['sheet', sheetTabs], ['doc', docTabs]]) {
      const localized = localizeRibbon(source, namespace, createOfficeTranslator(locale))
      for (const tab of localized) {
        assert.equal(tab.title, messages[locale][`tabs.${tab.id}`], `${locale} tab ${tab.id}`)
        for (const group of tab.groups) {
          assert.equal(group.title, messages[locale][`groups.${group.id}`], `${locale} group ${group.id}`)
          for (const item of group.items) {
            assert.equal(item.title, messages[locale][`items.${item.id}`], `${locale} item ${item.id}`)
            const original = source.find(candidate => candidate.id === tab.id)?.groups
              .find(candidate => candidate.id === group.id)?.items.find(candidate => candidate.id === item.id)
            for (const [index, option] of (item.options ?? []).entries()) {
              const key = `options.${namespace}.${item.id}.${String(option.value)}`
              const expected = messages[locale][key] ?? original?.options?.[index]?.label
              assert.equal(option.label, expected, `${locale} option ${item.id}`)
            }
          }
        }
      }
    }
  }
})

test('locale lookup uses feature translations and falls back to English before the caller text', () => {
  const arabic = createOfficeTranslator('ar-SA')
  const french = createOfficeTranslator('fr-FR')
  assert.equal(arabic('sheet.items.sheet.command.paste', '粘贴'), 'لصق')
  assert.equal(french('doc.items.doc-copy', '复制'), 'Copier')
  assert.equal(arabic('core.workspace.rename', 'Rename'), 'إعادة تسمية')
})
