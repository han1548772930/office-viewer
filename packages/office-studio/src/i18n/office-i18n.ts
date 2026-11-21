import type { InjectionKey, Ref } from 'vue'
import { createI18n } from 'vue-i18n'
import type { RibbonTab } from '../platform/ribbon/ribbon-schema.ts'
import { coreEnUS } from './core/locales/en-US.ts'
import { coreZhCN } from './core/locales/zh-CN.ts'
import { coreTranslations } from './core/locales/generated.ts'
import { reviewedCoreTranslations } from './core/locales/reviewed.ts'
import { docsEnUS } from './docs/locales/en-US.ts'
import { docsZhCN } from './docs/locales/zh-CN.ts'
import { docsTranslations } from './docs/locales/generated.ts'
import { reviewedDocsTranslations } from './docs/locales/reviewed.ts'
import { sheetsEnUS } from './sheets/locales/en-US.ts'
import { sheetsZhCN } from './sheets/locales/zh-CN.ts'
import { sheetsTranslations } from './sheets/locales/generated.ts'
import { reviewedSheetsTranslations } from './sheets/locales/reviewed.ts'
import { conditionalFormattingLocaleMessages } from './sheets/locales/conditional-formatting.ts'
import { sheetLocalearSA, sheetLocaleesES, sheetLocalecaES, sheetLocaledeDE, sheetLocalefaIR, sheetLocalefrFR, sheetLocaleidID, sheetLocaleitIT, sheetLocalejaJP, sheetLocalekoKR, sheetLocaleplPL, sheetLocaleptBR, sheetLocaleruRU, sheetLocaleskSK, sheetLocaleviVN, sheetLocalezhHK, sheetLocalezhTW } from './sheets/locales/univer-parity.ts'
import { docLocalearSA, docLocaleesES, docLocalecaES, docLocaledeDE, docLocalefaIR, docLocalefrFR, docLocaleidID, docLocaleitIT, docLocalejaJP, docLocalekoKR, docLocaleplPL, docLocaleptBR, docLocaleruRU, docLocaleskSK, docLocaleviVN, docLocalezhHK, docLocalezhTW } from './docs/locales/univer-parity.ts'

export const supportedOfficeLocales = [
  'ar-SA', 'ca-ES', 'de-DE', 'en-US', 'es-ES', 'fa-IR', 'fr-FR', 'id-ID', 'it-IT', 'ja-JP',
  'ko-KR', 'pl-PL', 'pt-BR', 'ru-RU', 'sk-SK', 'vi-VN', 'zh-CN', 'zh-HK', 'zh-TW',
] as const
export type OfficeLocale = (typeof supportedOfficeLocales)[number]
export type OfficeTranslator = (key: string, fallback: string) => string

export const OfficeLocaleKey: InjectionKey<Readonly<Ref<OfficeLocale>>> = Symbol('OfficeLocale')
export const OfficeTranslatorKey: InjectionKey<OfficeTranslator> = Symbol('OfficeTranslator')

const appLabels: Record<OfficeLocale, Record<string, string>> = {
  'ar-SA': { 'app.language': 'اللغة', 'app.chinese': 'الصينية المبسطة', 'app.english': 'الإنجليزية' },
  'ca-ES': { 'app.language': 'Idioma', 'app.chinese': 'Xinès simplificat', 'app.english': 'Anglès' },
  'de-DE': { 'app.language': 'Sprache', 'app.chinese': 'Vereinfachtes Chinesisch', 'app.english': 'Englisch' },
  'en-US': { 'app.language': 'Language', 'app.chinese': 'Simplified Chinese', 'app.english': 'English' },
  'es-ES': { 'app.language': 'Idioma', 'app.chinese': 'Chino simplificado', 'app.english': 'Inglés' },
  'fa-IR': { 'app.language': 'زبان', 'app.chinese': 'چینی ساده‌شده', 'app.english': 'انگلیسی' },
  'fr-FR': { 'app.language': 'Langue', 'app.chinese': 'Chinois simplifié', 'app.english': 'Anglais' },
  'id-ID': { 'app.language': 'Bahasa', 'app.chinese': 'Tionghoa Sederhana', 'app.english': 'Inggris' },
  'it-IT': { 'app.language': 'Lingua', 'app.chinese': 'Cinese semplificato', 'app.english': 'Inglese' },
  'ja-JP': { 'app.language': '言語', 'app.chinese': '簡体字中国語', 'app.english': '英語' },
  'ko-KR': { 'app.language': '언어', 'app.chinese': '중국어 간체', 'app.english': '영어' },
  'pl-PL': { 'app.language': 'Język', 'app.chinese': 'Chiński uproszczony', 'app.english': 'Angielski' },
  'pt-BR': { 'app.language': 'Idioma', 'app.chinese': 'Chinês simplificado', 'app.english': 'Inglês' },
  'ru-RU': { 'app.language': 'Язык', 'app.chinese': 'Упрощённый китайский', 'app.english': 'Английский' },
  'sk-SK': { 'app.language': 'Jazyk', 'app.chinese': 'Zjednodušená čínština', 'app.english': 'Angličtina' },
  'vi-VN': { 'app.language': 'Ngôn ngữ', 'app.chinese': 'Tiếng Trung giản thể', 'app.english': 'Tiếng Anh' },
  'zh-CN': { 'app.language': '界面语言', 'app.chinese': '简体中文', 'app.english': '英语' },
  'zh-HK': { 'app.language': '介面語言', 'app.chinese': '簡體中文', 'app.english': '英文' },
  'zh-TW': { 'app.language': '介面語言', 'app.chinese': '簡體中文', 'app.english': '英文' },
}

const noteLabels: Partial<Record<OfficeLocale, Record<string, string>>> = {
  'ar-SA': { 'items.doc-footnote': 'حاشية سفلية', 'items.doc-footnote-insert': 'إدراج حاشية سفلية', 'items.doc-endnote-insert': 'إدراج حاشية ختامية', 'items.doc-notes-manage': 'إدارة الملاحظات', 'doc.notes.title': 'الحواشي السفلية والختامية', 'doc.notes.footnotes': 'حواشي سفلية', 'doc.notes.endnotes': 'حواشي ختامية', 'doc.notes.empty': 'لا توجد ملاحظات من هذا النوع في المستند.', 'doc.notes.delete': 'حذف الملاحظة' },
  'ca-ES': { 'items.doc-footnote': 'Nota al peu', 'items.doc-footnote-insert': 'Insereix una nota al peu', 'items.doc-endnote-insert': 'Insereix una nota final', 'items.doc-notes-manage': 'Gestiona les notes', 'doc.notes.title': 'Notes al peu i notes finals', 'doc.notes.footnotes': 'Notes al peu', 'doc.notes.endnotes': 'Notes finals', 'doc.notes.empty': 'El document no conté notes d’aquest tipus.', 'doc.notes.delete': 'Suprimeix la nota' },
  'de-DE': { 'items.doc-footnote': 'Fußnote', 'items.doc-footnote-insert': 'Fußnote einfügen', 'items.doc-endnote-insert': 'Endnote einfügen', 'items.doc-notes-manage': 'Fußnoten und Endnoten verwalten', 'doc.notes.title': 'Fußnoten und Endnoten', 'doc.notes.footnotes': 'Fußnoten', 'doc.notes.endnotes': 'Endnoten', 'doc.notes.empty': 'Das Dokument enthält keine Notizen dieses Typs.', 'doc.notes.delete': 'Notiz löschen' },
  'es-ES': { 'items.doc-footnote': 'Nota al pie', 'items.doc-footnote-insert': 'Insertar nota al pie', 'items.doc-endnote-insert': 'Insertar nota al final', 'items.doc-notes-manage': 'Gestionar notas', 'doc.notes.title': 'Notas al pie y notas finales', 'doc.notes.footnotes': 'Notas al pie', 'doc.notes.endnotes': 'Notas finales', 'doc.notes.empty': 'El documento no contiene notas de este tipo.', 'doc.notes.delete': 'Eliminar nota' },
  'fa-IR': { 'items.doc-footnote': 'پانویس', 'items.doc-footnote-insert': 'افزودن پانویس', 'items.doc-endnote-insert': 'افزودن پی‌نوشت', 'items.doc-notes-manage': 'مدیریت یادداشت‌ها', 'doc.notes.title': 'پانویس و پی‌نوشت', 'doc.notes.footnotes': 'پانویس‌ها', 'doc.notes.endnotes': 'پی‌نوشت‌ها', 'doc.notes.empty': 'این سند یادداشتی از این نوع ندارد.', 'doc.notes.delete': 'حذف یادداشت' },
  'fr-FR': { 'items.doc-footnote': 'Note de bas de page', 'items.doc-footnote-insert': 'Insérer une note de bas de page', 'items.doc-endnote-insert': 'Insérer une note de fin', 'items.doc-notes-manage': 'Gérer les notes', 'doc.notes.title': 'Notes de bas de page et notes de fin', 'doc.notes.footnotes': 'Notes de bas de page', 'doc.notes.endnotes': 'Notes de fin', 'doc.notes.empty': 'Le document ne contient aucune note de ce type.', 'doc.notes.delete': 'Supprimer la note' },
  'id-ID': { 'items.doc-footnote': 'Catatan kaki', 'items.doc-footnote-insert': 'Sisipkan catatan kaki', 'items.doc-endnote-insert': 'Sisipkan catatan akhir', 'items.doc-notes-manage': 'Kelola catatan', 'doc.notes.title': 'Catatan kaki dan catatan akhir', 'doc.notes.footnotes': 'Catatan kaki', 'doc.notes.endnotes': 'Catatan akhir', 'doc.notes.empty': 'Dokumen tidak memiliki catatan jenis ini.', 'doc.notes.delete': 'Hapus catatan' },
  'it-IT': { 'items.doc-footnote': 'Nota a piè di pagina', 'items.doc-footnote-insert': 'Inserisci nota a piè di pagina', 'items.doc-endnote-insert': 'Inserisci nota di chiusura', 'items.doc-notes-manage': 'Gestisci le note', 'doc.notes.title': 'Note a piè di pagina e note di chiusura', 'doc.notes.footnotes': 'Note a piè di pagina', 'doc.notes.endnotes': 'Note di chiusura', 'doc.notes.empty': 'Il documento non contiene note di questo tipo.', 'doc.notes.delete': 'Elimina nota' },
  'ja-JP': { 'items.doc-footnote': '脚注', 'items.doc-footnote-insert': '脚注を挿入', 'items.doc-endnote-insert': '文末脚注を挿入', 'items.doc-notes-manage': '脚注と文末脚注を管理', 'doc.notes.title': '脚注と文末脚注', 'doc.notes.footnotes': '脚注', 'doc.notes.endnotes': '文末脚注', 'doc.notes.empty': 'この種類の注釈はありません。', 'doc.notes.delete': '注釈を削除' },
  'ko-KR': { 'items.doc-footnote': '각주', 'items.doc-footnote-insert': '각주 삽입', 'items.doc-endnote-insert': '미주 삽입', 'items.doc-notes-manage': '각주 및 미주 관리', 'doc.notes.title': '각주 및 미주', 'doc.notes.footnotes': '각주', 'doc.notes.endnotes': '미주', 'doc.notes.empty': '문서에 이 유형의 주석이 없습니다.', 'doc.notes.delete': '주석 삭제' },
  'pl-PL': { 'items.doc-footnote': 'Przypis dolny', 'items.doc-footnote-insert': 'Wstaw przypis dolny', 'items.doc-endnote-insert': 'Wstaw przypis końcowy', 'items.doc-notes-manage': 'Zarządzaj przypisami', 'doc.notes.title': 'Przypisy dolne i końcowe', 'doc.notes.footnotes': 'Przypisy dolne', 'doc.notes.endnotes': 'Przypisy końcowe', 'doc.notes.empty': 'Dokument nie zawiera przypisów tego typu.', 'doc.notes.delete': 'Usuń przypis' },
  'pt-BR': { 'items.doc-footnote': 'Nota de rodapé', 'items.doc-footnote-insert': 'Inserir nota de rodapé', 'items.doc-endnote-insert': 'Inserir nota de fim', 'items.doc-notes-manage': 'Gerenciar notas', 'doc.notes.title': 'Notas de rodapé e notas de fim', 'doc.notes.footnotes': 'Notas de rodapé', 'doc.notes.endnotes': 'Notas de fim', 'doc.notes.empty': 'O documento não contém notas deste tipo.', 'doc.notes.delete': 'Excluir nota' },
  'ru-RU': { 'items.doc-footnote': 'Сноска', 'items.doc-footnote-insert': 'Вставить сноску', 'items.doc-endnote-insert': 'Вставить концевую сноску', 'items.doc-notes-manage': 'Управление сносками', 'doc.notes.title': 'Сноски и концевые сноски', 'doc.notes.footnotes': 'Сноски', 'doc.notes.endnotes': 'Концевые сноски', 'doc.notes.empty': 'В документе нет сносок этого типа.', 'doc.notes.delete': 'Удалить сноску' },
  'sk-SK': { 'items.doc-footnote': 'Poznámka pod čiarou', 'items.doc-footnote-insert': 'Vložiť poznámku pod čiarou', 'items.doc-endnote-insert': 'Vložiť koncovú poznámku', 'items.doc-notes-manage': 'Spravovať poznámky', 'doc.notes.title': 'Poznámky pod čiarou a koncové poznámky', 'doc.notes.footnotes': 'Poznámky pod čiarou', 'doc.notes.endnotes': 'Koncové poznámky', 'doc.notes.empty': 'Dokument neobsahuje poznámky tohto typu.', 'doc.notes.delete': 'Odstrániť poznámku' },
  'vi-VN': { 'items.doc-footnote': 'Chú thích cuối trang', 'items.doc-footnote-insert': 'Chèn chú thích cuối trang', 'items.doc-endnote-insert': 'Chèn chú thích cuối văn bản', 'items.doc-notes-manage': 'Quản lý chú thích', 'doc.notes.title': 'Chú thích cuối trang và cuối văn bản', 'doc.notes.footnotes': 'Chú thích cuối trang', 'doc.notes.endnotes': 'Chú thích cuối văn bản', 'doc.notes.empty': 'Tài liệu không có chú thích loại này.', 'doc.notes.delete': 'Xóa chú thích' },
  'zh-HK': { 'items.doc-footnote': '註腳', 'items.doc-footnote-insert': '插入註腳', 'items.doc-endnote-insert': '插入尾註', 'items.doc-notes-manage': '管理註腳和尾註', 'doc.notes.title': '註腳和尾註', 'doc.notes.footnotes': '註腳', 'doc.notes.endnotes': '尾註', 'doc.notes.empty': '目前文件沒有此類註釋。', 'doc.notes.delete': '刪除註釋' },
  'zh-TW': { 'items.doc-footnote': '註腳', 'items.doc-footnote-insert': '插入註腳', 'items.doc-endnote-insert': '插入尾註', 'items.doc-notes-manage': '管理註腳和尾註', 'doc.notes.title': '註腳和尾註', 'doc.notes.footnotes': '註腳', 'doc.notes.endnotes': '尾註', 'doc.notes.empty': '目前文件沒有此類註釋。', 'doc.notes.delete': '刪除註釋' },
}

const additionalFeatureLabels: Partial<Record<OfficeLocale, Record<string, string>>> = {
  'ar-SA': { ...sheetLocalearSA, ...docLocalearSA },
  'ca-ES': { ...sheetLocalecaES, ...docLocalecaES },
  'de-DE': { ...sheetLocaledeDE, ...docLocaledeDE },
  'es-ES': { ...sheetLocaleesES, ...docLocaleesES },
  'fa-IR': { ...sheetLocalefaIR, ...docLocalefaIR },
  'fr-FR': { ...sheetLocalefrFR, ...docLocalefrFR },
  'id-ID': { ...sheetLocaleidID, ...docLocaleidID },
  'it-IT': { ...sheetLocaleitIT, ...docLocaleitIT },
  'ja-JP': { ...sheetLocalejaJP, ...docLocalejaJP },
  'ko-KR': { ...sheetLocalekoKR, ...docLocalekoKR },
  'pl-PL': { ...sheetLocaleplPL, ...docLocaleplPL },
  'pt-BR': { ...sheetLocaleptBR, ...docLocaleptBR },
  'ru-RU': { ...sheetLocaleruRU, ...docLocaleruRU },
  'sk-SK': { ...sheetLocaleskSK, ...docLocaleskSK },
  'vi-VN': { ...sheetLocaleviVN, ...docLocaleviVN },
  'zh-HK': { ...sheetLocalezhHK, ...docLocalezhHK },
  'zh-TW': { ...sheetLocalezhTW, ...docLocalezhTW },
}

export function createOfficeMessages(): Record<OfficeLocale, Record<string, string>> {
  const messages = {} as Record<OfficeLocale, Record<string, string>>
  const englishFallback = { ...coreEnUS, ...sheetsEnUS, ...docsEnUS }
  for (const locale of supportedOfficeLocales) {
    messages[locale] = {
      ...englishFallback,
      ...additionalFeatureLabels[locale],
      ...noteLabels[locale],
      ...coreTranslations[locale],
      ...reviewedCoreTranslations[locale],
      ...sheetsTranslations[locale],
      ...reviewedSheetsTranslations[locale],
      ...conditionalFormattingLocaleMessages[locale],
      ...docsTranslations[locale],
      ...reviewedDocsTranslations[locale],
      ...(locale === 'zh-CN' ? { ...coreZhCN, ...sheetsZhCN, ...docsZhCN } : {}),
      ...appLabels[locale],
    }
  }
  return messages
}

export function createOfficeI18n() {
  const stored = typeof localStorage === 'undefined' ? null : localStorage.getItem('office-studio.locale')
  const locale: OfficeLocale = supportedOfficeLocales.find(candidate => candidate === stored) ?? 'zh-CN'
  return createI18n({
    legacy: false,
    locale,
    fallbackLocale: 'en-US',
    flatJson: true,
    messages: createOfficeMessages(),
  })
}

export function createOfficeTranslator(locale: OfficeLocale): OfficeTranslator {
  const messagesByLocale = createOfficeMessages()
  const messages = messagesByLocale[locale]
  const englishMessages = messagesByLocale['en-US']
  const chineseMessages = messagesByLocale['zh-CN']
  return (key, fallback) => {
    const normalized = key.replace(/^(sheet|doc|core)\./, '')
    return messages[key] ?? messages[normalized]
      ?? englishMessages[key] ?? englishMessages[normalized]
      ?? chineseMessages[key] ?? chineseMessages[normalized]
      ?? fallback
  }
}

export function createRibbonTranslator(locale: OfficeLocale, _namespace: 'sheet' | 'doc'): OfficeTranslator {
  return createOfficeTranslator(locale)
}

export function localizeRibbon(tabs: RibbonTab[], namespace: 'sheet' | 'doc', translate: OfficeTranslator): RibbonTab[] {
  const localized = (key: string, fallback: string) => translate(`${namespace}.${key}`, fallback)
  return tabs.map(tab => ({
    ...tab,
    title: localized(`tabs.${tab.id}`, tab.title),
    groups: tab.groups.map(group => ({
      ...group,
      title: localized(`groups.${group.id}`, group.title ?? group.id),
      items: group.items.map(localizeItem),
    })),
  }))

  function localizeItem(item: RibbonTab['groups'][number]['items'][number]): RibbonTab['groups'][number]['items'][number] {
    const children = item.children?.map(localizeItem)
    const options = item.options?.map(option => ({
      ...option,
      label: localized(`options.${namespace}.${item.id}.${String(option.value)}`, option.label),
    }))
    const title = localized(`items.${item.id}`, item.title)
    return {
      ...item,
      title,
      ...(item.tooltip ? { tooltip: localized(`items.${item.id}.tooltip`, item.tooltip === item.title ? title : item.tooltip) } : {}),
      ...(item.placeholder ? { placeholder: localized(`items.${item.id}.placeholder`, item.placeholder) } : {}),
      ...(item.disabledReason ? { disabledReason: localized(`items.${item.id}.disabledReason`, item.disabledReason) } : {}),
      ...(children ? { children } : {}),
      ...(options ? { options } : {}),
    }
  }
}
