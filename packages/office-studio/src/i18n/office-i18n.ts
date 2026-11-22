import { computed, inject, provide, ref, type InjectionKey, type Ref } from 'vue'
import { supportedOfficeLocales, type OfficeLocale, type OfficeRegionalSettings } from './regional-settings.ts'
export { supportedOfficeLocales, type OfficeLocale } from './regional-settings.ts'
import { createI18n } from 'vue-i18n'
import type { RibbonTab } from '../platform/ribbon/ribbon-schema.ts'
import { coreEnUS } from './core/locales/en-US.ts'
import { coreZhCN } from './core/locales/zh-CN.ts'
import { coreTranslations } from './core/locales/generated.ts'
import { reviewedCoreTranslations } from './core/locales/reviewed.ts'
import { docsEnUS } from './docs/locales/en-US.ts'
import { documentTocLocales } from '../features/document/toc/locales.ts'
import { docsZhCN } from './docs/locales/zh-CN.ts'
import { docsTranslations } from './docs/locales/generated.ts'
import { reviewedDocsTranslations } from './docs/locales/reviewed.ts'
import { docPermissionLocaleMessages } from './docs/locales/permissions.ts'
import { docToolbarColorLocales } from './docs/locales/toolbar-colors.ts'
import { documentSettingsLocaleMessages } from './docs/locales/document-settings.ts'
import { documentCommentLocaleMessages } from './docs/locales/document-comments.ts'
import { documentTableInsertLocales } from './docs/locales/document-table-insert.ts'
import { sheetsEnUS } from './sheets/locales/en-US.ts'
import { sheetsZhCN } from './sheets/locales/zh-CN.ts'
import { sheetNumberFormatLocales } from './sheets/locales/number-format.ts'
import { sheetToolbarBorderApiStyleLocales, sheetToolbarColorLocales } from './sheets/locales/toolbar-colors.ts'
import { sheetsTranslations } from './sheets/locales/generated.ts'
import { reviewedSheetsTranslations } from './sheets/locales/reviewed.ts'
import { sheetFindReplaceLocales } from './sheets/locales/find-replace.ts'
import { sheetFormulaClipboardLocales } from './sheets/locales/formula-clipboard.ts'
import { sheetQuickValidationLocales } from './sheets/locales/quick-validation.ts'
import { sheetTabContextLocales } from './sheets/locales/sheet-tab-context.ts'
import { sheetFilterMenuLocales } from './sheets/locales/filter-menu.ts'
import { conditionalFormattingLocaleMessages } from './sheets/locales/conditional-formatting.ts'
import { sheetCrosshairHighlightLocales } from './sheets/locales/crosshair-highlight.ts'
import { sheetLocalearSA, sheetLocaleesES, sheetLocalecaES, sheetLocaledeDE, sheetLocalefaIR, sheetLocalefrFR, sheetLocaleidID, sheetLocaleitIT, sheetLocalejaJP, sheetLocalekoKR, sheetLocaleplPL, sheetLocaleptBR, sheetLocaleruRU, sheetLocaleskSK, sheetLocaleviVN, sheetLocalezhHK, sheetLocalezhTW } from './sheets/locales/workbench.ts'
import { docLocalearSA, docLocaleesES, docLocalecaES, docLocaledeDE, docLocalefaIR, docLocalefrFR, docLocaleidID, docLocaleitIT, docLocalejaJP, docLocalekoKR, docLocaleplPL, docLocaleptBR, docLocaleruRU, docLocaleskSK, docLocaleviVN, docLocalezhHK, docLocalezhTW } from './docs/locales/workbench.ts'

export type OfficeTranslator = (key: string, fallback: string) => string

export const OfficeLocaleKey: InjectionKey<Readonly<Ref<OfficeLocale>>> = Symbol('OfficeLocale')
export const OfficeRegionKey: InjectionKey<Readonly<Ref<OfficeLocale>>> = Symbol('OfficeRegion')
export function useOfficeRegion() {
  return inject(OfficeRegionKey, undefined) ?? inject(OfficeLocaleKey, ref<OfficeLocale>('zh-CN'))
}
export function provideWorkbenchRegionalSettings(settings: () => OfficeRegionalSettings | undefined) {
  const hostLocale = inject(OfficeLocaleKey, undefined), hostRegion = inject(OfficeRegionKey, undefined)
  const hostDirection = inject(OfficeDirectionKey, undefined)
  const locale = computed(() => hostLocale?.value ?? settings()?.locale.value ?? 'zh-CN')
  const region = computed(() => hostRegion?.value ?? (settings()?.regionPreference.value === 'auto' ? locale.value : settings()?.region.value ?? locale.value))
  const direction = computed(() => hostDirection?.value ?? (locale.value === 'ar-SA' || locale.value === 'fa-IR' ? 'rtl' : 'ltr'))
  const translator = computed(() => createOfficeTranslator(locale.value))
  provide(OfficeLocaleKey, locale)
  provide(OfficeRegionKey, region)
  provide(OfficeDirectionKey, direction)
  provide(OfficeTranslatorKey, (key, fallback) => translator.value(key, fallback))
  return { locale, region, direction }
}
export function provideWorkbenchLocales(settings: () => OfficeRegionalSettings | undefined) {
  return provideWorkbenchRegionalSettings(settings).locale
}
export const OfficeDirectionKey: InjectionKey<Readonly<Ref<'ltr' | 'rtl'>>> = Symbol('OfficeDirection')
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

const documentSearchLabels: Record<OfficeLocale, Record<string, string>> = {
  'ar-SA': { 'document.search.find-placeholder': 'بحث في المستند', 'document.search.replace-placeholder': 'استبدال بـ', 'document.search.previous': 'التطابق السابق', 'document.search.next': 'التطابق التالي', 'document.search.close': 'إغلاق البحث', 'document.search.replace': 'استبدال', 'document.search.replace-all': 'استبدال الكل', 'document.search.case-sensitive': 'مطابقة حالة الأحرف', 'document.search.whole-word': 'كلمات كاملة' },
  'ca-ES': { 'document.search.find-placeholder': 'Cerca al document', 'document.search.replace-placeholder': 'Substitueix per', 'document.search.previous': 'Coincidència anterior', 'document.search.next': 'Coincidència següent', 'document.search.close': 'Tanca la cerca', 'document.search.replace': 'Substitueix', 'document.search.replace-all': 'Substitueix-ho tot', 'document.search.case-sensitive': 'Distingeix majúscules', 'document.search.whole-word': 'Paraules completes' },
  'de-DE': { 'document.search.find-placeholder': 'Im Dokument suchen', 'document.search.replace-placeholder': 'Ersetzen durch', 'document.search.previous': 'Vorheriger Treffer', 'document.search.next': 'Nächster Treffer', 'document.search.close': 'Suche schließen', 'document.search.replace': 'Ersetzen', 'document.search.replace-all': 'Alle ersetzen', 'document.search.case-sensitive': 'Groß-/Kleinschreibung beachten', 'document.search.whole-word': 'Ganze Wörter' },
  'en-US': { 'document.search.find-placeholder': 'Find in document', 'document.search.replace-placeholder': 'Replace with', 'document.search.previous': 'Previous match', 'document.search.next': 'Next match', 'document.search.close': 'Close find', 'document.search.replace': 'Replace', 'document.search.replace-all': 'Replace all', 'document.search.case-sensitive': 'Match case', 'document.search.whole-word': 'Whole words' },
  'es-ES': { 'document.search.find-placeholder': 'Buscar en el documento', 'document.search.replace-placeholder': 'Reemplazar por', 'document.search.previous': 'Coincidencia anterior', 'document.search.next': 'Coincidencia siguiente', 'document.search.close': 'Cerrar búsqueda', 'document.search.replace': 'Reemplazar', 'document.search.replace-all': 'Reemplazar todo', 'document.search.case-sensitive': 'Distinguir mayúsculas', 'document.search.whole-word': 'Palabras completas' },
  'fa-IR': { 'document.search.find-placeholder': 'جستجو در سند', 'document.search.replace-placeholder': 'جایگزینی با', 'document.search.previous': 'نتیجه قبلی', 'document.search.next': 'نتیجه بعدی', 'document.search.close': 'بستن جستجو', 'document.search.replace': 'جایگزینی', 'document.search.replace-all': 'جایگزینی همه', 'document.search.case-sensitive': 'تطبیق بزرگی حروف', 'document.search.whole-word': 'کلمات کامل' },
  'fr-FR': { 'document.search.find-placeholder': 'Rechercher dans le document', 'document.search.replace-placeholder': 'Remplacer par', 'document.search.previous': 'Résultat précédent', 'document.search.next': 'Résultat suivant', 'document.search.close': 'Fermer la recherche', 'document.search.replace': 'Remplacer', 'document.search.replace-all': 'Tout remplacer', 'document.search.case-sensitive': 'Respecter la casse', 'document.search.whole-word': 'Mots entiers' },
  'id-ID': { 'document.search.find-placeholder': 'Cari dalam dokumen', 'document.search.replace-placeholder': 'Ganti dengan', 'document.search.previous': 'Hasil sebelumnya', 'document.search.next': 'Hasil berikutnya', 'document.search.close': 'Tutup pencarian', 'document.search.replace': 'Ganti', 'document.search.replace-all': 'Ganti semua', 'document.search.case-sensitive': 'Cocokkan huruf besar-kecil', 'document.search.whole-word': 'Kata utuh' },
  'it-IT': { 'document.search.find-placeholder': 'Cerca nel documento', 'document.search.replace-placeholder': 'Sostituisci con', 'document.search.previous': 'Risultato precedente', 'document.search.next': 'Risultato successivo', 'document.search.close': 'Chiudi ricerca', 'document.search.replace': 'Sostituisci', 'document.search.replace-all': 'Sostituisci tutto', 'document.search.case-sensitive': 'Maiuscole/minuscole', 'document.search.whole-word': 'Parole intere' },
  'ja-JP': { 'document.search.find-placeholder': 'ドキュメント内を検索', 'document.search.replace-placeholder': '置換後の文字列', 'document.search.previous': '前の一致', 'document.search.next': '次の一致', 'document.search.close': '検索を閉じる', 'document.search.replace': '置換', 'document.search.replace-all': 'すべて置換', 'document.search.case-sensitive': '大文字と小文字を区別', 'document.search.whole-word': '単語単位' },
  'ko-KR': { 'document.search.find-placeholder': '문서에서 찾기', 'document.search.replace-placeholder': '바꿀 내용', 'document.search.previous': '이전 일치 항목', 'document.search.next': '다음 일치 항목', 'document.search.close': '찾기 닫기', 'document.search.replace': '바꾸기', 'document.search.replace-all': '모두 바꾸기', 'document.search.case-sensitive': '대소문자 구분', 'document.search.whole-word': '전체 단어' },
  'pl-PL': { 'document.search.find-placeholder': 'Znajdź w dokumencie', 'document.search.replace-placeholder': 'Zamień na', 'document.search.previous': 'Poprzednie dopasowanie', 'document.search.next': 'Następne dopasowanie', 'document.search.close': 'Zamknij wyszukiwanie', 'document.search.replace': 'Zamień', 'document.search.replace-all': 'Zamień wszystko', 'document.search.case-sensitive': 'Uwzględnij wielkość liter', 'document.search.whole-word': 'Całe wyrazy' },
  'pt-BR': { 'document.search.find-placeholder': 'Localizar no documento', 'document.search.replace-placeholder': 'Substituir por', 'document.search.previous': 'Correspondência anterior', 'document.search.next': 'Próxima correspondência', 'document.search.close': 'Fechar localização', 'document.search.replace': 'Substituir', 'document.search.replace-all': 'Substituir tudo', 'document.search.case-sensitive': 'Diferenciar maiúsculas', 'document.search.whole-word': 'Palavras inteiras' },
  'ru-RU': { 'document.search.find-placeholder': 'Найти в документе', 'document.search.replace-placeholder': 'Заменить на', 'document.search.previous': 'Предыдущее совпадение', 'document.search.next': 'Следующее совпадение', 'document.search.close': 'Закрыть поиск', 'document.search.replace': 'Заменить', 'document.search.replace-all': 'Заменить все', 'document.search.case-sensitive': 'Учитывать регистр', 'document.search.whole-word': 'Слова целиком' },
  'sk-SK': { 'document.search.find-placeholder': 'Hľadať v dokumente', 'document.search.replace-placeholder': 'Nahradiť čím', 'document.search.previous': 'Predchádzajúca zhoda', 'document.search.next': 'Nasledujúca zhoda', 'document.search.close': 'Zavrieť hľadanie', 'document.search.replace': 'Nahradiť', 'document.search.replace-all': 'Nahradiť všetko', 'document.search.case-sensitive': 'Rozlišovať veľkosť písmen', 'document.search.whole-word': 'Celé slová' },
  'vi-VN': { 'document.search.find-placeholder': 'Tìm trong tài liệu', 'document.search.replace-placeholder': 'Thay thế bằng', 'document.search.previous': 'Kết quả trước', 'document.search.next': 'Kết quả tiếp theo', 'document.search.close': 'Đóng tìm kiếm', 'document.search.replace': 'Thay thế', 'document.search.replace-all': 'Thay thế tất cả', 'document.search.case-sensitive': 'Phân biệt hoa thường', 'document.search.whole-word': 'Từ hoàn chỉnh' },
  'zh-CN': { 'document.search.find-placeholder': '在文档中查找', 'document.search.replace-placeholder': '替换为', 'document.search.previous': '上一个匹配项', 'document.search.next': '下一个匹配项', 'document.search.close': '关闭查找', 'document.search.replace': '替换', 'document.search.replace-all': '全部替换', 'document.search.case-sensitive': '区分大小写', 'document.search.whole-word': '全字匹配' },
  'zh-HK': { 'document.search.find-placeholder': '在文件中尋找', 'document.search.replace-placeholder': '取代為', 'document.search.previous': '上一個相符項目', 'document.search.next': '下一個相符項目', 'document.search.close': '關閉尋找', 'document.search.replace': '取代', 'document.search.replace-all': '全部取代', 'document.search.case-sensitive': '區分大小寫', 'document.search.whole-word': '全字符合' },
  'zh-TW': { 'document.search.find-placeholder': '在文件中尋找', 'document.search.replace-placeholder': '取代為', 'document.search.previous': '上一個相符項目', 'document.search.next': '下一個相符項目', 'document.search.close': '關閉尋找', 'document.search.replace': '取代', 'document.search.replace-all': '全部取代', 'document.search.case-sensitive': '區分大小寫', 'document.search.whole-word': '全字符合' },
}

const documentSearchScopeLabels: Record<OfficeLocale, string> = {
  'ar-SA': 'البحث في التحديد', 'ca-ES': 'Cerca només la selecció', 'de-DE': 'Nur Auswahl', 'en-US': 'Selection only',
  'es-ES': 'Solo en la selección', 'fa-IR': 'فقط در گزیده', 'fr-FR': 'Sélection uniquement', 'id-ID': 'Hanya pilihan',
  'it-IT': 'Solo nella selezione', 'ja-JP': '選択範囲のみ', 'ko-KR': '선택 영역만', 'pl-PL': 'Tylko zaznaczenie',
  'pt-BR': 'Somente na seleção', 'ru-RU': 'Только выделение', 'sk-SK': 'Iba vo výbere', 'vi-VN': 'Chỉ trong vùng chọn',
  'zh-CN': '仅在选区中查找', 'zh-HK': '只在選取範圍尋找', 'zh-TW': '只在選取範圍尋找',
}

const paragraphStyleLabels: Record<OfficeLocale, Record<string, string>> = {
  'ar-SA': { 'options.doc.doc-block-style.normal': 'عادي', 'options.doc.doc-block-style.title': 'العنوان', 'options.doc.doc-block-style.subtitle': 'العنوان الفرعي', 'options.doc.doc-block-style.heading1': 'عنوان 1', 'options.doc.doc-block-style.heading2': 'عنوان 2', 'options.doc.doc-block-style.heading3': 'عنوان 3', 'options.doc.doc-block-style.heading4': 'عنوان 4', 'options.doc.doc-block-style.heading5': 'عنوان 5', 'options.doc.doc-block-style.H6': 'عنوان 6' },
  'ca-ES': { 'options.doc.doc-block-style.normal': 'Cos', 'options.doc.doc-block-style.title': 'Títol', 'options.doc.doc-block-style.subtitle': 'Subtítol', 'options.doc.doc-block-style.heading1': 'Encapçalament 1', 'options.doc.doc-block-style.heading2': 'Encapçalament 2', 'options.doc.doc-block-style.heading3': 'Encapçalament 3', 'options.doc.doc-block-style.heading4': 'Encapçalament 4', 'options.doc.doc-block-style.heading5': 'Encapçalament 5', 'options.doc.doc-block-style.H6': 'Encapçalament 6' },
  'de-DE': { 'options.doc.doc-block-style.normal': 'Standard', 'options.doc.doc-block-style.title': 'Titel', 'options.doc.doc-block-style.subtitle': 'Untertitel', 'options.doc.doc-block-style.heading1': 'Überschrift 1', 'options.doc.doc-block-style.heading2': 'Überschrift 2', 'options.doc.doc-block-style.heading3': 'Überschrift 3', 'options.doc.doc-block-style.heading4': 'Überschrift 4', 'options.doc.doc-block-style.heading5': 'Überschrift 5', 'options.doc.doc-block-style.H6': 'Überschrift 6' },
  'en-US': { 'options.doc.doc-block-style.normal': 'Normal', 'options.doc.doc-block-style.title': 'Title', 'options.doc.doc-block-style.subtitle': 'Subtitle', 'options.doc.doc-block-style.heading1': 'Heading 1', 'options.doc.doc-block-style.heading2': 'Heading 2', 'options.doc.doc-block-style.heading3': 'Heading 3', 'options.doc.doc-block-style.heading4': 'Heading 4', 'options.doc.doc-block-style.heading5': 'Heading 5', 'options.doc.doc-block-style.H6': 'Heading 6' },
  'es-ES': { 'options.doc.doc-block-style.normal': 'Estilo normal', 'options.doc.doc-block-style.title': 'Título', 'options.doc.doc-block-style.subtitle': 'Subtítulo', 'options.doc.doc-block-style.heading1': 'Encabezado 1', 'options.doc.doc-block-style.heading2': 'Encabezado 2', 'options.doc.doc-block-style.heading3': 'Encabezado 3', 'options.doc.doc-block-style.heading4': 'Encabezado 4', 'options.doc.doc-block-style.heading5': 'Encabezado 5', 'options.doc.doc-block-style.H6': 'Encabezado 6' },
  'fa-IR': { 'options.doc.doc-block-style.normal': 'عادی', 'options.doc.doc-block-style.title': 'عنوان', 'options.doc.doc-block-style.subtitle': 'زیرعنوان', 'options.doc.doc-block-style.heading1': 'سرفصل ۱', 'options.doc.doc-block-style.heading2': 'سرفصل ۲', 'options.doc.doc-block-style.heading3': 'سرفصل ۳', 'options.doc.doc-block-style.heading4': 'سرفصل ۴', 'options.doc.doc-block-style.heading5': 'سرفصل ۵', 'options.doc.doc-block-style.H6': 'سرفصل ۶' },
  'fr-FR': { 'options.doc.doc-block-style.normal': 'Style normal', 'options.doc.doc-block-style.title': 'Titre', 'options.doc.doc-block-style.subtitle': 'Sous-titre', 'options.doc.doc-block-style.heading1': 'Titre 1', 'options.doc.doc-block-style.heading2': 'Titre 2', 'options.doc.doc-block-style.heading3': 'Titre 3', 'options.doc.doc-block-style.heading4': 'Titre 4', 'options.doc.doc-block-style.heading5': 'Titre 5', 'options.doc.doc-block-style.H6': 'Titre 6' },
  'id-ID': { 'options.doc.doc-block-style.normal': 'Gaya normal', 'options.doc.doc-block-style.title': 'Judul', 'options.doc.doc-block-style.subtitle': 'Subjudul', 'options.doc.doc-block-style.heading1': 'Tajuk 1', 'options.doc.doc-block-style.heading2': 'Tajuk 2', 'options.doc.doc-block-style.heading3': 'Tajuk 3', 'options.doc.doc-block-style.heading4': 'Tajuk 4', 'options.doc.doc-block-style.heading5': 'Tajuk 5', 'options.doc.doc-block-style.H6': 'Tajuk 6' },
  'it-IT': { 'options.doc.doc-block-style.normal': 'Normale', 'options.doc.doc-block-style.title': 'Titolo', 'options.doc.doc-block-style.subtitle': 'Sottotitolo', 'options.doc.doc-block-style.heading1': 'Titolo 1', 'options.doc.doc-block-style.heading2': 'Titolo 2', 'options.doc.doc-block-style.heading3': 'Titolo 3', 'options.doc.doc-block-style.heading4': 'Titolo 4', 'options.doc.doc-block-style.heading5': 'Titolo 5', 'options.doc.doc-block-style.H6': 'Titolo 6' },
  'ja-JP': { 'options.doc.doc-block-style.normal': '標準', 'options.doc.doc-block-style.title': 'タイトル', 'options.doc.doc-block-style.subtitle': 'サブタイトル', 'options.doc.doc-block-style.heading1': '見出し 1', 'options.doc.doc-block-style.heading2': '見出し 2', 'options.doc.doc-block-style.heading3': '見出し 3', 'options.doc.doc-block-style.heading4': '見出し 4', 'options.doc.doc-block-style.heading5': '見出し 5', 'options.doc.doc-block-style.H6': '見出し 6' },
  'ko-KR': { 'options.doc.doc-block-style.normal': '보통', 'options.doc.doc-block-style.title': '제목', 'options.doc.doc-block-style.subtitle': '부제목', 'options.doc.doc-block-style.heading1': '제목 1', 'options.doc.doc-block-style.heading2': '제목 2', 'options.doc.doc-block-style.heading3': '제목 3', 'options.doc.doc-block-style.heading4': '제목 4', 'options.doc.doc-block-style.heading5': '제목 5', 'options.doc.doc-block-style.H6': '제목 6' },
  'pl-PL': { 'options.doc.doc-block-style.normal': 'Normalny', 'options.doc.doc-block-style.title': 'Tytuł', 'options.doc.doc-block-style.subtitle': 'Podtytuł', 'options.doc.doc-block-style.heading1': 'Nagłówek 1', 'options.doc.doc-block-style.heading2': 'Nagłówek 2', 'options.doc.doc-block-style.heading3': 'Nagłówek 3', 'options.doc.doc-block-style.heading4': 'Nagłówek 4', 'options.doc.doc-block-style.heading5': 'Nagłówek 5', 'options.doc.doc-block-style.H6': 'Nagłówek 6' },
  'pt-BR': { 'options.doc.doc-block-style.normal': 'Estilo normal', 'options.doc.doc-block-style.title': 'Título', 'options.doc.doc-block-style.subtitle': 'Subtítulo', 'options.doc.doc-block-style.heading1': 'Título 1', 'options.doc.doc-block-style.heading2': 'Título 2', 'options.doc.doc-block-style.heading3': 'Título 3', 'options.doc.doc-block-style.heading4': 'Título 4', 'options.doc.doc-block-style.heading5': 'Título 5', 'options.doc.doc-block-style.H6': 'Título 6' },
  'ru-RU': { 'options.doc.doc-block-style.normal': 'Обычный', 'options.doc.doc-block-style.title': 'Название', 'options.doc.doc-block-style.subtitle': 'Подзаголовок', 'options.doc.doc-block-style.heading1': 'Заголовок 1', 'options.doc.doc-block-style.heading2': 'Заголовок 2', 'options.doc.doc-block-style.heading3': 'Заголовок 3', 'options.doc.doc-block-style.heading4': 'Заголовок 4', 'options.doc.doc-block-style.heading5': 'Заголовок 5', 'options.doc.doc-block-style.H6': 'Заголовок 6' },
  'sk-SK': { 'options.doc.doc-block-style.normal': 'Normálne', 'options.doc.doc-block-style.title': 'Názov', 'options.doc.doc-block-style.subtitle': 'Podnadpis', 'options.doc.doc-block-style.heading1': 'Nadpis 1', 'options.doc.doc-block-style.heading2': 'Nadpis 2', 'options.doc.doc-block-style.heading3': 'Nadpis 3', 'options.doc.doc-block-style.heading4': 'Nadpis 4', 'options.doc.doc-block-style.heading5': 'Nadpis 5', 'options.doc.doc-block-style.H6': 'Nadpis 6' },
  'vi-VN': { 'options.doc.doc-block-style.normal': 'Bình thường', 'options.doc.doc-block-style.title': 'Tiêu đề', 'options.doc.doc-block-style.subtitle': 'Phụ đề', 'options.doc.doc-block-style.heading1': 'Đầu đề 1', 'options.doc.doc-block-style.heading2': 'Đầu đề 2', 'options.doc.doc-block-style.heading3': 'Đầu đề 3', 'options.doc.doc-block-style.heading4': 'Đầu đề 4', 'options.doc.doc-block-style.heading5': 'Đầu đề 5', 'options.doc.doc-block-style.H6': 'Đầu đề 6' },
  'zh-CN': { 'options.doc.doc-block-style.normal': '正文', 'options.doc.doc-block-style.title': '标题', 'options.doc.doc-block-style.subtitle': '副标题', 'options.doc.doc-block-style.heading1': '标题 1', 'options.doc.doc-block-style.heading2': '标题 2', 'options.doc.doc-block-style.heading3': '标题 3', 'options.doc.doc-block-style.heading4': '标题 4', 'options.doc.doc-block-style.heading5': '标题 5', 'options.doc.doc-block-style.H6': '标题 6' },
  'zh-HK': { 'options.doc.doc-block-style.normal': '正文', 'options.doc.doc-block-style.title': '標題', 'options.doc.doc-block-style.subtitle': '副標題', 'options.doc.doc-block-style.heading1': '標題 1', 'options.doc.doc-block-style.heading2': '標題 2', 'options.doc.doc-block-style.heading3': '標題 3', 'options.doc.doc-block-style.heading4': '標題 4', 'options.doc.doc-block-style.heading5': '標題 5', 'options.doc.doc-block-style.H6': '標題 6' },
  'zh-TW': { 'options.doc.doc-block-style.normal': '正文', 'options.doc.doc-block-style.title': '標題', 'options.doc.doc-block-style.subtitle': '副標題', 'options.doc.doc-block-style.heading1': '標題 1', 'options.doc.doc-block-style.heading2': '標題 2', 'options.doc.doc-block-style.heading3': '標題 3', 'options.doc.doc-block-style.heading4': '標題 4', 'options.doc.doc-block-style.heading5': '標題 5', 'options.doc.doc-block-style.H6': '標題 6' },
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

const sheetRibbonRuleGroupLabels: Record<OfficeLocale, string> = {
  'ar-SA': 'القواعد', 'ca-ES': 'Regles', 'de-DE': 'Regeln', 'en-US': 'Rules', 'es-ES': 'Reglas',
  'fa-IR': 'قواعد', 'fr-FR': 'Règles', 'id-ID': 'Aturan', 'it-IT': 'Regole', 'ja-JP': 'ルール',
  'ko-KR': '규칙', 'pl-PL': 'Reguły', 'pt-BR': 'Regras', 'ru-RU': 'Правила', 'sk-SK': 'Pravidlá',
  'vi-VN': 'Quy tắc', 'zh-CN': '规则', 'zh-HK': '規則', 'zh-TW': '規則',
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
      ...sheetFindReplaceLocales[locale],
      ...sheetFormulaClipboardLocales[locale],
      ...sheetQuickValidationLocales[locale],
      ...sheetTabContextLocales[locale],
      ...sheetFilterMenuLocales[locale],
      'groups.rules': sheetRibbonRuleGroupLabels[locale],
      ...conditionalFormattingLocaleMessages[locale],
      ...sheetCrosshairHighlightLocales[locale],
      ...sheetNumberFormatLocales[locale],
      ...sheetToolbarColorLocales[locale],
      ...sheetToolbarBorderApiStyleLocales[locale],
      ...docsTranslations[locale],
      ...documentTocLocales[locale as keyof typeof documentTocLocales],
      'items.doc-insert-table-of-contents': documentTocLocales[locale]['document.toc.insert'],
      'items.doc-insert-table-of-contents-automatic': documentTocLocales[locale]['document.toc.automatic'],
      'items.doc-insert-table-of-contents-custom': documentTocLocales[locale]['document.toc.custom'],
      'items.doc-update-table-of-contents': documentTocLocales[locale]['document.toc.update'],
      'items.doc-delete-table-of-contents': documentTocLocales[locale]['document.toc.delete'],
      'groups.table-of-contents-manage': documentTocLocales[locale]['document.toc.manage'],
      ...reviewedDocsTranslations[locale],
      ...docPermissionLocaleMessages[locale],
      ...docToolbarColorLocales[locale],
      ...documentSettingsLocaleMessages[locale],
      ...documentCommentLocaleMessages[locale],
      ...documentTableInsertLocales[locale],
      'groups.doc-access': locale === 'zh-CN' ? '权限' : locale === 'en-US' ? 'Access' : docPermissionLocaleMessages[locale]!['doc.permissions.scope.document']!,
      'items.doc-permissions': locale === 'zh-CN' ? '文档权限' : locale === 'en-US' ? 'Document Permissions' : docPermissionLocaleMessages[locale]!['doc.permissions.title']!,
      ...(locale === 'zh-CN' ? { ...coreZhCN, ...sheetsZhCN, ...docsZhCN } : {}),
      ...documentSearchLabels[locale],
      'document.search.selection-only': documentSearchScopeLabels[locale],
      ...paragraphStyleLabels[locale],
      ...Object.fromEntries([1, 2, 3, 4, 5].map(level => [
        `options.doc.doc-block-style.H${level}`,
        paragraphStyleLabels[locale][`options.doc.doc-block-style.heading${level}`],
      ])),
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
