import type { OfficeLocale } from '../../../i18n/regional-settings'

// Fixed v1.6.0 docs-ui/src/locale/*: docs-ui.heading labels.
const labels = {
  'ar-SA': ['طي', 'توسيع'], 'ca-ES': ['Redueix', 'Expandeix'], 'de-DE': ['Einklappen', 'Ausklappen'],
  'en-US': ['Collapse', 'Expand'], 'es-ES': ['Contraer', 'Expandir'], 'fa-IR': ['جمع کردن', 'باز کردن'],
  'fr-FR': ['Réduire', 'Développer'], 'id-ID': ['Ciutkan', 'Perluas'], 'it-IT': ['Comprimi', 'Espandi'],
  'ja-JP': ['折りたたむ', '展開する'], 'ko-KR': ['접기', '펼치기'], 'pl-PL': ['Zwiń', 'Rozwiń'],
  'pt-BR': ['Recolher', 'Expandir'], 'ru-RU': ['Свернуть', 'Развернуть'], 'sk-SK': ['Zbaliť', 'Rozbaliť'],
  'vi-VN': ['Thu gọn', 'Mở rộng'], 'zh-CN': ['收起', '展开'], 'zh-HK': ['收合', '展開'], 'zh-TW': ['收合', '展開'],
} satisfies Record<OfficeLocale, readonly [string, string]>

export function documentHeadingCollapseMessages(locale: OfficeLocale) {
  return { 'document.heading.collapse': labels[locale][0], 'document.heading.expand': labels[locale][1] }
}
