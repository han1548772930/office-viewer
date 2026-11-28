import type { DocumentCharacterLocale, DocumentSymbolCategory } from '../../features/document/model/document-character-catalog'
import { documentCharacterPickerLabels } from '../../features/document/services/document-special-characters'
import { documentCharacterPickerLocales } from '../../i18n/docs/locales/document-character-picker'

interface AdditionalLabels { skinTone: string; additional: string; noSymbols: string; loadError: string; retry: string }
const additionalLabels: Record<DocumentCharacterLocale, AdditionalLabels> = {
  'ar-SA': { skinTone: 'لون البشرة', additional: 'المزيد من الرموز', noSymbols: 'لم يتم العثور على رموز', loadError: 'تعذر تحميل أسماء الرموز التعبيرية. أعد المحاولة.', retry: 'إعادة المحاولة' },
  'ca-ES': { skinTone: 'To de pell', additional: 'Més símbols', noSymbols: 'No s’han trobat símbols', loadError: 'No s’han pogut carregar els noms dels emojis. Torna-ho a provar.', retry: 'Torna-ho a provar' },
  'de-DE': { skinTone: 'Hautfarbe', additional: 'Weitere Symbole', noSymbols: 'Keine Symbole gefunden', loadError: 'Emoji-Namen konnten nicht geladen werden. Bitte erneut versuchen.', retry: 'Erneut versuchen' },
  'en-US': { skinTone: 'Skin tone', additional: 'More symbols', noSymbols: 'No symbols found', loadError: 'Emoji names could not be loaded. Try again.', retry: 'Try again' },
  'es-ES': { skinTone: 'Tono de piel', additional: 'Más símbolos', noSymbols: 'No se encontraron símbolos', loadError: 'No se pudieron cargar los nombres de los emojis. Inténtalo de nuevo.', retry: 'Reintentar' },
  'fa-IR': { skinTone: 'رنگ پوست', additional: 'نمادهای بیشتر', noSymbols: 'نمادی پیدا نشد', loadError: 'نام شکلک‌ها بارگیری نشد. دوباره تلاش کنید.', retry: 'تلاش دوباره' },
  'fr-FR': { skinTone: 'Teinte de peau', additional: 'Autres symboles', noSymbols: 'Aucun symbole trouvé', loadError: 'Impossible de charger les noms des emojis. Réessayez.', retry: 'Réessayer' },
  'id-ID': { skinTone: 'Warna kulit', additional: 'Simbol lainnya', noSymbols: 'Tidak ada simbol ditemukan', loadError: 'Nama emoji tidak dapat dimuat. Coba lagi.', retry: 'Coba lagi' },
  'it-IT': { skinTone: 'Tonalità della pelle', additional: 'Altri simboli', noSymbols: 'Nessun simbolo trovato', loadError: 'Impossibile caricare i nomi degli emoji. Riprova.', retry: 'Riprova' },
  'ja-JP': { skinTone: '肌の色', additional: 'その他の記号', noSymbols: '記号が見つかりません', loadError: '絵文字の名前を読み込めませんでした。再試行してください。', retry: '再試行' },
  'ko-KR': { skinTone: '피부색', additional: '추가 기호', noSymbols: '기호를 찾을 수 없습니다', loadError: '이모지 이름을 불러오지 못했습니다. 다시 시도하세요.', retry: '다시 시도' },
  'pl-PL': { skinTone: 'Odcień skóry', additional: 'Więcej symboli', noSymbols: 'Nie znaleziono symboli', loadError: 'Nie udało się wczytać nazw emoji. Spróbuj ponownie.', retry: 'Spróbuj ponownie' },
  'pt-BR': { skinTone: 'Tom de pele', additional: 'Mais símbolos', noSymbols: 'Nenhum símbolo encontrado', loadError: 'Não foi possível carregar os nomes dos emojis. Tente novamente.', retry: 'Tentar novamente' },
  'ru-RU': { skinTone: 'Тон кожи', additional: 'Другие символы', noSymbols: 'Символы не найдены', loadError: 'Не удалось загрузить названия эмодзи. Повторите попытку.', retry: 'Повторить' },
  'sk-SK': { skinTone: 'Odtieň pokožky', additional: 'Ďalšie symboly', noSymbols: 'Nenašli sa žiadne symboly', loadError: 'Názvy emoji sa nepodarilo načítať. Skúste to znova.', retry: 'Skúsiť znova' },
  'vi-VN': { skinTone: 'Màu da', additional: 'Ký hiệu khác', noSymbols: 'Không tìm thấy ký hiệu', loadError: 'Không thể tải tên biểu tượng cảm xúc. Hãy thử lại.', retry: 'Thử lại' },
  'zh-CN': { skinTone: '肤色', additional: '更多符号', noSymbols: '未找到符号', loadError: '表情名称加载失败，请重试。', retry: '重试' },
  'zh-HK': { skinTone: '膚色', additional: '更多符號', noSymbols: '找不到符號', loadError: '表情名稱載入失敗，請重試。', retry: '重試' },
  'zh-TW': { skinTone: '膚色', additional: '更多符號', noSymbols: '找不到符號', loadError: '表情名稱載入失敗，請重試。', retry: '重試' },
}

export function documentCharacterUiLabels(locale: DocumentCharacterLocale) {
  const fixed = documentCharacterPickerLabels(locale), existing = documentCharacterPickerLocales[locale], extra = additionalLabels[locale]
  const symbolGroups: Readonly<Record<DocumentSymbolCategory, string>> = { mathematics: fixed.mathematics, greek: fixed.greek, common: fixed.common,
    arrows: existing['doc.character-picker.category.arrows']!, punctuation: existing['doc.character-picker.category.punctuation']!,
    currency: existing['doc.character-picker.category.currency']!, additional: extra.additional }
  return { ...fixed, ...extra, symbolGroups, category: existing['doc.character-picker.category']!,
    custom: existing['doc.character-picker.custom-character']!, customPlaceholder: existing['doc.character-picker.custom-placeholder']!,
    insert: existing['doc.character-picker.insert']!, insertCharacter: existing['doc.character-picker.insert-character']!,
    list: { emoji: existing['doc.character-picker.emoji-list']!, symbol: existing['doc.character-picker.symbol-list']! } }
}
