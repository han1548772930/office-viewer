import type { OfficeLocale } from '../../regional-settings.ts'

const colorKeys = [
  'reset-text', 'no-fill', 'automatic', 'unset-border', 'mixed', 'options',
  'more', 'cancel', 'confirm', 'standard', 'value', 'format', 'invalid',
] as const
type ColorLabels = readonly [string, string, string, string, string, string, string, string, string, string, string, string, string]

/** Sheets uses an opaque color palette, a separate custom draft and explicit reset commands. */
const labels: Record<OfficeLocale, ColorLabels> = {
  'ar-SA': ['إعادة تعيين لون النص', 'بلا تعبئة', 'لون نص تلقائي', 'لون الحدود غير محدد', 'مختلط', 'خيارات {title}', 'المزيد من الألوان', 'إلغاء', 'موافق', 'ألوان قياسية', 'قيمة اللون', 'تنسيق اللون', 'أدخل قيمة لون صالحة'],
  'ca-ES': ['Restableix el color del text', 'Sense emplenament', 'Color de text automàtic', 'Color de la vora sense definir', 'Mixt', 'Opcions de {title}', 'Més colors', 'Cancel·la', 'D’acord', 'Colors estàndard', 'Valor del color', 'Format del color', 'Introduïu un valor de color vàlid'],
  'de-DE': ['Textfarbe zurücksetzen', 'Keine Füllung', 'Automatische Textfarbe', 'Keine Rahmenfarbe festgelegt', 'Gemischt', '{title}: Optionen', 'Mehr Farben', 'Abbrechen', 'OK', 'Standardfarben', 'Farbwert', 'Farbformat', 'Geben Sie einen gültigen Farbwert ein'],
  'en-US': ['Reset Text Color', 'No Fill', 'Automatic Text Color', 'Border Color Not Set', 'Mixed', '{title} Options', 'More Colors', 'Cancel', 'OK', 'Standard Colors', 'Color Value', 'Color Format', 'Enter a valid color value'],
  'es-ES': ['Restablecer color de texto', 'Sin relleno', 'Color de texto automático', 'Color del borde sin definir', 'Mixto', 'Opciones de {title}', 'Más colores', 'Cancelar', 'Aceptar', 'Colores estándar', 'Valor de color', 'Formato de color', 'Introduce un valor de color válido'],
  'fa-IR': ['بازنشانی رنگ متن', 'بدون رنگ زمینه', 'رنگ متن خودکار', 'رنگ کادر تنظیم نشده', 'ترکیبی', 'گزینه‌های {title}', 'رنگ‌های بیشتر', 'لغو', 'تأیید', 'رنگ‌های استاندارد', 'مقدار رنگ', 'قالب رنگ', 'یک مقدار رنگ معتبر وارد کنید'],
  'fr-FR': ['Réinitialiser la couleur du texte', 'Aucun remplissage', 'Couleur du texte automatique', 'Couleur de bordure non définie', 'Mixte', 'Options de {title}', 'Autres couleurs', 'Annuler', 'OK', 'Couleurs standard', 'Valeur de couleur', 'Format de couleur', 'Saisissez une valeur de couleur valide'],
  'id-ID': ['Atur Ulang Warna Teks', 'Tanpa Isian', 'Warna Teks Otomatis', 'Warna Batas Belum Diatur', 'Campuran', 'Opsi {title}', 'Warna Lainnya', 'Batal', 'OK', 'Warna Standar', 'Nilai Warna', 'Format Warna', 'Masukkan nilai warna yang valid'],
  'it-IT': ['Reimposta colore testo', 'Nessun riempimento', 'Colore testo automatico', 'Colore bordo non impostato', 'Misto', 'Opzioni di {title}', 'Altri colori', 'Annulla', 'OK', 'Colori standard', 'Valore del colore', 'Formato del colore', 'Inserisci un valore di colore valido'],
  'ja-JP': ['文字色をリセット', '塗りつぶしなし', '自動文字色', '罫線の色は未設定', '混在', '{title}のオプション', 'その他の色', 'キャンセル', '確認', '標準の色', '色の値', '色の形式', '有効な色の値を入力してください'],
  'ko-KR': ['텍스트 색상 초기화', '채우기 없음', '자동 텍스트 색', '테두리 색상 미설정', '혼합', '{title} 옵션', '추가 색상', '취소', '확인', '표준 색상', '색상 값', '색상 형식', '유효한 색상 값을 입력하세요'],
  'pl-PL': ['Resetuj kolor tekstu', 'Bez wypełnienia', 'Automatyczny kolor tekstu', 'Kolor obramowania nieustawiony', 'Mieszane', 'Opcje: {title}', 'Więcej kolorów', 'Anuluj', 'OK', 'Kolory standardowe', 'Wartość koloru', 'Format koloru', 'Wprowadź prawidłową wartość koloru'],
  'pt-BR': ['Redefinir Cor do Texto', 'Sem Preenchimento', 'Cor do Texto Automática', 'Cor da Borda Não Definida', 'Misto', 'Opções de {title}', 'Mais Cores', 'Cancelar', 'OK', 'Cores Padrão', 'Valor da Cor', 'Formato da Cor', 'Insira um valor de cor válido'],
  'ru-RU': ['Сбросить цвет текста', 'Нет заливки', 'Автоматический цвет текста', 'Цвет границы не задан', 'Смешанный', 'Параметры: {title}', 'Другие цвета', 'Отмена', 'ОК', 'Стандартные цвета', 'Значение цвета', 'Формат цвета', 'Введите допустимое значение цвета'],
  'sk-SK': ['Obnoviť farbu textu', 'Bez výplne', 'Automatická farba textu', 'Farba orámovania nie je nastavená', 'Zmiešané', 'Možnosti: {title}', 'Ďalšie farby', 'Zrušiť', 'OK', 'Štandardné farby', 'Hodnota farby', 'Formát farby', 'Zadajte platnú hodnotu farby'],
  'vi-VN': ['Đặt lại màu văn bản', 'Không tô màu', 'Màu văn bản tự động', 'Chưa đặt màu đường viền', 'Hỗn hợp', 'Tùy chọn {title}', 'Màu khác', 'Hủy', 'Xác nhận', 'Màu tiêu chuẩn', 'Giá trị màu', 'Định dạng màu', 'Nhập giá trị màu hợp lệ'],
  'zh-CN': ['重置文字颜色', '无填充', '自动文字颜色', '未设置边框颜色', '混合', '{title}选项', '更多颜色', '取消', '确定', '标准颜色', '颜色值', '颜色格式', '请输入有效的颜色值'],
  'zh-HK': ['重設文字顏色', '無填滿', '自動文字顏色', '未設定框線顏色', '混合', '{title}選項', '更多顏色', '取消', '確定', '標準顏色', '顏色值', '顏色格式', '請輸入有效的顏色值'],
  'zh-TW': ['重設文字色彩', '無填滿', '自動文字色彩', '未設定框線色彩', '混合', '{title}選項', '其他色彩', '取消', '確定', '標準色彩', '色彩值', '色彩格式', '請輸入有效的色彩值'],
}

export const sheetToolbarColorLocales = Object.fromEntries(Object.entries(labels).map(([locale, values]) => [
  locale, Object.fromEntries(colorKeys.map((key, index) => [`sheet.color.${key}`, values[index]!])),
]))

const borderApiStyles: Record<OfficeLocale, readonly [string, string]> = {
  'ar-SA': ['بلا خط حدود', 'شرطة ونقطة مائلة'],
  'ca-ES': ['Sense línia de vora', 'Guió i punt inclinat'],
  'de-DE': ['Keine Rahmenlinie', 'Schräge Strichpunktlinie'],
  'en-US': ['No Border Line', 'Slanted Dash-Dot'],
  'es-ES': ['Sin línea de borde', 'Guion y punto inclinado'],
  'fa-IR': ['بدون خط کادر', 'خط و نقطه مایل'],
  'fr-FR': ['Aucune ligne de bordure', 'Tiret-point incliné'],
  'id-ID': ['Tanpa Garis Batas', 'Garis-Titik Miring'],
  'it-IT': ['Nessuna linea di bordo', 'Tratto-punto inclinato'],
  'ja-JP': ['罫線なし', '斜めの一点鎖線'],
  'ko-KR': ['테두리 선 없음', '기울어진 일점쇄선'],
  'pl-PL': ['Bez linii obramowania', 'Ukośna kreska-kropka'],
  'pt-BR': ['Sem Linha de Borda', 'Traço-Ponto Inclinado'],
  'ru-RU': ['Без линии границы', 'Наклонная штрихпунктирная'],
  'sk-SK': ['Bez čiary orámovania', 'Šikmá bodkočiarkovaná'],
  'vi-VN': ['Không có đường viền', 'Gạch chấm nghiêng'],
  'zh-CN': ['无边框线', '倾斜点划线'],
  'zh-HK': ['無框線', '傾斜點劃線'],
  'zh-TW': ['無框線', '傾斜點畫線'],
}
/** Display API-only pens without adding choices to the fixed twelve-style menu. */
export const sheetToolbarBorderApiStyleLocales = Object.fromEntries(Object.entries(borderApiStyles).map(([locale, values]) => [
  locale, { 'sheet.border.style.none': values[0], 'sheet.border.style.slant-dash-dot': values[1] },
]))
