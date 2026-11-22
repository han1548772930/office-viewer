const colorKeys = [
  'black', 'gray', 'white', 'red', 'orange', 'yellow', 'green', 'teal',
  'blue', 'navy', 'purple', 'pink', 'brown', 'light-gray', 'light-blue', 'light-yellow',
] as const

const colorValues = [
  '#1f2937', '#6b7280', '#ffffff', '#dc2626', '#ea580c', '#facc15', '#16a34a', '#0d9488',
  '#2563eb', '#1e3a8a', '#9333ea', '#db2777', '#92400e', '#d1d5db', '#bfdbfe', '#fef3c7',
] as const

const labels: Record<string, string[]> = {
  'ar-SA': ['أسود', 'رمادي', 'أبيض', 'أحمر', 'برتقالي', 'أصفر', 'أخضر', 'فيروزي', 'أزرق', 'كحلي', 'بنفسجي', 'وردي', 'بني', 'رمادي فاتح', 'أزرق فاتح', 'أصفر فاتح'],
  'ca-ES': ['Negre', 'Gris', 'Blanc', 'Vermell', 'Taronja', 'Groc', 'Verd', 'Turquesa', 'Blau', 'Blau marí', 'Lila', 'Rosa', 'Marró', 'Gris clar', 'Blau clar', 'Groc clar'],
  'de-DE': ['Schwarz', 'Grau', 'Weiß', 'Rot', 'Orange', 'Gelb', 'Grün', 'Türkis', 'Blau', 'Marineblau', 'Lila', 'Rosa', 'Braun', 'Hellgrau', 'Hellblau', 'Hellgelb'],
  'en-US': ['Black', 'Gray', 'White', 'Red', 'Orange', 'Yellow', 'Green', 'Teal', 'Blue', 'Navy', 'Purple', 'Pink', 'Brown', 'Light Gray', 'Light Blue', 'Light Yellow'],
  'es-ES': ['Negro', 'Gris', 'Blanco', 'Rojo', 'Naranja', 'Amarillo', 'Verde', 'Verde azulado', 'Azul', 'Azul marino', 'Morado', 'Rosa', 'Marrón', 'Gris claro', 'Azul claro', 'Amarillo claro'],
  'fa-IR': ['مشکی', 'خاکستری', 'سفید', 'قرمز', 'نارنجی', 'زرد', 'سبز', 'فیروزه‌ای', 'آبی', 'سرمه‌ای', 'بنفش', 'صورتی', 'قهوه‌ای', 'خاکستری روشن', 'آبی روشن', 'زرد روشن'],
  'fr-FR': ['Noir', 'Gris', 'Blanc', 'Rouge', 'Orange', 'Jaune', 'Vert', 'Sarcelle', 'Bleu', 'Bleu marine', 'Violet', 'Rose', 'Marron', 'Gris clair', 'Bleu clair', 'Jaune clair'],
  'id-ID': ['Hitam', 'Abu-abu', 'Putih', 'Merah', 'Oranye', 'Kuning', 'Hijau', 'Hijau kebiruan', 'Biru', 'Biru dongker', 'Ungu', 'Merah muda', 'Cokelat', 'Abu-abu muda', 'Biru muda', 'Kuning muda'],
  'it-IT': ['Nero', 'Grigio', 'Bianco', 'Rosso', 'Arancione', 'Giallo', 'Verde', 'Verde acqua', 'Blu', 'Blu navy', 'Viola', 'Rosa', 'Marrone', 'Grigio chiaro', 'Azzurro', 'Giallo chiaro'],
  'ja-JP': ['黒', 'グレー', '白', '赤', 'オレンジ', '黄', '緑', '青緑', '青', '濃紺', '紫', 'ピンク', '茶', 'ライトグレー', 'ライトブルー', 'ライトイエロー'],
  'ko-KR': ['검정', '회색', '흰색', '빨강', '주황', '노랑', '초록', '청록', '파랑', '남색', '보라', '분홍', '갈색', '밝은 회색', '밝은 파랑', '밝은 노랑'],
  'pl-PL': ['Czarny', 'Szary', 'Biały', 'Czerwony', 'Pomarańczowy', 'Żółty', 'Zielony', 'Morski', 'Niebieski', 'Granatowy', 'Fioletowy', 'Różowy', 'Brązowy', 'Jasnoszary', 'Jasnoniebieski', 'Jasnożółty'],
  'pt-BR': ['Preto', 'Cinza', 'Branco', 'Vermelho', 'Laranja', 'Amarelo', 'Verde', 'Verde-azulado', 'Azul', 'Azul-marinho', 'Roxo', 'Rosa', 'Marrom', 'Cinza-claro', 'Azul-claro', 'Amarelo-claro'],
  'ru-RU': ['Черный', 'Серый', 'Белый', 'Красный', 'Оранжевый', 'Желтый', 'Зеленый', 'Бирюзовый', 'Синий', 'Темно-синий', 'Фиолетовый', 'Розовый', 'Коричневый', 'Светло-серый', 'Светло-синий', 'Светло-желтый'],
  'sk-SK': ['Čierna', 'Sivá', 'Biela', 'Červená', 'Oranžová', 'Žltá', 'Zelená', 'Tyrkysová', 'Modrá', 'Námornícka modrá', 'Fialová', 'Ružová', 'Hnedá', 'Svetlosivá', 'Svetlomodrá', 'Svetložltá'],
  'vi-VN': ['Đen', 'Xám', 'Trắng', 'Đỏ', 'Cam', 'Vàng', 'Xanh lá', 'Xanh ngọc', 'Xanh dương', 'Xanh navy', 'Tím', 'Hồng', 'Nâu', 'Xám nhạt', 'Xanh nhạt', 'Vàng nhạt'],
  'zh-CN': ['黑色', '灰色', '白色', '红色', '橙色', '黄色', '绿色', '青色', '蓝色', '藏青色', '紫色', '粉色', '棕色', '浅灰色', '浅蓝色', '浅黄色'],
  'zh-HK': ['黑色', '灰色', '白色', '紅色', '橙色', '黃色', '綠色', '青色', '藍色', '藏青色', '紫色', '粉紅色', '棕色', '淺灰色', '淺藍色', '淺黃色'],
  'zh-TW': ['黑色', '灰色', '白色', '紅色', '橘色', '黃色', '綠色', '青色', '藍色', '深藍色', '紫色', '粉紅色', '棕色', '淺灰色', '淺藍色', '淺黃色'],
}

const resetLabels: Record<string, { fontColor: string; noBackground: string }> = {
  'ar-SA': { fontColor: 'إعادة تعيين لون الخط', noBackground: 'بلا خلفية' },
  'ca-ES': { fontColor: 'Restableix el color de la font', noBackground: 'Sense fons' },
  'de-DE': { fontColor: 'Schriftfarbe zurücksetzen', noBackground: 'Kein Hintergrund' },
  'en-US': { fontColor: 'Reset Font Color', noBackground: 'No Background' },
  'es-ES': { fontColor: 'Restablecer color de fuente', noBackground: 'Sin fondo' },
  'fa-IR': { fontColor: 'بازنشانی رنگ قلم', noBackground: 'بدون پس‌زمینه' },
  'fr-FR': { fontColor: 'Réinitialiser la couleur de police', noBackground: 'Sans arrière-plan' },
  'id-ID': { fontColor: 'Atur ulang warna font', noBackground: 'Tanpa latar belakang' },
  'it-IT': { fontColor: 'Reimposta colore carattere', noBackground: 'Nessuno sfondo' },
  'ja-JP': { fontColor: 'フォントの色をリセット', noBackground: '背景なし' },
  'ko-KR': { fontColor: '글꼴 색상 초기화', noBackground: '배경 없음' },
  'pl-PL': { fontColor: 'Resetuj kolor czcionki', noBackground: 'Brak tła' },
  'pt-BR': { fontColor: 'Redefinir cor da fonte', noBackground: 'Sem plano de fundo' },
  'ru-RU': { fontColor: 'Сбросить цвет шрифта', noBackground: 'Без фона' },
  'sk-SK': { fontColor: 'Obnoviť farbu písma', noBackground: 'Bez pozadia' },
  'vi-VN': { fontColor: 'Đặt lại màu phông chữ', noBackground: 'Không nền' },
  'zh-CN': { fontColor: '重置字体颜色', noBackground: '无背景色' },
  'zh-HK': { fontColor: '重設字體顏色', noBackground: '無背景色' },
  'zh-TW': { fontColor: '重設字體顏色', noBackground: '無背景色' },
}

/** Labels for the color palette and its custom-color stage. */
const pickerLabels: Record<string, { more: string; cancel: string; confirm: string; reset: string; automatic: string; standard: string; value: string; format: string; invalid: string }> = {
  'ar-SA': { more: 'المزيد من الألوان', cancel: 'إلغاء', confirm: 'موافق', reset: 'إعادة تعيين', automatic: 'لون نص تلقائي', standard: 'ألوان قياسية', value: 'قيمة اللون', format: 'تنسيق اللون', invalid: 'أدخل قيمة لون صالحة' },
  'ca-ES': { more: 'Més colors', cancel: 'cancel·la', confirm: "D'acord", reset: 'Restableix', automatic: 'Color de text automàtic', standard: 'Colors estàndard', value: 'Valor del color', format: 'Format del color', invalid: 'Introduïu un valor de color vàlid' },
  'de-DE': { more: 'Mehr Farben', cancel: 'Abbrechen', confirm: 'OK', reset: 'Zurücksetzen', automatic: 'Automatische Textfarbe', standard: 'Standardfarben', value: 'Farbwert', format: 'Farbformat', invalid: 'Geben Sie einen gültigen Farbwert ein' },
  'en-US': { more: 'More Colors', cancel: 'cancel', confirm: 'ok', reset: 'Reset', automatic: 'Automatic Text Color', standard: 'Standard Colors', value: 'Color Value', format: 'Color Format', invalid: 'Enter a valid color value' },
  'es-ES': { more: 'Más colores', cancel: 'Cancelar', confirm: 'Aceptar', reset: 'Restablecer', automatic: 'Color de texto automático', standard: 'Colores estándar', value: 'Valor de color', format: 'Formato de color', invalid: 'Introduce un valor de color válido' },
  'fa-IR': { more: 'رنگ‌های بیشتر', cancel: 'لغو', confirm: 'باشه', reset: 'بازنشانی', automatic: 'رنگ متن خودکار', standard: 'رنگ‌های استاندارد', value: 'مقدار رنگ', format: 'قالب رنگ', invalid: 'یک مقدار رنگ معتبر وارد کنید' },
  'fr-FR': { more: 'Plus de couleurs', cancel: 'annuler', confirm: 'OK', reset: 'Réinitialiser', automatic: 'Couleur du texte automatique', standard: 'Couleurs standard', value: 'Valeur de couleur', format: 'Format de couleur', invalid: 'Saisissez une valeur de couleur valide' },
  'id-ID': { more: 'Warna Lainnya', cancel: 'Batal', confirm: 'OK', reset: 'Atur Ulang', automatic: 'Warna Teks Otomatis', standard: 'Warna Standar', value: 'Nilai Warna', format: 'Format Warna', invalid: 'Masukkan nilai warna yang valid' },
  'it-IT': { more: 'Altri Colori', cancel: 'annulla', confirm: 'OK', reset: 'Reimposta', automatic: 'Colore del testo automatico', standard: 'Colori standard', value: 'Valore del colore', format: 'Formato del colore', invalid: 'Inserisci un valore di colore valido' },
  'ja-JP': { more: 'その他の色', cancel: 'キャンセル', confirm: '確認', reset: '既定の色に戻す', automatic: '自動文字色', standard: '標準の色', value: '色の値', format: '色の形式', invalid: '有効な色の値を入力してください' },
  'ko-KR': { more: '추가 색상', cancel: '취소', confirm: '확인', reset: '기본색으로 재설정', automatic: '자동 텍스트 색', standard: '표준 색상', value: '색상 값', format: '색상 형식', invalid: '유효한 색상 값을 입력하세요' },
  'pl-PL': { more: 'Więcej kolorów', cancel: 'anuluj', confirm: 'OK', reset: 'Resetuj', automatic: 'Automatyczny kolor tekstu', standard: 'Kolory standardowe', value: 'Wartość koloru', format: 'Format koloru', invalid: 'Wprowadź prawidłową wartość koloru' },
  'pt-BR': { more: 'Mais Cores', cancel: 'cancelar', confirm: 'OK', reset: 'Redefinir', automatic: 'Cor de Texto Automática', standard: 'Cores Padrão', value: 'Valor da Cor', format: 'Formato da Cor', invalid: 'Insira um valor de cor válido' },
  'ru-RU': { more: 'Больше цветов', cancel: 'отмена', confirm: 'ок', reset: 'Сбросить', automatic: 'Автоматический цвет текста', standard: 'Стандартные цвета', value: 'Значение цвета', format: 'Формат цвета', invalid: 'Введите допустимое значение цвета' },
  'sk-SK': { more: 'Viac farieb', cancel: 'zrušiť', confirm: 'OK', reset: 'Obnoviť', automatic: 'Automatická farba textu', standard: 'Štandardné farby', value: 'Hodnota farby', format: 'Formát farby', invalid: 'Zadajte platnú hodnotu farby' },
  'vi-VN': { more: 'Màu sắc khác', cancel: 'Hủy bỏ', confirm: 'Xác nhận', reset: 'Đặt lại màu', automatic: 'Màu văn bản tự động', standard: 'Màu tiêu chuẩn', value: 'Giá trị màu', format: 'Định dạng màu', invalid: 'Nhập giá trị màu hợp lệ' },
  'zh-CN': { more: '更多颜色', cancel: '取消', confirm: '确定', reset: '重置颜色', automatic: '自动文字颜色', standard: '标准颜色', value: '颜色值', format: '颜色格式', invalid: '请输入有效的颜色值' },
  'zh-HK': { more: '更多顏色', cancel: '取消', confirm: '確定', reset: '重置顏色', automatic: '自動文字顏色', standard: '標準顏色', value: '顏色值', format: '顏色格式', invalid: '請輸入有效的顏色值' },
  'zh-TW': { more: '更多顏色', cancel: '取消', confirm: '確定', reset: '重置顏色', automatic: '自動文字顏色', standard: '標準顏色', value: '顏色值', format: '顏色格式', invalid: '請輸入有效的顏色值' },
}

export const docToolbarColorLocales: Record<string, Record<string, string>> = Object.fromEntries(
  Object.entries(labels).map(([locale, names]) => {
    const picker = pickerLabels[locale] ?? pickerLabels['en-US']!
    const noBackground = resetLabels[locale]?.noBackground ?? 'No Background'
    const messages: Record<string, string> = {
      'items.doc-text-color-reset': resetLabels[locale]?.fontColor ?? 'Reset Font Color',
      'items.doc-highlight-reset': picker.reset,
      'items.doc-text-color-clear': picker.automatic,
      'items.doc-highlight-clear': noBackground,
      'document.color.noBackground': noBackground,
      ...Object.fromEntries(Object.entries(picker).map(([key, value]) => [`document.color.${key}`, value])),
    }
    colorKeys.forEach((key, index) => {
      const value = colorValues[index]
      const label = names[index]
      if (value && label) {
        messages[`colors.${key}`] = label
        messages[`items.doc-text-color-${value}`] = label
        messages[`items.doc-highlight-${value}`] = label
      }
    })
    return [locale, messages]
  }),
)
