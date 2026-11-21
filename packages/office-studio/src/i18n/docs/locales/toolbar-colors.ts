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

export const docToolbarColorLocales: Record<string, Record<string, string>> = Object.fromEntries(
  Object.entries(labels).map(([locale, names]) => {
    const messages: Record<string, string> = {
      'items.doc-text-color-reset': resetLabels[locale]?.fontColor ?? 'Reset Font Color',
      'items.doc-highlight-reset': resetLabels[locale]?.noBackground ?? 'No Background',
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
