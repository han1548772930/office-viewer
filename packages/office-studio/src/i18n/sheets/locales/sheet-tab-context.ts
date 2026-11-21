const labels = {
  'ar-SA': ['نسخ الورقة', 'لون علامة التبويب', 'بلا لون', 'أحمر', 'برتقالي', 'أصفر', 'أخضر', 'سماوي', 'أزرق', 'بنفسجي', 'إخفاء الصف', 'إخفاء العمود', 'إظهار الصفوف المخفية', 'إظهار الأعمدة المخفية', 'احتواء ارتفاع الصف تلقائيًا', 'احتواء عرض العمود تلقائيًا', 'إدراج صف أسفل', 'إدراج عمود إلى اليمين'],
  'ca-ES': ['Copia el full', 'Color de la pestanya', 'Sense color', 'Vermell', 'Taronja', 'Groc', 'Verd', 'Cian', 'Blau', 'Lila', 'Amaga la fila', 'Amaga la columna', 'Mostra les files ocultes', 'Mostra les columnes ocultes', 'Ajusta automàticament l’alçada', 'Ajusta automàticament l’amplada', 'Insereix fila a sota', 'Insereix columna a la dreta'],
  'de-DE': ['Blatt kopieren', 'Registerkartenfarbe', 'Keine Farbe', 'Rot', 'Orangefarben', 'Gelb', 'Grün', 'Türkis', 'Blau', 'Lila', 'Zeile ausblenden', 'Spalte ausblenden', 'Ausgeblendete Zeilen einblenden', 'Ausgeblendete Spalten einblenden', 'Zeilenhöhe automatisch anpassen', 'Spaltenbreite automatisch anpassen', 'Zeile darunter einfügen', 'Spalte rechts einfügen'],
  'en-US': ['Copy Sheet', 'Tab Color', 'No Color', 'Red', 'Orange', 'Yellow', 'Green', 'Cyan', 'Blue', 'Purple', 'Hide Row', 'Hide Column', 'Unhide Rows', 'Unhide Columns', 'AutoFit Row Height', 'AutoFit Column Width', 'Insert Row Below', 'Insert Column Right'],
  'es-ES': ['Copiar hoja', 'Color de pestaña', 'Sin color', 'Rojo', 'Naranja', 'Amarillo', 'Verde', 'Cian', 'Azul', 'Morado', 'Ocultar fila', 'Ocultar columna', 'Mostrar filas ocultas', 'Mostrar columnas ocultas', 'Autoajustar alto de fila', 'Autoajustar ancho de columna', 'Insertar fila debajo', 'Insertar columna a la derecha'],
  'fa-IR': ['کپی برگه', 'رنگ زبانه', 'بدون رنگ', 'قرمز', 'نارنجی', 'زرد', 'سبز', 'فیروزه‌ای', 'آبی', 'بنفش', 'پنهان کردن ردیف', 'پنهان کردن ستون', 'نمایش ردیف‌های پنهان', 'نمایش ستون‌های پنهان', 'تنظیم خودکار ارتفاع ردیف', 'تنظیم خودکار عرض ستون', 'درج ردیف در پایین', 'درج ستون در سمت راست'],
  'fr-FR': ['Copier la feuille', 'Couleur de l’onglet', 'Aucune couleur', 'Rouge', 'Orangé', 'Jaune', 'Vert', 'Turquoise', 'Bleu', 'Violet', 'Masquer la ligne', 'Masquer la colonne', 'Afficher les lignes masquées', 'Afficher les colonnes masquées', 'Ajuster automatiquement la hauteur', 'Ajuster automatiquement la largeur', 'Insérer une ligne en dessous', 'Insérer une colonne à droite'],
  'id-ID': ['Salin lembar', 'Warna tab', 'Tanpa warna', 'Merah', 'Oranye', 'Kuning', 'Hijau', 'Sian', 'Biru', 'Ungu', 'Sembunyikan baris', 'Sembunyikan kolom', 'Tampilkan baris tersembunyi', 'Tampilkan kolom tersembunyi', 'Sesuaikan tinggi baris otomatis', 'Sesuaikan lebar kolom otomatis', 'Sisipkan baris di bawah', 'Sisipkan kolom di kanan'],
  'it-IT': ['Copia foglio', 'Colore scheda', 'Nessun colore', 'Rosso', 'Arancione', 'Giallo', 'Verde', 'Ciano', 'Blu', 'Viola', 'Nascondi riga', 'Nascondi colonna', 'Mostra righe nascoste', 'Mostra colonne nascoste', 'Adatta automaticamente altezza riga', 'Adatta automaticamente larghezza colonna', 'Inserisci riga sotto', 'Inserisci colonna a destra'],
  'ja-JP': ['シートのコピー', 'タブの色', '色なし', '赤', 'オレンジ', '黄', '緑', 'シアン', '青', '紫', '行を非表示', '列を非表示', '非表示の行を再表示', '非表示の列を再表示', '行の高さを自動調整', '列の幅を自動調整', '下に行を挿入', '右に列を挿入'],
  'ko-KR': ['시트 복사', '탭 색', '색 없음', '빨강', '주황', '노랑', '초록', '청록', '파랑', '보라', '행 숨기기', '열 숨기기', '숨겨진 행 표시', '숨겨진 열 표시', '행 높이 자동 맞춤', '열 너비 자동 맞춤', '아래에 행 삽입', '오른쪽에 열 삽입'],
  'pl-PL': ['Kopiuj arkusz', 'Kolor karty', 'Brak koloru', 'Czerwony', 'Pomarańczowy', 'Żółty', 'Zielony', 'Cyjan', 'Niebieski', 'Fioletowy', 'Ukryj wiersz', 'Ukryj kolumnę', 'Odkryj ukryte wiersze', 'Odkryj ukryte kolumny', 'Autodopasuj wysokość wiersza', 'Autodopasuj szerokość kolumny', 'Wstaw wiersz poniżej', 'Wstaw kolumnę po prawej'],
  'pt-BR': ['Copiar planilha', 'Cor da guia', 'Sem cor', 'Vermelho', 'Laranja', 'Amarelo', 'Verde', 'Ciano', 'Azul', 'Roxo', 'Ocultar linha', 'Ocultar coluna', 'Reexibir linhas ocultas', 'Reexibir colunas ocultas', 'Ajustar altura da linha automaticamente', 'Ajustar largura da coluna automaticamente', 'Inserir linha abaixo', 'Inserir coluna à direita'],
  'ru-RU': ['Копировать лист', 'Цвет вкладки', 'Без цвета', 'Красный', 'Оранжевый', 'Желтый', 'Зеленый', 'Голубой', 'Синий', 'Фиолетовый', 'Скрыть строку', 'Скрыть столбец', 'Показать скрытые строки', 'Показать скрытые столбцы', 'Автоподбор высоты строки', 'Автоподбор ширины столбца', 'Вставить строку ниже', 'Вставить столбец справа'],
  'sk-SK': ['Kopírovať hárok', 'Farba karty', 'Bez farby', 'Červená', 'Oranžová', 'Žltá', 'Zelená', 'Azúrová', 'Modrá', 'Fialová', 'Skryť riadok', 'Skryť stĺpec', 'Zobraziť skryté riadky', 'Zobraziť skryté stĺpce', 'Automaticky prispôsobiť výšku', 'Automaticky prispôsobiť šírku', 'Vložiť riadok pod', 'Vložiť stĺpec vpravo'],
  'vi-VN': ['Sao chép trang tính', 'Màu thẻ', 'Không màu', 'Đỏ', 'Cam', 'Vàng', 'Xanh lá', 'Lục lam', 'Xanh dương', 'Tím', 'Ẩn hàng', 'Ẩn cột', 'Hiện hàng bị ẩn', 'Hiện cột bị ẩn', 'Tự động khớp chiều cao hàng', 'Tự động khớp chiều rộng cột', 'Chèn hàng bên dưới', 'Chèn cột bên phải'],
  'zh-CN': ['复制工作表', '标签颜色', '无颜色', '红色', '橙色', '黄色', '绿色', '青色', '蓝色', '紫色', '隐藏行', '隐藏列', '取消隐藏行', '取消隐藏列', '自动调整行高', '自动调整列宽', '在下方插入行', '在右侧插入列'],
  'zh-HK': ['複製工作表', '索引標籤色彩', '無色彩', '紅色', '橙色', '黃色', '綠色', '青色', '藍色', '紫色', '隱藏列', '隱藏欄', '取消隱藏列', '取消隱藏欄', '自動調整列高', '自動調整欄寬', '在下方插入列', '在右側插入欄'],
  'zh-TW': ['複製工作表', '索引標籤色彩', '無色彩', '紅色', '橙色', '黃色', '綠色', '青色', '藍色', '紫色', '隱藏列', '隱藏欄', '取消隱藏列', '取消隱藏欄', '自動調整列高', '自動調整欄寬', '在下方插入列', '在右側插入欄'],
} as const

const translationKeys = [
  'sheet.footer.copy-sheet', 'sheet.footer.tab-color', 'sheet.footer.no-color',
  'sheet.footer.red', 'sheet.footer.orange', 'sheet.footer.yellow', 'sheet.footer.green',
  'sheet.footer.cyan', 'sheet.footer.blue', 'sheet.footer.purple',
  'sheet.context-menu.hide-row', 'sheet.context-menu.hide-column',
  'sheet.context-menu.show-hidden-rows', 'sheet.context-menu.show-hidden-columns',
  'sheet.context-menu.auto-fit-row', 'sheet.context-menu.auto-fit-column',
  'sheet.context-menu.insert-row-below', 'sheet.context-menu.insert-column-right',
] as const

export const sheetTabContextLocales = Object.fromEntries(
  Object.entries(labels).map(([locale, values]) => [locale, Object.fromEntries(translationKeys.map((key, index) => [key, values[index]!]))]),
) as Record<keyof typeof labels, Record<(typeof translationKeys)[number], string>>
