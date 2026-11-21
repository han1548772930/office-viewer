interface DocumentTocLocaleText {
  insert: string
  contentsTitle: string
  manage: string
  automatic: string
  custom: string
  update: string
  updatePages: string
  updateAll: string
  delete: string
  levels: string
  format: string
  showPageNumbers: string
  rightAlignPageNumbers: string
  tabLeader: string
  leaderNone: string
  leaderDots: string
  leaderDashes: string
  leaderUnderline: string
  formatTemplate: string
  formatClassic: string
  formatModern: string
  formatSimple: string
  noHeadings: string
  preview: string
  previewHeading: string
}

const localeText: Record<string, DocumentTocLocaleText> = {
  'ar-SA': {
    insert: 'جدول المحتويات', contentsTitle: 'المحتويات', manage: 'إدارة', automatic: 'جدول تلقائي', custom: 'جدول محتويات مخصص…', update: 'تحديث جدول المحتويات',
    updatePages: 'تحديث أرقام الصفحات فقط', updateAll: 'تحديث الجدول بالكامل', delete: 'إزالة جدول المحتويات',
    levels: 'إظهار المستويات', format: 'التنسيقات', showPageNumbers: 'إظهار أرقام الصفحات', rightAlignPageNumbers: 'محاذاة أرقام الصفحات إلى اليمين',
    tabLeader: 'حرف تعبئة علامات الجدولة', leaderNone: 'بلا', leaderDots: 'نقاط', leaderDashes: 'شرطات', leaderUnderline: 'تسطير',
    formatTemplate: 'من القالب', formatClassic: 'كلاسيكي', formatModern: 'حديث', formatSimple: 'بسيط',
    noHeadings: 'لم يتم العثور على عناوين. طبّق أنماط العنوان 1–3 أو اختر المستويات المطابقة.', preview: 'معاينة الطباعة', previewHeading: 'عنوان',
  },
  'ca-ES': {
    insert: 'Taula de continguts', contentsTitle: 'Continguts', manage: 'Gestiona', automatic: 'Taula automàtica', custom: 'Taula de continguts personalitzada…', update: 'Actualitza la taula de continguts',
    updatePages: 'Actualitza només els números de pàgina', updateAll: 'Actualitza tota la taula', delete: 'Elimina la taula de continguts',
    levels: 'Mostra els nivells', format: 'Formats', showPageNumbers: 'Mostra els números de pàgina', rightAlignPageNumbers: 'Alinea els números de pàgina a la dreta',
    tabLeader: 'Caràcter de farciment', leaderNone: 'Cap', leaderDots: 'Punts', leaderDashes: 'Guions', leaderUnderline: 'Subratllat',
    formatTemplate: 'De la plantilla', formatClassic: 'Clàssic', formatModern: 'Modern', formatSimple: 'Simple',
    noHeadings: 'No s\'ha trobat cap títol. Apliqueu els estils Títol 1–3 o trieu els nivells corresponents.', preview: 'Visualització prèvia d\'impressió', previewHeading: 'Títol',
  },
  'de-DE': {
    insert: 'Inhaltsverzeichnis', contentsTitle: 'Inhalt', manage: 'Verwalten', automatic: 'Automatisches Verzeichnis', custom: 'Benutzerdefiniertes Inhaltsverzeichnis…', update: 'Inhaltsverzeichnis aktualisieren',
    updatePages: 'Nur Seitenzahlen aktualisieren', updateAll: 'Gesamtes Verzeichnis aktualisieren', delete: 'Inhaltsverzeichnis entfernen',
    levels: 'Ebenen anzeigen', format: 'Formate', showPageNumbers: 'Seitenzahlen anzeigen', rightAlignPageNumbers: 'Seitenzahlen rechtsbündig',
    tabLeader: 'Füllzeichen', leaderNone: 'Keine', leaderDots: 'Punkte', leaderDashes: 'Striche', leaderUnderline: 'Unterstrich',
    formatTemplate: 'Von Vorlage', formatClassic: 'Klassisch', formatModern: 'Modern', formatSimple: 'Einfach',
    noHeadings: 'Keine Überschriften gefunden. Wenden Sie Überschrift 1–3 an oder wählen Sie passende Ebenen.', preview: 'Druckvorschau', previewHeading: 'Überschrift',
  },
  'en-US': {
    insert: 'Table of contents', contentsTitle: 'Contents', manage: 'Manage', automatic: 'Automatic table', custom: 'Custom table of contents…', update: 'Update table of contents',
    updatePages: 'Update page numbers only', updateAll: 'Update entire table', delete: 'Remove table of contents',
    levels: 'Show levels', format: 'Formats', showPageNumbers: 'Show page numbers', rightAlignPageNumbers: 'Right align page numbers',
    tabLeader: 'Tab leader', leaderNone: 'None', leaderDots: 'Dots', leaderDashes: 'Dashes', leaderUnderline: 'Underline',
    formatTemplate: 'From template', formatClassic: 'Classic', formatModern: 'Modern', formatSimple: 'Simple',
    noHeadings: 'No headings were found. Apply Heading 1–3 styles first, or choose matching levels.', preview: 'Print preview', previewHeading: 'Heading',
  },
  'es-ES': {
    insert: 'Tabla de contenido', contentsTitle: 'Contenido', manage: 'Administrar', automatic: 'Tabla automática', custom: 'Tabla de contenido personalizada…', update: 'Actualizar tabla de contenido',
    updatePages: 'Actualizar solo los números de página', updateAll: 'Actualizar toda la tabla', delete: 'Quitar tabla de contenido',
    levels: 'Mostrar niveles', format: 'Formatos', showPageNumbers: 'Mostrar números de página', rightAlignPageNumbers: 'Alinear números de página a la derecha',
    tabLeader: 'Carácter de relleno', leaderNone: 'Ninguno', leaderDots: 'Puntos', leaderDashes: 'Guiones', leaderUnderline: 'Subrayado',
    formatTemplate: 'De la plantilla', formatClassic: 'Clásico', formatModern: 'Moderno', formatSimple: 'Sencillo',
    noHeadings: 'No se encontraron títulos. Aplique los estilos Título 1–3 o elija los niveles correspondientes.', preview: 'Vista previa de impresión', previewHeading: 'Título',
  },
  'fa-IR': {
    insert: 'فهرست مطالب', contentsTitle: 'مطالب', manage: 'مدیریت', automatic: 'فهرست خودکار', custom: 'فهرست مطالب سفارشی…', update: 'به‌روزرسانی فهرست مطالب',
    updatePages: 'فقط به‌روزرسانی شماره صفحات', updateAll: 'به‌روزرسانی کل فهرست', delete: 'حذف فهرست مطالب',
    levels: 'نمایش سطوح', format: 'قالب‌ها', showPageNumbers: 'نمایش شماره صفحات', rightAlignPageNumbers: 'تراز شماره صفحات به راست',
    tabLeader: 'نویسهٔ راهنمای تب', leaderNone: 'هیچ‌کدام', leaderDots: 'نقطه', leaderDashes: 'خط تیره', leaderUnderline: 'زیرخط',
    formatTemplate: 'از الگو', formatClassic: 'کلاسیک', formatModern: 'مدرن', formatSimple: 'ساده',
    noHeadings: 'عنوانی یافت نشد. سبک‌های عنوان ۱ تا ۳ را اعمال کنید یا سطوح منطبق را انتخاب کنید.', preview: 'پیش‌نمایش چاپ', previewHeading: 'عنوان',
  },
  'fr-FR': {
    insert: 'Table des matières', contentsTitle: 'Sommaire', manage: 'Gérer', automatic: 'Table automatique', custom: 'Table des matières personnalisée…', update: 'Mettre à jour la table des matières',
    updatePages: 'Mettre à jour uniquement les numéros de page', updateAll: 'Mettre à jour toute la table', delete: 'Supprimer la table des matières',
    levels: 'Afficher les niveaux', format: 'Formats', showPageNumbers: 'Afficher les numéros de page', rightAlignPageNumbers: 'Aligner les numéros de page à droite',
    tabLeader: 'Caractères de suite', leaderNone: 'Aucun', leaderDots: 'Points', leaderDashes: 'Tirets', leaderUnderline: 'Soulignement',
    formatTemplate: 'Depuis le modèle', formatClassic: 'Classique', formatModern: 'Moderne', formatSimple: 'Simple',
    noHeadings: 'Aucun titre trouvé. Appliquez les styles Titre 1–3 ou choisissez les niveaux correspondants.', preview: 'Aperçu avant impression', previewHeading: 'Titre',
  },
  'id-ID': {
    insert: 'Daftar isi', contentsTitle: 'Daftar isi', manage: 'Kelola', automatic: 'Daftar otomatis', custom: 'Daftar isi kustom…', update: 'Perbarui daftar isi',
    updatePages: 'Perbarui nomor halaman saja', updateAll: 'Perbarui seluruh daftar', delete: 'Hapus daftar isi',
    levels: 'Tampilkan tingkat', format: 'Format', showPageNumbers: 'Tampilkan nomor halaman', rightAlignPageNumbers: 'Ratakan nomor halaman ke kanan',
    tabLeader: 'Pengisi tab', leaderNone: 'Tidak ada', leaderDots: 'Titik', leaderDashes: 'Garis putus-putus', leaderUnderline: 'Garis bawah',
    formatTemplate: 'Dari templat', formatClassic: 'Klasik', formatModern: 'Modern', formatSimple: 'Sederhana',
    noHeadings: 'Tidak ada judul yang ditemukan. Terapkan gaya Judul 1–3 atau pilih tingkat yang sesuai.', preview: 'Pratinjau cetak', previewHeading: 'Judul',
  },
  'it-IT': {
    insert: 'Sommario', contentsTitle: 'Sommario', manage: 'Gestisci', automatic: 'Sommario automatico', custom: 'Sommario personalizzato…', update: 'Aggiorna sommario',
    updatePages: 'Aggiorna solo i numeri di pagina', updateAll: 'Aggiorna intero sommario', delete: 'Rimuovi sommario',
    levels: 'Mostra livelli', format: 'Formati', showPageNumbers: 'Mostra numeri di pagina', rightAlignPageNumbers: 'Allinea i numeri di pagina a destra',
    tabLeader: 'Carattere di riempimento', leaderNone: 'Nessuno', leaderDots: 'Punti', leaderDashes: 'Trattini', leaderUnderline: 'Sottolineatura',
    formatTemplate: 'Da modello', formatClassic: 'Classico', formatModern: 'Moderno', formatSimple: 'Semplice',
    noHeadings: 'Nessun titolo trovato. Applica gli stili Titolo 1–3 o scegli i livelli corrispondenti.', preview: 'Anteprima di stampa', previewHeading: 'Titolo',
  },
  'ja-JP': {
    insert: '目次', contentsTitle: '目次', manage: '管理', automatic: '自動作成の目次', custom: 'ユーザー設定の目次…', update: '目次の更新',
    updatePages: 'ページ番号だけを更新する', updateAll: '目次をすべて更新する', delete: '目次の削除',
    levels: '表示するレベル', format: '書式', showPageNumbers: 'ページ番号を表示する', rightAlignPageNumbers: 'ページ番号を右揃えにする',
    tabLeader: 'タブリーダー', leaderNone: 'なし', leaderDots: '点線', leaderDashes: '破線', leaderUnderline: '下線',
    formatTemplate: 'テンプレートから', formatClassic: 'クラシック', formatModern: 'モダン', formatSimple: 'シンプル',
    noHeadings: '見出しが見つかりません。見出し 1～3 のスタイルを適用するか、対応するレベルを選択してください。', preview: '印刷プレビュー', previewHeading: '見出し',
  },
  'ko-KR': {
    insert: '목차', contentsTitle: '목차', manage: '관리', automatic: '자동 목차', custom: '사용자 지정 목차…', update: '목차 업데이트',
    updatePages: '페이지 번호만 업데이트', updateAll: '목차 전체 업데이트', delete: '목차 제거',
    levels: '표시할 수준', format: '서식', showPageNumbers: '페이지 번호 표시', rightAlignPageNumbers: '페이지 번호 오른쪽 맞춤',
    tabLeader: '탭 채움선', leaderNone: '없음', leaderDots: '점선', leaderDashes: '파선', leaderUnderline: '밑줄',
    formatTemplate: '서식 파일에서', formatClassic: '클래식', formatModern: '모던', formatSimple: '간단',
    noHeadings: '제목을 찾을 수 없습니다. 제목 1–3 스타일을 적용하거나 해당 수준을 선택하세요.', preview: '인쇄 미리 보기', previewHeading: '제목',
  },
  'pl-PL': {
    insert: 'Spis treści', contentsTitle: 'Spis treści', manage: 'Zarządzaj', automatic: 'Automatyczny spis treści', custom: 'Niestandardowy spis treści…', update: 'Aktualizuj spis treści',
    updatePages: 'Aktualizuj tylko numery stron', updateAll: 'Aktualizuj cały spis', delete: 'Usuń spis treści',
    levels: 'Pokaż poziomy', format: 'Formaty', showPageNumbers: 'Pokaż numery stron', rightAlignPageNumbers: 'Wyrównaj numery stron do prawej',
    tabLeader: 'Znak wiodący', leaderNone: 'Brak', leaderDots: 'Kropki', leaderDashes: 'Kreski', leaderUnderline: 'Podkreślenie',
    formatTemplate: 'Z szablonu', formatClassic: 'Klasyczny', formatModern: 'Nowoczesny', formatSimple: 'Prosty',
    noHeadings: 'Nie znaleziono nagłówków. Zastosuj style Nagłówek 1–3 lub wybierz odpowiednie poziomy.', preview: 'Podgląd wydruku', previewHeading: 'Nagłówek',
  },
  'pt-BR': {
    insert: 'Sumário', contentsTitle: 'Sumário', manage: 'Gerenciar', automatic: 'Sumário automático', custom: 'Sumário personalizado…', update: 'Atualizar sumário',
    updatePages: 'Atualizar apenas os números de página', updateAll: 'Atualizar o sumário inteiro', delete: 'Remover sumário',
    levels: 'Mostrar níveis', format: 'Formatos', showPageNumbers: 'Mostrar números de página', rightAlignPageNumbers: 'Alinhar números de página à direita',
    tabLeader: 'Preenchimento de tabulação', leaderNone: 'Nenhum', leaderDots: 'Pontos', leaderDashes: 'Traços', leaderUnderline: 'Sublinhado',
    formatTemplate: 'Do modelo', formatClassic: 'Clássico', formatModern: 'Moderno', formatSimple: 'Simples',
    noHeadings: 'Nenhum título encontrado. Aplique os estilos Título 1–3 ou escolha os níveis correspondentes.', preview: 'Visualização de impressão', previewHeading: 'Título',
  },
  'ru-RU': {
    insert: 'Оглавление', contentsTitle: 'Содержание', manage: 'Управление', automatic: 'Автособираемое оглавление', custom: 'Настраиваемое оглавление…', update: 'Обновить оглавление',
    updatePages: 'Обновить только номера страниц', updateAll: 'Обновить целиком', delete: 'Удалить оглавление',
    levels: 'Показать уровни', format: 'Форматы', showPageNumbers: 'Показать номера страниц', rightAlignPageNumbers: 'Номера страниц по правому краю',
    tabLeader: 'Заполнитель', leaderNone: 'Нет', leaderDots: 'Точки', leaderDashes: 'Тире', leaderUnderline: 'Подчёркивание',
    formatTemplate: 'Из шаблона', formatClassic: 'Классический', formatModern: 'Современный', formatSimple: 'Простой',
    noHeadings: 'Заголовки не найдены. Примените стили «Заголовок 1–3» или выберите соответствующие уровни.', preview: 'Предварительный просмотр', previewHeading: 'Заголовок',
  },
  'sk-SK': {
    insert: 'Obsah', contentsTitle: 'Obsah', manage: 'Spravovať', automatic: 'Automatický obsah', custom: 'Vlastný obsah…', update: 'Aktualizovať obsah',
    updatePages: 'Aktualizovať iba čísla strán', updateAll: 'Aktualizovať celý obsah', delete: 'Odstrániť obsah',
    levels: 'Zobraziť úrovne', format: 'Formáty', showPageNumbers: 'Zobraziť čísla strán', rightAlignPageNumbers: 'Zarovnať čísla strán doprava',
    tabLeader: 'Vodiaci znak', leaderNone: 'Žiadny', leaderDots: 'Bodky', leaderDashes: 'Pomlčky', leaderUnderline: 'Podčiarknutie',
    formatTemplate: 'Zo šablóny', formatClassic: 'Klasický', formatModern: 'Moderný', formatSimple: 'Jednoduchý',
    noHeadings: 'Nenašli sa žiadne nadpisy. Použite štýly Nadpis 1–3 alebo vyberte zodpovedajúce úrovne.', preview: 'Ukážka pred tlačou', previewHeading: 'Nadpis',
  },
  'vi-VN': {
    insert: 'Mục lục', contentsTitle: 'Mục lục', manage: 'Quản lý', automatic: 'Mục lục tự động', custom: 'Mục lục tùy chỉnh…', update: 'Cập nhật mục lục',
    updatePages: 'Chỉ cập nhật số trang', updateAll: 'Cập nhật toàn bộ mục lục', delete: 'Xóa mục lục',
    levels: 'Hiển thị cấp độ', format: 'Định dạng', showPageNumbers: 'Hiển thị số trang', rightAlignPageNumbers: 'Căn phải số trang',
    tabLeader: 'Ký tự dẫn tab', leaderNone: 'Không có', leaderDots: 'Dấu chấm', leaderDashes: 'Dấu gạch ngang', leaderUnderline: 'Gạch dưới',
    formatTemplate: 'Từ mẫu', formatClassic: 'Cổ điển', formatModern: 'Hiện đại', formatSimple: 'Đơn giản',
    noHeadings: 'Không tìm thấy tiêu đề. Hãy áp dụng kiểu Tiêu đề 1–3 hoặc chọn cấp độ phù hợp.', preview: 'Xem trước khi in', previewHeading: 'Tiêu đề',
  },
  'zh-CN': {
    insert: '插入目录', contentsTitle: '目录', manage: '管理', automatic: '自动目录', custom: '自定义目录…', update: '更新目录',
    updatePages: '只更新页码', updateAll: '更新整个目录', delete: '删除目录',
    levels: '显示级别', format: '格式', showPageNumbers: '显示页码', rightAlignPageNumbers: '页码右对齐',
    tabLeader: '制表符前导符', leaderNone: '无', leaderDots: '点线', leaderDashes: '短划线', leaderUnderline: '下划线',
    formatTemplate: '来自模板', formatClassic: '经典', formatModern: '现代', formatSimple: '简洁',
    noHeadings: '未找到标题。请先应用标题 1–3 样式，或选择匹配的标题级别。', preview: '打印预览', previewHeading: '标题',
  },
  'zh-HK': {
    insert: '插入目錄', contentsTitle: '目錄', manage: '管理', automatic: '自動目錄', custom: '自訂目錄…', update: '更新目錄',
    updatePages: '只更新頁碼', updateAll: '更新整個目錄', delete: '刪除目錄',
    levels: '顯示級別', format: '格式', showPageNumbers: '顯示頁碼', rightAlignPageNumbers: '頁碼靠右對齊',
    tabLeader: '定位點前置字元', leaderNone: '無', leaderDots: '點線', leaderDashes: '短劃線', leaderUnderline: '底線',
    formatTemplate: '來自範本', formatClassic: '傳統', formatModern: '現代', formatSimple: '簡潔',
    noHeadings: '找不到標題。請先套用標題 1–3 樣式，或選擇相符的標題級別。', preview: '預覽列印', previewHeading: '標題',
  },
  'zh-TW': {
    insert: '目錄', contentsTitle: '目錄', manage: '管理', automatic: '自動目錄', custom: '自訂目錄…', update: '更新目錄',
    updatePages: '只更新頁碼', updateAll: '更新整個目錄', delete: '移除目錄',
    levels: '顯示階層', format: '格式', showPageNumbers: '顯示頁碼', rightAlignPageNumbers: '頁碼靠右對齊',
    tabLeader: '定位點前置字元', leaderNone: '無', leaderDots: '點線', leaderDashes: '虛線', leaderUnderline: '底線',
    formatTemplate: '來自範本', formatClassic: '傳統', formatModern: '現代', formatSimple: '簡單',
    noHeadings: '找不到標題。請先套用標題 1–3 樣式，或選擇相符的階層。', preview: '預覽列印', previewHeading: '標題',
  },
}

const translationKeys: Record<keyof DocumentTocLocaleText, string> = {
  insert: 'insert', contentsTitle: 'contents-title', manage: 'manage', automatic: 'automatic', custom: 'custom', update: 'update', updatePages: 'update-pages',
  updateAll: 'update-all', delete: 'delete', levels: 'levels', format: 'format', showPageNumbers: 'show-page-numbers',
  rightAlignPageNumbers: 'right-align-page-numbers', tabLeader: 'tab-leader', leaderNone: 'leader-none', leaderDots: 'leader-dots',
  leaderDashes: 'leader-dashes', leaderUnderline: 'leader-underline', formatTemplate: 'format-template', formatClassic: 'format-classic',
  formatModern: 'format-modern', formatSimple: 'format-simple', noHeadings: 'no-headings', preview: 'preview', previewHeading: 'preview-heading',
}

export const documentTocLocales: Record<string, Record<string, string>> = Object.fromEntries(
  Object.entries(localeText).map(([locale, strings]) => [
    locale,
    Object.fromEntries(Object.entries(strings).map(([key, value]) => [
      `document.toc.${translationKeys[key as keyof DocumentTocLocaleText]}`,
      value,
    ])),
  ]),
)
