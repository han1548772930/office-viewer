import type { OfficeLocale } from '../../regional-settings.ts'

const labels = {
  'ar-SA': ['بحث عن الميزات', 'اكتب اسم ميزة أو قائمة…', 'لم يتم العثور على ميزات متاحة', 'الشريط'],
  'ca-ES': ['Cerca funcions', 'Escriviu una funció o un nom de menú…', 'No s’han trobat funcions disponibles', 'Cinta'],
  'de-DE': ['Funktionen suchen', 'Funktion oder Menüname eingeben…', 'Keine verfügbaren Funktionen gefunden', 'Menüband'],
  'en-US': ['Search features', 'Type a feature or menu name…', 'No available features found', 'Ribbon'],
  'es-ES': ['Buscar funciones', 'Escribe una función o nombre de menú…', 'No se encontraron funciones disponibles', 'Cinta'],
  'fa-IR': ['جستجوی قابلیت‌ها', 'نام قابلیت یا منو را وارد کنید…', 'هیچ قابلیت در دسترسی یافت نشد', 'نوار ابزار'],
  'fr-FR': ['Rechercher des fonctionnalités', 'Saisissez une fonctionnalité ou un nom de menu…', 'Aucune fonctionnalité disponible trouvée', 'Ruban'],
  'id-ID': ['Cari fitur', 'Ketik nama fitur atau menu…', 'Tidak ada fitur yang tersedia', 'Pita'],
  'it-IT': ['Cerca funzionalità', 'Digita una funzionalità o il nome di un menu…', 'Nessuna funzionalità disponibile trovata', 'Barra multifunzione'],
  'ja-JP': ['機能を検索', '機能またはメニュー名を入力…', '利用可能な機能が見つかりません', 'リボン'],
  'ko-KR': ['기능 검색', '기능 또는 메뉴 이름 입력…', '사용 가능한 기능을 찾을 수 없습니다', '리본'],
  'pl-PL': ['Wyszukaj funkcje', 'Wpisz nazwę funkcji lub menu…', 'Nie znaleziono dostępnych funkcji', 'Wstążka'],
  'pt-BR': ['Pesquisar recursos', 'Digite um recurso ou nome de menu…', 'Nenhum recurso disponível encontrado', 'Faixa de opções'],
  'ru-RU': ['Поиск функций', 'Введите название функции или меню…', 'Доступные функции не найдены', 'Лента'],
  'sk-SK': ['Vyhľadať funkcie', 'Zadajte názov funkcie alebo ponuky…', 'Nenašli sa žiadne dostupné funkcie', 'Pás s nástrojmi'],
  'vi-VN': ['Tìm kiếm tính năng', 'Nhập tên tính năng hoặc menu…', 'Không tìm thấy tính năng khả dụng', 'Dải băng'],
  'zh-CN': ['搜索功能', '输入功能或菜单名称…', '未找到当前可用的功能', '功能区'],
  'zh-HK': ['搜尋功能', '輸入功能或選單名稱…', '找不到目前可用的功能', '功能區'],
  'zh-TW': ['搜尋功能', '輸入功能或選單名稱…', '找不到目前可用的功能', '功能區'],
} satisfies Record<OfficeLocale, readonly [string, string, string, string]>

export function featureSearchMessages(locale: OfficeLocale) {
  const [title, placeholder, empty, ribbon] = labels[locale]
  return { 'feature-search.title': title, 'feature-search.placeholder': placeholder,
    'feature-search.empty': empty, 'feature-search.ribbon': ribbon }
}
