import type { OfficeLocale } from '../../office-i18n.ts'

const labels: Record<OfficeLocale, Record<string, string>> = {
  'ar-SA': { 'doc.table-insert.title': 'إدراج جدول', 'doc.table-insert.rows': 'صفوف', 'doc.table-insert.columns': 'أعمدة', 'common.insert': 'إدراج' },
  'ca-ES': { 'doc.table-insert.title': 'Insereix una taula', 'doc.table-insert.rows': 'Files', 'doc.table-insert.columns': 'Columnes', 'common.insert': 'Insereix' },
  'de-DE': { 'doc.table-insert.title': 'Tabelle einfügen', 'doc.table-insert.rows': 'Zeilen', 'doc.table-insert.columns': 'Spalten', 'common.insert': 'Einfügen' },
  'en-US': { 'doc.table-insert.title': 'Insert Table', 'doc.table-insert.rows': 'Rows', 'doc.table-insert.columns': 'Columns', 'common.insert': 'Insert' },
  'es-ES': { 'doc.table-insert.title': 'Insertar tabla', 'doc.table-insert.rows': 'Filas', 'doc.table-insert.columns': 'Columnas', 'common.insert': 'Insertar' },
  'fa-IR': { 'doc.table-insert.title': 'درج جدول', 'doc.table-insert.rows': 'سطرها', 'doc.table-insert.columns': 'ستون‌ها', 'common.insert': 'درج' },
  'fr-FR': { 'doc.table-insert.title': 'Insérer un tableau', 'doc.table-insert.rows': 'Lignes', 'doc.table-insert.columns': 'Colonnes', 'common.insert': 'Insérer' },
  'id-ID': { 'doc.table-insert.title': 'Sisipkan Tabel', 'doc.table-insert.rows': 'Baris', 'doc.table-insert.columns': 'Kolom', 'common.insert': 'Sisipkan' },
  'it-IT': { 'doc.table-insert.title': 'Inserisci tabella', 'doc.table-insert.rows': 'Righe', 'doc.table-insert.columns': 'Colonne', 'common.insert': 'Inserisci' },
  'ja-JP': { 'doc.table-insert.title': '表を挿入', 'doc.table-insert.rows': '行', 'doc.table-insert.columns': '列', 'common.insert': '挿入' },
  'ko-KR': { 'doc.table-insert.title': '표 삽입', 'doc.table-insert.rows': '행', 'doc.table-insert.columns': '열', 'common.insert': '삽입' },
  'pl-PL': { 'doc.table-insert.title': 'Wstaw tabelę', 'doc.table-insert.rows': 'Wiersze', 'doc.table-insert.columns': 'Kolumny', 'common.insert': 'Wstaw' },
  'pt-BR': { 'doc.table-insert.title': 'Inserir tabela', 'doc.table-insert.rows': 'Linhas', 'doc.table-insert.columns': 'Colunas', 'common.insert': 'Inserir' },
  'ru-RU': { 'doc.table-insert.title': 'Вставить таблицу', 'doc.table-insert.rows': 'Строки', 'doc.table-insert.columns': 'Столбцы', 'common.insert': 'Вставить' },
  'sk-SK': { 'doc.table-insert.title': 'Vložiť tabuľku', 'doc.table-insert.rows': 'Riadky', 'doc.table-insert.columns': 'Stĺpce', 'common.insert': 'Vložiť' },
  'vi-VN': { 'doc.table-insert.title': 'Chèn bảng', 'doc.table-insert.rows': 'Hàng', 'doc.table-insert.columns': 'Cột', 'common.insert': 'Chèn' },
  'zh-CN': { 'doc.table-insert.title': '插入表格', 'doc.table-insert.rows': '行数', 'doc.table-insert.columns': '列数', 'common.insert': '插入' },
  'zh-HK': { 'doc.table-insert.title': '插入表格', 'doc.table-insert.rows': '行數', 'doc.table-insert.columns': '列數', 'common.insert': '插入' },
  'zh-TW': { 'doc.table-insert.title': '插入表格', 'doc.table-insert.rows': '行數', 'doc.table-insert.columns': '欄數', 'common.insert': '插入' },
}

export const documentTableInsertLocales: Record<OfficeLocale, Record<string, string>> = labels
