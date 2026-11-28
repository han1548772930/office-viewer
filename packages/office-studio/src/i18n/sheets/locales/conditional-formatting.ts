const keys = [
  'items.sheet.operation.open-conditional-formatting', 'conditional.title', 'conditional.rules', 'conditional.add',
  'conditional.range', 'conditional.operator', 'conditional.value', 'conditional.second-value', 'conditional.fill',
  'conditional.font-color', 'conditional.bold', 'conditional.stop-if-true', 'conditional.save-rule', 'conditional.update',
  'conditional.delete', 'conditional.move-up', 'conditional.move-down', 'conditional.error.invalid', 'conditional.error.range',
  'conditional.operator.equal', 'conditional.operator.notEqual', 'conditional.operator.greaterThan',
  'conditional.operator.greaterThanOrEqual', 'conditional.operator.lessThan', 'conditional.operator.lessThanOrEqual',
  'conditional.operator.between', 'conditional.operator.notBetween',
] as const

const localeValues: Record<string, readonly string[]> = {
  'ar-SA': ['التنسيق الشرطي','التنسيق الشرطي','القواعد','قاعدة جديدة','تطبيق على','الشرط','القيمة','القيمة الثانية','التعبئة','لون الخط','عريض','إيقاف عند تحقق الشرط','حفظ القاعدة','تحديث القاعدة','حذف القاعدة','نقل القاعدة لأعلى','نقل القاعدة لأسفل','قاعدة التنسيق الشرطي غير صالحة.','أدخل نطاقاً صالحاً واحداً أو أكثر.','يساوي','لا يساوي','أكبر من','أكبر من أو يساوي','أصغر من','أصغر من أو يساوي','بين','ليس بين'],
  'ca-ES': ['Format condicional','Format condicional','Regles','Regla nova','Aplica a','Condició','Valor','Segon valor','Farciment','Color del text','Negreta','Atura si es compleix','Desa la regla','Actualitza la regla','Suprimeix la regla','Mou la regla amunt','Mou la regla avall','La regla de format condicional no és vàlida.','Introdueix un o més intervals de cel·les vàlids.','És igual a','No és igual a','Més gran que','Més gran o igual que','Més petit que','Més petit o igual que','Entre','Fora de'],
  'de-DE': ['Bedingte Formatierung','Bedingte Formatierung','Regeln','Neue Regel','Anwenden auf','Bedingung','Wert','Zweiter Wert','Füllung','Schriftfarbe','Fett','Bei Treffer stoppen','Regel speichern','Regel aktualisieren','Regel löschen','Regel nach oben','Regel nach unten','Die Regel für bedingte Formatierung ist ungültig.','Geben Sie einen oder mehrere gültige Zellbereiche ein.','Gleich','Ungleich','Größer als','Größer oder gleich','Kleiner als','Kleiner oder gleich','Zwischen','Nicht zwischen'],
  'es-ES': ['Formato condicional','Formato condicional','Reglas','Nueva regla','Aplicar a','Condición','Valor','Segundo valor','Relleno','Color de fuente','Negrita','Detener si se cumple','Guardar regla','Actualizar regla','Eliminar regla','Subir regla','Bajar regla','La regla de formato condicional no es válida.','Introduce uno o varios rangos de celdas válidos.','Igual a','Distinto de','Mayor que','Mayor o igual que','Menor que','Menor o igual que','Entre','No está entre'],
  'fa-IR': ['قالب‌بندی شرطی','قالب‌بندی شرطی','قواعد','قاعده جدید','اعمال به','شرط','مقدار','مقدار دوم','رنگ زمینه','رنگ قلم','پررنگ','در صورت تطبیق متوقف شود','ذخیره قاعده','به‌روزرسانی قاعده','حذف قاعده','انتقال قاعده به بالا','انتقال قاعده به پایین','قاعده قالب‌بندی شرطی نامعتبر است.','یک یا چند محدوده سلولی معتبر وارد کنید.','برابر با','نابرابر با','بزرگ‌تر از','بزرگ‌تر یا مساوی','کوچک‌تر از','کوچک‌تر یا مساوی','بین','خارج از بازه'],
  'fr-FR': ['Mise en forme conditionnelle','Mise en forme conditionnelle','Règles','Nouvelle règle','Appliquer à','Opérateur','Valeur','Deuxième valeur','Remplissage','Couleur de police','Gras','Arrêter si vrai','Enregistrer la règle','Mettre à jour la règle','Supprimer la règle','Monter la règle','Descendre la règle','La règle de mise en forme conditionnelle est invalide.','Saisissez une ou plusieurs plages de cellules valides.','Égal à','Différent de','Supérieur à','Supérieur ou égal à','Inférieur à','Inférieur ou égal à','Entre','Pas entre'],
  'id-ID': ['Pemformatan Bersyarat','Pemformatan Bersyarat','Aturan','Aturan Baru','Terapkan Ke','Kondisi','Nilai','Nilai Kedua','Isian','Warna Font','Tebal','Hentikan Jika Benar','Simpan Aturan','Perbarui Aturan','Hapus Aturan','Pindahkan Aturan ke Atas','Pindahkan Aturan ke Bawah','Aturan pemformatan bersyarat tidak valid.','Masukkan satu atau beberapa rentang sel yang valid.','Sama Dengan','Tidak Sama Dengan','Lebih Dari','Lebih Dari atau Sama Dengan','Kurang Dari','Kurang Dari atau Sama Dengan','Di Antara','Tidak Di Antara'],
  'it-IT': ['Formattazione condizionale','Formattazione condizionale','Regole','Nuova regola','Applica a','Condizione','Valore','Secondo valore','Riempimento','Colore carattere','Grassetto','Interrompi se vero','Salva regola','Aggiorna regola','Elimina regola','Sposta regola in alto','Sposta regola in basso','La regola di formattazione condizionale non è valida.','Inserisci uno o più intervalli di celle validi.','Uguale a','Diverso da','Maggiore di','Maggiore o uguale a','Minore di','Minore o uguale a','Compreso tra','Non compreso tra'],
  'ja-JP': ['条件付き書式','条件付き書式','ルール','新しいルール','適用先','条件','値','2番目の値','塗りつぶし','文字色','太字','条件を満たしたら停止','ルールを保存','ルールを更新','ルールを削除','ルールを上へ移動','ルールを下へ移動','条件付き書式ルールが無効です。','有効なセル範囲を1つ以上入力してください。','次の値に等しい','次の値に等しくない','次の値より大きい','次の値以上','次の値より小さい','次の値以下','次の値の間','次の値の間以外'],
  'ko-KR': ['조건부 서식','조건부 서식','규칙','새 규칙','적용 대상','조건','값','두 번째 값','채우기','글꼴 색','굵게','참이면 중지','규칙 저장','규칙 업데이트','규칙 삭제','규칙 위로 이동','규칙 아래로 이동','조건부 서식 규칙이 올바르지 않습니다.','유효한 셀 범위를 하나 이상 입력하세요.','같음','같지 않음','보다 큼','크거나 같음','보다 작음','작거나 같음','사이','사이가 아님'],
  'pl-PL': ['Formatowanie warunkowe','Formatowanie warunkowe','Reguły','Nowa reguła','Zastosuj do','Warunek','Wartość','Druga wartość','Wypełnienie','Kolor czcionki','Pogrubienie','Zatrzymaj, gdy prawda','Zapisz regułę','Zaktualizuj regułę','Usuń regułę','Przenieś regułę w górę','Przenieś regułę w dół','Reguła formatowania warunkowego jest nieprawidłowa.','Wprowadź co najmniej jeden prawidłowy zakres komórek.','Równe','Różne od','Większe niż','Większe lub równe','Mniejsze niż','Mniejsze lub równe','Pomiędzy','Poza zakresem'],
  'pt-BR': ['Formatação Condicional','Formatação Condicional','Regras','Nova Regra','Aplicar a','Condição','Valor','Segundo Valor','Preenchimento','Cor da Fonte','Negrito','Parar se Verdadeiro','Salvar Regra','Atualizar Regra','Excluir Regra','Mover Regra para Cima','Mover Regra para Baixo','A regra de formatação condicional é inválida.','Digite um ou mais intervalos de células válidos.','Igual a','Diferente de','Maior que','Maior ou igual a','Menor que','Menor ou igual a','Entre','Fora do intervalo'],
  'ru-RU': ['Условное форматирование','Условное форматирование','Правила','Новое правило','Применить к','Условие','Значение','Второе значение','Заливка','Цвет шрифта','Полужирный','Остановить при совпадении','Сохранить правило','Обновить правило','Удалить правило','Переместить правило вверх','Переместить правило вниз','Правило условного форматирования недопустимо.','Введите один или несколько допустимых диапазонов ячеек.','Равно','Не равно','Больше','Больше или равно','Меньше','Меньше или равно','Между','Вне диапазона'],
  'sk-SK': ['Podmienené formátovanie','Podmienené formátovanie','Pravidlá','Nové pravidlo','Použiť na','Podmienka','Hodnota','Druhá hodnota','Výplň','Farba písma','Tučné','Zastaviť pri zhode','Uložiť pravidlo','Aktualizovať pravidlo','Odstrániť pravidlo','Presunúť pravidlo nahor','Presunúť pravidlo nadol','Pravidlo podmieneného formátovania je neplatné.','Zadajte jeden alebo viac platných rozsahov buniek.','Rovná sa','Nerovná sa','Väčšie ako','Väčšie alebo rovné','Menšie ako','Menšie alebo rovné','Medzi','Mimo rozsahu'],
  'vi-VN': ['Định dạng có điều kiện','Định dạng có điều kiện','Quy tắc','Quy tắc mới','Áp dụng cho','Điều kiện','Giá trị','Giá trị thứ hai','Tô màu','Màu phông chữ','Đậm','Dừng nếu đúng','Lưu quy tắc','Cập nhật quy tắc','Xóa quy tắc','Di chuyển quy tắc lên','Di chuyển quy tắc xuống','Quy tắc định dạng có điều kiện không hợp lệ.','Nhập một hoặc nhiều phạm vi ô hợp lệ.','Bằng','Không bằng','Lớn hơn','Lớn hơn hoặc bằng','Nhỏ hơn','Nhỏ hơn hoặc bằng','Trong khoảng','Ngoài khoảng'],
  'zh-HK': ['條件格式','條件格式','規則','新增規則','套用到','條件','數值','第二個數值','填滿','文字顏色','粗體','符合時停止','儲存規則','更新規則','刪除規則','上移規則','下移規則','條件格式規則無效。','請輸入一個或多個有效儲存格範圍。','等於','不等於','大於','大於或等於','小於','小於或等於','介乎','不介乎'],
  'zh-TW': ['條件格式','條件格式','規則','新增規則','套用到','條件','數值','第二個數值','填滿','文字顏色','粗體','符合時停止','儲存規則','更新規則','刪除規則','上移規則','下移規則','條件格式規則無效。','請輸入一個或多個有效儲存格範圍。','等於','不等於','大於','大於或等於','小於','小於或等於','介於','不介於'],
}

const quickKeys = ['data-bar', 'icon-set', 'rank', 'shape', 'direction', 'create', 'clear'] as const
const quickValues: Record<string, readonly string[]> = {
  'ar-SA': ['شريط البيانات', 'مجموعة أيقونات', 'الترتيب', 'الشكل', 'الاتجاه', 'إنشاء تنسيق شرطي', 'مسح القواعد للنطاق المحدد'],
  'ca-ES': ['Barra de dades', "Conjunt d'icones", 'Rang', 'Forma', 'Direcció', 'Crea format condicional', 'Esborra les regles del rang seleccionat'],
  'de-DE': ['Datenbalken', 'Symbolsatz', 'Rang', 'Form', 'Richtung', 'Bedingte Formatierung erstellen', 'Regeln für ausgewählten Bereich löschen'],
  'en-US': ['Data Bar', 'Icon Set', 'Rank', 'Shape', 'Direction', 'Create Conditional Formatting', 'Clear Rules for Selected Range'],
  'es-ES': ['Barra de datos', 'Conjunto de iconos', 'Rango', 'Forma', 'Dirección', 'Crear formato condicional', 'Borrar reglas del rango seleccionado'],
  'fa-IR': ['نوار داده', 'مجموعه آیکون', 'رتبه', 'شکل', 'جهت', 'ایجاد قالب‌بندی شرطی', 'پاک کردن قوانین برای محدوده انتخاب شده'],
  'fr-FR': ['Barre de données', "Jeu d'icônes", 'Rang', 'Forme', 'Direction', 'Créer une mise en forme conditionnelle', 'Effacer les règles pour la plage sélectionnée'],
  'id-ID': ['Bilah Data', 'Set Ikon', 'Peringkat', 'Bentuk', 'Arah', 'Buat Format Bersyarat', 'Hapus Aturan untuk Rentang yang Dipilih'],
  'it-IT': ['Barra Dati', 'Set di Icone', 'Grado', 'Forma', 'Direzione', 'Crea Formattazione Condizionale', 'Cancella Regole per Intervallo Selezionato'],
  'ja-JP': ['データバー', 'アイコンセット', '評価', '図形', '方向', '新しいルール', '選択したセルからルールをクリア'],
  'ko-KR': ['데이터 막대', '아이콘 집합', '순위', '모양', '방향', '조건부 서식 만들기', '선택 영역 규칙 지우기'],
  'pl-PL': ['Pasek danych', 'Zestaw ikon', 'Ranga', 'Kształt', 'Kierunek', 'Utwórz formatowanie warunkowe', 'Wyczyść reguły dla zaznaczonego zakresu'],
  'pt-BR': ['Barra de Dados', 'Conjunto de Ícones', 'Classificação', 'Forma', 'Direção', 'Criar Formatação Condicional', 'Limpar Regras do Intervalo Selecionado'],
  'ru-RU': ['Гистограмма данных', 'Набор иконок', 'Ранг', 'Форма', 'Направление', 'Создание условного форматирования', 'Очистить правила для выбранного диапазона'],
  'sk-SK': ['Dátový pruh', 'Súprava ikon', 'Poradie', 'Tvar', 'Smer', 'Vytvoriť podmienené formátovanie', 'Vymazať pravidlá pre vybraný rozsah'],
  'vi-VN': ['Thanh dữ liệu', 'Tập hợp biểu tượng', 'Hạng', 'Hình dạng', 'Hướng', 'Tạo định dạng có điều kiện mới', 'Xóa quy tắc cho vùng chọn'],
  'zh-CN': ['数据条', '图标集', '等级', '形状', '方向', '新建条件格式', '清除所选区域的规则'],
  'zh-HK': ['資料條', '圖示集', '等級', '形狀', '方向', '新建條件格式', '清除所選區域的規則'],
  'zh-TW': ['資料條', '圖示集', '等級', '形狀', '方向', '新建條件格式', '清除所選區域的規則'],
}

const scopeKeys = ['scope.worksheet', 'scope.selection', 'scope.label', 'clear-list', 'empty-selection'] as const
const scopeValues: Record<string, readonly string[]> = {
  'ar-SA': ['ورقة العمل', 'التحديد الحالي', 'إظهار قواعد', 'مسح القواعد المعروضة', 'لا توجد قواعد تنسيق شرطي للتحديد الحالي.'],
  'ca-ES': ['Full de càlcul', 'Selecció actual', 'Mostra les regles de', 'Esborra les regles de la llista', 'La selecció actual no té regles de format condicional.'],
  'de-DE': ['Arbeitsblatt', 'Aktuelle Auswahl', 'Regeln anzeigen für', 'Regeln in der Liste löschen', 'Für die aktuelle Auswahl sind keine Regeln für bedingte Formatierung vorhanden.'],
  'en-US': ['Worksheet', 'Current Selection', 'Show Rules for', 'Clear Listed Rules', 'The current selection has no conditional formatting rules.'],
  'es-ES': ['Hoja de cálculo', 'Selección actual', 'Mostrar reglas para', 'Borrar las reglas de la lista', 'La selección actual no tiene reglas de formato condicional.'],
  'fa-IR': ['کاربرگ', 'انتخاب فعلی', 'نمایش قواعد برای', 'پاک کردن قواعد فهرست', 'انتخاب فعلی قاعده قالب‌بندی شرطی ندارد.'],
  'fr-FR': ['Feuille de calcul', 'Sélection actuelle', 'Afficher les règles pour', 'Effacer les règles de la liste', 'La sélection actuelle ne contient aucune règle de mise en forme conditionnelle.'],
  'id-ID': ['Lembar Kerja', 'Pilihan Saat Ini', 'Tampilkan Aturan untuk', 'Hapus Aturan dalam Daftar', 'Pilihan saat ini tidak memiliki aturan pemformatan bersyarat.'],
  'it-IT': ['Foglio di lavoro', 'Selezione corrente', 'Mostra regole per', 'Cancella le regole elencate', 'La selezione corrente non contiene regole di formattazione condizionale.'],
  'ja-JP': ['ワークシート', '現在の選択範囲', 'ルールの表示対象', '一覧のルールをクリア', '現在の選択範囲に条件付き書式ルールはありません。'],
  'ko-KR': ['워크시트', '현재 선택 영역', '규칙 표시 대상', '목록의 규칙 지우기', '현재 선택 영역에 조건부 서식 규칙이 없습니다.'],
  'pl-PL': ['Arkusz', 'Bieżące zaznaczenie', 'Pokaż reguły dla', 'Wyczyść reguły na liście', 'Bieżące zaznaczenie nie ma reguł formatowania warunkowego.'],
  'pt-BR': ['Planilha', 'Seleção Atual', 'Mostrar Regras para', 'Limpar Regras da Lista', 'A seleção atual não tem regras de formatação condicional.'],
  'ru-RU': ['Лист', 'Текущее выделение', 'Показать правила для', 'Очистить правила в списке', 'Для текущего выделения нет правил условного форматирования.'],
  'sk-SK': ['Hárok', 'Aktuálny výber', 'Zobraziť pravidlá pre', 'Vymazať pravidlá v zozname', 'Aktuálny výber nemá pravidlá podmieneného formátovania.'],
  'vi-VN': ['Trang tính', 'Vùng chọn hiện tại', 'Hiển thị quy tắc cho', 'Xóa quy tắc trong danh sách', 'Vùng chọn hiện tại không có quy tắc định dạng có điều kiện.'],
  'zh-CN': ['工作表', '当前选区', '显示规则范围', '清除列表规则', '当前选区没有条件格式规则。'],
  'zh-HK': ['工作表', '目前選取範圍', '顯示規則範圍', '清除清單規則', '目前選取範圍沒有條件格式規則。'],
  'zh-TW': ['工作表', '目前選取範圍', '顯示規則範圍', '清除清單規則', '目前選取範圍沒有條件格式規則。'],
}

const permissionKeys = ['drag', 'protected-rules', 'error.permission'] as const
const permissionValues: Record<string, readonly string[]> = {
  'ar-SA': ['اسحب لإعادة ترتيب القواعد', 'بعض القواعد محمية؛ تظهر القواعد القابلة للتحرير فقط.', 'لا يمكنك تحرير جميع نطاقات القاعدة أو تغيير الأولويات المتأثرة.'],
  'ca-ES': ['Arrossega per reordenar les regles', 'Algunes regles estan protegides; només es mostren les regles editables.', 'No pots editar tots els intervals de la regla ni canviar les prioritats afectades.'],
  'de-DE': ['Regeln durch Ziehen neu anordnen', 'Einige Regeln sind geschützt. Nur bearbeitbare Regeln werden angezeigt.', 'Sie können nicht alle Regelbereiche bearbeiten oder die betroffenen Prioritäten ändern.'],
  'en-US': ['Drag to reorder rules', 'Some rules are protected. Only editable rules are shown.', 'You cannot edit every range of the rule or change the affected priorities.'],
  'es-ES': ['Arrastrar para ordenar las reglas', 'Algunas reglas están protegidas. Solo se muestran las reglas editables.', 'No puedes editar todos los rangos de la regla ni cambiar las prioridades afectadas.'],
  'fa-IR': ['برای مرتب‌سازی قواعد بکشید', 'برخی قواعد محافظت شده‌اند؛ فقط قواعد قابل ویرایش نمایش داده می‌شوند.', 'اجازه ویرایش همه محدوده‌های قاعده یا تغییر اولویت‌های مرتبط را ندارید.'],
  'fr-FR': ['Faire glisser pour réordonner les règles', 'Certaines règles sont protégées. Seules les règles modifiables sont affichées.', 'Vous ne pouvez pas modifier toutes les plages de la règle ou les priorités concernées.'],
  'id-ID': ['Seret untuk mengurutkan aturan', 'Sebagian aturan dilindungi. Hanya aturan yang dapat diedit yang ditampilkan.', 'Anda tidak dapat mengedit semua rentang aturan atau mengubah prioritas terkait.'],
  'it-IT': ['Trascina per riordinare le regole', 'Alcune regole sono protette. Vengono mostrate solo quelle modificabili.', 'Non puoi modificare tutti gli intervalli della regola o le priorità interessate.'],
  'ja-JP': ['ドラッグしてルールを並べ替え', '保護されたルールがあります。編集可能なルールのみ表示します。', 'ルールのすべての適用範囲または関連する優先順位を変更する権限がありません。'],
  'ko-KR': ['드래그하여 규칙 순서 변경', '일부 규칙은 보호되어 있습니다. 편집 가능한 규칙만 표시됩니다.', '규칙의 모든 적용 범위 또는 관련 우선순위를 변경할 권한이 없습니다.'],
  'pl-PL': ['Przeciągnij, aby zmienić kolejność reguł', 'Niektóre reguły są chronione. Wyświetlane są tylko reguły edytowalne.', 'Nie możesz edytować wszystkich zakresów reguły ani zmieniać powiązanych priorytetów.'],
  'pt-BR': ['Arraste para reordenar as regras', 'Algumas regras estão protegidas. Apenas regras editáveis são mostradas.', 'Você não pode editar todos os intervalos da regra nem alterar as prioridades afetadas.'],
  'ru-RU': ['Перетащите для изменения порядка правил', 'Некоторые правила защищены. Показаны только доступные для редактирования правила.', 'Нет прав на изменение всех диапазонов правила или затронутых приоритетов.'],
  'sk-SK': ['Potiahnutím zmeniť poradie pravidiel', 'Niektoré pravidlá sú chránené. Zobrazujú sa iba upraviteľné pravidlá.', 'Nemôžete upraviť všetky rozsahy pravidla ani súvisiace priority.'],
  'vi-VN': ['Kéo để sắp xếp lại quy tắc', 'Một số quy tắc được bảo vệ. Chỉ hiển thị các quy tắc có thể chỉnh sửa.', 'Bạn không có quyền chỉnh sửa tất cả phạm vi của quy tắc hoặc các mức ưu tiên liên quan.'],
  'zh-CN': ['拖动调整规则顺序', '部分规则受保护，仅显示可编辑的规则。', '当前用户没有编辑规则全部应用范围或调整相关优先级的权限。'],
  'zh-HK': ['拖動調整規則次序', '部分規則受保護，只顯示可編輯的規則。', '目前使用者沒有編輯規則全部套用範圍或調整相關優先次序的權限。'],
  'zh-TW': ['拖曳調整規則順序', '部分規則受保護，只顯示可編輯的規則。', '目前使用者沒有編輯規則全部套用範圍或調整相關優先順序的權限。'],
}

export const conditionalFormattingLocaleMessages: Partial<Record<string, Record<string, string>>> = Object.fromEntries(
  Object.entries(quickValues).map(([locale, values]) => [locale, Object.fromEntries([
    ...(localeValues[locale] ?? []).map((value, index) => [keys[index]!, value]),
    ...values.map((value, index) => [`conditional.quick.${quickKeys[index]!}`, value]),
    ...scopeValues[locale]!.map((value, index) => [`conditional.${scopeKeys[index]!}`, value]),
    ...permissionValues[locale]!.map((value, index) => [`conditional.${permissionKeys[index]!}`, value]),
  ])]),
)
