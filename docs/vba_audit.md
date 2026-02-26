## Modules inventory

Источник: `extracted_vba/modules/*.bas` и события из `extracted_vba/classes/*.cls`.

### `Module1.bas`

- **Public Sub (в стандартном модуле `Sub` без модификатора считается Public)**:
  - `UpdateAllStampsFromConfig`
  - `UpdateStampFlexible(sheetName As String, needLogo As Boolean, logoCell As String, needStamp As Boolean, stampCell As String, needSignDir As Boolean, signDirCell As String, needSignClient As Boolean, clientSignCell As String)`
- **Public Function**:
  - `GetCompanyFiles(companyName As String, ByRef stampPath As String, ByRef logoPath As String, ByRef signDirPath As String, ByRef signClientPath As String) As Boolean`
- **Event / не для ручного запуска**:
  - `Private Sub Worksheet_Change(ByVal Target As Range)` (внутри `Me.Range("B11")` → вызов `UpdateAllStampsFromConfig`)

Файловые/листовые зависимости в этом модуле (точные имена листов):

- Читает `ThisWorkbook.Sheets("Конфиг")` (таблица конфигурации по строкам)
- Читает `ThisWorkbook.Sheets("ввод данных").Range("B11")` (имя компании)
- Читает `ThisWorkbook.Sheets("компании")` (пути к файлам картинок)

### `Module2.bas`

- **Public Function**:
  - `SpellNumberEN(value, [major], [minor], [britishAnd], [titleCase]) As String` (число → слова EN)
  - `SpellUSD(n) As String` (обёртка над `SpellNumberEN`)
  - `NumberToWordsENOnly(n) As String` (обёртка над `SpellNumberEN`)
- **Private Function**:
  - `ChunkToWordsEN(n, includeAnd) As String` (вспомогательная)

### Class modules (`extracted_vba/classes/*.cls`)

Код (события) найден только в `extracted_vba/classes/Лист3.cls`:

- `Private Sub Worksheet_Change(ByVal Target As Range)` → при изменении `B11` вызывает `UpdateAllStampsFromConfig`
- `Private Sub Worksheet_Calculate()` → отслеживает изменение `B11` после пересчёта и вызывает `UpdateAllStampsFromConfig`
- `Private Sub Worksheet_Activate()` → при активации листа вызывает `UpdateAllStampsFromConfig`

`extracted_vba/classes/ЭтаКнига.cls` (ThisWorkbook) — без обработчиков событий.

## Entrypoint candidates

Кандидаты на “главный макрос” (верхний уровень логики):

- **Рекомендуемый**: `Module1.UpdateAllStampsFromConfig`
  - Массово проходит по конфигу, выбирает листы, решает какие картинки вставлять, и вызывает всю остальную логику.
- **Альтернативный (event-driven triggers, не для runner’а)**: события листа `Лист3`:
  - `Лист3.Worksheet_Change` / `Лист3.Worksheet_Calculate` / `Лист3.Worksheet_Activate`
  - Все три просто вызывают `UpdateAllStampsFromConfig` при изменении/пересчёте `B11` или при входе на лист.
- **Низкоуровневый**: `Module1.UpdateStampFlexible(...)`
  - Обновляет один лист по параметрам (удобно для ручной отладки, но для runner’а неудобно из‑за большого списка аргументов).

## Selected entrypoint

Entrypoint: **`Module1.UpdateAllStampsFromConfig`**

Why this:

- Это единственная процедура верхнего уровня, которая **обходит весь список документов/листов** через `ThisWorkbook.Sheets("Конфиг")` и запускает обновление для каждого листа.
- Внутри она вызывает `UpdateStampFlexible`, где сосредоточена основная логика: чтение `ввод данных`!`B11`, выбор путей на листе `компании`, удаление старых картинок и вставка новых через `Shapes.AddPicture`.
- События листа `Лист3` (Change/Calculate/Activate) фактически являются лишь “триггерами” для этой процедуры, но для автоматизации они менее управляемы.
- В извлечённом VBA не обнаружено альтернативных “Generate/Export/Print/PDF” макросов — функционально книга автоматизирует именно **обновление логотипов/печатей/подписей** по конфигурации.

