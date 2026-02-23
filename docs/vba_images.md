## Image insertion methods

Источник: `extracted_vba/modules/Module1.bas`, процедура `UpdateStampFlexible`.

Используется один метод вставки изображений:

- `Worksheet.Shapes.AddPicture` с параметрами:
  - `LinkToFile:=msoFalse`, `SaveWithDocument:=msoTrue`
  - Позиционирование по **координатам ячейки**: `Left:=Range.Left`, `Top:=Range.Top`
  - Размеры фиксированные (в пикселях/поинтах Excel):
    - logo: `Width:=200`, `Height:=100`
    - stamp: `Width:=150`, `Height:=150`
    - director sign: `Width:=150`, `Height:=150`
    - client sign: `Width:=200`, `Height:=100`

Перед вставкой делается очистка:

- удаляются **все** фигуры типа `msoPicture` на целевом листе:
  - `For Each s In ws.Shapes: If s.Type = msoPicture Then s.Delete`

## Image source (where paths come from)

Пути к файлам берутся **не из кода**, а из листов книги:

1) Выбор компании:

- `companyName = ThisWorkbook.Sheets("ввод данных").Range("B11").Value`

2) Поиск строкой на листе `компании` (колонка A) и извлечение путей:

- `stampPath = Cells(i, 4)` (колонка D)
- `logoPath = Cells(i, 5)` (колонка E)
- `signDirPath = Cells(i, 6)` (колонка F)
- `signClientPath = Cells(i, 7)` (колонка G)

Перед вставкой файл проверяется через `Dir(path) <> ""`. Если файла нет — вставка **молча пропускается**.

## Image anchors (cells/shapes/names)

Якоря (куда вставлять) задаются на листе `Конфиг`:

- колонка A: `sheetName` — имя листа, куда вставлять (должно совпадать с именем вкладки)
- колонка B: `needLogo` (TRUE/FALSE)
- колонка C: `logoCell` (адрес ячейки, например `B2`)
- колонка D: `needStamp`
- колонка E: `stampCell`
- колонка F: `needSignDir`
- колонка G: `signDirCell`
- колонка H: `needSignClient`
- колонка I: `clientSignCell`

Важно: `sheetName` должен быть **точно** таким же, как имя листа в книге. Список точных имён листов см. в `extracted_vba/workbook_map.md` (например: `ввод данных`, `договор`, `компании`, `Конфиг`, `т-банк`, `Райф` и т.д.).

## What Excel Runner must provide (files on disk vs storage)

Runner должен обеспечить:

- **Файлы картинок на диске** по путям из листа `компании` (колонки D–G).
  - Формат не ограничен кодом (Excel обычно принимает PNG/JPG), но фактически ориентируйтесь на PNG.
  - Рекомендуется использовать **абсолютные пути** (иначе относительные пути зависят от текущей директории процесса Excel).
- **Корректную таблицу `Конфиг`**:
  - перечисление листов, флаги, адреса якорных ячеек.
- **Заполненную `ввод данных`!B11** (название компании должно совпадать с `компании`!A).

