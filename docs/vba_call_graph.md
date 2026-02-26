## Selected entrypoint call graph

Entrypoint: `Module1.UpdateAllStampsFromConfig` (файл `extracted_vba/modules/Module1.bas`)

### Call tree (3+ levels)

- `Module1.UpdateAllStampsFromConfig` (`Module1.bas`)
  - **Что делает**: читает таблицу на листе `Конфиг`, и для каждой строки конфигурации вызывает обновление печатей/подписей/логотипа на заданном листе.
  - **I/O**:
    - Читает: `ThisWorkbook.Sheets("Конфиг")`, колонки A:I (1..9), строки 2..lastRow
    - Для каждой строки: `sheetName` (A), флаги `needLogo/needStamp/needSignDir/needSignClient` (B/D/F/H), якорные ячейки (C/E/G/I)
  - Calls:
    - `Module1.UpdateStampFlexible(sheetName, needLogo, logoCell, needStamp, stampCell, needSignDir, signDirCell, needSignClient, clientSignCell)` (`Module1.bas`)

  - `Module1.UpdateStampFlexible(...)` (`Module1.bas`)
    - **Что делает**: для конкретного листа удаляет старые picture-shapes и (по флагам) вставляет 4 изображения: logo / stamp / director sign / client sign.
    - **I/O**:
      - Читает: `ThisWorkbook.Sheets("ввод данных").Range("B11")` → `companyName`
      - Проверяет существование листа: `ThisWorkbook.Sheets(sheetName)` (имя должно совпадать с именем вкладки)
      - Якори вставки: `ws.Range(logoCell/stampCell/signDirCell/clientSignCell)` (адреса из `Конфиг`)
      - Файлы: проверка `Dir(pathX) <> ""` перед вставкой
    - Calls:
      - `Module1.GetCompanyFiles(companyName, pathStamp, pathLogo, pathSignDir, pathSignClient)` (`Module1.bas`)
    - Leaf-операции (внутри процедуры):
      - **Удаление картинок**: `For Each s In ws.Shapes: If s.Type = msoPicture Then s.Delete`
      - **Вставка картинок**:
        - `ws.Shapes.AddPicture Filename:=pathLogo, Left:=logoRange.Left, Top:=logoRange.Top, Width:=200, Height:=100`
        - `ws.Shapes.AddPicture Filename:=pathStamp, Left:=stampRange.Left, Top:=stampRange.Top, Width:=150, Height:=150`
        - `ws.Shapes.AddPicture Filename:=pathSignDir, Left:=signDirRange.Left, Top:=signDirRange.Top, Width:=150, Height:=150`
        - `ws.Shapes.AddPicture Filename:=pathSignClient, Left:=clientSignRange.Left, Top:=clientSignRange.Top, Width:=200, Height:=100`

    - `Module1.GetCompanyFiles(...) As Boolean` (`Module1.bas`)
      - **Что делает**: на листе `компании` ищет строку по имени компании и возвращает пути к 4 файлам.
      - **I/O**:
        - Читает: `ThisWorkbook.Sheets("компании")`
        - Поиск: колонка A (1) со строки 2 до `lastRow`
        - Возвращает пути из колонок:
          - D (4): печать (`stampPath`)
          - E (5): логотип (`logoPath`)
          - F (6): подпись директора (`signDirPath`)
          - G (7): подпись клиента (`signClientPath`)

### Event-driven triggers (не основной entrypoint для runner’а)

Файл: `extracted_vba/classes/Лист3.cls`

- `Лист3.Worksheet_Change` → если меняется `B11`, вызывает `UpdateAllStampsFromConfig`
- `Лист3.Worksheet_Calculate` → если после пересчёта изменился `B11`, вызывает `UpdateAllStampsFromConfig`
- `Лист3.Worksheet_Activate` → при переходе на лист вызывает `UpdateAllStampsFromConfig`

Примечание по соответствию листа: модуль с кодовым именем `Лист3` обычно соответствует 3‑й вкладке книги; в `extracted_vba/workbook_map.md` 3‑й лист — `договор`. Если это критично для automation, проверяйте соответствие в самом Excel (Developer → Properties).

