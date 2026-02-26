Attribute VB_Name = "Module1"
'  Срабатывает при изменении B11
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B11")) Is Nothing Then
        Call UpdateAllStampsFromConfig
    End If
End Sub
'  Читает лист "Конфиг" и обновляет все листы
Sub UpdateAllStampsFromConfig()
    Dim cfgSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    Dim needLogo As Boolean, logoCell As String
    Dim needStamp As Boolean, stampCell As String
    Dim needSignDir As Boolean, signDirCell As String
    Dim needSignClient As Boolean, clientSignCell As String
    ' Лист с конфигурацией
    Set cfgSheet = ThisWorkbook.Sheets("Конфиг")
    ' Сколько строк в таблице
    lastRow = cfgSheet.Cells(cfgSheet.Rows.Count, 1).End(xlUp).Row
    ' Проходим по всем строкам
    For i = 2 To lastRow
        sheetName = Trim(cfgSheet.Cells(i, 1).value)
        needLogo = (UCase(cfgSheet.Cells(i, 2).value) = "TRUE")
        logoCell = Trim(cfgSheet.Cells(i, 3).value)
        needStamp = (UCase(cfgSheet.Cells(i, 4).value) = "TRUE")
        stampCell = Trim(cfgSheet.Cells(i, 5).value)
        needSignDir = (UCase(cfgSheet.Cells(i, 6).value) = "TRUE")
        signDirCell = Trim(cfgSheet.Cells(i, 7).value)
        needSignClient = (UCase(cfgSheet.Cells(i, 8).value) = "TRUE")
        clientSignCell = Trim(cfgSheet.Cells(i, 9).value)
        ' Обновляем конкретный лист по этим настройкам
        Call UpdateStampFlexible(sheetName, _
                                 needLogo, logoCell, _
                                 needStamp, stampCell, _
                                 needSignDir, signDirCell, _
                                 needSignClient, clientSignCell)
    Next i
End Sub
'  Обновляет конкретный лист (лого, печать, подпись директора, подпись клиента)
Sub UpdateStampFlexible(sheetName As String, _
                        needLogo As Boolean, logoCell As String, _
                        needStamp As Boolean, stampCell As String, _
                        needSignDir As Boolean, signDirCell As String, _
                        needSignClient As Boolean, clientSignCell As String)
    Dim ws As Worksheet, wsInput As Worksheet
    Dim logoRange As Range, stampRange As Range
    Dim signDirRange As Range, clientSignRange As Range
    Dim shp As Shape
    Dim companyName As String
    ' Пути к файлам
    Dim pathStamp As String, pathLogo As String
    Dim pathSignDir As String, pathSignClient As String
    ' Читаем название компании из B11
    Set wsInput = ThisWorkbook.Sheets("ввод данных")
    companyName = Trim(wsInput.Range("B11").value)
    '  Получаем пути для этой компании с листа "компании"
    If Not GetCompanyFiles(companyName, pathStamp, pathLogo, pathSignDir, pathSignClient) Then
        Exit Sub ' если компания не найдена – ничего не делаем
    End If
    ' Проверяем наличие листа
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    '  Удаляем все старые картинки (Shapes)
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Type = msoPicture Then s.Delete
    Next s
    ' --- ЛОГО ---
    If needLogo And logoCell <> "" And pathLogo <> "" Then
        Set logoRange = ws.Range(logoCell)
        If Dir(pathLogo) <> "" Then
            ws.Shapes.AddPicture Filename:=pathLogo, _
                LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                Left:=logoRange.Left, Top:=logoRange.Top, _
                Width:=200, Height:=100
        End If
    End If
    ' --- ПЕЧАТЬ ---
    If needStamp And stampCell <> "" And pathStamp <> "" Then
        Set stampRange = ws.Range(stampCell)
        If Dir(pathStamp) <> "" Then
            ws.Shapes.AddPicture Filename:=pathStamp, _
                LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                Left:=stampRange.Left, Top:=stampRange.Top, _
                Width:=150, Height:=150
        End If
    End If
    
      ' --- ПОДПИСЬ ДИРЕКТОРА ---
    If needSignDir And signDirCell <> "" And pathSignDir <> "" Then
        Set signDirRange = ws.Range(signDirCell)
        If Dir(pathSignDir) <> "" Then
            ws.Shapes.AddPicture Filename:=pathSignDir, _
                LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                Left:=signDirRange.Left, Top:=signDirRange.Top, _
                Width:=150, Height:=150
        End If
    End If
    ' --- ПОДПИСЬ КЛИЕНТА ---
    If needSignClient And clientSignCell <> "" And pathSignClient <> "" Then
        Set clientSignRange = ws.Range(clientSignCell)
        If Dir(pathSignClient) <> "" Then
            ws.Shapes.AddPicture Filename:=pathSignClient, _
                LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                Left:=clientSignRange.Left, Top:=clientSignRange.Top, _
                Width:=200, Height:=100
        End If
    End If
End Sub
'  Ищет компанию на листе "компании" и возвращает пути
Function GetCompanyFiles(companyName As String, _
                         ByRef stampPath As String, _
                         ByRef logoPath As String, _
                         ByRef signDirPath As String, _
                         ByRef signClientPath As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("компании")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    GetCompanyFiles = False
    For i = 2 To lastRow
        If StrComp(Trim(ws.Cells(i, 1).value), companyName, vbTextCompare) = 0 Then
            stampPath = Trim(ws.Cells(i, 4).value)        ' Печать
            logoPath = Trim(ws.Cells(i, 5).value)         ' Логотип
            signDirPath = Trim(ws.Cells(i, 6).value)      ' Подпись Директора
            signClientPath = Trim(ws.Cells(i, 7).value)   ' Подпись Клиента
            GetCompanyFiles = True
            Exit Function
        End If
    Next i
End Function



