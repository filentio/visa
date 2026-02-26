Attribute VB_Name = "Лист3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B11")) Is Nothing Then
        If Application.EnableEvents = False Then Application.EnableEvents = True
        Call UpdateAllStampsFromConfig
    End If
End Sub
'  Запуск при изменении B11 формулой (пересчёт листа)
Private Sub Worksheet_Calculate()
    Static lastValue As String
    Dim currentValue As String
    currentValue = Me.Range("B11").value
    If currentValue <> lastValue Then
        lastValue = currentValue
        If Application.EnableEvents = False Then Application.EnableEvents = True
        Call UpdateAllStampsFromConfig
    End If
End Sub
'  Запуск при активации листа (перешли на вкладку)
Private Sub Worksheet_Activate()
    If Application.EnableEvents = False Then Application.EnableEvents = True
    Call UpdateAllStampsFromConfig
End Sub
