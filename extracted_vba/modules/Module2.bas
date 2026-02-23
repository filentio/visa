Attribute VB_Name = "Module2"
Option Explicit
Private Function ChunkToWordsEN(ByVal n As Long, ByVal includeAnd As Boolean) As String
    Dim ones As Variant, tens As Variant
    ones = Array("zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", _
                 "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
    tens = Array("", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
    Dim s As String, h As Long, t As Long, u As Long, last2 As Long
    h = n \ 100: t = (n Mod 100) \ 10: u = n Mod 10
    If h > 0 Then
        s = ones(h) & " hundred"
        If (n Mod 100) > 0 Then s = s & IIf(includeAnd, " and ", " ")
    End If
    last2 = n Mod 100
    If last2 > 0 Or n = 0 Then
        If last2 < 20 Then
            If Not (n = 0 And h > 0) Then s = s & ones(last2)
        Else
            s = s & tens(t)
            If u > 0 Then s = s & "-" & ones(u)
        End If
    End If
    ChunkToWordsEN = s
End Function
Public Function SpellNumberEN(ByVal value As Variant, _
    Optional ByVal major As String = "dollar", _
    Optional ByVal minor As String = "cent", _
    Optional ByVal britishAnd As Boolean = False, _
    Optional ByVal titleCase As Boolean = True) As String
    Dim n As Double
    If Not IsNumeric(value) Then SpellNumberEN = CVErr(xlErrValue): Exit Function
    n = CDbl(value)
    Dim sign As String: If n < 0 Then sign = "minus ": n = Abs(n)
    Dim ip As Double: ip = Fix(n)
    Dim cents As Long: cents = CLng(WorksheetFunction.Round((n - ip) * 100, 0))
    If cents = 100 Then ip = ip + 1: cents = 0
    Dim s As String
    Dim b As Long, m As Long, t As Long, r As Long
    b = Int(ip / 1000000000#): ip = ip - b * 1000000000#
    m = Int(ip / 1000000#):     ip = ip - m * 1000000#
    t = Int(ip / 1000#):        r = ip - t * 1000#
    If b > 0 Then s = s & ChunkToWordsEN(b, britishAnd) & " billion"
    If m > 0 Then s = s & IIf(s <> "", " ", "") & ChunkToWordsEN(m, britishAnd) & " million"
    If t > 0 Then s = s & IIf(s <> "", " ", "") & ChunkToWordsEN(t, britishAnd) & " thousand"
    If r > 0 Or s = "" Then s = s & IIf(s <> "", " ", "") & ChunkToWordsEN(r, britishAnd)
    If major <> "" Then
        s = s & " " & IIf((b + m + t + r) = 1, major, major & "s")
    End If
    If minor <> "" Then
        s = s & " and " & IIf(cents > 0, ChunkToWordsEN(cents, False) & " " & IIf(cents = 1, minor, minor & "s"), "zero " & minor & "s")
    ElseIf cents > 0 Then
        s = s & " point " & ChunkToWordsEN(cents, False)
    End If
    s = sign & s
    If titleCase Then s = StrConv(s, vbProperCase)
    SpellNumberEN = s
End Function
' ??????? «????????» ???????:
Public Function SpellUSD(n As Variant) As String
    SpellUSD = SpellNumberEN(n, "dollar", "cent", False, True)
End Function
Public Function NumberToWordsENOnly(n As Variant) As String
    NumberToWordsENOnly = SpellNumberEN(n, "", "", False, True)
End Function

