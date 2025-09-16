Attribute VB_Name = "Module1"
Option Explicit

' =========================
'  Normalisasi angka IDR
' =========================
Private Function ParseIDR(ByVal amount As Variant) As Currency
    On Error GoTo Fallback
    If IsNumeric(amount) Then
        ParseIDR = CCur(amount)
        Exit Function
    End If
    
    Dim s As String
    s = CStr(amount)
    s = Replace(s, "Rp", "", , , vbTextCompare)
    s = Replace(s, "IDR", "", , , vbTextCompare)
    s = Replace(s, ChrW(160), "") ' non-breaking space dari PDF
    s = Replace(s, " ", "")

    Dim hasDot As Boolean, hasComma As Boolean
    hasDot = InStr(s, ".") > 0
    hasComma = InStr(s, ",") > 0
    
    ' Format Indonesia: titik = ribuan, koma = desimal
    If hasDot And hasComma Then
        s = Replace(s, ".", "")
        s = Replace(s, ",", ".")
    ElseIf hasDot And Not hasComma Then
        If (Len(s) - Len(Replace(s, ".", ""))) > 1 Then s = Replace(s, ".", "")
    ElseIf hasComma And Not hasDot Then
        s = Replace(s, ",", ".")
    End If
    
    ParseIDR = CCur(Val(s))
    Exit Function
Fallback:
    ParseIDR = 0
End Function

' Util untuk bagi & sisa terhadap 1000 yang aman untuk Currency
Private Function Div1000(ByVal n As Currency) As Currency
    Div1000 = Fix(n / 1000)
End Function
Private Function Mod1000(ByVal n As Currency) As Long
    Mod1000 = CLng(n - 1000 * Fix(n / 1000))
End Function

' =========================
'  TERBILANG – INDONESIA
' =========================
Public Function TerbilangIDR(ByVal amount As Variant) As String
    Dim n As Currency, intPart As Currency, fracPart As Integer
    n = ParseIDR(amount)
    If n < 0 Then TerbilangIDR = "minus " & TerbilangIDR(-n): Exit Function
    
    intPart = Fix(n)
    fracPart = CInt(Round((n - intPart) * 100, 0))
    
    Dim words As String
    If intPart = 0 Then
        words = "nol"
    Else
        words = Trim(IDR_IntToWords(intPart))
    End If
    
    If fracPart > 0 Then
        TerbilangIDR = words & " rupiah dan " & IDR_IntToWords(fracPart) & " sen"
    Else
        TerbilangIDR = words & " rupiah"
    End If
End Function

Private Function IDR_IntToWords(ByVal n As Currency) As String
    Dim satuan() As String
    satuan = Split("nol satu dua tiga empat lima enam tujuh delapan sembilan sepuluh sebelas")
    
    If n < 12 Then
        IDR_IntToWords = satuan(n)
    ElseIf n < 20 Then
        IDR_IntToWords = IDR_IntToWords(n - 10) & " belas"
    ElseIf n < 100 Then
        IDR_IntToWords = IDR_IntToWords(Fix(n / 10)) & " puluh" & IIf(n Mod 10 > 0, " " & IDR_IntToWords(n Mod 10), "")
    ElseIf n < 200 Then
        IDR_IntToWords = "seratus" & IIf(n - 100 > 0, " " & IDR_IntToWords(n - 100), "")
    ElseIf n < 1000 Then
        IDR_IntToWords = IDR_IntToWords(Fix(n / 100)) & " ratus" & IIf(n Mod 100 > 0, " " & IDR_IntToWords(n Mod 100), "")
    ElseIf n < 2000 Then
        IDR_IntToWords = "seribu" & IIf(n - 1000 > 0, " " & IDR_IntToWords(n - 1000), "")
    ElseIf n < 1000000 Then
        IDR_IntToWords = IDR_IntToWords(Fix(n / 1000)) & " ribu" & IIf(n Mod 1000 > 0, " " & IDR_IntToWords(n Mod 1000), "")
    ElseIf n < 1000000000# Then
        IDR_IntToWords = IDR_IntToWords(Fix(n / 1000000)) & " juta" & IIf(n Mod 1000000 > 0, " " & IDR_IntToWords(n Mod 1000000), "")
    ElseIf n < 1000000000000# Then
        IDR_IntToWords = IDR_IntToWords(Fix(n / 1000000000#)) & " miliar" & IIf(n - Fix(n / 1000000000#) * 1000000000# > 0, " " & IDR_IntToWords(n - Fix(n / 1000000000#) * 1000000000#), "")
    ElseIf n < 1E+15 Then
        IDR_IntToWords = IDR_IntToWords(Fix(n / 1000000000000#)) & " triliun" & IIf(n - Fix(n / 1000000000000#) * 1000000000000# > 0, " " & IDR_IntToWords(n - Fix(n / 1000000000000#) * 1000000000000#), "")
    Else
        IDR_IntToWords = "terlalu besar"
    End If
End Function

' =========================
'  AMOUNT IN WORDS – EN
' =========================
Public Function AmountInWordsIDR_EN(ByVal amount As Variant) As String
    Dim n As Currency, intPart As Currency, fracPart As Integer
    n = ParseIDR(amount)
    If n < 0 Then AmountInWordsIDR_EN = "minus " & AmountInWordsIDR_EN(-n): Exit Function
    
    intPart = Fix(n)
    fracPart = CInt(Round((n - intPart) * 100, 0))
    
    Dim words As String
    If intPart = 0 Then
        words = "zero"
    Else
        words = Trim(EN_IntToWords(intPart))
    End If
    
    If fracPart > 0 Then
        AmountInWordsIDR_EN = words & " rupiah and " & EN_IntToWords(fracPart) & " cents"
    Else
        AmountInWordsIDR_EN = words & " rupiah"
    End If
End Function

Private Function EN_IntToWords(ByVal n As Currency) As String
    Dim units() As String, tens() As String, scales() As String
    units = Split("zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen nineteen")
    tens = Split("zero ten twenty thirty forty fifty sixty seventy eighty ninety")
    scales = Split(" thousand million billion trillion")

    
    If n < 20 Then
        EN_IntToWords = units(n)
        Exit Function
    ElseIf n < 100 Then
        Dim t As String
        t = tens(Fix(n / 10))
        If n Mod 10 > 0 Then t = t & "-" & units(n Mod 10)
        EN_IntToWords = t
        Exit Function
    ElseIf n < 1000 Then
        Dim h As String
        h = units(Fix(n / 100)) & " hundred"
        If n Mod 100 > 0 Then h = h & " " & EN_IntToWords(n Mod 100)
        EN_IntToWords = h
        Exit Function
    End If
    
    Dim i As Integer, chunk As Long, result As String, remainder As Currency
    i = 0: result = ""
    Do While n > 0 And i <= 4
        chunk = Mod1000(n)
        If chunk > 0 Then
            Dim cw As String
            cw = EN_IntToWords(chunk)
            If i > 0 Then cw = cw & " " & scales(i)
            If Len(result) > 0 Then
                result = cw & " " & result
            Else
                result = cw
            End If
        End If
        n = Div1000(n)
        i = i + 1
    Loop
    EN_IntToWords = Trim(result)
End Function

