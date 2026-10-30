Attribute VB_Name = "DuplicateAudit"
Option Explicit
' Excel LTSC 2024 / Windows. Harici kitaplik veya ag islemi kullanilmaz.
Private Const RAW As String = "01_Ham_Veri"
Private Const CRIT As String = "02_Kriterler"
Private Const OUT As String = "03_Kontrol"
Private Const SUMMARY As String = "04_Ozet"
Private Const MAX_ROWS As Long = 50000
Private Const MAX_COLS As Long = 100

Private Function UniqueStatus() As String
    UniqueStatus = "TEK" & ChrW(304) & "L"
End Function
Private Function OriginalStatus() As String
    OriginalStatus = "OR" & ChrW(304) & "J" & ChrW(304) & "NAL"
End Function
Private Function DuplicateStatus() As String
    DuplicateStatus = "M" & ChrW(220) & "KERRER"
End Function
Private Function LastRow(ByVal s As Worksheet) As Long
    Dim c As Range
    Set c = s.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not c Is Nothing Then LastRow = c.Row
End Function
Private Function LastCol(ByVal s As Worksheet) As Long
    Dim c As Range
    Set c = s.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not c Is Nothing Then LastCol = c.Column
End Function
Private Function TurkishLower(ByVal s As String) As String
    s = Replace(s, "I", ChrW(305))
    s = Replace(s, ChrW(304), "i")
    TurkishLower = LCase$(s)
End Function
Private Function Token(ByVal v As Variant, ByVal clean As Boolean, ByVal fold As Boolean) As String
    Dim s As String, typ As String
    If IsError(v) Then Err.Raise vbObjectError + 1, , ("H" & ChrW(252) & "cre hatas" & ChrW(305) & " var. Kaynak veriyi d" & ChrW(252) & "zeltin.")
    Select Case VarType(v)
    Case vbEmpty, vbNull: typ = "E": s = ""
    Case vbString
        s = CStr(v)
        If clean Then s = Trim$(s)
        If fold Then s = TurkishLower(s)
        If Len(s) = 0 Then typ = "E" Else typ = "S"
    Case vbBoolean: typ = "B": s = CStr(v)
    Case Else: typ = "N": s = CStr(v)
    End Select
    Token = typ & CStr(Len(s)) & ":" & s & ";"
End Function
Private Function KeyFor(ByRef data As Variant, ByVal r As Long, ByRef flags As Variant, ByVal clean As Boolean, ByVal fold As Boolean) As String
    Dim c As Long, k As String
    For c = 1 To UBound(data, 2)
        If flags(c, 1) = 1 Then k = k & Token(data(r, c), clean, fold)
    Next c
    KeyFor = k
End Function
' Dizi motoru: cikti sutunlari = durum, grup, adet, orijinal kaynak satiri, kaynak satiri.
Public Function AuditArray(ByRef data As Variant, ByRef flags As Variant, Optional ByVal clean As Boolean = False, Optional ByVal fold As Boolean = False) As Variant
    Dim counts As Object, firsts As Object, keys() As String, result() As Variant
    Dim r As Long, c As Long, n As Long, selected As Long, k As String, first As Long
    For c = 1 To UBound(data, 2)
        If flags(c, 1) <> 0 And flags(c, 1) <> 1 Then Err.Raise vbObjectError + 2, , ("Kriterler yaln" & ChrW(305) & "zca 0 veya 1 olmal" & ChrW(305) & ".")
        selected = selected + flags(c, 1)
    Next c
    If selected = 0 Then Err.Raise vbObjectError + 3, , ("En az bir kriter se" & ChrW(231) & "in.")
    n = UBound(data, 1)
    ReDim keys(1 To n): ReDim result(1 To n, 1 To 5)
    Set counts = CreateObject("Scripting.Dictionary"): counts.CompareMode = vbBinaryCompare
    Set firsts = CreateObject("Scripting.Dictionary"): firsts.CompareMode = vbBinaryCompare
    For r = 1 To n
        k = KeyFor(data, r, flags, clean, fold): keys(r) = k
        If counts.Exists(k) Then
            counts(k) = counts(k) + 1
        Else
            counts.Add k, 1: firsts.Add k, r
        End If
    Next r
    For r = 1 To n
        k = keys(r): first = firsts(k)
        If counts(k) = 1 Then
            result(r, 1) = UniqueStatus()
        ElseIf r = first Then
            result(r, 1) = OriginalStatus()
        Else
            result(r, 1) = DuplicateStatus()
        End If
        result(r, 2) = "GRP-" & Format$(first + 1, "000000")
        result(r, 3) = counts(k): result(r, 4) = first + 1: result(r, 5) = r + 1
    Next r
    AuditArray = result
End Function
Private Sub WriteLiteral(ByVal target As Range, ByRef values As Variant)
    ' =, +, -, @ ile baslayan metinlerin formule donusmesini engelle. Sayisal turleri koru.
    Dim safe As Variant, r As Long, c As Long
    safe = values
    For r = 1 To UBound(safe, 1)
        For c = 1 To UBound(safe, 2)
            If VarType(safe(r, c)) = vbString Then safe(r, c) = "'" & safe(r, c)
        Next c
    Next r
    target.Value2 = safe
End Sub
Public Sub AnalizEt()
    Dim s As Worksheet, cfg As Worksheet, dst As Worksheet, data As Variant, flags As Variant, result As Variant
    Dim nr As Long, nc As Long, c As Long, calc As XlCalculation, started As Double
    On Error GoTo Failed
    calc = Application.Calculation
    Set s = ThisWorkbook.Worksheets(RAW): Set cfg = ThisWorkbook.Worksheets(CRIT): Set dst = ThisWorkbook.Worksheets(OUT)
    nr = LastRow(s) - 1: nc = LastCol(s)
    If nr < 1 Or nr > MAX_ROWS Or nc < 1 Or nc > MAX_COLS Then Err.Raise vbObjectError + 4, , ("S" & ChrW(305) & "n" & ChrW(305) & "r: 1-50.000 kay" & ChrW(305) & "t, 1-100 s" & ChrW(252) & "tun.")
    data = s.Range(s.Cells(2, 1), s.Cells(nr + 1, nc)).Value2
    If nr = 1 And nc = 1 Then
        ReDim data(1 To 1, 1 To 1): data(1, 1) = s.Cells(2, 1).Value2
    End If
    flags = cfg.Range("B7").Resize(nc, 1).Value2
    If nc = 1 Then
        ReDim flags(1 To 1, 1 To 1): flags(1, 1) = cfg.Range("B7").Value2
    End If
    started = Timer
    result = AuditArray(data, flags, cfg.Range("F7").Value2 = 1, cfg.Range("F8").Value2 = 1)
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    If dst.AutoFilterMode Then dst.AutoFilterMode = False
    dst.Cells.Clear
    dst.Columns.Hidden = False
    s.Range(s.Cells(1, 1), s.Cells(1, nc)).Copy dst.Cells(1, 1)
    WriteLiteral dst.Cells(2, 1).Resize(nr, nc), data
    ' Metin kodlarini sayiya cevirmeden kaynak tarih ve kod gorunum bicimlerini koru.
    For c = 1 To nc
        s.Range(s.Cells(2, c), s.Cells(nr + 1, c)).Copy
        dst.Cells(2, c).PasteSpecial xlPasteFormats
    Next c
    Application.CutCopyMode = False
    dst.Cells(1, nc + 1).Resize(1, 5).Value = Array(("Kay" & ChrW(305) & "t Durumu"), ("E" & ChrW(351) & "le" & ChrW(351) & "me Grubu Kimli" & ChrW(287) & "i"), ("E" & ChrW(351) & "le" & ChrW(351) & "me Adedi"), ("Orijinal Kay" & ChrW(305) & "t Referans" & ChrW(305)), ("Kaynak Sat" & ChrW(305) & "r Numaras" & ChrW(305)))
    dst.Cells(2, nc + 1).Resize(nr, 5).Value2 = result
    StyleResult dst, nr, nc
    UpdateSummary result, flags, Timer - started
    cfg.Range("E10").Value = ("VBA anl" & ChrW(305) & "k sonu" & ChrW(231) & ": veri/kriter de" & ChrW(287) & "i" & ChrW(351) & "ince AnalizEt " & ChrW(231) & "al" & ChrW(305) & ChrW(351) & "t" & ChrW(305) & "r" & ChrW(305) & "n.")
    ThisWorkbook.Worksheets("00_Baslangic").Range("B15").Value = ("VBA analizi tamamland" & ChrW(305) & ": ") & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    MsgBox nr & (" kay" & ChrW(305) & "t analiz edildi. Hi" & ChrW(231) & "bir kay" & ChrW(305) & "t silinmedi."), vbInformation
Done:
    Application.Calculation = calc: Application.EnableEvents = True: Application.ScreenUpdating = True
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Duplicate Audit Studio"
    Resume Done
End Sub
Private Sub StyleResult(ByVal s As Worksheet, ByVal nr As Long, ByVal nc As Long)
    Dim area As Range, col As String, fc As FormatCondition
    Set area = s.Range(s.Cells(2, 1), s.Cells(nr + 1, nc + 5))
    col = Split(s.Cells(1, nc + 1).Address, "$")(1)
    area.FormatConditions.Delete
    Set fc = area.FormatConditions.Add(xlExpression, , "=$" & col & "2=""" & OriginalStatus() & """"): fc.Interior.Color = RGB(223, 242, 230)
    Set fc = area.FormatConditions.Add(xlExpression, , "=$" & col & "2=""" & DuplicateStatus() & """"): fc.Interior.Color = RGB(252, 225, 225)
    With s.Range(s.Cells(1, 1), s.Cells(1, nc + 5))
        .Interior.Color = RGB(20, 63, 54): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 32
    End With
    s.Range(s.Cells(1, 1), s.Cells(nr + 1, nc + 5)).AutoFilter
    s.Columns.ColumnWidth = 20
    s.Activate: ActiveWindow.FreezePanes = False: s.Range("B2").Select: ActiveWindow.FreezePanes = True
End Sub
Private Sub UpdateSummary(ByRef result As Variant, ByRef flags As Variant, ByVal elapsed As Double)
    Dim r As Long, u As Long, o As Long, d As Long, selected As Long, s As Worksheet
    For r = 1 To UBound(result, 1)
        Select Case result(r, 1)
        Case UniqueStatus(): u = u + 1
        Case OriginalStatus(): o = o + 1
        Case DuplicateStatus(): d = d + 1
        End Select
    Next r
    For r = 1 To UBound(flags, 1): selected = selected + flags(r, 1): Next r
    Set s = ThisWorkbook.Worksheets(SUMMARY)
    s.Range("B4").Value = UBound(result, 1): s.Range("B5").Value = u: s.Range("B6").Value = o
    s.Range("B7").Value = d: s.Range("B8").Value = o: s.Range("B9").Value = selected
    s.Range("B10").Value = d / UBound(result, 1): s.Range("B11").Value = UBound(result, 1) - u - o - d
    s.Range("B12").Value = IIf(elapsed < 0, elapsed + 86400, elapsed)
End Sub
Public Sub DosyaYukle()
    Dim path As Variant, wb As Workbook, s As Worksheet, target As Worksheet, nr As Long, nc As Long
    Dim answer As Variant, values As Variant, headers As Object, h As String, base As String, c As Long, suffix As Long
    Dim oldSecurity As Long, oldEvents As Boolean, oldCalc As XlCalculation, chosen As Long, i As Long, names As String
    oldSecurity = Application.AutomationSecurity: oldEvents = Application.EnableEvents: oldCalc = Application.Calculation
    On Error GoTo Failed
    path = Application.GetOpenFilename("Excel veya CSV (*.xlsx;*.csv),*.xlsx;*.csv")
    If VarType(path) = vbBoolean Then Exit Sub
    If CStr(path) = ThisWorkbook.FullName Then Err.Raise vbObjectError + 8, , ("Analiz kitab" & ChrW(305) & "n" & ChrW(305) & "n kendisini y" & ChrW(252) & "klemeyin.")
    If LCase$(Right$(CStr(path), 5)) <> ".xlsx" And LCase$(Right$(CStr(path), 4)) <> ".csv" Then Err.Raise vbObjectError + 5, , ("Yaln" & ChrW(305) & "zca XLSX veya CSV se" & ChrW(231) & "in.")
    If LCase$(Right$(CStr(path), 4)) = ".csv" Then
        MsgBox ("CSV i" & ChrW(231) & "in Excel Veri > Metin/CSV'den yolunu kullan" & ChrW(305) & "n. Kodlamay" & ChrW(305) & " UTF-8 veya Windows-1254, ayrac" & ChrW(305) & " uygun se" & ChrW(231) & "in. Kod s" & ChrW(252) & "tunlar" & ChrW(305) & "n" & ChrW(305) & " MET" & ChrW(304) & "N yap" & ChrW(305) & "n. Sonucu XLSX kaydedip y" & ChrW(252) & "kleyin."), vbInformation
        Exit Sub
    End If
    Application.AutomationSecurity = 3: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    Set wb = Workbooks.Open(CStr(path), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For i = 1 To wb.Worksheets.Count: names = names & i & " - " & wb.Worksheets(i).Name & vbCrLf: Next i
    answer = Application.InputBox(("Sayfa numaras" & ChrW(305) & ":") & vbCrLf & Left$(names, 180), ("Veri sayfas" & ChrW(305)), 1, Type:=1)
    If VarType(answer) = vbBoolean Then GoTo Done
    If answer <> Fix(answer) Or answer < 1 Or answer > wb.Worksheets.Count Then Err.Raise vbObjectError + 6, , ("Ge" & ChrW(231) & "ersiz sayfa.")
    chosen = CLng(answer): Set s = wb.Worksheets(chosen)
    nr = LastRow(s): nc = LastCol(s)
    If nr < 2 Or nr > MAX_ROWS + 1 Or nc > MAX_COLS Then Err.Raise vbObjectError + 7, , ("S" & ChrW(305) & "n" & ChrW(305) & "r: 50.000 kay" & ChrW(305) & "t / 100 s" & ChrW(252) & "tun. " & ChrW(304) & "lk sat" & ChrW(305) & "r ba" & ChrW(351) & "l" & ChrW(305) & "k olmal" & ChrW(305) & ".")
    If MsgBox(("Mevcut ham veri ve analiz de" & ChrW(287) & "i" & ChrW(351) & "ecek. Devam edilsin mi?"), vbYesNo + vbQuestion) <> vbYes Then GoTo Done
    values = s.Range(s.Cells(1, 1), s.Cells(nr, nc)).Value2
    Set headers = CreateObject("Scripting.Dictionary"): headers.CompareMode = vbBinaryCompare
    For c = 1 To nc
        base = Trim$(CStr(values(1, c))): If base = "" Then base = ("S" & ChrW(252) & "tun ") & c
        h = base: suffix = 2
        Do While headers.Exists(h): h = base & " (" & suffix & ")": suffix = suffix + 1: Loop
        headers.Add h, True: values(1, c) = h
    Next c
    Set target = ThisWorkbook.Worksheets(RAW)
    If target.AutoFilterMode Then target.AutoFilterMode = False
    target.Cells.Clear
    target.Columns.Hidden = False
    WriteLiteral target.Range("A1").Resize(nr, nc), values
    s.Range(s.Cells(1, 1), s.Cells(nr, nc)).Copy
    target.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    With ThisWorkbook.Worksheets(CRIT)
        .Range("A7:B106").ClearContents
        For c = 1 To nc: .Cells(c + 6, 1).Value = values(1, c): .Cells(c + 6, 2).Value = 1: Next c
        .Range("E10").Value = ("Yeni veri. AnalizEt " & ChrW(231) & "al" & ChrW(305) & ChrW(351) & "t" & ChrW(305) & "r" & ChrW(305) & "n.")
    End With
    ThisWorkbook.Worksheets(OUT).Cells.Clear
    ThisWorkbook.Worksheets(SUMMARY).Range("B4:B12").ClearContents
    ThisWorkbook.Worksheets("00_Baslangic").Range("B15").Value = ("Yeni dosya y" & ChrW(252) & "klendi; " & ChrW(246) & "nceki analiz ge" & ChrW(231) & "ersiz.")
    MsgBox nr - 1 & (" kay" & ChrW(305) & "t / ") & nc & (" s" & ChrW(252) & "tun y" & ChrW(252) & "klendi. Kriterleri se" & ChrW(231) & "ip AnalizEt " & ChrW(231) & "al" & ChrW(305) & ChrW(351) & "t" & ChrW(305) & "r" & ChrW(305) & "n."), vbInformation
Done:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
    Resume Done
End Sub
Public Sub ExcelAktar()
    Dim path As Variant, wb As Workbook, s As Worksheet
    If LastRow(ThisWorkbook.Worksheets(OUT)) < 2 Then MsgBox (ChrW(214) & "nce analiz yap" & ChrW(305) & "n."), vbExclamation: Exit Sub
    If MsgBox(("VBA modunda son veri/kriter de" & ChrW(287) & "i" & ChrW(351) & "ikli" & ChrW(287) & "inden sonra AnalizEt " & ChrW(231) & "al" & ChrW(305) & ChrW(351) & "t" & ChrW(305) & "rd" & ChrW(305) & "n" & ChrW(305) & "z m" & ChrW(305) & "?"), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    path = Application.GetSaveAsFilename("Duplicate_Audit_Sonuc.xlsx", "Excel (*.xlsx),*.xlsx")
    If VarType(path) = vbBoolean Then Exit Sub
    On Error GoTo Failed
    ThisWorkbook.Worksheets(Array(RAW, CRIT, OUT, SUMMARY)).Copy
    Set wb = ActiveWorkbook
    For Each s In wb.Worksheets
        s.UsedRange.Copy: s.UsedRange.PasteSpecial xlPasteValues
    Next s
    Application.CutCopyMode = False
    wb.SaveAs Filename:=CStr(path), FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    MsgBox ("Statik anl" & ChrW(305) & "k sonu" & ChrW(231) & " kaydedildi. Kaynak veri silinmedi."), vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
    If Not wb Is Nothing Then wb.Close False
End Sub
Public Sub GrubuIncele()
    Dim s As Worksheet, nc As Long, group As String
    Set s = ThisWorkbook.Worksheets(OUT): nc = LastCol(ThisWorkbook.Worksheets(RAW))
    group = InputBox(("Grup kimli" & ChrW(287) & "i (" & ChrW(246) & "rnek GRP-000002). Bo" & ChrW(351) & " b" & ChrW(305) & "rakmak filtreyi kald" & ChrW(305) & "r" & ChrW(305) & "r."))
    If LastRow(s) < 2 Then Exit Sub
    s.Activate
    If s.FilterMode Then s.ShowAllData
    If Len(group) > 0 Then s.Range(s.Cells(1, 1), s.Cells(LastRow(s), nc + 5)).AutoFilter Field:=nc + 2, Criteria1:=group
End Sub
Public Sub Kurulum()
    Dim s As Worksheet, button As Shape, names As Variant, actions As Variant, i As Long
    Set s = ThisWorkbook.Worksheets("00_Baslangic")
    names = Array(("Dosya y" & ChrW(252) & "kle"), "Analiz et", "Grubu incele", "Excel aktar", ("Testleri " & ChrW(231) & "al" & ChrW(305) & ChrW(351) & "t" & ChrW(305) & "r"), ("Performans " & ChrW(246) & "l" & ChrW(231)))
    actions = Array("DosyaYukle", "AnalizEt", "GrubuIncele", "ExcelAktar", "TestleriCalistir", "PerformansOlc")
    For i = s.Shapes.Count To 1 Step -1
        If Left$(s.Shapes(i).Name, 9) = "DA_Button" Then s.Shapes(i).Delete
    Next i
    For i = 0 To 5
        Set button = s.Shapes.AddShape(msoShapeRoundedRectangle, 35 + (i Mod 3) * 220, 380 + (i \ 3) * 65, 205, 48)
        button.Name = "DA_Button" & i
        button.TextFrame2.TextRange.Text = names(i): button.TextFrame2.TextRange.Font.Size = 13
        button.Fill.ForeColor.RGB = RGB(20, 63, 54): button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & actions(i)
    Next i
    MsgBox ("D" & ChrW(252) & ChrW(287) & "meler haz" & ChrW(305) & "r. Dosyay" & ChrW(305) & " XLSM olarak kaydedin."), vbInformation
End Sub
Public Sub YanYanaKarsilastir()
    Dim s As Worksheet, view As Worksheet, nc As Long, r As Long, c As Long, col As Long, group As String, v As Variant
    Set s = ThisWorkbook.Worksheets(OUT): nc = LastCol(ThisWorkbook.Worksheets(RAW))
    If LastRow(s) < 2 Then MsgBox (ChrW(214) & "nce analiz yap" & ChrW(305) & "n."), vbExclamation: Exit Sub
    group = InputBox(("Yan yana kar" & ChrW(351) & ChrW(305) & "la" & ChrW(351) & "t" & ChrW(305) & "r" & ChrW(305) & "lacak grup kimli" & ChrW(287) & "i (" & ChrW(246) & "rne" & ChrW(287) & "in GRP-000002):"))
    If Len(group) = 0 Then Exit Sub
    On Error Resume Next
    Set view = ThisWorkbook.Worksheets("06_Karsilastirma")
    On Error GoTo 0
    If view Is Nothing Then
        Set view = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        view.Name = "06_Karsilastirma"
    End If
    view.Cells.Clear: view.Cells(1, 1).Value = ("Alan / Kaynak sat" & ChrW(305) & "r"): col = 2
    For c = 1 To nc + 5: view.Cells(c + 1, 1).Value = "'" & CStr(s.Cells(1, c).Value): Next c
    For r = 2 To LastRow(s)
        If s.Cells(r, nc + 2).Value = group Then
            If col > 21 Then Exit For
            view.Cells(1, col).Value = s.Cells(r, nc + 5).Value
            For c = 1 To nc + 5
                v = s.Cells(r, c).Value2
                If VarType(v) = vbString Then v = "'" & v
                view.Cells(c + 1, col).Value2 = v
                view.Cells(c + 1, col).NumberFormat = s.Cells(r, c).NumberFormat
            Next c
            col = col + 1
        End If
    Next r
    view.Columns.ColumnWidth = 24: view.Rows(1).Font.Bold = True: view.Activate
    MsgBox ("En fazla ilk 20 kay" & ChrW(305) & "t yan yana g" & ChrW(246) & "sterilir. T" & ChrW(252) & "m grup i" & ChrW(231) & "in GrubuIncele kullan" & ChrW(305) & "n."), vbInformation
End Sub
