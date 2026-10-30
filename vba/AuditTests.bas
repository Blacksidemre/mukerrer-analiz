Attribute VB_Name = "AuditTests"
Option Explicit
Private passed As Long
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 100, , ("BA" & ChrW(350) & "ARISIZ: ") & label
    passed = passed + 1
End Sub
Public Sub TestleriCalistir()
    Dim d(1 To 4, 1 To 2) As Variant, f(1 To 2, 1 To 1) As Variant, a As Variant, b As Variant
    Dim failed As Boolean, s As Worksheet
    On Error GoTo Fail
    passed = 0
    d(1, 1) = "D001": d(1, 2) = 50000
    d(2, 1) = "D001": d(2, 2) = 50000
    d(3, 1) = "D001": d(3, 2) = "50000"
    d(4, 1) = "D001": d(4, 2) = 50000
    f(1, 1) = 1: f(2, 1) = 1
    a = DuplicateAudit.AuditArray(d, f)
    Check a(1, 3) = 3, (ChrW(231) & "oklu kriter / " & ChrW(252) & ChrW(231) & " kay" & ChrW(305) & "tl" & ChrW(305) & " grup")
    Check a(2, 4) = 2 And a(4, 4) = 2, ("orijinal kaynak referans" & ChrW(305))
    Check a(3, 3) = 1, ("say" & ChrW(305) & " ve metin ayr" & ChrW(305) & "m" & ChrW(305))
    Check UBound(a, 1) = 4, ("b" & ChrW(252) & "t" & ChrW(252) & "n sat" & ChrW(305) & "rlar korunuyor")
    f(2, 1) = 0: b = DuplicateAudit.AuditArray(d, f)
    Check b(1, 3) = 4, ("tek kriter / kriter de" & ChrW(287) & "i" & ChrW(351) & "ikli" & ChrW(287) & "i")
    f(1, 1) = 0
    On Error Resume Next
    b = DuplicateAudit.AuditArray(d, f): failed = Err.Number <> 0: Err.Clear
    On Error GoTo Fail
    Check failed, ("s" & ChrW(305) & "f" & ChrW(305) & "r kriter reddediliyor")
    f(1, 1) = 1: f(2, 1) = 1
    d(1, 1) = "AB": d(1, 2) = "C": d(2, 1) = "A": d(2, 2) = "BC"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1 And a(2, 3) = 1, ("anahtar birle" & ChrW(351) & "tirmesinde " & ChrW(231) & "ak" & ChrW(305) & ChrW(351) & "ma yok")
    d(1, 1) = Empty: d(2, 1) = "": d(1, 2) = "x": d(2, 2) = "x"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, ("bo" & ChrW(351) & " h" & ChrW(252) & "cre ve bo" & ChrW(351) & " metin")
    d(1, 1) = "001": d(2, 1) = "1"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1, ("ba" & ChrW(351) & ChrW(305) & "nda s" & ChrW(305) & "f" & ChrW(305) & "r bulunan kodlar")
    d(1, 1) = " a ": d(2, 1) = "a"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1, ("kesin bo" & ChrW(351) & "luk kar" & ChrW(351) & ChrW(305) & "la" & ChrW(351) & "t" & ChrW(305) & "rmas" & ChrW(305))
    a = DuplicateAudit.AuditArray(d, f, True): Check a(1, 3) = 2, ("kenar bo" & ChrW(351) & "lu" & ChrW(287) & "u temizleme se" & ChrW(231) & "ene" & ChrW(287) & "i")
    d(1, 1) = ChrW(304): d(2, 1) = "i"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1, ("varsay" & ChrW(305) & "lan harfe duyarl" & ChrW(305) & " kar" & ChrW(351) & ChrW(305) & "la" & ChrW(351) & "t" & ChrW(305) & "rma")
    a = DuplicateAudit.AuditArray(d, f, False, True): Check a(1, 3) = 2, ("T" & ChrW(252) & "rk" & ChrW(231) & "e noktal" & ChrW(305) & " " & ChrW(304))
    d(1, 1) = "I": d(2, 1) = ChrW(305)
    a = DuplicateAudit.AuditArray(d, f, False, True): Check a(1, 3) = 2, ("T" & ChrW(252) & "rk" & ChrW(231) & "e noktas" & ChrW(305) & "z I")
    d(1, 1) = CDbl(DateSerial(2024, 1, 1)): d(2, 1) = CDbl(DateSerial(2024, 1, 1))
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, "Excel seri tarihleri"
    d(1, 1) = "*?~": d(2, 1) = "*?~"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, ("joker karakterler d" & ChrW(252) & "z metindir")
    d(1, 1) = String$(1000, "x"): d(2, 1) = String$(1000, "x")
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, "uzun anahtarlar"
    Set s = ThisWorkbook.Worksheets("05_Testler")
    s.Range("B4").Value = passed & (" do" & ChrW(287) & "rulama BA" & ChrW(350) & "ARILI"): s.Range("B5").Value = Now
    MsgBox passed & (" do" & ChrW(287) & "rulama BA" & ChrW(350) & "ARILI"), vbInformation
    Exit Sub
Fail:
    ThisWorkbook.Worksheets("05_Testler").Range("B4").Value = Err.Description
    MsgBox Err.Description, vbCritical
End Sub
Public Sub PerformansOlc()
    Dim sizes As Variant, size As Variant, d() As Variant, f(1 To 4, 1 To 1) As Variant, a As Variant
    Dim r As Long, c As Long, line As Long, started As Double, elapsed As Double, s As Worksheet
    On Error GoTo Fail
    sizes = Array(500, 5000, 20000, 50000): line = 9
    Set s = ThisWorkbook.Worksheets("05_Testler")
    For c = 1 To 4: f(c, 1) = 1: Next c
    For Each size In sizes
        ReDim d(1 To CLng(size), 1 To 4)
        For r = 1 To CLng(size)
            d(r, 1) = "D" & (r \ 2): d(r, 2) = "B" & (r \ 2)
            d(r, 3) = "Kredi": d(r, 4) = 50000
        Next r
        started = Timer: a = DuplicateAudit.AuditArray(d, f): elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400
        s.Cells(line, 1).Value = size: s.Cells(line, 2).Value = elapsed
        s.Cells(line, 3).Value = UBound(a, 1): s.Cells(line, 4).Value = Now
        line = line + 1
    Next size
    MsgBox ("Motor " & ChrW(246) & "l" & ChrW(231) & ChrW(252) & "mleri 05_Testler sayfas" & ChrW(305) & "na yaz" & ChrW(305) & "ld" & ChrW(305) & ". Dosya y" & ChrW(252) & "kleme ve bi" & ChrW(231) & "imlendirme s" & ChrW(252) & "resi dahil de" & ChrW(287) & "il."), vbInformation
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation
End Sub
