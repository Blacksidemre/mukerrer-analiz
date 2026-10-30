Attribute VB_Name = "AuditTests"
Option Explicit
Private passed As Long
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 100, , "FAIL: " & label
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
    Check a(1, 3) = 3, "multi criterion / triple group"
    Check a(2, 4) = 2 And a(4, 4) = 2, "original source reference"
    Check a(3, 3) = 1, "number and text distinct"
    Check UBound(a, 1) = 4, "no rows dropped"
    f(2, 1) = 0: b = DuplicateAudit.AuditArray(d, f)
    Check b(1, 3) = 4, "single criterion / criteria change"
    f(1, 1) = 0
    On Error Resume Next
    b = DuplicateAudit.AuditArray(d, f): failed = Err.Number <> 0: Err.Clear
    On Error GoTo Fail
    Check failed, "zero criteria rejected"
    f(1, 1) = 1: f(2, 1) = 1
    d(1, 1) = "AB": d(1, 2) = "C": d(2, 1) = "A": d(2, 2) = "BC"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1 And a(2, 3) = 1, "collision safe tuples"
    d(1, 1) = Empty: d(2, 1) = "": d(1, 2) = "x": d(2, 2) = "x"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, "blank and empty string"
    d(1, 1) = "001": d(2, 1) = "1"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1, "leading zero codes"
    d(1, 1) = " a ": d(2, 1) = "a"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1, "strict whitespace"
    a = DuplicateAudit.AuditArray(d, f, True): Check a(1, 3) = 2, "trim option"
    d(1, 1) = ChrW(304): d(2, 1) = "i"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 1, "case sensitive default"
    a = DuplicateAudit.AuditArray(d, f, False, True): Check a(1, 3) = 2, "Turkish dotted I"
    d(1, 1) = "I": d(2, 1) = ChrW(305)
    a = DuplicateAudit.AuditArray(d, f, False, True): Check a(1, 3) = 2, "Turkish dotless I"
    d(1, 1) = CDbl(DateSerial(2024, 1, 1)): d(2, 1) = CDbl(DateSerial(2024, 1, 1))
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, "Excel serial dates"
    d(1, 1) = "*?~": d(2, 1) = "*?~"
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, "wildcards are literal"
    d(1, 1) = String$(1000, "x"): d(2, 1) = String$(1000, "x")
    a = DuplicateAudit.AuditArray(d, f): Check a(1, 3) = 2, "long keys"
    Set s = ThisWorkbook.Worksheets("05_Testler")
    s.Range("B4").Value = passed & " assertion PASS": s.Range("B5").Value = Now
    MsgBox passed & " assertion PASS", vbInformation
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
    MsgBox "Motor olcumleri 05_Testler sayfasina yazildi. Dosya yukleme ve bicimlendirme suresi dahil degil.", vbInformation
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation
End Sub
