Attribute VB_Name = "WorkbookSupport"
Option Explicit

Public Function AppSheet(ByVal sheetName As String) As Worksheet
    Set AppSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function LastDataRow(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim col As Long, candidate As Long
    LastDataRow = 1
    For col = firstCol To lastCol
        candidate = ws.Cells(ws.Rows.Count, col).End(xlUp).row
        If candidate > LastDataRow Then LastDataRow = candidate
    Next col
End Function

Public Sub ClearRows(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim lastRow As Long
    lastRow = LastDataRow(ws, firstCol, lastCol)
    If lastRow >= firstRow Then ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).ClearContents
End Sub

Public Function TryDateParts(ByVal yearText As Variant, ByVal monthText As Variant, ByVal dayText As Variant, ByRef result As Date) As Boolean
    Dim y As Double, m As Double, d As Double, candidate As Date
    On Error GoTo Invalid
    If Not IsNumeric(yearText) Or Not IsNumeric(monthText) Or Not IsNumeric(dayText) Then Exit Function
    y = CDbl(yearText): m = CDbl(monthText): d = CDbl(dayText)
    If y <> Fix(y) Or m <> Fix(m) Or d <> Fix(d) Then Exit Function
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    candidate = DateSerial(CInt(y), CInt(m), CInt(d))
    If Year(candidate) <> y Or Month(candidate) <> m Or Day(candidate) <> d Then Exit Function
    result = candidate
    TryDateParts = True
Invalid:
End Function

Public Function TryNumber(ByVal value As Variant, ByRef result As Double, Optional ByVal allowZero As Boolean = False) As Boolean
    On Error GoTo Invalid
    If Len(Trim$(CStr(value))) = 0 Or Not IsNumeric(value) Then Exit Function
    result = CDbl(value)
    If result < 0 Or (result = 0 And Not allowZero) Then Exit Function
    TryNumber = True
Invalid:
End Function

Public Function TryMoney(ByVal value As Variant, ByRef result As Currency, Optional ByVal allowZero As Boolean = False) As Boolean
    Dim number As Double
    On Error GoTo Invalid
    If Not TryNumber(value, number, allowZero) Then Exit Function
    result = CCur(number)
    If result = 0 And Not allowZero Then Exit Function
    TryMoney = True
Invalid:
End Function
