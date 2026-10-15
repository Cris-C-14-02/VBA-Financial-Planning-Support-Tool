Attribute VB_Name = "FinanceData"
Option Explicit

Public Function IsFinanceCategory(ByVal category As String) As Boolean
    Select Case category
        Case "Shopping", "Bills", "Income", "Entertainment", "Other": IsFinanceCategory = True
    End Select
End Function

Public Sub SaveTransaction(ByVal enteredOn As Date, ByVal item As String, ByVal amount As Currency, ByVal category As String, ByVal description As String)
    Dim ws As Worksheet, row As Long
    If Len(Trim$(item)) = 0 Or amount <= 0 Or Not IsFinanceCategory(category) Then Err.Raise 5, , "Invalid transaction."
    Set ws = AppSheet("Expenses&Incomes")
    row = LastDataRow(ws, 1, 5) + 1
    ws.Cells(row, 1).value = enteredOn
    ws.Cells(row, 1).NumberFormat = "yyyy-mm-dd"
    ws.Cells(row, 2).NumberFormat = "@"
    ws.Cells(row, 2).Value2 = Trim$(item)
    ws.Cells(row, 3).value = amount
    ws.Cells(row, 3).NumberFormat = "$#,##0.00"
    ws.Cells(row, 4).Value2 = category
    ws.Cells(row, 5).NumberFormat = "@"
    ws.Cells(row, 5).Value2 = description
End Sub

Public Sub SaveGoal(ByVal begins As Date, ByVal ends As Date, ByVal opening As Currency, ByVal target As Currency)
    Dim ws As Worksheet, row As Long
    If ends < begins Or opening < 0 Or target <= 0 Then Err.Raise 5, , "Invalid savings goal."
    Set ws = AppSheet("Goals")
    row = LastDataRow(ws, 1, 4) + 1
    ws.Cells(row, 1).value = begins
    ws.Cells(row, 2).value = ends
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 2)).NumberFormat = "yyyy-mm-dd"
    ws.Cells(row, 3).value = opening
    ws.Cells(row, 4).value = target
    ws.Range(ws.Cells(row, 3), ws.Cells(row, 4)).NumberFormat = "$#,##0.00"
End Sub
