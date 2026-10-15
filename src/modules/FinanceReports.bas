Attribute VB_Name = "FinanceReports"
Option Explicit

Public Sub ClearFinancialReport()
    Dim ws As Worksheet, index As Long
    Set ws = AppSheet("Output")
    ClearRows ws, 2, 5, 10
    ws.Range("M2:N6").ClearContents
    ws.Range("A2:B2,A4,A6:B8").ClearContents
    For index = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(index).Name = "SavingsBalance" Or ws.ChartObjects(index).Name = "SpendingBreakdown" Then ws.ChartObjects(index).Delete
    Next index
End Sub

Public Sub RefreshFinancialReport()
    Dim source As Worksheet, goals As Worksheet, output As Worksheet
    Dim goalRow As Long, sourceRow As Long, outputRow As Long, lastRow As Long, row As Long
    Dim begins As Date, ends As Date, enteredOn As Date
    Dim balance As Currency, target As Currency, amount As Currency
    Dim totals(0 To 3) As Currency, totalSpend As Currency, categories As Variant
    Dim category As String, categoryIndex As Long
    Set source = AppSheet("Expenses&Incomes")
    Set goals = AppSheet("Goals")
    Set output = AppSheet("Output")
    goalRow = LastDataRow(goals, 1, 4)
    If goalRow < 2 Then Err.Raise 5, , "Add a savings goal before generating the report."
    If Not IsDate(goals.Cells(goalRow, 1).value) Or Not IsDate(goals.Cells(goalRow, 2).value) Then Err.Raise 5, , "The latest goal has invalid dates."
    begins = CDate(goals.Cells(goalRow, 1).value): ends = CDate(goals.Cells(goalRow, 2).value)
    If ends < begins Then Err.Raise 5, , "The goal end date must be on or after its start date."
    If Not TryMoney(goals.Cells(goalRow, 3).value, balance, True) Or Not TryMoney(goals.Cells(goalRow, 4).value, target) Then Err.Raise 5, , "The latest goal has invalid amounts."
    lastRow = LastDataRow(source, 1, 5)
    ' Validate before replacing the previous output.
    For row = 2 To lastRow
        If Application.CountA(source.Range(source.Cells(row, 1), source.Cells(row, 5))) > 0 Then
            If Not IsDate(source.Cells(row, 1).value) Then Err.Raise 5, , "Invalid transaction date at row " & row
            If Not TryMoney(source.Cells(row, 3).value, amount) Then Err.Raise 5, , "Invalid transaction amount at row " & row
            If Not IsFinanceCategory(CStr(source.Cells(row, 4).value)) Then Err.Raise 5, , "Invalid category at row " & row
        End If
    Next row
    ClearFinancialReport
    output.Range("A2").value = begins: output.Range("A4").value = ends
    output.Range("A2,A4").NumberFormat = "yyyy-mm-dd"
    output.Range("B2").value = balance: output.Range("B2").NumberFormat = "$#,##0.00"
    outputRow = 2
    For sourceRow = 2 To lastRow
        If IsDate(source.Cells(sourceRow, 1).value) Then
            enteredOn = CDate(source.Cells(sourceRow, 1).value)
            If DateValue(enteredOn) >= begins And DateValue(enteredOn) <= ends Then
                output.Range(output.Cells(outputRow, 5), output.Cells(outputRow, 9)).value = source.Range(source.Cells(sourceRow, 1), source.Cells(sourceRow, 5)).value
                outputRow = outputRow + 1
            End If
        End If
    Next sourceRow
    If outputRow > 2 Then output.Range("E1:I" & outputRow - 1).Sort Key1:=output.Range("E2"), Order1:=xlAscending, Header:=xlYes
    categories = Array("Shopping", "Entertainment", "Bills", "Other")
    For row = 2 To outputRow - 1
        amount = CCur(output.Cells(row, 7).value)
        category = CStr(output.Cells(row, 8).value)
        If category = "Income" Then
            balance = balance + amount
        Else
            balance = balance - amount
            totalSpend = totalSpend + amount
            For categoryIndex = 0 To 3
                If category = categories(categoryIndex) Then totals(categoryIndex) = totals(categoryIndex) + amount
            Next categoryIndex
        End If
        output.Cells(row, 10).value = balance
    Next row
    output.Range("E2:E" & Application.Max(2, outputRow - 1)).NumberFormat = "yyyy-mm-dd"
    output.Range("G2:G" & Application.Max(2, outputRow - 1)).NumberFormat = "$#,##0.00"
    output.Range("J2:J" & Application.Max(2, outputRow - 1)).NumberFormat = "$#,##0.00"
    For categoryIndex = 0 To 3
        output.Cells(categoryIndex + 2, 12).value = categories(categoryIndex)
        output.Cells(categoryIndex + 2, 13).value = totals(categoryIndex)
        If totalSpend > 0 Then output.Cells(categoryIndex + 2, 14).value = totals(categoryIndex) / totalSpend
    Next categoryIndex
    output.Range("L6").value = "Total spending": output.Range("M6").value = totalSpend
    output.Range("M2:M6").NumberFormat = "$#,##0.00": output.Range("N2:N5").NumberFormat = "0.0%"
    ' Keep goal progress separate from the transaction-derived balance.
    output.Range("A6").value = "Savings target": output.Range("B6").value = target
    output.Range("A7").value = "Closing balance": output.Range("B7").value = balance
    output.Range("A8").value = "Remaining to goal": output.Range("B8").value = Application.Max(0, target - balance)
    output.Range("B6:B8").NumberFormat = "$#,##0.00"
    If outputRow > 2 Then FinanceCharts.UpdateSavingsChart outputRow - 1
    If totalSpend > 0 Then FinanceCharts.UpdateSpendingChart
End Sub

Public Sub ShowSpendingSummary()
    Dim ws As Worksheet, row As Long, largest As Currency, category As String
    RefreshFinancialReport
    Set ws = AppSheet("Output")
    For row = 2 To 5
        If CCur(ws.Cells(row, 13).value) > largest Then
            largest = CCur(ws.Cells(row, 13).value): category = CStr(ws.Cells(row, 12).value)
        End If
    Next row
    If largest = 0 Then
        MsgBox "No spending in the selected goal period.", vbInformation
    Else
        MsgBox "Largest spending category: " & category & " (" & FormatCurrency(largest) & ").", vbInformation
    End If
End Sub
