Attribute VB_Name = "SmokeTests"
Option Explicit
Private Sub Require(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 901, "SmokeTests", description
End Sub

Public Function RunSmokeTests() As String
    On Error GoTo Failed
    Dim d As Date, money As Currency, outsider As Workbook
    Dim target As Worksheet, countBefore As Long
    Load AddItemForm
    Require AddItemForm.cboxCategory.ListCount = 5, "Transaction categories did not initialize"
    Unload AddItemForm
    Load AddSavingGoalForm
    Unload AddSavingGoalForm
    Load OutputForm
    Unload OutputForm
    Require TryDateParts("2024", "2", "29", d), "Leap day rejected"
    Require Not TryDateParts("2023", "2", "29", d), "Invalid leap day accepted"
    Require Not TryDateParts("2024", "2.5", "1", d), "Fractional month accepted"
    Require Not TryMoney("x", money), "Text money accepted"
    Require Not TryMoney(-1, money), "Negative money accepted"
    Require Not TryMoney(0, money), "Zero transaction accepted"
    Require TryMoney(0, money, True), "Zero opening balance rejected"
    ClearExpensesIncomes
    ClearSavingGoal
    SaveGoal DateSerial(2026, 1, 1), DateSerial(2026, 1, 31), 100, 500
    SaveTransaction DateSerial(2025, 12, 31), "Outside before", 999, "Income", "Test"
    SaveTransaction DateSerial(2026, 1, 1), "Income", 200, "Income", "Test"
    SaveTransaction DateSerial(2026, 1, 15), "Groceries", 50, "Shopping", "Test"
    SaveTransaction DateSerial(2026, 1, 31), "Utilities", 25, "Bills", "Test"
    SaveTransaction DateSerial(2026, 2, 1), "Outside after", 999, "Bills", "Test"
    Set outsider = Application.Workbooks.Add
    RefreshFinancialReport
    Set target = AppSheet("Output")
    Require target.Range("B7").Value = 225, "Closing balance should be 225"
    Require target.Range("B8").Value = 275, "Remaining goal should be 275"
    Require target.Range("M6").Value = 75, "Income included in spending"
    Require Abs(target.Range("N2").Value - 2# / 3#) < 0.000001, "Shopping share incorrect"
    Require Abs(Application.Sum(target.Range("N2:N5")) - 1) < 0.000001, "Shares do not sum to one"
    Require LastDataRow(target, 5, 10) = 4, "Goal period not inclusive or outside rows included"
    Require target.ChartObjects.Count = 2, "Expected exactly two charts"
    RefreshFinancialReport
    Require target.ChartObjects.Count = 2, "Repeated report added charts"
    Require Application.CountA(outsider.Worksheets(1).UsedRange) = 0, "Other workbook was modified"
    outsider.Close False
    ClearExpensesIncomes
    SaveTransaction DateSerial(2026, 1, 4), "Income only", 20, "Income", "Test"
    RefreshFinancialReport
    Require target.Range("M6").Value = 0, "Income-only spending not zero"
    Require Application.CountA(target.Range("N2:N5")) = 0, "Zero denominator generated percentages"
    Require target.ChartObjects.Count = 1, "Empty spending chart remains"
    Require Len(CStr(target.Range("E4").Value)) = 0, "Old report rows remain"
    ClearExpensesIncomes
    RefreshFinancialReport
    Require target.Range("B7").Value = 100, "Empty report lost opening balance"
    Require target.ChartObjects.Count = 0, "Empty report retains charts"
    SaveGoal DateSerial(2026, 2, 1), DateSerial(2026, 2, 28), 300, 600
    RefreshFinancialReport
    Require target.Range("B7").Value = 300, "Latest goal not active"
    ClearSavingGoal
    Require Application.CountA(target.Range("A6:B8")) = 0, "Goal summary not cleared"
    RunSmokeTests = "PASS: calendar and amount validation; inclusive goal period; balance 225; spending 75; shares sum to 100%; zero/empty data; latest goal; stale-row cleanup; chart repeatability; other-workbook isolation."
    Exit Function
Failed:
    RunSmokeTests = "FAIL: " & Err.Number & " " & Err.Description
End Function

Public Sub BuildDemo()
    ClearExpensesIncomes
    ClearSavingGoal
    SaveGoal DateSerial(2026, 10, 1), DateSerial(2026, 10, 31), 1200, 2000
    SaveTransaction DateSerial(2026, 10, 2), "Pay", 1400, "Income", "Synthetic demo"
    SaveTransaction DateSerial(2026, 10, 4), "Rent", 500, "Bills", "Synthetic demo"
    SaveTransaction DateSerial(2026, 10, 8), "Groceries", 120, "Shopping", "Synthetic demo"
    SaveTransaction DateSerial(2026, 10, 12), "Cinema", 40, "Entertainment", "Synthetic demo"
    SaveTransaction DateSerial(2026, 10, 15), "Books", 60, "Shopping", "Synthetic demo"
    SaveTransaction DateSerial(2026, 10, 20), "Transport", 80, "Other", "Synthetic demo"
    RefreshFinancialReport
End Sub
