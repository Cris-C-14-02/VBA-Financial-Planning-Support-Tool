Attribute VB_Name = "Module1"
Option Explicit

Public Sub ClearExpensesIncomes()
    ClearRows AppSheet("Expenses&Incomes"), 2, 1, 5
    ClearFinancialReport
End Sub
Public Sub ClearOutput()
    ClearFinancialReport
End Sub
Public Sub ClearSavingGoal()
    ClearRows AppSheet("Goals"), 2, 1, 4
    ClearFinancialReport
End Sub
