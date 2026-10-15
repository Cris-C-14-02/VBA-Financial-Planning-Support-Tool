Attribute VB_Name = "Module2"
Option Explicit

Public Sub Button4_Click()
    OutputForm.Show
End Sub
Public Sub Button3_Click()
    AddItemForm.Show
End Sub
Public Sub ButtonAddSavingGoal_Click()
    AddSavingGoalForm.Show
End Sub
Public Sub FinancialAdvice_Button3_Click()
    On Error GoTo Failed
    ShowSpendingSummary
    Exit Sub
Failed:
    MsgBox Err.description, vbExclamation
End Sub
