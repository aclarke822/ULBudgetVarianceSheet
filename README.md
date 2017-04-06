# ULBudgetVarianceSheet
Add these lines to the Budget Entry sheet module from the VBA editor:
Private Sub Worksheet_Change(ByVal Target As Range)
  ActiveSheet.Calculate
End Sub
