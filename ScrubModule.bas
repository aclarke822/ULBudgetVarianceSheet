Attribute VB_Name = "ScrubModule"
'Global variables
Dim SummarySheet As Integer
Dim ActualSheet As Integer
Dim Budget_EntrySheet As Integer
Dim FirstDifferenceSheet As Integer

Dim rowStartBE As Variant
Dim rowEndBE As Variant
Dim columnStartBE As Variant
Dim columnEndBE As Variant

Dim rowStartS As Variant
Dim rowEndS As Variant
Dim columnStartS As Variant
Dim columnEndS As Variant

Dim SummaryTotalBudgetedColumn As Integer
Dim BudgetEntryTotalColumn As Integer
Dim TotalSections As Integer
Public Sub ScrubData()

OnEntry

'Cell start and end references for budget matrix. For modularity
rowStartBE = Array(7, 34, 51, 58, 86, 102, 108, 116, 126, 131)
rowEndBE = Array(27, 47, 54, 82, 98, 104, 112, 122, 127, 135)
columnStartBE = Array(5, 5, 5, 5, 5, 5, 5, 5, 5, 5)
columnEndBE = Array(16, 16, 16, 16, 16, 16, 16, 16, 16, 16)

rowStartS = Array(8, 35, 52, 59, 87, 103, 109, 117, 127, 132)
rowEndS = Array(28, 48, 55, 83, 99, 105, 113, 123, 128, 136)
columnStartS = Array(5, 5, 5, 5, 5, 5, 5, 5, 5, 5)
columnEndBS = Array(16, 16, 16, 16, 16, 16, 16, 16, 16, 16)

SummaryTotalBudgetedColumn = 18

TotalSections = 10

SummarySheet = 1
ActualSheet = 2
Budget_EntrySheet = 3
FirstDifferenceSheet = 4

ScrubSBESheets

DeleteAllSavedEntries

BeforeExit

End Sub
Private Sub ScrubSBESheets()

For i = 0 To TotalSections - 1
        Call SetRangeToZero(Budget_EntrySheet, CInt(rowStartBE(i)), CInt(columnStartBE(i)), CInt(rowEndBE(i)), CInt(columnEndBE(i)))
        Call SetRangeToZero(SummarySheet, CInt(rowStartS(i)), 18, CInt(rowEndS(i)), 18)
Next i

End Sub
Private Sub DeleteAllSavedEntries()

While Sheets.Count > Budget_EntrySheet
    Sheets(FirstDifferenceSheet).Delete
Wend

End Sub
Private Sub BeforeExit()
    Sheets(ActualSheet).Activate
    Calculate
    TurnOffAutomaticCalculation
    Call TurnOnScreenUpdating
    Call TurnOnDisplayAlerts
End Sub
Private Sub OnEntry()
    Calculate
    TurnOffAutomaticCalculation
    Call TurnOffScreenUpdating
    Call TurnOffDisplayAlerts
End Sub
