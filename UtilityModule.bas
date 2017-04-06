Attribute VB_Name = "UtilityModule"
Public Function CopyRowOfFiniteLength(Sheet1 As Integer, sheet2 As Integer, row1Start As Integer, row2Start As Integer, column1 As Integer, column2 As Integer, numberOfRows As Integer)
    For i = 0 To numberOfRows - 1
        Sheets(Sheet1).Cells(row1Start + i, column1).value = Sheets(sheet2).Cells(row2Start + i, column2).value
    Next i
End Function
Public Function SetRangeToZero(sheet As Integer, rowStart As Integer, columnStart As Integer, rowEnd As Integer, columnEnd As Integer)
    Sheets(sheet).Activate
    Sheets(sheet).Range(Cells(rowStart, columnStart), Cells(rowEnd, columnEnd)).value = 0
End Function
Public Function WriteValueToCell(sheet As Integer, row As Integer, column As Integer, value As Double)
    Sheets(sheet).Cells(row, column).value = value
End Function
Public Function TurnOnScreenUpdating()
    Application.ScreenUpdating = True
End Function
Public Function TurnOffScreenUpdating()
    Application.ScreenUpdating = False
End Function
Public Function TurnOffDisplayAlerts()
    Application.DisplayAlerts = False
End Function
Public Function TurnOnDisplayAlerts()
    Application.DisplayAlerts = True
End Function
Public Function GetSheetName(number As Long) As String
    GetSheetName = Sheets(number).Name
End Function
Public Sub TurnOffAutomaticCalculation()
    Application.Calculation = xlCalculationManual
End Sub
Public Sub TurnOnAutomaticCalculation()
    Application.Calculation = xlCalculationAutomatic
End Sub
Public Sub Calculate()
    Application.Calculate
End Sub
