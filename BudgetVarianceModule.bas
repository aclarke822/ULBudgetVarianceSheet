Attribute VB_Name = "BudgetVarianceModule"
'Global variables
Dim SummarySheet As Integer
Dim ActualSheet As Integer
Dim Budget_EntrySheet As Integer
Dim FirstDifferenceSheet As Integer
Dim FirstEntrySheet As Integer
Dim SecondDifferenceSheet As Integer
Dim SecondEntrySheet As Integer

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
Sub Submit_Entry()

Application.Calculate
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

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
BudgetEntryTotalColumn = 17

TotalSections = 10

SummarySheet = 1
ActualSheet = 2
Budget_EntrySheet = 3
FirstDifferenceSheet = 4
FirstEntrySheet = 5
SecondDifferenceSheet = 6
SecondEntrySheet = 7

'Prevents user from creating entries too quickly, thereby creating 2 sheets with the same name
If Sheets.Count > Budget_EntrySheet Then
    If Sheets(FirstEntrySheet).Name = Format(DateTime.Now, "MMddyy_hhmmss") + "E" Then
        MsgBox "Slow Down!"
        Call BeforeExit
        Exit Sub
    End If
End If

For i = 0 To TotalSections - 1
    IsValid = CheckRangeIsNumeric(Budget_EntrySheet, CInt(rowStartBE(i)), CInt(columnStartBE(i)), CInt(rowEndBE(i) - rowStartBE(i) + 1), CInt(columnEndBE(i) - columnStartBE(i) + 1))
    If Not IsValid Then
        MsgBox "Please check that there are only numeric values in your entry!"
        Call BeforeExit
        Exit Sub
    End If
Next i

Call CreateEntryAndDifference

Call CalculateDifference

Call BeforeExit
End Sub
Public Function GetSheetName(number As Long) As String
    GetSheetName = Sheets(number).Name
End Function
'return the string representation a cell reference 'cell' with its row shifted by 'shift' amount
Public Function AdjustCell(cell As Range, shift As Long) As String
    'must tell VBA what type of variable this is because it is passed to another function
    Dim cellAddress As String
    'get the string representation of the cell reference
    cellAddress = cell.Address
    'remove the '$' in the cell address
    cellAddress = Replace(cellAddress, "$", "")
    'get the number portion of the cell reference for the row
    cellRow = OnlyNumbers(cellAddress)
    'remove the numbers from the cellAddress to get the column
    cellColumn = Replace(cellAddress, cellRow, "")
    'shift the row by 'shift' amount
    cellRow = cellRow + shift
    'return the string that represents the cell reference shifted by 'shift' amount
    AdjustCell = "!" & cellColumn & cellRow

End Function
'return only the numbers in their original order from a string of alphanumerica characers as Double
Function OnlyNumbers(alphaNum As String) As Double
    Dim Char As String
    Dim x As Integer
    Dim numbers As String
    
    'create a blank string for numbers to be added
    numbers = ""
    'iterate through all of the characters in the string
    For x = 1 To Len(alphaNum)
        Char = Mid(alphaNum, x, 1)
        'check to see if their Ascii value corresponds to that of a number
        If Asc(Char) >= 48 And Asc(Char) <= 57 Then
            'put the numbers together in order as they are detected
            numbers = numbers & Char
        End If
    Next
    'return the numbers as a value (of type Double)
    OnlyNumbers = Val(numbers)
End Function
Public Function CalculateDifference()
    

    If Sheets.Count < 6 Then
        'sets the value of a range of cells to 0 (sheet, rowstart, columnstart, rowend, columnend)
        For i = 0 To TotalSections - 1
            Call SetRangeToZero(FirstDifferenceSheet, _
            CInt(rowStartBE(i)), CInt(columnStartBE(i)), _
            CInt(rowEndBE(i)), CInt(columnEndBE(i)))
        Next i
        
    
        'copies a row of a specified length (sheet1, sheet2, row1start, row2start, sheet1column, sheet2column, numberofrowstocopy)
        For i = 0 To TotalSections - 1
            Call CopyRowOfFiniteLength(SummarySheet, FirstEntrySheet, _
            CInt(rowStartS(i)), CInt(rowStartBE(i)), _
            SummaryTotalBudgetedColumn, BudgetEntryTotalColumn, _
            CInt(rowEndBE(i) - rowStartBE(i) + 1))
        Next i
    
        Exit Function
    End If
    
    For i = 0 To TotalSections - 1
        Call GetDifferenceBetweenTwoRanges(FirstEntrySheet, SecondEntrySheet, _
        CInt(rowStartBE(i)), CInt(rowStartBE(i)), CInt(columnStartBE(i)), CInt(columnStartBE(i)), _
        CInt(rowEndBE(i) - rowStartBE(i) + 1), CInt(columnEndBE(i) - columnStartBE(i) + 1))
    Next i

End Function
Function CopyRowOfFiniteLength(Sheet1 As Integer, sheet2 As Integer, row1Start As Integer, row2Start As Integer, column1 As Integer, column2 As Integer, numberOfRows As Integer)
    For i = 0 To numberOfRows - 1
        Sheets(Sheet1).Cells(row1Start + i, column1).value = Sheets(sheet2).Cells(row2Start + i, column2).value
    Next i
End Function
Function SetRangeToZero(sheet As Integer, rowStart As Integer, columnStart As Integer, rowEnd As Integer, columnEnd As Integer)
    Sheets(sheet).Range(Cells(rowStart, columnStart), Cells(rowEnd, columnEnd)).value = 0
End Function
Function GetDifferenceBetweenTwoRanges(Sheet1 As Integer, sheet2 As Integer, row1Start As Integer, row2Start As Integer, column1Start As Integer, column2Start As Integer, numberOfRows As Integer, numberOfColumns As Integer)

    For i = 0 To numberOfRows - 1
        For j = 0 To numberOfColumns - 1
            difference = -(CDbl(Sheets(sheet2).Cells(row2Start + i, column2Start + j)) - CDbl(Sheets(Sheet1).Cells(row1Start + i, column1Start + j)))
            
            If Abs(difference) < 0.01 Then
                difference = 0#
                Call WriteValueToCell(FirstDifferenceSheet, row1Start + i, column2Start + j, CDbl(difference))
            Else
                'Application.ScreenUpdating = True
                Sheets(FirstDifferenceSheet).Activate
                Cells(row1Start + i, column2Start + j).Activate
                
                
                'Comment = InputBox("This budget value is different from the last, why is that?")
                'If Comment = vbNullString Then
                    'Comment = "User declined to comment"
                'End If
                
                Sheets(FirstEntrySheet).Cells(row1Start + i, column2Start + j).AddComment Comment
                Call WriteValueToCell(FirstDifferenceSheet, row1Start + i, column2Start + j, CDbl(difference))
                Application.ScreenUpdating = False
            End If
        Next j
    Next i
End Function
Function CheckRangeIsNumeric(sheet As Integer, rowStart As Integer, columnStart As Integer, numberOfRows As Integer, numberOfColumns As Integer) As Boolean
    CheckRangeIsNumeric = True
    For i = 0 To numberOfRows - 1
        For j = 0 To numberOfColumns - 1
            If Not IsNumeric(Sheets(sheet).Cells(rowStart + i, columnStart + j).value) Then
                    CheckRangeIsNumeric = False
                    Application.ScreenUpdating = True
                    Sheets(Budget_EntrySheet).Activate
                    Cells(rowStart + i, columnStart + j).Activate
                Exit Function
            End If
        Next j
    Next i
End Function
Function WriteValueToCell(sheet As Integer, row As Integer, column As Integer, value As Double)
    Sheets(sheet).Cells(row, column).value = value
End Function
Sub TurnOffAutomaticCalculation()
Application.Calculation = xlCalculationManual
End Sub
Sub TurnOnAutomaticCalculation()
Application.Calculation = xlCalculationAutomatic
End Sub
Sub Calculate()
Application.Calculate
End Sub
Function BeforeExit()
Sheets(Budget_EntrySheet).Activate
'Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.Calculate
End Function
Function CreateEntryAndDifference()
'Activate the Budget_Entry sheet and copy as is
Sheets(Budget_EntrySheet).Activate
ActiveSheet.Copy After:=Sheets(Budget_EntrySheet)

'Set the variable Entry to the new sheet. Name the sheet and delete that macro button
Set Entry = Sheets(FirstDifferenceSheet)
Entry.Name = Format(DateTime.Now, "MMddyy_hhmmss") + "E"
Entry.DrawingObjects.Delete

'Create difference sheet. Replace the E denoting Entry with D denoting Difference to differentiate the sheets
Entry.Copy After:=Sheets(Budget_EntrySheet)
Set difference = Sheets(FirstDifferenceSheet)
difference.Name = Replace(Entry.Name, "E", "D")
End Function
