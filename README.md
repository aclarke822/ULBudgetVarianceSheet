# ULBudgetVarianceSheet
Download Git from https://git-scm.com/download/win  

Install with all default options

Create a new folder somewhere with the name ULBudgetVarianceSheet

Right click in the folder and click "Git Bash Here"

Run these commands in order, you can copy but must right-click and click paste to paste:
1. git init
2. git remote add origin https://github.com/aclarke822/ULBudgetVarianceSheet.git
3. git fetch
4. git checkout master


Open the Excel sheet and add these lines to the Budget Entry sheet module from the VBA editor:

Private Sub Worksheet_Change(ByVal Target As Range)  
  ActiveSheet.Calculate  
End Sub  
