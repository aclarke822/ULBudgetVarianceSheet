# ULBudgetVarianceSheet

If the sheet is working the way you want it and you don't have time to test out these changes then you don't have to do any of this.
The only thing I think should definitely be done if you are otherwise happy with the sheet is step number 7.  
Those 3 lines will keep the Budget Sheet calculating for the people entering the number but it won't update any other sheet in the workbook.  
Right now the whole workbook only calculates when you run any macros or do it manually. Just FYI.   

Follow steps IN ORDER:  
1. Download the ZIP from the "Clone or Download" button. Extract.  
2. If you intend on using an excel sheet with different entry arrays, then go to the VBA Editor and copy them somewhere. You could copy them to a Sheet's module or a text file.  
3. Go to VBA Editor and click File>Import>GitModule  
4. Create a button linked to the Import and Export subroutines within the GitModule (on the Actual sheet I assume)  
5. Import modules with the button  
6. Copy the entry array if/that you saved earlier back into BudgetEntryModule AND into ScrubModule
7. Copy these three lines into the Budget Entry sheet:  
Private Sub Worksheet_Change(ByVal Target As Range)  
  ActiveSheet.Calculate  
End Sub  
8. Test it thoroughly!  

Using Git:  

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

To send any changes back to the repository for me to see:  
1. Make sure you scrub the data!    
2. Export all of the modules with the GitModule subroutines    
3. Right click in the folder and click "Git Bash Here" and run the following commands  
4. git add .  
5. git commit -m "put a descriptive message here"  
6. git push origin master  

All done!  
