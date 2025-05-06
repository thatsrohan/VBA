Attribute VB_Name = "Module2"
Sub automate_2()
Dim temp As Integer
Dim targetRow As Range
Dim sourceRow As Range
Dim lastRow As Long
Sheet2.Range("Q6").Value = Application.InputBox("Enter the Temperature")
Set sourceRow = Sheets("auto2").Rows(2)
sourceRow.Calculate
Application.CalculateFullRebuild
lastRow = Sheets("auto2").Cells(Rows.Count, 1).End(xlUp).Row
Set targetRow = Sheets("auto2").Rows(lastRow + 1)
sourceRow.UnMerge
targetRow.UnMerge
 sourceRow.Copy
 targetRow.PasteSpecial Paste:=xlPasteValues
 Application.CutCopyMode = False
 

End Sub
