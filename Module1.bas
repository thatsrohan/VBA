Attribute VB_Name = "Module1"
Sub automate()

Dim ws As Worksheet
Dim wsname As String
Dim sourceSheet As Worksheet
Dim userInput As Integer
Set sourceSheet = ThisWorkbook.Sheets("PFD Calculation")
wsname = Application.InputBox("Enter the name of the new Worksheet")
Set ws = ThisWorkbook.Sheets.Add
ws.Name = wsname
sourceSheet.UsedRange.Copy Destination:=ws.Range("A1")
Application.CutCopyMode = False
userInput = Application.InputBox("Enter the Temperature")
ws.Range("C6").Value = userInput

End Sub

