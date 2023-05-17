Attribute VB_Name = "FilterSchools"
Sub filter()
Application.ScreenUpdating = False
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim w As Long
Dim v As Variant

'specify sheet name in which the data is stored
sht = "Raw Data"

'Workbook where VBA code resides
Set Workbk = ThisWorkbook

'change filter column in the following code
last = Workbk.Sheets(sht).Cells(Rows.Count, "F").End(xlUp).Row

With Workbk.Sheets(sht)
Set rng = .Range("A1:DI" & last)
End With

Workbk.Sheets(sht).Range("F1:F" & last).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("DL1"), Unique:=True

' Loop through unique values in column
For Each x In Workbk.Sheets(sht).Range([DL2], Cells(Rows.Count, "DL").End(xlUp))

With rng
.AutoFilter
.AutoFilter Field:=6, Criteria1:=x.Value
.SpecialCells(xlCellTypeVisible).Copy

'Add New Workbook in loop
Set newBook = Workbooks.Add(xlWBATWorksheet)

newBook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Data"
newBook.Activate
ActiveSheet.Paste
End With
                                                 
'Save new workbook
newBook.SaveAs _
        Filename:="C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"

'Close workbook
newBook.Close SaveChanges:=False

Next x

' Turn off filter
Workbk.Sheets(sht).AutoFilterMode = False
Workbk.Sheets(sht).ShowAllData

With Application
.CutCopyMode = False
.ScreenUpdating = True
End With

End Sub



