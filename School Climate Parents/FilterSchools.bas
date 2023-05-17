Attribute VB_Name = "FilterSchools"
Sub Filter()
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
sht = "Data"

'Workbook where VBA code resides
Set Workbk = ThisWorkbook

'change filter column in the following code
last = Workbk.Sheets(sht).Cells(Rows.Count, "B").End(xlUp).Row

With Workbk.Sheets(sht)
Set rng = .Range("A1:CA" & last)
End With

Workbk.Sheets(sht).Range("B1:B" & last).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("CD1"), Unique:=True

' Loop through unique values in column
For Each x In Workbk.Sheets(sht).Range([CD2], Cells(Rows.Count, "CD").End(xlUp))

With rng
.AutoFilter
.AutoFilter Field:=2, Criteria1:=x.Value
.SpecialCells(xlCellTypeVisible).Copy

'Add New Workbook in loop
Set newBook = Workbooks.Add(xlWBATWorksheet)

newBook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Data"
newBook.Activate
ActiveSheet.Paste
End With
                                                                                       
'Save new workbook
newBook.SaveAs _
        Filename:="C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Parents Report 2022.xlsx"

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





