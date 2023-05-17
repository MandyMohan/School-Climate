Attribute VB_Name = "ExportPDF"
Sub Export()
Dim y As Worksheet
Dim x As Range
Dim last As Long

last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "BJ").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("BJ2:BJ" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Teachers Report 2022.xlsx"
    Worksheets("Data").Visible = False
    For Each y In ActiveWorkbook.Worksheets
        With y.PageSetup
          .Orientation = xlPortrait
          .Zoom = False
          .FitToPagesTall = 1
          .FitToPagesWide = 1
        End With
    Next y
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Teachers Report 2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False

   'Save workbook
    ActiveWorkbook.Save
        
    'Close workbook
    ActiveWorkbook.Close
Next x
End Sub

