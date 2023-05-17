Attribute VB_Name = "ExportPDF"
Sub Export()
Dim y As Worksheet
Dim x As Range
Dim lst As Long

lst = ActiveWorkbook.Sheets("Sheet1").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Sheet1").Range("DL2:DL" & lst)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    Worksheets("Data").Visible = False
    Worksheets("TransformData").Visible = False
    Worksheets("Score Results").Visible = False
    For Each y In ActiveWorkbook.Worksheets
        With y.PageSetup
          .Orientation = xlPortrait
          .Zoom = False
          .FitToPagesTall = 1
          .FitToPagesWide = 1
        End With
    Next y
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Student Report 2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False

   'Save workbook
    ActiveWorkbook.Save
        
    'Close workbook
    ActiveWorkbook.Close
Next x
End Sub
