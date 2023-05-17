Attribute VB_Name = "GangActivity"
Sub Gang()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim c1 As Long
Dim d As Long
Dim e As Long
Dim g As Long
Dim h As Long
Dim k As Long
Dim l As Long
Dim j As Long
Dim m As Long
Dim t As Long
Dim i As Long
Dim w As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "BJ").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("BJ2:BJ" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Teachers Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:BF" & m).Value
    End With
    
                                      'Gang Activity
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Gang Activity"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Gang Activity"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c5 = t
    ActiveSheet.Range("A" & t).Value = v(1, 42)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Yes"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP2:AP" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Yes") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "No"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "No") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I don't know"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Don't Know") / w * 100, 2) & "%"
    t = t + 1
    c = t
    ActiveSheet.Range("A" & t).Value = v(1, 43)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Yes"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ2:AQ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "Yes") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "No"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "No") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I don't know"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "Don't Know") / w * 100, 2) & "%"
    
    With ActiveSheet.Range("A" & c5 & ":B" & c5)
         .Font.Size = 18
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 100
    End With
    ActiveSheet.Range("A" & c5 & ":B" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c5 + 1 & ":B" & t).Font.Size = 18
    ActiveSheet.Range("A1").ColumnWidth = 48.57
    ActiveSheet.Range("A" & c5 - 1).RowHeight = 60
    ActiveSheet.Range("A" & c5 + 1 & ":A" & t).RowHeight = 80
    ActiveSheet.Range("A" & c5 & ":B" & t).WrapText = True
    ActiveSheet.Range("B1").ColumnWidth = 20
    ActiveSheet.Range("C1").ColumnWidth = 4.71
    ActiveSheet.Range("A" & c5 & ":B" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & c5 & ":B" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B" & c5 & ":B" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B" & c5 & ":B" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 2))
         .Font.Size = 18
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 100
    End With

    
    
         'Chart (gANG aCTIVITY)

  Set Ws = Worksheets("Gang Activity")
  Set Rang = Ws.Range("A" & c5 & ":B" & c - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .ChartTitle.Text = "Are there gangs at your school this year?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .ChartColor = 22

    With .Parent
           .Left = Sheets("Gang Activity").Range("D" & c5).Left
           .Top = Sheets("Gang Activity").Range("D" & c5).Top
           .Width = Sheets("Gang Activity").Range("D" & c5 & ":K" & c5).Width - 0.5
           .Height = Sheets("Gang Activity").Range("D" & c5 & ":K" & c - 1).Height
    End With
End With
                                          

  Set Ws = Worksheets("Gang Activity")
  Set Rang = Ws.Range("A" & c & ":B" & t)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "Have gangs caused problems at your school this year (such as fights or sale of drugs)?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(250, 172, 114)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlNone
        .Axes(xlCategory).ReversePlotOrder = True
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh

    With .Parent
           .Left = Sheets("Gang Activity").Range("D" & c).Left
           .Top = Sheets("Gang Activity").Range("D" & c).Top
           .Width = Sheets("Gang Activity").Range("D" & c & ":K" & c).Width - 0.5
           .Height = Sheets("Gang Activity").Range("D" & c & ":K" & t).Height
    End With
End With
    t = 0
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close

Next x
End Sub

