Attribute VB_Name = "TeachersExpectations"
Sub Expectations()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim i As Long
Dim t As Long
Dim d As Long
Dim c As Long
Dim e As Long
Dim f As Long
Dim g As Long
Dim a As Long
Dim w As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    
                                         'Academic Expectations Subscale
                                         
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Expectations"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Expectations: Teacher Expectations"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Agree"
    ActiveSheet.Range("E" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 41)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO2:AO" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 42)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP2:AP" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 43)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ2:AQ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AQ1:AQ" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 44)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR2:AR" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    c = t
    ActiveSheet.Range("A" & t).Value = "Expectations: Instructional Practices"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Agree"
    ActiveSheet.Range("E" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 45)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS2:AS" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 46)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT2:AT" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 47)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU2:AU" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 48)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV2:AV" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 49)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW2:AW" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW2:AW" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW2:AW" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW2:AW" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW2:AW" & m), "Strongly Agree") / w * 100, 2) & "%"
    With ActiveSheet.Range("A1:E1")
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A1:E" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A2:E" & t).Font.Size = 16
    ActiveSheet.Range("A1:A" & t).RowHeight = 60
    ActiveSheet.Range("A1:B1").ColumnWidth = 40
    ActiveSheet.Range("A1:E" & t).WrapText = True
    ActiveSheet.Range("C1:F1").ColumnWidth = 20
    ActiveSheet.Range("A1:A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A1:A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B1:E" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B1:E" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 5))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    a = c - 1
    d = t + 2
    e = d + (c - 2)
    f = d + 19
    g = f + (t - c)
    
     'Chart (Teacher Expectations)

    Dim Ws As Worksheet
    Dim Rang As Range
    Dim MyChart As Object
    Dim rngData As Range
    
    Range("A1:A" & a).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("C1:C" & a).Copy Range(Cells(d, 3), Cells(e, 3))
    Range("B1:B" & a).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("D1:E" & a).Copy Range(Cells(d, 5), Cells(e, 6))
    Worksheets("Expectations").Cells(d, 2).Value = "Strongly Disagree"
    Worksheets("Expectations").Range(Cells(d + 1, 2), Cells(e, 2)).Value = 0
    Range(Cells(d, 1), Cells(e, 6)).Font.Color = vbWhite
    Set rngData = Worksheets("Expectations").Range(Cells(d + 1, 2), Cells(e, 4))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(d, 1), Cells(e, 6)).Borders.LineStyle = xlNone
    Range(Cells(d, 1), Cells(e, 6)).Interior.Color = xlNone
    Range(Cells(d, 1), Cells(e, 6)).RowHeight = 15
    
    
    Set Ws = Worksheets("Expectations")
    Set Rang = Ws.Range(Cells(d, 1), Cells(e, 6))
    Set MyChart = Ws.Shapes.AddChart2
    
    With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Expectations: Teacher Expectations"   'Title
        .ChartTitle.Font.Size = 20
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .PlotBy = xlColumns                  'Switch row/column
        .Axes(xlValue).MinimumScale = -1    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14
        .PlotArea.Border.LineStyle = xlContinuous
        .PlotArea.Border.Color = RGB(165, 165, 165)
        .Legend.Position = xlLegendPositionTop
        .Legend.Width = 150
        .Legend.Left = 190
        .Legend.Top = 35
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Legend.LegendEntries(3).Select
        Selection.Delete
       
        
    With .Parent
        .Left = Sheets("Expectations").Range("A" & d).Left
        .Top = Sheets("Expectations").Range("A" & d).Top
        .Width = Sheets("Expectations").Range(Cells(d, 1), Cells(d, 6)).Width - 0.5
        .Height = Sheets("Expectations").Range(Cells(d, 1), Cells(e + d, 6)).Height
    End With
    
    End With
    
                                'Chart (Instructional Practices)
    
    Range(Cells(c, 1), Cells(t, 1)).Copy Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Range(Cells(c, 3), Cells(t, 3)).Copy Range(Cells(f, 3), Cells(g, 3))
    Range(Cells(c, 2), Cells(t, 2)).Copy Range(Cells(f, 4), Cells(g, 4))
    Range(Cells(c, 4), Cells(t, 5)).Copy Range(Cells(f, 5), Cells(g, 6))
    Worksheets("Expectations").Cells(f, 2).Value = "Strongly Disagree"
    Worksheets("Expectations").Range(Cells(f + 1, 2), Cells(g, 2)).Value = 0
    Range(Cells(f, 1), Cells(g, 6)).Font.Color = vbWhite
    Set rngData = Worksheets("Expectations").Range(Cells(f + 1, 2), Cells(g, 4))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(f, 1), Cells(g, 6)).Borders.LineStyle = xlNone
    Range(Cells(f, 1), Cells(g, 6)).Interior.Color = xlNone
    Range(Cells(f, 1), Cells(g, 6)).RowHeight = 15
    
    With ActiveSheet.Range("B1:B" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 2)).Merge
    Next i
    
  Set Ws = Worksheets("Expectations")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 6))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Expectations: Instructional Practices"   'Title
        .ChartTitle.Font.Size = 20
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .PlotBy = xlColumns                  'Switch row/column
        .Axes(xlValue).MinimumScale = -1    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14
        .PlotArea.Border.LineStyle = xlContinuous
        .PlotArea.Border.Color = RGB(165, 165, 165)
        .Legend.Position = xlLegendPositionTop
        .Legend.Width = 150
        .Legend.Left = 190
        .Legend.Top = 35
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Legend.LegendEntries(3).Select
        Selection.Delete

    With .Parent
           .Left = Sheets("Expectations").Range("A" & f).Left
           .Top = Sheets("Expectations").Range("A" & f).Top
           .Width = Sheets("Expectations").Range(Cells(f, 1), Cells(f, 6)).Width - 0.5
           .Height = Sheets("Expectations").Range(Cells(f, 1), Cells(f + 19, 6)).Height
    End With

End With

    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0

Next x
End Sub
