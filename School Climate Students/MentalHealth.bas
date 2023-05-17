Attribute VB_Name = "MentalHealth"
Sub Health()
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
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Mental Health"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Mental Health"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c1 = t
    ActiveSheet.Range("A" & t).Value = "In the past 30 days how often did you .."
    ActiveSheet.Range("B" & t).Value = "Never"
    ActiveSheet.Range("C" & t).Value = "Seldom"
    ActiveSheet.Range("D" & t).Value = "Sometimes"
    ActiveSheet.Range("E" & t).Value = "Often"
    ActiveSheet.Range("F" & t).Value = "Always"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 105)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DA2:DA" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DA1:DA" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DA1:DA" & m), "Seldom") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DA1:DA" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DA1:DA" & m), "Often") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DA1:DA" & m), "Always") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 106)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DB2:DB" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DB1:DB" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DB1:DB" & m), "Seldom") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DB1:DB" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DB1:DB" & m), "Often") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DB1:DB" & m), "Always") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 107)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DC2:DC" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DC1:DC" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DC1:DC" & m), "Seldom") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DC1:DC" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DC1:DC" & m), "Often") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DC1:DC" & m), "Always") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 108)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DD2:DD" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DD2:DD" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DD2:DD" & m), "Seldom") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DD2:DD" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DD2:DD" & m), "Often") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DD2:DD" & m), "Always") / w * 100, 2) & "%"
   With ActiveSheet.Range("A" & c1 & ":F" & c1)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A" & c1 & ":F" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c1 & ":F" & t).Font.Size = 16
    ActiveSheet.Range("A" & c1 & ":A" & t).RowHeight = 60
    ActiveSheet.Range("A" & c1 & ":H" & c1).ColumnWidth = 20
    ActiveSheet.Range("A" & c1 & ":A" & t).WrapText = True
    ActiveSheet.Range("A" & c1 & ":A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & c1 & ":A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B" & c1 & ":F" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B" & c1 & ":F" & t).VerticalAlignment = xlVAlignCenter
    d = t + 2
    e = d + (t - 3)
    g = e + d + 5
    h = g
    
      'Chart (Mental Health)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  Dim rnge1 As Range
  Dim rnge2 As Range
  
    Range("A" & c1 & ":A" & t).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("C" & c1 & ":C" & t).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("B" & c1 & ":B" & t).Copy Range(Cells(d, 5), Cells(e, 5))
    Range("E" & c1 & ":F" & t).Copy Range(Cells(d, 7), Cells(e, 8))
    Range("D" & c1 & ":D" & t).Copy Range(Cells(d, 6), Cells(e, 6))
    Range("D" & c1 & ":D" & t).Copy Range(Cells(d, 2), Cells(e, 2))
    Worksheets("Mental Health").Cells(d, 3).Value = "Never"
    Worksheets("Mental Health").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Range(Cells(d, 1), Cells(e, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Mental Health").Range(Cells(d + 1, 2), Cells(e, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Mental Health").Range(Cells(d + 1, 6), Cells(e, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Mental Health").Range(Cells(d + 1, 2), Cells(e, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(d, 1), Cells(e, 8)).Borders.LineStyle = xlNone
    Range(Cells(d, 1), Cells(e, 8)).Interior.Color = xlNone
    Range(Cells(d, 1), Cells(e, 8)).RowHeight = 15
    
    With ActiveSheet.Range("B" & c1 & ":C" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = c1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 3)).Merge
    Next i
    
    
  Set Ws = Worksheets("Mental Health")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "In the past 30 days how often did you .."   'Title
        .ChartTitle.Font.Size = 20
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .PlotBy = IIf(.PlotBy = xlRows, xlColumns, xlRows) 'Switch row/column
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
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Width = 150
        .Legend.Left = 140
        .Legend.Top = 28
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .Legend.LegendEntries(1).Select
        Selection.Delete
        .Legend.LegendEntries(3).Select
        Selection.Delete

    With .Parent
           .Left = Sheets("Mental Health").Range("A" & d).Left
           .Top = Sheets("Mental Health").Range("A" & d).Top
           .Width = Sheets("Mental Health").Range(Cells(d, 1), Cells(d, 8)).Width - 0.5
           .Height = Sheets("Mental Health").Range(Cells(d, 1), Cells(e + d + 2, 8)).Height
    End With

End With

    ActiveSheet.Range("A" & g).Value = "How often did you feel this way when you arrived at school?"
    ActiveSheet.Range("B" & g).Value = "Never"
    ActiveSheet.Range("C" & g).Value = "Sometimes"
    ActiveSheet.Range("D" & g).Value = "Almost every day"
    ActiveSheet.Range("E" & g).Value = "Every day"
    g = g + 1
    ActiveSheet.Range("A" & g).Value = v(1, 109)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DE2:DE" & m), "<>" & "")
    ActiveSheet.Range("B" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DE1:DE" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DE1:DE" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DE1:DE" & m), "Almost every day") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DE1:DE" & m), "Every day") / w * 100, 2) & "%"
    g = g + 1
    ActiveSheet.Range("A" & g).Value = v(1, 110)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DF2:DF" & m), "<>" & "")
    ActiveSheet.Range("B" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DF1:DF" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DF1:DF" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DF1:DF" & m), "Almost every day") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DF1:DF" & m), "Every day") / w * 100, 2) & "%"
    g = g + 1
    ActiveSheet.Range("A" & g).Value = v(1, 111)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("DG2:DG" & m), "<>" & "")
    ActiveSheet.Range("B" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DG1:DG" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DG1:DG" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DG1:DG" & m), "Almost every day") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & g).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("DG1:DG" & m), "Every day") / w * 100, 2) & "%"
    With ActiveSheet.Range("A" & h & ":E" & h)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A" & h & ":E" & g).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A2" & ":E" & g).Font.Size = 16
    ActiveSheet.Range("A" & h & ":A" & g).RowHeight = 60
    ActiveSheet.Range("A" & g + 1).RowHeight = 18.75
    ActiveSheet.Range("A2" & ":E" & g).WrapText = True
    ActiveSheet.Range("A" & h & ":A" & g).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & h & ":A" & g).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B" & h & ":E" & g).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B" & h & ":E" & g).VerticalAlignment = xlVAlignCenter
    k = g + 2
    l = k + 3
    Range("A" & h & ":A" & g).Copy Range(Cells(k, 1), Cells(l, 1))  'Table w/ -ve values
    Range("C" & h & ":C" & g).Copy Range(Cells(k, 3), Cells(l, 3))
    Range("B" & h & ":B" & g).Copy Range(Cells(k, 4), Cells(l, 4))
    Range("D" & h & ":E" & g).Copy Range(Cells(k, 5), Cells(l, 6))
    Worksheets("Mental Health").Cells(k, 2).Value = "Never"
      Worksheets("Mental Health").Range(Cells(k + 1, 2), Cells(l, 2)).Value = 0
      Range(Cells(k, 1), Cells(l, 6)).Font.Color = vbWhite
      Set rngData = Worksheets("Mental Health").Range(Cells(k + 1, 2), Cells(l, 4))
      rngData = Evaluate(rngData.Address & "*-1")
      Range(Cells(k, 1), Cells(l, 6)).Borders.LineStyle = xlNone
      Range(Cells(k, 1), Cells(l, 6)).Interior.Color = xlNone
      Range(Cells(k, 1), Cells(l, 6)).RowHeight = 18.75
    
    With ActiveSheet.Range("B" & h & ":D" & g)
         .Insert Shift:=xlToRight
    End With
    
    For j = h To g
        ActiveSheet.Range(Cells(j, 1), Cells(j, 4)).Merge
    Next j
      
      Set Ws = Worksheets("Mental Health")
      Set Rang = Ws.Range(Cells(k, 1), Cells(l, 6))
      Set MyChart = Ws.Shapes.AddChart2
      
      With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "How often did you feel this way when you arrived at school?"   'Title
        .ChartTitle.Font.Size = 20
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .PlotBy = IIf(.PlotBy = xlRows, xlColumns, xlRows) 'Switch row/column
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
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Width = 150
        .Legend.Left = 100
        .Legend.Top = 30
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Legend.LegendEntries(3).Select
        Selection.Delete
    
        With .Parent
               .Left = Sheets("Mental Health").Range("A" & k).Left
               .Top = Sheets("Mental Health").Range("A" & k).Top
               .Width = Sheets("Mental Health").Range(Cells(k, 1), Cells(l, 8)).Width - 0.5
               .Height = Sheets("Mental Health").Range(Cells(k, 1), Cells(k + 18, 8)).Height
        End With
    
    End With
    
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub



