Attribute VB_Name = "StudentSupport"
Sub Support()
Dim x As Range
Dim rng As Range
Dim sht As String
Dim m As Long
Dim c As Long
Dim a As Long
Dim t As Long
Dim i As Long
Dim d As Long
Dim f As Long
Dim g As Long
Dim e As Long
Dim w As Long
Dim last As Long
Dim v As Variant

last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "BJ").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("BJ2:BJ" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Teachers Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:BF" & m).Value
    End With

                                            'Respect for Students Subscale
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Relations Students & Adults"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Relationships between Students and Adults: Respect for Students"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 3)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("C2:C" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 4)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("D2:D" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 5)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("E2:E" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
    
                                             'Willingness to seek help
                                                  
    ActiveSheet.Range("A" & t).Value = "Relationships between Students and Adults: Willingness to Seek Help"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 7)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("G2:G" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("G1:G" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("G1:G" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("G1:G" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("G1:G" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("G1:G" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("G1:G" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 8)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("H2:H" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("H1:H" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("H1:H" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("H1:H" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("H1:H" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("H1:H" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("H1:H" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 9)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("I2:I" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 10)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("J2:J" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 11)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("K2:K" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 12)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("L2:L" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Strongly Agree") / w * 100, 2) & "%"
    With ActiveSheet.Range("A1:G1")
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A1:G" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A2:G" & t).Font.Size = 16
    ActiveSheet.Range("A1:A" & t).RowHeight = 60
    ActiveSheet.Range("B1:I1").ColumnWidth = 20
    ActiveSheet.Range("A1:G" & t).WrapText = True
    ActiveSheet.Range("A1:A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A1:A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B1:G" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B1:G" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 7))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    a = c - 1
    d = t + 3
    e = d + (c - 2)
    f = e + d + 4
    g = f + (t - c)
    
       'Chart (Respect for Students)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  
    Range("A1:A" & a).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("B1:B" & a).Copy Range(Cells(d, 5), Cells(e, 5))
    Range("C1:C" & a).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("D1:D" & a).Copy Range(Cells(d, 2), Cells(e, 2))
    Range("E1:G" & a).Copy Range(Cells(d, 7), Cells(e, 9))
    Worksheets("Relations Students & Adults").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("Relations Students & Adults").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Worksheets("Relations Students & Adults").Cells(d, 6).Value = "Somewhat Disagree"
    Worksheets("Relations Students & Adults").Range(Cells(d + 1, 6), Cells(e, 6)).Value = 0
    Set rngData = Worksheets("Relations Students & Adults").Range(Cells(d + 1, 2), Cells(e, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(d, 1), Cells(e, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
    
  Set Ws = Worksheets("Relations Students & Adults")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Relationships between Students and Adults: Respect for Students"   'Title
        .ChartTitle.Font.Size = 20
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Bold = True
        .PlotBy = IIf(.PlotBy = xlRows, xlColumns, xlRows)                  'Switch row/column
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
        .Legend.Width = 240
        .Legend.Left = 110
        .Legend.Top = 15
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(6).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(6).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(7).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(8).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Legend.LegendEntries(1).Select
        Selection.Delete
        .Legend.LegendEntries(3).Select
        Selection.Delete
        
    With .Parent
           .Left = Sheets("Relations Students & Adults").Range("A" & d).Left
           .Top = Sheets("Relations Students & Adults").Range("A" & d).Top
           .Width = Sheets("Relations Students & Adults").Range(Cells(d, 1), Cells(d, 9)).Width - 0.5
           .Height = Sheets("Relations Students & Adults").Range(Cells(d, 1), Cells(e + d + 1, 9)).Height
    End With

End With
    
      'Chart (Willingness to seek help)
  
    Range(Cells(c, 1), Cells(t, 1)).Copy Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Range(Cells(c, 2), Cells(t, 2)).Copy Range(Cells(f, 5), Cells(g, 5))
    Range(Cells(c, 3), Cells(t, 3)).Copy Range(Cells(f, 4), Cells(g, 4))
    Range(Cells(c, 4), Cells(t, 4)).Copy Range(Cells(f, 2), Cells(g, 2))
    Range(Cells(c, 5), Cells(t, 7)).Copy Range(Cells(f, 7), Cells(g, 9))
    Worksheets("Relations Students & Adults").Cells(f, 3).Value = "Strongly Disagree"
    Worksheets("Relations Students & Adults").Range(Cells(f + 1, 3), Cells(g, 3)).Value = 0
    Worksheets("Relations Students & Adults").Cells(f, 6).Value = "Somewhat Disagree"
    Worksheets("Relations Students & Adults").Range(Cells(f + 1, 6), Cells(g, 6)).Value = 0
    Set rngData = Worksheets("Relations Students & Adults").Range(Cells(f + 1, 2), Cells(g, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(f, 1), Cells(g, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
    
  Set Ws = Worksheets("Relations Students & Adults")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Relationships between Students & Adults: Willingness to seek help"   'Title
        .ChartTitle.Font.Size = 18
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
        .Legend.Width = 230
        .Legend.Left = 155
        .Legend.Top = 7
        .Legend.Font.Size = 14
        '.Legend.Font.Color = vbBlack
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(6).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(7).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(8).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Legend.LegendEntries(1).Select
        Selection.Delete
        .Legend.LegendEntries(3).Select
        Selection.Delete
        
        With ActiveSheet.Range("B1:C" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 3)).Merge
    Next i

    With .Parent
           .Left = Sheets("Relations Students & Adults").Range("A" & f).Left
           .Top = Sheets("Relations Students & Adults").Range("A" & f).Top
           .Width = Sheets("Relations Students & Adults").Range(Cells(f, 1), Cells(f, 9)).Width - 0.5
           .Height = Sheets("Relations Students & Adults").Range(Cells(f, 1), Cells(f + 22, 9)).Height
    End With

End With
    
    
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub

