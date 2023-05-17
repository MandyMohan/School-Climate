Attribute VB_Name = "PrevalenceToTeasingAndBullying"
Sub Bullying()
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
    
                                  'Bullying Subscale
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Bullying"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Bullying: Prevalence to Teasing and Bullying"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 18)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("R2:R" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 19)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("S2:S" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("S1:S" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("S1:S" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("S1:S" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("S1:S" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("S1:S" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("S1:S" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 20)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("T2:T" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 21)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("U2:U" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("U1:U" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("U1:U" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("U1:U" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("U1:U" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("U1:U" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("U1:U" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 22)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("V2:V" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 23)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("W2:W" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("W1:W" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("W1:W" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("W1:W" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("W1:W" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("W1:W" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("W1:W" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
                                                      'Bullying by Adult
    ActiveSheet.Range("A" & t).Value = "Bullying: Victimization by Adults"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 25)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y2:Y" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y1:Y" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y1:Y" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y1:Y" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y1:Y" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y1:Y" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Y1:Y" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 26)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z2:Z" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z1:Z" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z1:Z" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z1:Z" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z1:Z" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z1:Z" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Z1:Z" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 27)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA2:AA" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA1:AA" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA1:AA" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA1:AA" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA1:AA" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA1:AA" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AA1:AA" & m), "Strongly Agree") / w * 100, 2) & "%"
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
    f = e + d + 3
    g = f + (t - c)
    
      'Chart (Bullying)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  
    Range("A1:A" & a).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("B1:B" & a).Copy Range(Cells(d, 5), Cells(e, 5))
    Range("C1:C" & a).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("D1:D" & a).Copy Range(Cells(d, 2), Cells(e, 2))
    Range("E1:G" & a).Copy Range(Cells(d, 7), Cells(e, 9))
    Worksheets("Bullying").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("Bullying").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Worksheets("Bullying").Cells(d, 6).Value = "Somewhat Disagree"
    Worksheets("Bullying").Range(Cells(d + 1, 6), Cells(e, 6)).Value = 0
    Set rngData = Worksheets("Bullying").Range(Cells(d + 1, 2), Cells(e, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(d, 1), Cells(e, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
  Set Ws = Worksheets("Bullying")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Bullying: Prevalence to Teasing and Bullying"   'Title
        .ChartTitle.Font.Size = 20
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
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
        .Legend.Width = 230
        .Legend.Left = 155
        .Legend.Top = 7
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
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
           .Left = Sheets("Bullying").Range("A" & d).Left
           .Top = Sheets("Bullying").Range("A" & d).Top
           .Width = Sheets("Bullying").Range(Cells(d, 1), Cells(d, 9)).Width - 0.5
           .Height = Sheets("Bullying").Range(Cells(d, 1), Cells(d + 21, 9)).Height
    End With

 End With
 
                                 'Chart (Bullying by Adult)
  
    Range(Cells(c, 1), Cells(t, 1)).Copy Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Range(Cells(c, 2), Cells(t, 2)).Copy Range(Cells(f, 5), Cells(g, 5))
    Range(Cells(c, 3), Cells(t, 3)).Copy Range(Cells(f, 4), Cells(g, 4))
    Range(Cells(c, 4), Cells(t, 4)).Copy Range(Cells(f, 2), Cells(g, 2))
    Range(Cells(c, 5), Cells(t, 7)).Copy Range(Cells(f, 7), Cells(g, 9))
    Worksheets("Bullying").Cells(f, 3).Value = "Strongly Disagree"
    Worksheets("Bullying").Range(Cells(f + 1, 3), Cells(g, 3)).Value = 0
    Worksheets("Bullying").Cells(f, 6).Value = "Somewhat Disagree"
    Worksheets("Bullying").Range(Cells(f + 1, 6), Cells(g, 6)).Value = 0
    Set rngData = Worksheets("Bullying").Range(Cells(f + 1, 2), Cells(g, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(f, 1), Cells(g, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    

    
  Set Ws = Worksheets("Bullying")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Bullying: Victimization by Adults"   'Title
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
        .Legend.Width = 230
        .Legend.Left = 155
        .Legend.Top = 11
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
           .Left = Sheets("Bullying").Range("A" & f).Left
           .Top = Sheets("Bullying").Range("A" & f).Top
           .Width = Sheets("Bullying").Range(Cells(f, 1), Cells(f, 9)).Width - 0.5
           .Height = Sheets("Bullying").Range(Cells(f, 1), Cells(f + 20, 9)).Height
    End With

End With
                                                     
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
Next x
End Sub


