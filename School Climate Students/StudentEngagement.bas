Attribute VB_Name = "StudentEngagement"
Sub Engagement()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim d As Long
Dim i As Long
Dim f As Long
Dim g As Long
Dim h As Long
Dim j As Long
Dim e As Long
Dim w As Long
Dim c As Long
Dim c1 As Long
Dim a As Long
Dim b As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    
                                              'Affective Engagement Subscale
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Student Engagement"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Student Engagement: Affective Engagement"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Neutral"
    ActiveSheet.Range("E" & t).Value = "Agree"
    ActiveSheet.Range("F" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 9)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("I2:I" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 10)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("J2:J" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("J1:J" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 11)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("K2:K" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    c = t

                                                        'Cognitive Engagement Subscale
                                                        
    ActiveSheet.Range("A" & t).Value = "Student Engagement: Cognitive Engagement"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Neutral"
    ActiveSheet.Range("E" & t).Value = "Agree"
    ActiveSheet.Range("F" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 12)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("L2:L" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("L1:L" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 13)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("M2:M" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 14)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("N2:N" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    c1 = t
                                                    'Behavioural Engagement Subscale
                                                              
    ActiveSheet.Range("A" & t).Value = "Student Engagement: Behavioural Engagement"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Neutral"
    ActiveSheet.Range("E" & t).Value = "Agree"
    ActiveSheet.Range("F" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 15)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("O2:O" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 16)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("P2:P" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 17)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q2:Q" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Strongly Agree") / w * 100, 2) & "%"
    With ActiveSheet.Range("A1:F1")
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A1:F" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A2:F" & t).Font.Size = 16
    ActiveSheet.Range("A1:A" & t).RowHeight = 60
    ActiveSheet.Range("B1:F1").ColumnWidth = 20
    ActiveSheet.Range("A1").ColumnWidth = 60
    ActiveSheet.Range("A1:F" & t).WrapText = True
    ActiveSheet.Range("A1:A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A1:A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B1:F" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B1:F" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 6))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    With ActiveSheet.Range(Cells(c1, 1), Cells(c1, 6))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    a = c - 1
    b = c1 - 1
    d = 2
    e = c
    f = e + d + 23
    g = f + (c1 - c - 1)
    h = f + 28
    j = h + (t - c1)
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Student Engagement1"
    ActiveSheet.Range("A1:H1").ColumnWidth = 20
    
       'Chart (Affective Engagement)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  Dim rnge1 As Range
  Dim rnge2 As Range
  Dim rnge3 As Range
  Dim rnge4 As Range
  
    Sheets("Student Engagement").Range("A1:A" & a).Copy Sheets("Student Engagement1").Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Sheets("Student Engagement").Range("C1:C" & a).Copy Sheets("Student Engagement1").Range(Cells(d, 4), Cells(e, 4))
    Sheets("Student Engagement").Range("B1:B" & a).Copy Sheets("Student Engagement1").Range(Cells(d, 5), Cells(e, 5))
    Sheets("Student Engagement").Range("E1:F" & a).Copy Sheets("Student Engagement1").Range(Cells(d, 7), Cells(e, 8))
    Sheets("Student Engagement").Range("D1:D" & a).Copy Sheets("Student Engagement1").Range(Cells(d, 6), Cells(e, 6))
    Sheets("Student Engagement").Range("D1:D" & a).Copy Sheets("Student Engagement1").Range(Cells(d, 2), Cells(e, 2))
    Worksheets("Student Engagement1").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("Student Engagement1").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Sheets("Student Engagement1").Range(Cells(d, 1), Cells(e, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Student Engagement1").Range(Cells(d + 1, 2), Cells(e, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Student Engagement1").Range(Cells(d + 1, 6), Cells(e, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Student Engagement1").Range(Cells(d + 1, 2), Cells(e, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Sheets("Student Engagement1").Range(Cells(d, 1), Cells(e, 8)).Borders.LineStyle = xlNone
    Sheets("Student Engagement1").Range(Cells(d, 1), Cells(e, 8)).Interior.Color = xlNone
    Sheets("Student Engagement1").Range(Cells(d, 1), Cells(e, 8)).RowHeight = 15
    
  Set Ws = Worksheets("Student Engagement1")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Student Engagement: Affective Engagement"   'Title
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
        .Legend.Width = 165
        .Legend.Left = 155
        .Legend.Top = 15
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
           .Left = Sheets("Student Engagement1").Range("A" & d).Left
           .Top = Sheets("Student Engagement1").Range("A" & d).Top
           .Width = Sheets("Student Engagement1").Range(Cells(d, 1), Cells(d, 8)).Width - 0.5
           .Height = Sheets("Student Engagement1").Range(Cells(d, 1), Cells(e + d + 18, 8)).Height
    End With

End With
    
      'Chart (Cognitive Engagement)
  
    Sheets("Student Engagement").Range("A" & c & ":A" & b).Copy Sheets("Student Engagement1").Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Sheets("Student Engagement").Range("C" & c & ":C" & b).Copy Sheets("Student Engagement1").Range(Cells(f, 4), Cells(g, 4))
    Sheets("Student Engagement").Range("B" & c & ":B" & b).Copy Sheets("Student Engagement1").Range(Cells(f, 5), Cells(g, 5))
    Sheets("Student Engagement").Range("E" & c & ":F" & b).Copy Sheets("Student Engagement1").Range(Cells(f, 7), Cells(g, 8))
    Sheets("Student Engagement").Range("D" & c & ":D" & b).Copy Sheets("Student Engagement1").Range(Cells(f, 6), Cells(g, 6))
    Sheets("Student Engagement").Range("D" & c & ":D" & b).Copy Sheets("Student Engagement1").Range(Cells(f, 2), Cells(g, 2))
    Worksheets("Student Engagement1").Cells(f, 3).Value = "Strongly Disagree"
    Worksheets("Student Engagement1").Range(Cells(f + 1, 3), Cells(g, 3)).Value = 0
    Sheets("Student Engagement1").Range(Cells(f, 1), Cells(g, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Student Engagement1").Range(Cells(f + 1, 2), Cells(g, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Student Engagement1").Range(Cells(f + 1, 6), Cells(g, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Student Engagement1").Range(Cells(f + 1, 2), Cells(g, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Sheets("Student Engagement1").Range(Cells(f, 1), Cells(g, 8)).Borders.LineStyle = xlNone
    Sheets("Student Engagement1").Range(Cells(f, 1), Cells(g, 8)).Interior.Color = xlNone
    Sheets("Student Engagement1").Range(Cells(f, 1), Cells(g, 8)).RowHeight = 15
    
  Set Ws = Worksheets("Student Engagement1")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Student Engagement: Cognitive Engagement"   'Title
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
        .Legend.Width = 180
        .Legend.Left = 155
        .Legend.Top = 15
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
           .Left = Sheets("Student Engagement1").Range("A" & f).Left
           .Top = Sheets("Student Engagement1").Range("A" & f).Top
           .Width = Sheets("Student Engagement1").Range(Cells(f, 1), Cells(f, 8)).Width - 0.5
           .Height = Sheets("Student Engagement1").Range(Cells(f, 1), Cells(f + 23, 8)).Height
    End With

End With

    'Chart (Behavioural Engagement)

    Sheets("Student Engagement").Range("A" & c1 & ":A" & t).Copy Sheets("Student Engagement1").Range(Cells(h, 1), Cells(j, 1))  'Table w/ -ve values
    Sheets("Student Engagement").Range("C" & c1 & ":C" & t).Copy Sheets("Student Engagement1").Range(Cells(h, 4), Cells(j, 4))
    Sheets("Student Engagement").Range("B" & c1 & ":B" & t).Copy Sheets("Student Engagement1").Range(Cells(h, 5), Cells(j, 5))
    Sheets("Student Engagement").Range("E" & c1 & ":F" & t).Copy Sheets("Student Engagement1").Range(Cells(h, 7), Cells(j, 8))
    Sheets("Student Engagement").Range("D" & c1 & ":D" & t).Copy Sheets("Student Engagement1").Range(Cells(h, 6), Cells(j, 6))
    Sheets("Student Engagement").Range("D" & c1 & ":D" & t).Copy Sheets("Student Engagement1").Range(Cells(h, 2), Cells(j, 2))
    Worksheets("Student Engagement1").Cells(h, 3).Value = "Strongly Disagree"
    Worksheets("Student Engagement1").Range(Cells(h + 1, 3), Cells(j, 3)).Value = 0
    Sheets("Student Engagement1").Range(Cells(h, 1), Cells(j, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Student Engagement1").Range(Cells(h + 1, 2), Cells(j, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Student Engagement1").Range(Cells(h + 1, 6), Cells(j, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Student Engagement1").Range(Cells(h + 1, 2), Cells(j, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Sheets("Student Engagement1").Range(Cells(h, 1), Cells(j, 8)).Borders.LineStyle = xlNone
    Sheets("Student Engagement1").Range(Cells(h, 1), Cells(j, 8)).Interior.Color = xlNone
    Sheets("Student Engagement1").Range(Cells(h, 1), Cells(j, 8)).RowHeight = 15
    
    
    Set Ws = Worksheets("Student Engagement1")
  Set Rang = Ws.Range(Cells(h, 1), Cells(j, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Student Engagement: Behavioural Engagement"   'Title
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
        .Legend.Width = 180
        .Legend.Left = 185
        .Legend.Top = 15
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
           .Left = Sheets("Student Engagement1").Range("A" & h).Left
           .Top = Sheets("Student Engagement1").Range("A" & h).Top
           .Width = Sheets("Student Engagement1").Range(Cells(h, 1), Cells(h, 8)).Width - 0.5
           .Height = Sheets("Student Engagement1").Range(Cells(h, 1), Cells(h + 23, 8)).Height
    End With

End With
    
    
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub
