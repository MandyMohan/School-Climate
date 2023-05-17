Attribute VB_Name = "StudentRelationships"
Sub Relationships()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim d As Long
Dim i As Long
Dim e As Long
Dim w As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "BJ").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("BJ2:BJ" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Teachers Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:BF" & m).Value
    End With
    
                                        'Relationship Among Adults
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Relationship Among Adults"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Relationships Among Adults: Collegiality"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 13)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("M2:M" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("M1:M" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 14)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("N2:N" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("N1:N" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 15)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("O2:O" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("O1:O" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 16)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("P2:P" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("P1:P" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 17)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q2:Q" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("Q1:Q" & m), "Strongly Agree") / w * 100, 2) & "%"
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
    d = t + 5
    e = d + (t - 1)
    
    'Chart (Relationship Among Adults)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  
    Range("A1:A" & t).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("B1:B" & t).Copy Range(Cells(d, 5), Cells(e, 5))
    Range("C1:C" & t).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("D1:D" & t).Copy Range(Cells(d, 2), Cells(e, 2))
    Range("E1:G" & t).Copy Range(Cells(d, 7), Cells(e, 9))
    Worksheets("Relationship Among Adults").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("Relationship Among Adults").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Worksheets("Relationship Among Adults").Cells(d, 6).Value = "Somewhat Disagree"
    Worksheets("Relationship Among Adults").Range(Cells(d + 1, 6), Cells(e, 6)).Value = 0
    Set rngData = Worksheets("Relationship Among Adults").Range(Cells(d + 1, 2), Cells(e, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(d, 1), Cells(e, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
  Set Ws = Worksheets("Relationship Among Adults")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Relationships Among Adults: Collegiality"   'Title
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
        .Legend.Left = 175
        .Legend.Top = 1
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
        
        With ActiveSheet.Range("B1:C" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 3)).Merge
    Next i
        
    With .Parent
           .Left = Sheets("Relationship Among Adults").Range("A" & d).Left
           .Top = Sheets("Relationship Among Adults").Range("A" & d).Top
           .Width = Sheets("Relationship Among Adults").Range(Cells(d, 1), Cells(d, 9)).Width - 0.5
           .Height = Sheets("Relationship Among Adults").Range(Cells(d, 1), Cells(e + d + 13, 9)).Height
    End With

End With

    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0

Next x
End Sub
