Attribute VB_Name = "PrevalenceToTeasingAndBullying"
Sub Bullying()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim i As Long
Dim d As Long
Dim e As Long
Dim t As Long
Dim w As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Bullying"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Prevalence to Teasing and Bullying"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Neutral"
    ActiveSheet.Range("E" & t).Value = "Agree"
    ActiveSheet.Range("F" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 65)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BM2:BM" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BM1:BM" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BM1:BM" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BM1:BM" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BM1:BM" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BM1:BM" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 66)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BN2:BN" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BN1:BN" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BN1:BN" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BN1:BN" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BN1:BN" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BN1:BN" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 67)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BO2:BO" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BO1:BO" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BO1:BO" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BO1:BO" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BO1:BO" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BO1:BO" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 68)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP2:BP" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 69)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ2:BQ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Strongly Agree") / w * 100, 2) & "%"
    With ActiveSheet.Range("A1:F1")
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A1:F" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A2:F" & t).Font.Size = 16
    ActiveSheet.Range("A1:A" & t).RowHeight = 60
    ActiveSheet.Range("A1:H1").ColumnWidth = 20
    ActiveSheet.Range("A1:F" & t).WrapText = True
    ActiveSheet.Range("A1:A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A1:A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B1:F" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B1:F" & t).VerticalAlignment = xlVAlignCenter
    d = t + 5
    e = d + (t - 1)
    
      'Chart (Bullying)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  Dim rnge1 As Range
  Dim rnge2 As Range
  
    Range("A1:A" & t).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("C1:C" & t).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("B1:B" & t).Copy Range(Cells(d, 5), Cells(e, 5))
    Range("E1:F" & t).Copy Range(Cells(d, 7), Cells(e, 8))
    Range("D1:D" & t).Copy Range(Cells(d, 6), Cells(e, 6))
    Range("D1:D" & t).Copy Range(Cells(d, 2), Cells(e, 2))
    Worksheets("Bullying").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("Bullying").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Range(Cells(d, 1), Cells(e, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Bullying").Range(Cells(d + 1, 2), Cells(e, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Bullying").Range(Cells(d + 1, 6), Cells(e, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Bullying").Range(Cells(d + 1, 2), Cells(e, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(d, 1), Cells(e, 8)).Borders.LineStyle = xlNone
    Range(Cells(d, 1), Cells(e, 8)).Interior.Color = xlNone
    Range(Cells(d, 1), Cells(e, 8)).RowHeight = 15
    
    With ActiveSheet.Range("B1:C" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 3)).Merge
    Next i
    
  Set Ws = Worksheets("Bullying")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Prevalence to Teasing and Bullying"   'Title
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
        .Legend.Position = xlLegendPositionTop
        .Legend.Width = 170
        .Legend.Left = 165
        .Legend.Top = 8
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
           .Left = Sheets("Bullying").Range("A" & d).Left
           .Top = Sheets("Bullying").Range("A" & d).Top
           .Width = Sheets("Bullying").Range(Cells(d, 1), Cells(d, 8)).Width - 0.5
           .Height = Sheets("Bullying").Range(Cells(d, 1), Cells(e + d + 15, 8)).Height
    End With

 End With
                                                     
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
Next x
End Sub


