Attribute VB_Name = "SocialEmotional"
Sub Learning()
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
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Social Emotional Learning"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Social-Emotional Learning"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Agree"
    ActiveSheet.Range("E" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 36)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ2:AJ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 37)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK2:AK" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 38)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL2:AL" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 39)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM2:AM" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 40)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN2:AN" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Strongly Agree") / w * 100, 2) & "%"
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
    d = t + 5
    e = d + (t - 1)
    
      'Chart (Social Emotional Learning)

      Dim Ws As Worksheet
      Dim Rang As Range
      Dim MyChart As Object
      Dim ser As Series
      
      Dim rngData As Range
      Dim rnge1 As Range
      Dim rnge2 As Range
      
      Range("A1:A" & t).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
      Range("C1:C" & t).Copy Range(Cells(d, 3), Cells(e, 3))
      Range("B1:B" & t).Copy Range(Cells(d, 4), Cells(e, 4))
      Range("D1:E" & t).Copy Range(Cells(d, 5), Cells(e, 6))
      Worksheets("Social Emotional Learning").Cells(d, 2).Value = "Strongly Disagree"
      Worksheets("Social Emotional Learning").Range(Cells(d + 1, 2), Cells(e, 2)).Value = 0
      Range(Cells(d, 1), Cells(e, 6)).Font.Color = vbWhite
      Set rngData = Worksheets("Social Emotional Learning").Range(Cells(d + 1, 2), Cells(e, 4))
      rngData = Evaluate(rngData.Address & "*-1")
      Range(Cells(d, 1), Cells(e, 6)).Borders.LineStyle = xlNone
      Range(Cells(d, 1), Cells(e, 6)).Interior.Color = xlNone
      Range(Cells(d, 1), Cells(e, 6)).RowHeight = 15
    
    With ActiveSheet.Range("B1:B" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 2)).Merge
    Next i
      
      Set Ws = Worksheets("Social Emotional Learning")
      Set Rang = Ws.Range(Cells(d, 1), Cells(e, 6))
      Set MyChart = Ws.Shapes.AddChart2
      
      With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Social-Emotional Learning"   'Title
        .ChartTitle.Font.Size = 20
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .PlotBy = IIf(.PlotBy = xlRows, xlColumns, xlRows) 'Switch row/column
        .Axes(xlValue).MinimumScale = -1    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
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
        .Legend.Top = 16
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Legend.LegendEntries(3).Select
        Selection.Delete
    
        With .Parent
               .Left = Sheets("Social Emotional Learning").Range("A" & d).Left
               .Top = Sheets("Social Emotional Learning").Range("A" & d).Top
               .Width = Sheets("Social Emotional Learning").Range(Cells(d, 1), Cells(d, 6)).Width - 0.5
               .Height = Sheets("Social Emotional Learning").Range(Cells(d, 1), Cells(e + d + 15, 6)).Height
        End With
    
    End With
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub

