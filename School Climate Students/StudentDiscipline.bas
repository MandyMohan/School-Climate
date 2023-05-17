Attribute VB_Name = "StudentDiscipline"
Sub Discipline()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim i As Long
Dim d As Long
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
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "School Disciplinary Structure"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "School Disciplinary Structure"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Neutral"
    ActiveSheet.Range("E" & t).Value = "Agree"
    ActiveSheet.Range("F" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 50)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX2:AX" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 51)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY2:AY" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 52)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ2:AZ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 53)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BA2:BA" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BA1:BA" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BA1:BA" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BA1:BA" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BA1:BA" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BA1:BA" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 54)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BB2:BB" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BB1:BB" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BB1:BB" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BB1:BB" & m), "Neutral") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BB1:BB" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BB1:BB" & m), "Strongly Agree") / w * 100, 2) & "%"
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
    
    
    'Chart (School Disciplinary Structure)

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
    Worksheets("School Disciplinary Structure").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("School Disciplinary Structure").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Range(Cells(d, 1), Cells(e, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("School Disciplinary Structure").Range(Cells(d + 1, 2), Cells(e, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("School Disciplinary Structure").Range(Cells(d + 1, 6), Cells(e, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("School Disciplinary Structure").Range(Cells(d + 1, 2), Cells(e, 5))
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
    
      Set Ws = Worksheets("School Disciplinary Structure")
      Set Rang = Ws.Range(Cells(d, 1), Cells(e, 8))
      Set MyChart = Ws.Shapes.AddChart2
      
      With MyChart.Chart
            .SetSourceData Rang
            .ChartType = xlBarStacked
            .ChartTitle.Text = "School Disciplinary Structure"   'Title
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
            .Legend.Left = 175
            .Legend.Top = 9
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
               .Left = Sheets("School Disciplinary Structure").Range("A" & d).Left
               .Top = Sheets("School Disciplinary Structure").Range("A" & d).Top
               .Width = Sheets("School Disciplinary Structure").Range(Cells(d, 1), Cells(d, 8)).Width - 0.5
               .Height = Sheets("School Disciplinary Structure").Range(Cells(d, 1), Cells(e + d + 15, 8)).Height
        End With
    
     End With
                                                     
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
Next x
End Sub
