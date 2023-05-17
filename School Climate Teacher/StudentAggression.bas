Attribute VB_Name = "StudentAggression"
Sub Aggression()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim c As Long
Dim m As Long
Dim a As Long
Dim i As Long
Dim t As Long
Dim d As Long
Dim f As Long
Dim g As Long
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
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Aggression"
    t = t + 1
                             
                                      'Student Aggression
                                      
    ActiveSheet.Range("A" & t).Value = "Aggression: Student Aggression Toward Adults"
    ActiveSheet.Range("B" & t).Value = "No"
    ActiveSheet.Range("C" & t).Value = "One Time"
    ActiveSheet.Range("D" & t).Value = "More than Once"
    ActiveSheet.Range("E" & t).Value = "Many Times"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 34)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AH2:AH" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AH1:AH" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AH1:AH" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AH1:AH" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AH1:AH" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 35)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI2:AI" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 36)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ2:AJ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 37)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK2:AK" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AK1:AK" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 38)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL2:AL" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AL1:AL" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
                                               'Reaction to Aggression
                                               
    ActiveSheet.Range("A" & t).Value = "Aggression: Adult Reactions to Student Aggression"
    ActiveSheet.Range("B" & t).Value = "Not true"
    ActiveSheet.Range("C" & t).Value = "A little true"
    ActiveSheet.Range("D" & t).Value = "Somewhat true"
    ActiveSheet.Range("E" & t).Value = "Definitely true"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 39)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM2:AM" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Not true") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "A little true") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Somewhat true") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AM1:AM" & m), "Definitely true") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 40)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN2:AN" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Not true") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "A little true") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Somewhat true") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AN1:AN" & m), "Definitely true") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 41)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO2:AO" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Not true") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "A little true") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Somewhat true") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Definitely true") / w * 100, 2) & "%"
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
    ActiveSheet.Range("A2:E" & t).WrapText = True
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
    d = t + 3
    e = d + (c - 2)
    f = d + 22
    g = f + (t - c)
    
       'Chart (Student Aggression)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  
    Range("A1:A" & a).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("C1:C" & a).Copy Range(Cells(d, 3), Cells(e, 3))
    Range("B1:B" & a).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("D1:E" & a).Copy Range(Cells(d, 5), Cells(e, 6))
    Worksheets("Aggression").Cells(d, 2).Value = "No"
    Worksheets("Aggression").Range(Cells(d + 1, 2), Cells(e, 2)).Value = 0
    Set rngData = Worksheets("Aggression").Range(Cells(d + 1, 2), Cells(e, 4))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(d, 1), Cells(e, 6))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
  Set Ws = Worksheets("Aggression")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 6))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Aggression: Student Aggression Toward Adults"   'Title
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
        .Legend.Width = 150
        .Legend.Left = 175
        .Legend.Top = 30
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Legend.LegendEntries(3).Select
        Selection.Delete
        
    With .Parent
           .Left = Sheets("Aggression").Range("A" & d).Left
           .Top = Sheets("Aggression").Range("A" & d).Top
           .Width = Sheets("Aggression").Range(Cells(d, 1), Cells(d, 6)).Width - 0.5
           .Height = Sheets("Aggression").Range(Cells(d, 1), Cells(d + 19, 6)).Height
    End With

End With
    
      'Chart (Reaction to Aggression)
  
    Range(Cells(c, 1), Cells(t, 1)).Copy Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Range(Cells(c, 3), Cells(t, 3)).Copy Range(Cells(f, 3), Cells(g, 3))
    Range(Cells(c, 2), Cells(t, 2)).Copy Range(Cells(f, 4), Cells(g, 4))
    Range(Cells(c, 4), Cells(t, 5)).Copy Range(Cells(f, 5), Cells(g, 6))
    Worksheets("Aggression").Cells(f, 2).Value = "Not True"
    Worksheets("Aggression").Range(Cells(f + 1, 2), Cells(g, 2)).Value = 0
    Set rngData = Worksheets("Aggression").Range(Cells(f + 1, 2), Cells(g, 4))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(f, 1), Cells(g, 6))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
    With ActiveSheet.Range("B1:B" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 2)).Merge
    Next i
    
  Set Ws = Worksheets("Aggression")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 6))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Aggression: Adult Reactions to Student Aggression"
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
        .Legend.Width = 150
        .Legend.Left = 175
        .Legend.Top = 30
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Legend.LegendEntries(3).Select
        Selection.Delete

    With .Parent
           .Left = Sheets("Aggression").Range("A" & f).Left
           .Top = Sheets("Aggression").Range("A" & f).Top
           .Width = Sheets("Aggression").Range(Cells(f, 1), Cells(f, 6)).Width - 0.5
           .Height = Sheets("Aggression").Range(Cells(f, 1), Cells(f + 19, 6)).Height
           
    End With

End With
    
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub

