Attribute VB_Name = "ParentInvolvement"
Sub Involvement()
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
Dim c As Long
Dim v As Variant

last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "CD").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("CD2:CD" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Parents Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:CA" & m).Value
    End With
    
                                        'Communication
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Parental Involvement"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Parental Involvement: Communication"
    ActiveSheet.Range("B" & t).Value = "Almost never"
    ActiveSheet.Range("C" & t).Value = "Once or twice per year"
    ActiveSheet.Range("D" & t).Value = "Every few months"
    ActiveSheet.Range("E" & t).Value = "Monthly"
    ActiveSheet.Range("F" & t).Value = "Weekly or more"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 3)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("C2:C" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Almost never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Once or twice per year") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Every few months") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Monthly") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("C1:C" & m), "Weekly or more") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 5)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("E2:E" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Almost never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Once or twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Every few months") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Monthly") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("E1:E" & m), "Weekly or more") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
                                                     'Parental Support
    ActiveSheet.Range("A" & t).Value = "Parental Support"
    ActiveSheet.Range("B" & t).Value = "Almost never"
    ActiveSheet.Range("C" & t).Value = "Once in a while"
    ActiveSheet.Range("D" & t).Value = "Sometimes"
    ActiveSheet.Range("E" & t).Value = "Frequently"
    ActiveSheet.Range("F" & t).Value = "Almost all the time"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 18)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("R2:R" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Almost never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Once in a while") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Frequently ") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("R1:R" & m), "Almost all the time") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 20)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("T2:T" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Almost never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Once in a while ") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Frequently") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("T1:T" & m), "Almost all the time") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 22)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("V2:V" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Almost never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Once in a while") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Sometimes") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Frequently") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("V1:V" & m), "Almost all the time") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 24)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("X2:X" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("X1:X" & m), "Almost never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("X1:X" & m), "Once in a while ") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("X1:X" & m), "Sometimes ") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("X1:X" & m), "Frequently") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("X1:X" & m), "Almost all the time") / w * 100, 2) & "%"
                                                     
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
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 6))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    a = c - 1
    d = t + 3
    e = d + (c - 2)
    f = d + 23
    g = f + (t - c)
   
       'Chart (Communication)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  Dim rnge1 As Range
  Dim rnge2 As Range
  
    Range("A1:A" & a).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("C1:C" & a).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("B1:B" & a).Copy Range(Cells(d, 5), Cells(e, 5))
    Range("E1:F" & a).Copy Range(Cells(d, 7), Cells(e, 8))
    Range("D1:D" & a).Copy Range(Cells(d, 6), Cells(e, 6))
    Range("D1:D" & a).Copy Range(Cells(d, 2), Cells(e, 2))
    Worksheets("Parental Involvement").Cells(d, 3).Value = "Almost never"
    Worksheets("Parental Involvement").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Range(Cells(d, 1), Cells(e, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Parental Involvement").Range(Cells(d + 1, 2), Cells(e, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Parental Involvement").Range(Cells(d + 1, 6), Cells(e, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Parental Involvement").Range(Cells(d + 1, 2), Cells(e, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(d, 1), Cells(e, 8)).Borders.LineStyle = xlNone
    Range(Cells(d, 1), Cells(e, 8)).Interior.Color = xlNone
    Range(Cells(d, 1), Cells(e, 8)).RowHeight = 15
    
  Set Ws = Worksheets("Parental Involvement")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Parental Involvement: Communication"   'Title
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
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Width = 250
        .Legend.Left = 155
        .Legend.Top = 12
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
           .Left = Sheets("Parental Involvement").Range("A" & d).Left
           .Top = Sheets("Parental Involvement").Range("A" & d).Top
           .Width = Sheets("Parental Involvement").Range(Cells(d, 1), Cells(d, 8)).Width - 0.5
           .Height = Sheets("Parental Involvement").Range(Cells(d, 1), Cells(d + 19, 8)).Height
    End With
 End With
    
    'Chart (Parent Support)
  
    Range(Cells(c, 1), Cells(t, 1)).Copy Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Range(Cells(c, 3), Cells(t, 3)).Copy Range(Cells(f, 4), Cells(g, 4))
    Range(Cells(c, 2), Cells(t, 2)).Copy Range(Cells(f, 5), Cells(g, 5))
    Range(Cells(c, 5), Cells(t, 6)).Copy Range(Cells(f, 7), Cells(g, 8))
    Range(Cells(c, 4), Cells(t, 4)).Copy Range(Cells(f, 6), Cells(g, 6))
    Range(Cells(c, 4), Cells(t, 4)).Copy Range(Cells(f, 2), Cells(g, 2))
    Worksheets("Parental Involvement").Cells(f, 3).Value = "Almost never"
    Worksheets("Parental Involvement").Range(Cells(f + 1, 3), Cells(g, 3)).Value = 0
    Range(Cells(d, 1), Cells(e, 8)).Font.Color = vbWhite
    Set rnge1 = Worksheets("Parental Involvement").Range(Cells(f + 1, 2), Cells(g, 2))
    rnge1 = Evaluate(rnge1.Address & "/2")
    Set rnge2 = Worksheets("Parental Involvement").Range(Cells(f + 1, 6), Cells(g, 6))
    rnge2 = Evaluate(rnge2.Address & "/2")
    Set rngData = Worksheets("Parental Involvement").Range(Cells(f + 1, 2), Cells(g, 5))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(f, 1), Cells(g, 8)).Borders.LineStyle = xlNone
    Range(Cells(f, 1), Cells(g, 8)).Interior.Color = xlNone
    Range(Cells(f, 1), Cells(g, 8)).RowHeight = 15
    
    
    With ActiveSheet.Range("B1:C" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 3)).Merge
    Next i
    
    Set Ws = Worksheets("Parental Involvement")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 8))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Parental Support"   'Title
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
        .Legend.Width = 200
        .Legend.Left = 175
        .Legend.Top = 10
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
           .Left = Sheets("Parental Involvement").Range("A" & f).Left
           .Top = Sheets("Parental Involvement").Range("A" & f).Top
           .Width = Sheets("Parental Involvement").Range(Cells(f, 1), Cells(f, 8)).Width - 0.5
           .Height = Sheets("Parental Involvement").Range(Cells(f, 1), Cells(f + 28, 8)).Height
    End With
  End With

      
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub


