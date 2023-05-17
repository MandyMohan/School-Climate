Attribute VB_Name = "VictimExperiences"
Sub Victimization()
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
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Victimization"
    t = t + 1
                             
                                      'Bullying Victimization
                                      
    ActiveSheet.Range("A" & t).Value = "Victimization: Bullying Experiences"
    ActiveSheet.Range("B" & t).Value = "Never"
    ActiveSheet.Range("C" & t).Value = "Once or Twice"
    ActiveSheet.Range("D" & t).Value = "About Once per Week"
    ActiveSheet.Range("E" & t).Value = "More than Once per Week"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 70)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR2:BR" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Once or Twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "About Once per Week") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "More than Once per Week") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 71)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS2:BS" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Once or Twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "About Once per Week") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "More than Once per Week") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 72)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT2:BT" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "Once or Twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "About Once per Week") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "More than Once per Week") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 73)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BU2:BU" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BU1:BU" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BU1:BU" & m), "Once or Twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BU1:BU" & m), "About Once per Week") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BU1:BU" & m), "More than Once per Week") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 74)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BV2:BV" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BV1:BV" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BV1:BV" & m), "Once or Twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BV1:BV" & m), "About Once per Week") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BV1:BV" & m), "More than Once per Week") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 75)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BW2:BW" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BW1:BW" & m), "Never") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BW1:BW" & m), "Once or Twice") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BW1:BW" & m), "About Once per Week") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BW1:BW" & m), "More than Once per Week") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
                                               'Victim Experiences
                                               
    ActiveSheet.Range("A" & t).Value = "Victimization: Victim Experiences"
    ActiveSheet.Range("B" & t).Value = "No"
    ActiveSheet.Range("C" & t).Value = "One Time"
    ActiveSheet.Range("D" & t).Value = "More than Once"
    ActiveSheet.Range("E" & t).Value = "Many Times"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 76)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BX2:BX" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BX1:BX" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BX1:BX" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BX1:BX" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BX1:BX" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 77)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BY2:BY" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BY1:BY" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BY1:BY" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BY1:BY" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BY1:BY" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 78)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BZ2:BZ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BZ1:BZ" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BZ1:BZ" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BZ1:BZ" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BZ1:BZ" & m), "Many Times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 79)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CA2:CA" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CA2:CA" & m), "No") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CA2:CA" & m), "One Time") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CA2:CA" & m), "More than Once") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CA2:CA" & m), "Many Times") / w * 100, 2) & "%"
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
    ActiveSheet.Range("B1:F1").WrapText = True
    ActiveSheet.Range("A1:E" & t).WrapText = True
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
    d = t + 2
    e = d + (c - 2)
    f = d + 18
    g = f + (t - c)
    
       'Chart (Bullying Victimisation)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  Dim rnge1 As Range
  Dim rnge2 As Range
  Dim rnge3 As Range
  Dim rnge4 As Range
  
    Range("A1:A" & a).Copy Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Range("C1:C" & a).Copy Range(Cells(d, 3), Cells(e, 3))
    Range("B1:B" & a).Copy Range(Cells(d, 4), Cells(e, 4))
    Range("D1:E" & a).Copy Range(Cells(d, 5), Cells(e, 6))
    Worksheets("Victimization").Cells(d, 2).Value = "Never"
    Worksheets("Victimization").Range(Cells(d + 1, 2), Cells(e, 2)).Value = 0
    Range(Cells(d, 1), Cells(e, 6)).Font.Color = vbWhite
    Set rngData = Worksheets("Victimization").Range(Cells(d + 1, 2), Cells(e, 4))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(d, 1), Cells(e, 6)).Borders.LineStyle = xlNone
    Range(Cells(d, 1), Cells(e, 6)).Interior.Color = xlNone
    Range(Cells(d, 1), Cells(e, 6)).RowHeight = 15
    
  Set Ws = Worksheets("Victimization")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 6))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Victimization: Bullying Experiences"   'Title
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
        .Legend.Width = 220
        .Legend.Left = 155
        .Legend.Top = 35
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Legend.LegendEntries(3).Select
        Selection.Delete
        
    With .Parent
           .Left = Sheets("Victimization").Range("A" & d).Left
           .Top = Sheets("Victimization").Range("A" & d).Top
           .Width = Sheets("Victimization").Range(Cells(d, 1), Cells(d, 6)).Width - 0.5
           .Height = Sheets("Victimization").Range(Cells(d, 1), Cells(d + 16, 6)).Height
    End With

End With
    
      'Chart (Victim Experiences)
  
    Range(Cells(c, 1), Cells(t, 1)).Copy Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Range(Cells(c, 3), Cells(t, 3)).Copy Range(Cells(f, 3), Cells(g, 3))
    Range(Cells(c, 2), Cells(t, 2)).Copy Range(Cells(f, 4), Cells(g, 4))
    Range(Cells(c, 4), Cells(t, 5)).Copy Range(Cells(f, 5), Cells(g, 6))
    Worksheets("Victimization").Cells(f, 2).Value = "No"
    Worksheets("Victimization").Range(Cells(f + 1, 2), Cells(g, 2)).Value = 0
    Range(Cells(f, 1), Cells(g, 6)).Font.Color = vbWhite
    Set rngData = Worksheets("Victimization").Range(Cells(f + 1, 2), Cells(g, 4))
    rngData = Evaluate(rngData.Address & "*-1")
    Range(Cells(f, 1), Cells(g, 6)).Borders.LineStyle = xlNone
    Range(Cells(f, 1), Cells(g, 6)).Interior.Color = xlNone
    Range(Cells(f, 1), Cells(g, 6)).RowHeight = 15
    
    With ActiveSheet.Range("B1:B" & t)
         .Insert Shift:=xlToRight
    End With
    
    For i = 1 To t
        ActiveSheet.Range(Cells(i, 1), Cells(i, 2)).Merge
    Next i
    
  Set Ws = Worksheets("Victimization")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 6))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Victimization: Victim Experiences"   'Title
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
        .Legend.Width = 150
        .Legend.Left = 170
        .Legend.Top = 35
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 195, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
        .Legend.LegendEntries(3).Select
        Selection.Delete

    With .Parent
           .Left = Sheets("Victimization").Range("A" & f).Left
           .Top = Sheets("Victimization").Range("A" & f).Top
           .Width = Sheets("Victimization").Range(Cells(f, 1), Cells(f, 6)).Width - 0.5
           .Height = Sheets("Victimization").Range(Cells(f, 1), Cells(f + 16, 6)).Height
    End With

End With
    
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub

