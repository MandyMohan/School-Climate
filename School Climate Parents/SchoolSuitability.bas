Attribute VB_Name = "SchoolSuitability"
Sub Suitability()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim c As Long
Dim w As Long
Dim v As Variant

last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "CD").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("CD2:CD" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Parents Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:CA" & m).Value
    End With
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Suitability"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Suitability"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c5 = t
                                           'Table 1 Suitability
    
    ActiveSheet.Range("A" & t).Value = "School Suitability"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely good"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("K2:K" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Extremely good") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite good"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Quite good") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat good"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Somewhat good") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly good"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Slightly good ") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not good at all"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("K1:K" & m), "Not good at all ") / w * 100, 2) & "%"
    t = t + 1
    c = t
    'Chart (Chart 1)

  Set Ws = Worksheets("Suitability")
  Set Rang = Ws.Range("A" & c5 & ":B" & c - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "Given your child’s cultural background (ideas, customs, social behaviour), how good a fit is his/her school?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 204, 255)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlNone
        .Axes(xlCategory).ReversePlotOrder = True
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh
    
       

    With .Parent
           .Left = Sheets("Suitability").Range("D" & c5).Left
           .Top = Sheets("Suitability").Range("D" & c5).Top
           .Width = Sheets("Suitability").Range("D" & c5 & ":L" & c5).Width - 0.5
           .Height = Sheets("Suitability").Range("D" & c5 & ":L" & c - 1).Height
    End With
End With

                                                 
                                             'Table 2 Suitability
    ActiveSheet.Range("A" & t).Value = "Sense of belonging"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Great amount of belonging"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("I2:I" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Great amount of  belonging") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite a bit of belonging"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Quite a bit of belonging") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Some belonging"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "Some belonging") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "A little bit of belonging"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "A little bit of belonging") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "No belonging at all"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("I1:I" & m), "No belonging at all") / w * 100, 2) & "%"
   
    c1 = t
    
  Set Ws = Worksheets("Suitability")
  Set Rang = Ws.Range("A" & c & ":B" & c1)
  Set MyChart = Ws.Shapes.AddChart2
  
    With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How much of a sense of belonging does your child feel at his/her school?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(153, 204, 255)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlNone
        .Axes(xlCategory).ReversePlotOrder = True
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh

    With .Parent
           .Left = Sheets("Suitability").Range("D" & c).Left
           .Top = Sheets("Suitability").Range("D" & c).Top
           .Width = Sheets("Suitability").Range("D" & c & ":L" & c - 1).Width - 0.5
           .Height = Sheets("Suitability").Range("D" & c & ":L" & c1).Height
    End With
End With
 With ActiveSheet.Range("A" & c5 & ":B" & c5)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    ActiveSheet.Range("A" & c5 & ":B" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c5 + 1 & ":B" & t).Font.Size = 16
    ActiveSheet.Range("A1").ColumnWidth = 38.86
    ActiveSheet.Range("A" & c5 + 1 & ":A" & t).RowHeight = 40
    ActiveSheet.Range("A" & c5 & ":A" & t).WrapText = True
    ActiveSheet.Range("B1").ColumnWidth = 20
    ActiveSheet.Range("C1").ColumnWidth = 3
    ActiveSheet.Range("A" & c5 & ":A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & c5 & ":A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B" & c5 & ":B" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B" & c5 & ":B" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 2))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    
    'Table 1 Learning Environment
    
    t = t + 2
    ActiveSheet.Range("A" & t).Value = "Institutional Environment"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c4 = t
    
    ActiveSheet.Range("A" & t).Value = "Learning Environment"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely well"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT2:AT" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Extremely well") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite well"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Quite well") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat well"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Somewhat well") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly well"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Slightly well") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not well at all"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Not well at all") / w * 100, 2) & "%"
    t = t + 1
    c3 = t
    'Chart (Learning Environment)

  Set Ws = Worksheets("Suitability")
  Set Rang = Ws.Range("A" & c4 & ":B" & c3 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How well does  your child’s school create a school environment that helps children learn?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(153, 153, 255)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlNone
        .Axes(xlCategory).ReversePlotOrder = True
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh
    
       

    With .Parent
           .Left = Sheets("Suitability").Range("D" & c4).Left
           .Top = Sheets("Suitability").Range("D" & c4).Top
           .Width = Sheets("Suitability").Range("D" & c4 & ":L" & c4).Width - 0.5
           .Height = Sheets("Suitability").Range("D" & c4 & ":L" & c3 - 1).Height
    End With
End With

                                                 
                                             'Table 2 Learning Environment
    ActiveSheet.Range("A" & t).Value = "Student Enjoyment"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Enjoy a tremendous amount"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP2:AP" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Enjoy a tremendous amount") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Enjoy quite a bit"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Enjoy quite a bit") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Enjoy somewhat"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Enjoy somewhat") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Enjoy a little bit"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Enjoy a little bit") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Do not enjoy at all"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AP1:AP" & m), "Do not enjoy at all") / w * 100, 2) & "%"
    c2 = t
    
  Set Ws = Worksheets("Suitability")
  Set Rang = Ws.Range("A" & c3 & ":B" & c2)
  Set MyChart = Ws.Shapes.AddChart2
  
     With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Text = "To what extent do you think that children enjoy going to your child's school?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        .Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .ChartColor = 25
    
    With .Parent
           .Left = Sheets("Suitability").Range("D" & c3).Left
           .Top = Sheets("Suitability").Range("D" & c3).Top
           .Width = Sheets("Suitability").Range("D" & c3 & ":L" & c3 - 1).Width - 0.5
           .Height = Sheets("Suitability").Range("D" & c3 & ":L" & c2).Height
    End With
End With
 
         'Formating
 With ActiveSheet.Range("A" & c4 & ":B" & c4)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    ActiveSheet.Range("A" & c4 & ":B" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c4 + 1 & ":B" & t).Font.Size = 16
    ActiveSheet.Range("A" & c4 + 1 & ":A" & t).RowHeight = 40
    ActiveSheet.Range("A" & c4 & ":A" & t).WrapText = True
    ActiveSheet.Range("A" & c4 & ":A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & c4 & ":A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B" & c4 & ":B" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B" & c4 & ":B" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c3, 1), Cells(c3, 2))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    
    
   
        
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub



