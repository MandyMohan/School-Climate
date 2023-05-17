Attribute VB_Name = "ParentInvolvement3"
Sub Involvement3()
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
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Parental Involvement 3"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Parental Involvement"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c5 = t
                                           'Parental Involvement: Participation
    
    ActiveSheet.Range("A" & t).Value = "Involvement in School Activites"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely involved"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("D2:D" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Extremely involved") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite involved"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Quite involved") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat involved"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Somewhat involved") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly involved"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Slightly involved") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not at all involved"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("D1:D" & m), "Not at all involved") / w * 100, 2) & "%"
    c = t
    
    'Chart (Parental Involvement: Participation)

  Set Ws = Worksheets("Parental Involvement 3")
  Set Rang = Ws.Range("A" & c5 & ":B" & c)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How involved have you been at your child's school (e.g. with parents' groups, fund raising etc.)?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(153, 204, 0)
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
           .Left = Sheets("Parental Involvement 3").Range("D" & c5).Left
           .Top = Sheets("Parental Involvement 3").Range("D" & c5).Top
           .Width = Sheets("Parental Involvement 3").Range("D" & c5 & ":L" & c5).Width - 0.5
           .Height = Sheets("Parental Involvement 3").Range("D" & c5 & ":L" & c).Height
    End With
End With

            'Formating

With ActiveSheet.Range("A" & c5 & ":B" & c5)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    ActiveSheet.Range("A" & c5 & ":B" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c5 + 1 & ":B" & t).Font.Size = 16
    ActiveSheet.Range("A1").ColumnWidth = 44.29
    ActiveSheet.Range("C1").ColumnWidth = 3
    ActiveSheet.Range("A" & c5 + 1 & ":A" & t).RowHeight = 40
    ActiveSheet.Range("A" & c5 & ":A" & t).WrapText = True
    ActiveSheet.Range("B1").ColumnWidth = 20
    ActiveSheet.Range("A" & c5 & ":A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & c5 & ":A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B" & c5 & ":B" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B" & c5 & ":B" & t).VerticalAlignment = xlVAlignCenter
    
    'Student Engagement
    
    t = t + 3
    ActiveSheet.Range("A" & t).Value = "Student Engagement"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c4 = t
    
    'Table 1
    
    ActiveSheet.Range("A" & t).Value = " Effort put into School - Related Tasks"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "A tremendous amount of effort"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI2:AI" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "A tremendous amount of effort") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite a bit of effort"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "Quite a bit of effort") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Some effort"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "Some effort") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "A little bit of effort"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "A little bit of effort") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Almost no effort"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AI1:AI" & m), "Almost no effort") / w * 100, 2) & "%"
    t = t + 1
    c3 = t
    'Chart (Chart 1)

  Set Ws = Worksheets("Parental Involvement 3")
  Set Rang = Ws.Range("A" & c4 & ":B" & c3 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How much effort does your child put into school-related tasks?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 204, 153)
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
           .Left = Sheets("Parental Involvement 3").Range("D" & c4).Left
           .Top = Sheets("Parental Involvement 3").Range("D" & c4).Top
           .Width = Sheets("Parental Involvement 3").Range("D" & c4 & ":L" & c4).Width - 0.5
           .Height = Sheets("Parental Involvement 3").Range("D" & c4 & ":L" & c3 - 1).Height
    End With
End With

                                                 
                                             'Table 2
    ActiveSheet.Range("A" & t).Value = "Motivation to Learn"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely motivated"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ2:AJ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Extremely motivated") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite motivated"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Quite motivated") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat motivated"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Somewhat motivated") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly motivated"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Slightly motivated") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not at all motivated"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AJ1:AJ" & m), "Not at all motivated") / w * 100, 2) & "%"
    t = t + 1
    c1 = t
    
  Set Ws = Worksheets("Parental Involvement 3")
  Set Rang = Ws.Range("A" & c3 & ":B" & c1 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
    With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How motivated is your child to learn the topics covered in class?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(204, 255, 102)
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
           .Left = Sheets("Parental Involvement 3").Range("D" & c3).Left
           .Top = Sheets("Parental Involvement 3").Range("D" & c3).Top
           .Width = Sheets("Parental Involvement 3").Range("D" & c3 & ":L" & c3 - 1).Width - 0.5
           .Height = Sheets("Parental Involvement 3").Range("D" & c3 & ":L" & c1 - 1).Height
    End With
End With
 
                                             'Table 3

    ActiveSheet.Range("A" & t).Value = "Determination"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Almost all the time"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO2:AO" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Almost all the time") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Frequently"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Frequently") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Sometimes"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Sometimes") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Once in a while"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Once in a while") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Almost Never"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AO1:AO" & m), "Almost never") / w * 100, 2) & "%"
    c2 = t
    
    Set Ws = Worksheets("Parental Involvement 3")
  Set Rang = Ws.Range("A" & c1 & ":B" & c2)
  Set MyChart = Ws.Shapes.AddChart2
  
 With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .ChartTitle.Text = "How often does your child give up on learning activities that s/he finds hard?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .ChartColor = 21
    
    With .Parent
           .Left = Sheets("Parental Involvement 3").Range("D" & c1).Left
           .Top = Sheets("Parental Involvement 3").Range("D" & c1).Top
           .Width = Sheets("Parental Involvement 3").Range("D" & c1 & ":L" & c1 - 1).Width - 0.5
           .Height = Sheets("Parental Involvement 3").Range("D" & c1 & ":L" & c2).Height
    End With
End With

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
    With ActiveSheet.Range(Cells(c1, 1), Cells(c1, 2))
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
    
