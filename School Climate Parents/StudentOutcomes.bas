Attribute VB_Name = "StudentOutcomes"
Sub Outcomes()
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
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Student Outcomes"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Student Outcomes"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c5 = t
                                           'Violence
    
    ActiveSheet.Range("A" & t).Value = "Violence"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Almost always"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP2:BP" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Almost always") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Frequently"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Frequently") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Sometimes"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Sometimes") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Once in a while"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Once in a while") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Almost never"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BP1:BP" & m), "Almost never") / w * 100, 2) & "%"
    t = t + 1
    c = t
    'Chart (Violence)

  Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c5 & ":B" & c - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How often do you worry about violence at your child's school?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 217, 102)
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
           .Left = Sheets("Student Outcomes").Range("D" & c5).Left
           .Top = Sheets("Student Outcomes").Range("D" & c5).Top
           .Width = Sheets("Student Outcomes").Range("D" & c5 & ":M" & c5).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c5 & ":M" & c - 1).Height
    End With
End With
 
        
                                             'Safety
    ActiveSheet.Range("A" & t).Value = "Safety"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not at all unsafe"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS2:BS" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Not at all unsafe") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly unsafe"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Slightly unsafe") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat unsafe"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Somewhat unsafe") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite unsafe "
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Quite unsafe ") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely unsafe"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BS1:BS" & m), "Extremely unsafe") / w * 100, 2) & "%"
    t = t + 1
    c1 = t
    
  Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c & ":B" & c1 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  'Chart (Safety)
    With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "Overall, how unsafe does your child feel at school?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 153, 204)
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
           .Left = Sheets("Student Outcomes").Range("D" & c).Left
           .Top = Sheets("Student Outcomes").Range("D" & c).Top
           .Width = Sheets("Student Outcomes").Range("D" & c & ":M" & c - 1).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c & ":M" & c1 - 1).Height
    End With
End With

   'Drugs
    ActiveSheet.Range("A" & t).Value = "Drugs"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not a problem at all"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT2:BT" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "Not a problem at all") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "A little bit of a problem"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "A little bit of a problem") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "A moderate problem"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "A moderate problem") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite a problem"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "Quite a problem") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "A tremendous problem"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BT1:BT" & m), "A tremendous problem") / w * 100, 2) & "%"
    t = t + 1
    c2 = t
    
  Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c1 & ":B" & c2 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  'Chart (Drugs)
    With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "To what extent are drugs a problem at your child's school?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(206, 95, 86)
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
           .Left = Sheets("Student Outcomes").Range("D" & c1).Left
           .Top = Sheets("Student Outcomes").Range("D" & c1).Top
           .Width = Sheets("Student Outcomes").Range("D" & c1 & ":M" & c1 - 1).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c1 & ":M" & c2 - 1).Height
    End With
End With
    
    'Bullying 1

    ActiveSheet.Range("A" & t).Value = "Bullying: Accessibility of aid for victims"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not at all difficult"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ2:BQ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Not at all difficult") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly difficult"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Slightly difficult") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat difficult"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Somewhat difficult") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite difficult "
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Quite difficult ") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely difficult"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BQ1:BQ" & m), "Extremely difficult") / w * 100, 2) & "%"
    t = t + 1
    c3 = t
    
    'Chart (bullying 1)
    Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c2 & ":B" & c3 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
 With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .ChartTitle.Text = "If a student is bullied at your child’s school, how difficult is it for him/her to get help from an adult?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .ChartColor = 22
    
    With .Parent
           .Left = Sheets("Student Outcomes").Range("D" & c2).Left
           .Top = Sheets("Student Outcomes").Range("D" & c2).Top
           .Width = Sheets("Student Outcomes").Range("D" & c2 & ":M" & c2 - 1).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c2 & ":M" & c3 - 1).Height
    End With
End With

 'Bullying 2

    ActiveSheet.Range("A" & t).Value = "Bullying: Occurence of cyber bullying"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not at all likely"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR2:BR" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Not at all likely") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Slightly likely"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Slightly likely") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Somewhat likely"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Somewhat likely") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Quite likely"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Quite likely") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Extremely likely"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("BR1:BR" & m), "Extremely likely") / w * 100, 2) & "%"
    
    c4 = t
    
    'Chart (bullying 2)
    Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c3 & ":B" & c4)
  Set MyChart = Ws.Shapes.AddChart2
  
 With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Bold = True
        .ChartTitle.Text = "How likely is it that someone from your child’s school will bully him/her online?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .ChartColor = 18
    
    With .Parent
           .Left = Sheets("Student Outcomes").Range("D" & c3).Left
           .Top = Sheets("Student Outcomes").Range("D" & c3).Top
           .Width = Sheets("Student Outcomes").Range("D" & c3 & ":M" & c3 - 1).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c3 & ":M" & c4).Height
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
    ActiveSheet.Range("A1").ColumnWidth = 48.29
    ActiveSheet.Range("A" & c5 + 1 & ":A" & t).RowHeight = 40
    ActiveSheet.Range("A" & c5 & ":A" & t).WrapText = True
    ActiveSheet.Range("B1").ColumnWidth = 20
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
    With ActiveSheet.Range(Cells(c1, 1), Cells(c1, 2))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    With ActiveSheet.Range(Cells(c2, 1), Cells(c2, 2))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
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




