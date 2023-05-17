Attribute VB_Name = "RiskyBehaviours"
Sub Behaviour()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim w As Long
Dim c As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    
                                              'Alcohol
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Risky Behaviours"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Risky Behaviours"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c5 = t
    ActiveSheet.Range("A" & t).Value = v(1, 99)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "0 days"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU2:CU" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "0 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "1 or 2 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "1 or 2 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "3 to 5 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "3 to 5 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "6 to 9 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "6 to 9 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "10 to 19 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "10 to 19 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "20 to 29 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "20 to 29 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "All 30 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CU1:CU" & m), "All 30 days") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
    'Chart (Alcohol)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Set Ws = Worksheets("Risky Behaviours")
  Set Rang = Ws.Range("A" & c5 & ":B" & c - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "During the past 30 days, on how many days did you have at least one drink of alcohol?"   'Title
        .ChartTitle.Font.Size = 16
        .ChartTitle.Font.Bold = True
        '.ChartTitle.Font.Color = vbBlack
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14

    With .Parent
           .Left = Sheets("Risky Behaviours").Range("D" & c5).Left
           .Top = Sheets("Risky Behaviours").Range("D" & c5).Top
           .Width = Sheets("Risky Behaviours").Range("D" & c5 & ":M" & c5).Width - 0.5
           .Height = Sheets("Risky Behaviours").Range("D" & c5 & ":M" & c - 1).Height
    End With
End With
    
                                           'Marijuana
                                             
    ActiveSheet.Range("A" & t).Value = v(1, 100)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "0 times"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV2:CV" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV1:CV" & m), "0 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "1 to 2 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV1:CV" & m), "1 to 2 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "3 to 9 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV1:CV" & m), "3 to 9 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "10 to 19 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV1:CV" & m), "10 to 19 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "20 to 39 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV1:CV" & m), "20 to 39 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "40 or more times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CV1:CV" & m), "40 or more times") / w * 100, 2) & "%"
    t = t + 1
    c1 = t
    
    
      'Chart (Marijuana)

  Set Ws = Worksheets("Risky Behaviours")
  Set Rang = Ws.Range("A" & c & ":B" & c1 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "During the past 30 days, how many times did you use marijuana?" 'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 16
        .ChartColor = "26"
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14
    

    With .Parent
           .Left = Sheets("Risky Behaviours").Range("D" & c).Left
           .Top = Sheets("Risky Behaviours").Range("D" & c).Top
           .Width = Sheets("Risky Behaviours").Range("D" & c & ":M" & c).Width - 0.5
           .Height = Sheets("Risky Behaviours").Range("D" & c & ":M" & c1 - 1).Height
    End With
End With
    
                                              'Weapon
                                              
    ActiveSheet.Range("A" & t).Value = v(1, 101)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "0 days"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CW2:CW" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CW1:CW" & m), "0 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "1 day"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CW1:CW" & m), "1 day") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "2 or 3 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CW1:CW" & m), "2 or 3 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "4 or 5 days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CW1:CW" & m), "4 or 5 days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "6 or more days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CW1:CW" & m), "6 or more days") / w * 100, 2) & "%"
    t = t + 1
    c2 = t
    
    
     'Chart (Weapon)

  Set Ws = Worksheets("Risky Behaviours")
  Set Rang = Ws.Range("A" & c1 & ":B" & c2 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "During the past 30 days, on how many days did you carry a weapon such as a gun, knife or club?" 'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 16
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(112, 48, 160)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14

    With .Parent
           .Left = Sheets("Risky Behaviours").Range("D" & c1).Left
           .Top = Sheets("Risky Behaviours").Range("D" & c1).Top
           .Width = Sheets("Risky Behaviours").Range("D" & c1 & ":M" & c1).Width - 0.5
           .Height = Sheets("Risky Behaviours").Range("D" & c1 & ":M" & c2 - 1).Height
    End With
End With
    
    
    
                                                 'Fight
                                                 
    ActiveSheet.Range("A" & t).Value = v(1, 102)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "0 times"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX2:CX" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "0 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "1 time"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "1 time") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "2 or 3 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "2 or 3 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "4 or 5 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "4 or 5 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "6 or 7 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "6 or 7 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "8 or 9 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "8 or 9 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "10 or 11 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "10 or 11 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "12 or more times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CX1:CX" & m), "12 or more times") / w * 100, 2) & "%"
    t = t + 1
    c3 = t
    
    
     'Chart (Fight)

  Set Ws = Worksheets("Risky Behaviours")
  Set Rang = Ws.Range("A" & c2 & ":B" & c3 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "During the past 30 days, how many times were you in a physical fight on school property?" 'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 16
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14

    With .Parent
           .Left = Sheets("Risky Behaviours").Range("D" & c2).Left
           .Top = Sheets("Risky Behaviours").Range("D" & c2).Top
           .Width = Sheets("Risky Behaviours").Range("D" & c2 & ":M" & c2).Width - 0.5
           .Height = Sheets("Risky Behaviours").Range("D" & c2 & ":M" & c3 - 1).Height
    End With
End With
    
                                                 'Think Suicide
                                                
    ActiveSheet.Range("A" & t).Value = v(1, 103)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Yes"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CY2:CY" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CY1:CY" & m), "Yes") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "No"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CY1:CY" & m), "No") / w * 100, 2) & "%"
    t = t + 1
    c4 = t
    
    
         'Chart (Think Suicide)

  Set Ws = Worksheets("Risky Behaviours")
  Set Rang = Ws.Range("A" & c3 & ":B" & c4 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 16
        .ChartTitle.Font.Bold = True
        .ChartTitle.Text = "During the past 12 months, did you ever seriously consider attempting suicide?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14

    With .Parent
           .Left = Sheets("Risky Behaviours").Range("D" & c3).Left
           .Top = Sheets("Risky Behaviours").Range("D" & c3).Top
           .Width = Sheets("Risky Behaviours").Range("D" & c3 & ":M" & c3).Width - 0.5
           .Height = Sheets("Risky Behaviours").Range("D" & c3 & ":M" & c4 - 1).Height
    End With
End With
    
                                                'Attempt Suicide
    ActiveSheet.Range("A" & t).Value = v(1, 104)
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "0 times"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CZ2:CZ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CZ1:CZ" & m), "0 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "1 time"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CZ1:CZ" & m), "1 time") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "2 or 3 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CZ1:CZ" & m), "2 or 3 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "4 or 5 times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CZ1:CZ" & m), "4 or 5 times") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "6 or more times"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CZ1:CZ" & m), "6 or more times") / w * 100, 2) & "%"
    With ActiveSheet.Range("A" & c5 & ":B" & c5)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    ActiveSheet.Range("A" & c5 & ":B" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c5 + 1 & ":B" & t).Font.Size = 16
    ActiveSheet.Range("A1").ColumnWidth = 60
    ActiveSheet.Range("A" & c5 + 1 & ":A" & t).RowHeight = 31.5
    ActiveSheet.Range("A" & c5 & ":B" & t).WrapText = True
    ActiveSheet.Range("B1").ColumnWidth = 20
    ActiveSheet.Range("A" & c5 & ":B" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A" & c5 & ":B" & t).VerticalAlignment = xlVAlignCenter
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
    With ActiveSheet.Range(Cells(c4, 1), Cells(c4, 2))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    
    
     'Chart (Attempt Suicide)

  Set Ws = Worksheets("Risky Behaviours")
  Set Rang = Ws.Range("A" & c4 & ":B" & t)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlColumnClustered
        .ChartTitle.Text = "During the past 12 months, how many times did you actually attempt suicide?" 'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Size = 16
        .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(199, 41, 191)
        .SeriesCollection(1).HasDataLabels = True
        '.SeriesCollection(1).DataLabels.Font.ColorIndex = 1
        .SeriesCollection(1).DataLabels.Font.Size = 14
        .Axes(xlValue).MinimumScale = 0    'Adjust scale
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%;0%;0%"     'Remove decimals from scale
        .Axes(xlValue).MajorGridlines.Delete    'Remove Gridlines
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        '.Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 14
        '.Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 14

    With .Parent
           .Left = Sheets("Risky Behaviours").Range("D" & c4).Left
           .Top = Sheets("Risky Behaviours").Range("D" & c4).Top
           .Width = Sheets("Risky Behaviours").Range("D" & c4 & ":M" & c4).Width - 0.5
           .Height = Sheets("Risky Behaviours").Range("D" & c4 & ":M" & t).Height
    End With
End With
    t = 0
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close

Next x
End Sub

