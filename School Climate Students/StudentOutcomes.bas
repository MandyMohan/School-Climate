Attribute VB_Name = "StudentOutcomes"
Sub Aspirations()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim c As Long
Dim w As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Raw Data").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Raw Data").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:DI" & m).Value
    End With
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Student Outcomes"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Student Outcomes"
    ActiveSheet.Range("A" & t).Font.Size = 28
    t = t + 2
    c5 = t
                                           'Suspension
    
    ActiveSheet.Range("A" & t).Value = "Student Outcomes: Suspension"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I have not been suspended from school this year"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG2:CG" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG1:CG" & m), "I have not been suspended from school this year") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I have been suspended for one day"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG1:CG" & m), "I have been suspended for one day") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I have been suspended for two days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG1:CG" & m), "I have been suspended for two days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I have been suspended for three days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG1:CG" & m), "I have been suspended for three days ") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I have been suspended for four days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG1:CG" & m), "I have been suspended for four days") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I have been suspended for five or more days"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CG1:CG" & m), "I have been suspended for five or more days") / w * 100, 2) & "%"
    t = t + 1
    c = t
    'Chart (Suspension)

  Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c5 & ":B" & c - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "How many days have you been suspended out of school this year?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 18
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
        .Axes(xlValue).TickLabels.Font.Size = 14
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh
       

    With .Parent
           .Left = Sheets("Student Outcomes").Range("D" & c5).Left
           .Top = Sheets("Student Outcomes").Range("D" & c5).Top
           .Width = Sheets("Student Outcomes").Range("D" & c5 & ":M" & c5).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c5 & ":M" & c - 1).Height
    End With
End With

                                                 
                                             'Absenteeism

    ActiveSheet.Range("A" & t).Value = "Student Outcomes: Absenteeism"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Never or almost never"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CH2:CH" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CH1:CH" & m), "Never or almost never") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Once a week"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CH1:CH" & m), "Once a week") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Once every two weeks"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CH1:CH" & m), "Once every two weeks") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Once a month"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CH1:CH" & m), "Once a month") / w * 100, 2) & "%"
    t = t + 1
    c1 = t
    
  Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c & ":B" & c1 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
   With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlPie
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 18
        .ChartTitle.Text = "How often are you absent from school?" 'Title
        .SetElement (msoElementLegendRight)    'Add Legend to the Top
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .ChartColor = 21
    
    With .Parent
           .Left = Sheets("Student Outcomes").Range("D" & c).Left
           .Top = Sheets("Student Outcomes").Range("D" & c).Top
           .Width = Sheets("Student Outcomes").Range("D" & c & ":M" & c - 1).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c & ":M" & c1 - 1).Height
    End With
End With
 
                                             'Absenteeism

    ActiveSheet.Range("A" & t).Value = "What are the reasons you are most often absent?"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Illness/Injury"
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI2:CQ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Illness/Injury") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "No transport"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "No transport") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Caring for family member"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Caring for family member") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Family issues"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Family issues") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Issues with teachers"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Issues with teachers") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Not prepared for class"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Not prepared for class") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Lack of money"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Lack of money") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Don't feel like coming to school"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Don't feel like coming to school") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Employment"
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CI1:CQ" & m), "Employment") / w * 100, 2) & "%"
    t = t + 1
    c2 = t
    
    Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c1 & ":B" & c2 - 1)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "What are the reasons you are most often absent?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 18
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
        .Axes(xlValue).TickLabels.Font.Size = 14
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh

    With .Parent
           .Left = Sheets("Student Outcomes").Range("D" & c1).Left
           .Top = Sheets("Student Outcomes").Range("D" & c1).Top
           .Width = Sheets("Student Outcomes").Range("D" & c1 & ":M" & c1 - 1).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c1 & ":M" & c2 - 1).Height
    End With
End With

                                          'Academic Aspiration Subscale
                                           
    ActiveSheet.Range("A" & t).Value = "Student Outcomes: Academic Aspirations"
    ActiveSheet.Range("B" & t).Value = "% Respondents"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I do not have plans after Secondary School."
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT2:CT" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT1:CT" & m), "I do not have plans after Secondary School.") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I expect to attend a technical school, learn a trade, pursue an apprenticeship, or other educational opportunity."
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT1:CT" & m), "I expect to attend a technical school, learn a trade, pursue an apprenticeship, or other educational opportunity.") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I expect to attend university."
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT1:CT" & m), "I expect to attend university.") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I expect to get a job."
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT1:CT" & m), "I expect to get a job.") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I expect to go to Form 6."
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT1:CT" & m), "I expect to go to Form 6.") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "I expect to join the military/ police service/ fire service."
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("CT1:CT" & m), "I expect to join the military/ police service/ fire service.") / w * 100, 2) & "%"
     With ActiveSheet.Range("A" & c5 & ":B" & c5)
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
         .RowHeight = 60
    End With
    ActiveSheet.Range("A" & c5 & ":B" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A" & c5 + 1 & ":B" & t).Font.Size = 16
    ActiveSheet.Range("A1").ColumnWidth = 65
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
    
    Set Ws = Worksheets("Student Outcomes")
  Set Rang = Ws.Range("A" & c2 & ":B" & t)
  Set MyChart = Ws.Shapes.AddChart2
  
   With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarClustered
        .ChartTitle.Text = "What are your plans after Secondary school?"   'Title
        '.ChartTitle.Font.Color = vbBlack
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 18
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(102, 153, 255)
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
        .Axes(xlValue).TickLabels.Font.Size = 14
        .Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh
    

    With .Parent
           .Left = Sheets("Student Outcomes").Range("D" & c2).Left
           .Top = Sheets("Student Outcomes").Range("D" & c2).Top
           .Width = Sheets("Student Outcomes").Range("D" & c2 & ":M" & t).Width - 0.5
           .Height = Sheets("Student Outcomes").Range("D" & c2 & ":M" & t).Height
    End With
End With

    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
    
Next x
End Sub

