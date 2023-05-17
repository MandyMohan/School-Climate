Attribute VB_Name = "DisciplineSafety"
Sub Safety()
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
last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "BJ").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("BJ2:BJ" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Teachers Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Data")
         m = .Range("A" & .Rows.Count).End(xlUp).Row
         v = .Range("A1:BF" & m).Value
    End With
    
                                                 'Safety & Discipline
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Discipline & Safety"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = "Discipline: Concerns about Discipline and Safety"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 28)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB2:AB" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB1:AB" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB1:AB" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB1:AB" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB1:AB" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB1:AB" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AB1:AB" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 29)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC2:AC" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC1:AC" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC1:AC" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC1:AC" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC1:AC" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC1:AC" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AC1:AC" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 31)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD2:AD" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD1:AD" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD1:AD" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD1:AD" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD1:AD" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD1:AD" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AD1:AD" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 31)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE2:AE" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE1:AE" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE1:AE" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE1:AE" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE1:AE" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE1:AE" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AE1:AE" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 32)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF2:AF" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF1:AF" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF1:AF" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF1:AF" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF1:AF" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF1:AF" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AF1:AF" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 33)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG2:AG" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG1:AG" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG1:AG" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG1:AG" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG1:AG" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG1:AG" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AG1:AG" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    c = t
    
                                                      'Discipline Structure
    ActiveSheet.Range("A" & t).Value = "Discipline: School Discipline Structure"
    ActiveSheet.Range("B" & t).Value = "Strongly Disagree"
    ActiveSheet.Range("C" & t).Value = "Disagree"
    ActiveSheet.Range("D" & t).Value = "Somewhat Disagree"
    ActiveSheet.Range("E" & t).Value = "Somewhat Agree"
    ActiveSheet.Range("F" & t).Value = "Agree"
    ActiveSheet.Range("G" & t).Value = "Strongly Agree"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 44)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR2:AR" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AR1:AR" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 45)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS2:AS" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AS1:AS" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 46)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT2:AT" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AT1:AT" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 47)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU2:AU" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AU1:AU" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 48)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV2:AV" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AV1:AV" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 49)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW2:AW" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW1:AW" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW1:AW" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW1:AW" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW1:AW" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW1:AW" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AW1:AW" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 50)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX2:AX" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AX1:AX" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 51)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY2:AY" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AY1:AY" & m), "Strongly Agree") / w * 100, 2) & "%"
    t = t + 1
    ActiveSheet.Range("A" & t).Value = v(1, 52)
    w = Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ2:AZ" & m), "<>" & "")
    ActiveSheet.Range("B" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Strongly Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("C" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("D" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Somewhat Disagree") / w * 100, 2) & "%"
    ActiveSheet.Range("E" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Somewhat Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("F" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Agree") / w * 100, 2) & "%"
    ActiveSheet.Range("G" & t).Value = Round(Application.WorksheetFunction.CountIf(Sheets("Data").Range("AZ1:AZ" & m), "Strongly Agree") / w * 100, 2) & "%"
    With ActiveSheet.Range("A1:G1")
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    ActiveSheet.Range("A1:G" & t).Borders.LineStyle = xlContinuous
    ActiveSheet.Range("A2:G" & t).Font.Size = 16
    ActiveSheet.Range("A1:A" & t).RowHeight = 60
    ActiveSheet.Range("A1").ColumnWidth = 48.43
    ActiveSheet.Range("B1:G1").ColumnWidth = 20
    ActiveSheet.Range("A1:G" & t).WrapText = True
    ActiveSheet.Range("A1:A" & t).HorizontalAlignment = xlHAlignLeft
    ActiveSheet.Range("A1:A" & t).VerticalAlignment = xlVAlignCenter
    ActiveSheet.Range("B1:G" & t).HorizontalAlignment = xlHAlignCenter
    ActiveSheet.Range("B1:G" & t).VerticalAlignment = xlVAlignCenter
    With ActiveSheet.Range(Cells(c, 1), Cells(c, 7))
         .Font.Size = 16
         .Font.Color = vbBlack
         .Font.Bold = True
         .Interior.Color = RGB(165, 165, 165)
    End With
    a = c - 1
    d = 2
    e = c
    f = e + d + 30
    g = f + (t - c)
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Discipline & Safety1"
    ActiveSheet.Range("B1:I1").ColumnWidth = 20
    ActiveSheet.Range("A1").ColumnWidth = 8.43
    
      'Chart (Safety & Discipline)

  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Dim rngData As Range
  
    Sheets("Discipline & Safety").Range("A1:A" & a).Copy Sheets("Discipline & Safety1").Range(Cells(d, 1), Cells(e, 1))  'Table w/ -ve values
    Sheets("Discipline & Safety").Range("B1:B" & a).Copy Sheets("Discipline & Safety1").Range(Cells(d, 5), Cells(e, 5))
    Sheets("Discipline & Safety").Range("C1:C" & a).Copy Sheets("Discipline & Safety1").Range(Cells(d, 4), Cells(e, 4))
    Sheets("Discipline & Safety").Range("D1:D" & a).Copy Sheets("Discipline & Safety1").Range(Cells(d, 2), Cells(e, 2))
    Sheets("Discipline & Safety").Range("E1:G" & a).Copy Sheets("Discipline & Safety1").Range(Cells(d, 7), Cells(e, 9))
    Worksheets("Discipline & Safety1").Cells(d, 3).Value = "Strongly Disagree"
    Worksheets("Discipline & Safety1").Range(Cells(d + 1, 3), Cells(e, 3)).Value = 0
    Worksheets("Discipline & Safety1").Cells(d, 6).Value = "Somewhat Disagree"
    Worksheets("Discipline & Safety1").Range(Cells(d + 1, 6), Cells(e, 6)).Value = 0
    Set rngData = Worksheets("Discipline & Safety1").Range(Cells(d + 1, 2), Cells(e, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(d, 1), Cells(e, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    
  Set Ws = Worksheets("Discipline & Safety1")
  Set Rang = Ws.Range(Cells(d, 1), Cells(e, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Discipline: Concerns about Discipline and Safety"   'Title
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
        .Legend.Width = 230
        .Legend.Height = 20
        .Legend.Left = 145
        .Legend.Top = 10
        '.Legend.Font.Color = vbBlack
        .Legend.Font.Size = 14
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(6).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(7).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(8).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Legend.LegendEntries(1).Select
        Selection.Delete
        .Legend.LegendEntries(3).Select
        Selection.Delete
    With .Parent
           .Left = Sheets("Discipline & Safety1").Range("A" & d).Left
           .Top = Sheets("Discipline & Safety1").Range("A" & d).Top
           .Width = Sheets("Discipline & Safety1").Range(Cells(d, 1), Cells(d, 9)).Width - 0.5
           .Height = Sheets("Discipline & Safety1").Range(Cells(d, 1), Cells(e + d + 25, 9)).Height
    End With

 End With
 
                                 'Chart (Discpline Structure)
  
    Sheets("Discipline & Safety").Range("A" & c & ":A" & t).Copy Sheets("Discipline & Safety1").Range(Cells(f, 1), Cells(g, 1))  'Table w/ -ve values
    Sheets("Discipline & Safety").Range("B" & c & ":B" & t).Copy Sheets("Discipline & Safety1").Range(Cells(f, 5), Cells(g, 5))
    Sheets("Discipline & Safety").Range("C" & c & ":C" & t).Copy Sheets("Discipline & Safety1").Range(Cells(f, 4), Cells(g, 4))
    Sheets("Discipline & Safety").Range("D" & c & ":D" & t).Copy Sheets("Discipline & Safety1").Range(Cells(f, 2), Cells(g, 2))
    Sheets("Discipline & Safety").Range("E" & c & ":G" & t).Copy Sheets("Discipline & Safety1").Range(Cells(f, 7), Cells(g, 9))
    Worksheets("Discipline & Safety1").Cells(f, 3).Value = "Strongly Disagree"
    Worksheets("Discipline & Safety1").Range(Cells(f + 1, 3), Cells(g, 3)).Value = 0
    Worksheets("Discipline & Safety1").Cells(f, 6).Value = "Somewhat Disagree"
    Worksheets("Discipline & Safety1").Range(Cells(f + 1, 6), Cells(g, 6)).Value = 0
    Set rngData = Worksheets("Discipline & Safety1").Range(Cells(f + 1, 2), Cells(g, 6))
    rngData = Evaluate(rngData.Address & "*-1")
    With Range(Cells(f, 1), Cells(g, 9))
        .Font.Color = vbWhite
        .Borders.LineStyle = xlNone
        .Interior.Color = xlNone
        .RowHeight = 15
    End With
    

    
  Set Ws = Worksheets("Discipline & Safety1")
  Set Rang = Ws.Range(Cells(f, 1), Cells(g, 9))
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Rang
        .ChartType = xlBarStacked
        .ChartTitle.Text = "Discipline: School Discipline Structure"   'Title
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
        .Legend.Width = 230
        .Legend.Height = 20
        .Legend.Left = 155
        .Legend.Top = 7
        .Legend.Font.Size = 14
        '.Legend.Font.Color = vbBlack
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(5).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
        .SeriesCollection(6).Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
        .SeriesCollection(7).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(8).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Legend.LegendEntries(1).Select
        Selection.Delete
        .Legend.LegendEntries(3).Select
        Selection.Delete
        

    With .Parent
           .Left = Sheets("Discipline & Safety1").Range("A" & f).Left
           .Top = Sheets("Discipline & Safety1").Range("A" & f).Top
           .Width = Sheets("Discipline & Safety1").Range(Cells(f, 1), Cells(f, 9)).Width - 0.5
           .Height = Sheets("Discipline & Safety1").Range(Cells(f, 1), Cells(f + 40, 9)).Height
    End With

End With
                                                     
    'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    t = 0
Next x
End Sub

