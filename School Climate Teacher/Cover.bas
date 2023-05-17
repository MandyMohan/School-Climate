Attribute VB_Name = "Cover"
Sub CoverPage()
Dim x As Range
Dim rng As Range
Dim last As Long
Dim s As Shape
Dim sht As String
Dim m As Long
Dim c As Long
Dim a As Long
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
    With ActiveWorkbook.Sheets("Sheet1")
         .Cells.Interior.ColorIndex = 2
         .Range("A1").ColumnWidth = 50
         .Range("B1").ColumnWidth = 80
         .Range("A1").Value = x.Value
         .Range("A1").Font.Size = 36
         .Range("A2").Value = "School Climate Survey 2022 (Teachers)"
         .Range("A2").Font.Size = 28
         .Range("A4").Value = "School Climate Scales"
         .Range("A4").VerticalAlignment = xlVAlignCenter
         .Range("A4").Font.Size = 22
         .Range("A4").Font.Bold = True
         .Range("A4").Font.Underline = xlUnderlineStyleSingle
         Set s = .Shapes.AddTextbox(msoTextOrientationHorizontal, .Range("A6").Left, .Range("A6").Top, .Range("A6:B6").Width - 0.5, .Range("A6:B10").Height)
         s.TextFrame.Characters.Text = "Below lists the ten (10) key scales from the School Climate Survey 2022 that were completed by teachers. Each scale is composed of a series of items and responses were given based on a 4 or 6 point Likert scale."
         s.TextFrame.Characters.Font.Size = 16
         s.Line.Visible = msoFalse
         .Range("A11").Value = "Key Scales"
         .Range("B11").Value = "Description"
         .Range("A12").Value = "Relationships between Students and Adults: Respect for Students"
         .Range("B12").Value = "Staff perceived as supportive and respectful of students."
         .Range("A13").Value = "Relationships between Students and Adults: Willingness to Seek Help"
         .Range("B13").Value = "Staff perceived as supportive and helpful."
         .Range("A14").Value = "Relationships Among Adults: Collegiality"
         .Range("B14").Value = "Staff have positive, cooperative interactions with each other."
         .Range("A15").Value = "Bullying: Prevalence of Teasing and Bullying"
         .Range("B15").Value = "Perception that bullying and teasing occurs."
         .Range("A16").Value = "Bullying: Victimization by Adults"
         .Range("B16").Value = "Students have experienced bullying or been victimized by an adult at school."
         .Range("A17").Value = "Aggression: Student Aggression Toward Adults"
         .Range("B17").Value = "Staff have experienced student aggression at school."
         .Range("A18").Value = "Aggression: Adult Reactions to Student Aggression"
         .Range("B18").Value = "Staff reactions to student aggression at school."
         .Range("A19").Value = "Discipline: Concerns about Discipline and Safety"
         .Range("B19").Value = "Staff feel safe at school and are satisfied with displicinary pratices."
         .Range("A20").Value = "Discipline: School Disciplinary Structure"
         .Range("B20").Value = "School rules are strict but fair and not discriminatory"
         .Range("A21").Value = "Gang Activity"
         .Range("B21").Value = "Students are involved in gang activity."
          With .Range("A11:B11")
                .Font.Size = 20
                .Font.Color = vbBlack
                .Font.Bold = True
                .Interior.Color = RGB(165, 165, 165)
          End With
          .Range("A11:B21").Borders.LineStyle = xlContinuous
          .Range("A12:B21").Font.Size = 16
          .Range("A11:B21").RowHeight = 70
          .Range("A12:B21").WrapText = True
          .Range("A12:A21").Font.Bold = True
          .Range("A11:B21").HorizontalAlignment = xlHAlignLeft
          .Range("A11:B21").VerticalAlignment = xlVAlignCenter
        
    End With
    
    Sheets("Sheet1").Name = "Key Scales"
    
     'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    Next x
End Sub
