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

last = ActiveWorkbook.Sheets("Data").Cells(Rows.Count, "CD").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Data").Range("CD2:CD" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Parents Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Sheet1")
         .Cells.Interior.ColorIndex = 2
         .Range("A1").ColumnWidth = 50
         .Range("B1").ColumnWidth = 80
         .Range("A1").Value = x.Value
         .Range("A1").Font.Size = 36
         .Range("A2").Value = "School Climate Survey 2022 (Parents)"
         .Range("A2").Font.Size = 28
         .Range("A4").Value = "School Climate Scales"
         .Range("A4").VerticalAlignment = xlVAlignCenter
         .Range("A4").Font.Size = 22
         .Range("A4").Font.Bold = True
         .Range("A4").Font.Underline = xlUnderlineStyleSingle
         Set s = .Shapes.AddTextbox(msoTextOrientationHorizontal, .Range("A6").Left, .Range("A6").Top, .Range("A6:B6").Width - 0.5, .Range("A6:B10").Height)
         s.TextFrame.Characters.Text = "Below lists the nine (9) key scales from the School Climate Survey 2022 that were completed by parents. Each scale is composed of a series of items and responses were given based on a 4 or 6 point Likert scale."
         s.TextFrame.Characters.Font.Size = 16
         s.Line.Visible = msoFalse
         .Range("A11").Value = "Key Scales"
         .Range("B11").Value = "Description"
         .Range("A12").Value = "Parental Involvement"
         .Range("B12").Value = "Parents involvement in school: Communication, Issues and Participation."
         .Range("A13").Value = "Parental Support"
         .Range("B13").Value = "Parents' ability to support students at home."
         .Range("A14").Value = "Student Engagement"
         .Range("B14").Value = "Parents perceive students as liking their school."
         .Range("A15").Value = "Parent Support at School"
         .Range("B15").Value = "Parents' confidence in supporting students at school."
         .Range("A16").Value = "Respect Between Adult and Student"
         .Range("B16").Value = "Perception that there is mutual respect between staff and students."
         .Range("A17").Value = "Student Outcomes"
         .Range("B17").Value = "Perception of Violence, Safety, Drugs and Bullying at schools."
         .Range("A18").Value = "School Suitability"
         .Range("B18").Value = "The school is suitable for the student."
         .Range("A19").Value = "School Compatibility"
         .Range("B19").Value = "The school is compatible with the student."
         .Range("A20").Value = "Institutional Environment"
         .Range("B20").Value = "The school provides suitable learning and enjoyment for students"
         With .Range("A11:B11")
                .Font.Size = 20
                .Font.Color = vbBlack
                .Font.Bold = True
                .Interior.Color = RGB(165, 165, 165)
          End With
          .Range("A11:B20").Borders.LineStyle = xlContinuous
          .Range("A12:B20").Font.Size = 16
          .Range("A11:B20").RowHeight = 70
          .Range("A12:B20").WrapText = True
          .Range("A12:A20").Font.Bold = True
          .Range("A11:B20").HorizontalAlignment = xlHAlignLeft
          .Range("A11:B20").VerticalAlignment = xlVAlignCenter
        
    End With
    
    Sheets("Sheet1").Name = "Key Scales"
    
     'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    Next x
End Sub

