Attribute VB_Name = "Results"
Sub ScaleResults()
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
last = ActiveWorkbook.Sheets("Sheet1").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Sheet1").Range("DL2:DL" & last)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    With ActiveWorkbook.Sheets("Sheet1")
         .Cells.Interior.ColorIndex = 2
         .Range("A1").ColumnWidth = 50
         .Range("B1").ColumnWidth = 80
         .Range("C1").ColumnWidth = 15
         .Range("A1").Value = x.Value
         .Range("A1").Font.Size = 48
         .Range("A2").Value = "School Climate Survey 2022"
         .Range("A2").Font.Size = 28
         .Range("A3:A4").RowHeight = 30
         .Range("A5").Value = "Scale Results"
         .Range("A5").VerticalAlignment = xlVAlignCenter
         .Range("A5").Font.Size = 18
         .Range("A5").Font.Bold = True
         .Range("A5").Font.Underline = xlUnderlineStyleSingle
         .Range("A7").RowHeight = 35
         Set s = .Shapes.AddTextbox(msoTextOrientationHorizontal, .Range("A6").Left, .Range("A6").Top, .Range("A6:C9").Width - 0.5, .Range("A6:C9").Height)
         s.TextFrame.Characters.Text = "Here are the results for all scales from the surveys completed by students. Each scale is composed of a series of items that are averaged into an overall score for your school. Scores were standardized so that the mean score for the total sample is 10 and the standard deviation is 1. Thus, scores between 9 and 11 are within 1 standard deviation of the sample mean. Higher scores indicate a more favourable school climate."
         s.TextFrame.Characters.Font.Size = 14
         s.Line.Visible = msoFalse
         .Range("A11").Value = "Key Scales"
         .Range("B11").Value = "Discription"
         .Range("C11").Value = "Score"
         .Range("A12").Value = "Student Support: Respect for Students"
         .Range("B12").Value = "Staff perceived as supportive and respectful of students"
         .Range("C12").Value = Sheets("Score Results").Range("B1").Value
         .Range("A13").Value = "Student Support: Willingness to seek help"
         .Range("B13").Value = "Staff perceived as supportive and helpful"
         .Range("C13").Value = Sheets("Score Results").Range("B2").Value
         .Range("A14").Value = "Student Engagement: Affective Engagement"
         .Range("B14").Value = "Students feel a sense of belonging to the school"
         .Range("C14").Value = Sheets("Score Results").Range("B3").Value
         .Range("A15").Value = "Student Engagement: Cognitive Engagement"
         .Range("B15").Value = "Students are academically motivated"
         .Range("C15").Value = Sheets("Score Results").Range("B4").Value
         .Range("A16").Value = "Student Engagement: Behavioural Engagement"
         .Range("B16").Value = "Student put effort into school work and activities"
         .Range("C16").Value = Sheets("Score Results").Range("B5").Value
         .Range("A17").Value = "Relationship Among Students"
         .Range("B17").Value = "Students have positive interactions with each other and are respectful to one another"
         .Range("C17").Value = Sheets("Score Results").Range("B6").Value
         .Range("A18").Value = "Cultural Competence"
         .Range("B18").Value = "Students are treated equally despite demographic characteristics"
         .Range("C18").Value = Sheets("Score Results").Range("B7").Value
         .Range("A19").Value = "Social-Emotional Learning"
         .Range("B19").Value = "Students can recognize and control their emotions to make better decisions"
         .Range("C19").Value = Sheets("Score Results").Range("B8").Value
         .Range("A20").Value = "Expectations: Teacher Expectations"
         .Range("B20").Value = "Teachers have high expectations for student learning"
         .Range("C20").Value = Sheets("Score Results").Range("B9").Value
         .Range("A21").Value = "Expectations: Instructional Practices"
         .Range("B21").Value = "Use strategies to improve student learning"
         .Range("C21").Value = Sheets("Score Results").Range("B10").Value
         .Range("A22").Value = "School Disciplinary Structure"
         .Range("B22").Value = "School rules are strict but fair and not discriminatory"
         .Range("C22").Value = Sheets("Score Results").Range("B11").Value
         .Range("A23").Value = "Leadership"
         .Range("B23").Value = "Perception that the school’s administration provides effective leadership"
         .Range("C23").Value = Sheets("Score Results").Range("B12").Value
         .Range("A24").Value = "Personal Safety"
         .Range("B24").Value = "Students feel safe at school"
         .Range("C24").Value = Sheets("Score Results").Range("B13").Value
         .Range("A25").Value = "Prevalence of Teasing and Bullying"
         .Range("B25").Value = "Perception that bullying and teasing occurs."
         .Range("C25").Value = Sheets("Score Results").Range("B14").Value
         .Range("A26").Value = "Victimization: Bullying Experiences"
         .Range("B26").Value = "Students have experienced bullying at school."
         .Range("C26").Value = Sheets("Score Results").Range("B15").Value
         .Range("A27").Value = "Victimization: Victim Experiences"
         .Range("B27").Value = "Students have experienced victimization at school."
         .Range("C27").Value = Sheets("Score Results").Range("B16").Value
         .Range("A28").Value = "Physical Climate"
         .Range("B28").Value = "The physical infrastructure at the school is comfortable and conducive to learning."
         .Range("C28").Value = Sheets("Score Results").Range("B17").Value
         .Range("A29").Value = "Mental Health"
         .Range("B29").Value = "Students experience mental health issues."
         .Range("C29").Value = Sheets("Score Results").Range("B18").Value
         .Range("A30").Value = "Risky Behaviour"
         .Range("B30").Value = "Students engage in risky behaviour such as substance use, fighting etc."
         .Range("C30").Value = Sheets("Score Results").Range("B19").Value
         .Range("A31").Value = "Student Outcomes: Suspension"
         .Range("B31").Value = "Students have been suspended from school"
         .Range("C31").Value = Sheets("Score Results").Range("B20").Value
         .Range("A32").Value = "Student Outcomes: Absenteeism"
         .Range("B32").Value = "Students are often absent from school"
         .Range("C32").Value = Sheets("Score Results").Range("B21").Value
         .Range("A33").Value = "Student Outcomes: Academic Apirations"
         .Range("B33").Value = "Students have plans to further their education"
         .Range("C33").Value = Sheets("Score Results").Range("B22").Value
          With .Range("A11:C11")
                .Font.Size = 16
                .Font.Color = vbBlack
                .Font.Bold = True
                .Interior.Color = RGB(165, 165, 165)
          End With
          .Range("A11:C33").Borders.LineStyle = xlContinuous
          .Range("A12:C33").Font.Size = 14
          .Range("A11:C33").RowHeight = 40
          .Range("A12:C33").WrapText = True
          .Range("A12:A33").Font.Bold = True
          .Range("A11:B33").HorizontalAlignment = xlHAlignLeft
          .Range("A11:C33").VerticalAlignment = xlVAlignCenter
          .Range("C11:C33").HorizontalAlignment = xlHAlignCenter
        
    End With
    
    Sheets("Sheet1").Name = "Scale Results"
    
     'Save workbook
    ActiveWorkbook.Save
    
    'Close workbook
    ActiveWorkbook.Close
    
    Next x
End Sub

