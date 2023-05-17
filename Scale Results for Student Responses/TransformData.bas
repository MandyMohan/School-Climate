Attribute VB_Name = "TransformData"
Sub Transform()
Dim x As Range
Dim y As Range
Dim z As Range
Dim rng As Range
Dim last As Long
Dim sht As String
Dim m As Long
Dim t As Long
Dim w As Long
Dim v As Variant
last = ActiveWorkbook.Sheets("Sheet1").Cells(Rows.Count, "F").End(xlUp).Row


For Each x In ActiveWorkbook.Sheets("Sheet1").Range("A1:BW" & last)
    
     If x.Value = "Strongly Agree" Or x.Value = "Never" Then
        x.Value = 5
     End If
     If x.Value = "Agree" Or x.Value = "Seldom" Then
        x.Value = 4
     End If
     If x.Value = "Neutral" Or x.Value = "Sometimes" Then
        x.Value = 3
     End If
     If x.Value = "Disagree" Or x.Value = "Often" Then
        x.Value = 2
     End If
     If x.Value = "Strongly Disagree" Or x.Value = "Always" Then
        x.Value = 1
     End If
     If x.Value = "12 or more times" Or x.Value = "Yes" Or x.Value = "All 30 days" Or x.Value = "Once a month" Or x.Value = "I do not have plans after Secondary School." Or x.Value = "I have been suspended for five or more days" Or x.Value = "Don't feel like coming to school" Then
        x.Value = 1
     End If
     If x.Value = "10 or 11 times" Or x.Value = "20 to 29 days" Or x.Value = "Not prepared for class" Or x.Value = "No" Or x.Value = "Once every two weeks" Or x.Value = "I expect to get a job." Or x.Value = "I have been suspended for four days" Then
        x.Value = 2
     End If
     If x.Value = "8 or 9 times" Or x.Value = "10 to 19 days" Or x.Value = "Issues with teachers" Or x.Value = "Once a week" Or x.Value = "I expect to join the military/ police service/ fire service." Or x.Value = "I have been suspended for three days " Then
        x.Value = 3
     End If
     If x.Value = "6 or 7 times" Or x.Value = "6 to 9 days" Or x.Value = "Employment" Or x.Value = "Never or almost never" Or x.Value = "I expect to attend a technical school, learn a trade, pursue an apprenticeship, or other educational opportunity." Or x.Value = "I have been suspended for two days" Then
        x.Value = 4
     End If
     If x.Value = "4 or 5 times" Or x.Value = "3 to 5 days" Or x.Value = "Family issues" Or x.Value = "I expect to go to Form 6." Or x.Value = "I have been suspended for one day" Then
        x.Value = 5
     End If
     If x.Value = "2 or 3 times" Or x.Value = "1 or 2 days" Or x.Value = "Caring for family member" Or x.Value = "I expect to attend university." Or x.Value = "I do not have plans after Secondary School." Or x.Value = "I have not been suspended from school this year" Then
        x.Value = 6
     End If
     If x.Value = "1 time" Or x.Value = "0 days" Or x.Value = "Lack of money" Then
        x.Value = 7
     End If
     If x.Value = "0 times" Or x.Value = "No transport" Then
        x.Value = 8
     End If
     If x.Value = "Illness/Injury" Then
        x.Value = 9
     End If
Next x
   
   For Each x In ActiveWorkbook.Sheets("Sheet1").Range("BX1:BX" & last)
        If x.Value = "6 or more times" Then
        x.Value = 1
        End If
        If x.Value = "4 or 5 times" Then
           x.Value = 2
        End If
        If x.Value = "2 or 3 times" Then
           x.Value = 3
        End If
        If x.Value = "1 time" Then
           x.Value = 4
        End If
        If x.Value = "0 times" Then
           x.Value = 5
        End If
        Next x

    For Each x In ActiveWorkbook.Sheets("Sheet1").Range("BY1:BZ" & last)
        If x.Value = "40 or more times" Or x.Value = "6 or more days" Then
        x.Value = 1
        End If
        If x.Value = "20 to 39 times" Or x.Value = "4 or 5 days" Then
           x.Value = 2
        End If
        If x.Value = "10 to 19 times" Or x.Value = "2 or 3 days" Then
           x.Value = 3
        End If
        If x.Value = "3 to 9 times" Or x.Value = "1 day" Then
           x.Value = 4
        End If
        If x.Value = "1 to 2 times" Or x.Value = "0 days" Then
           x.Value = 5
        End If
        If x.Value = "0 times" Then
           x.Value = 6
        End If
    Next x

    For Each z In ActiveWorkbook.Sheets("Sheet1").Range(Cells(1, 81), Cells(last, 107))
            If z.Value = "Strongly Agree" Or z.Value = "Never" Or z.Value = "No" Then
               z.Value = 4
            End If
            If z.Value = "Agree" Or z.Value = "Once or Twice" Or z.Value = "One Time" Or z.Value = "Sometimes" Then
               z.Value = 3
            End If
            If z.Value = "Disagree" Or z.Value = "About Once per Week" Or z.Value = "More than Once" Or z.Value = "Almost every day" Then
               z.Value = 2
            End If
            If z.Value = "Strongly Disagree" Or z.Value = "More than Once per Week" Or z.Value = "Many Times" Or z.Value = "Every day" Then
               z.Value = 1
            End If
    Next z

    For Each y In ActiveWorkbook.Sheets("Sheet1").Range(Cells(1, 108), Cells(last, 113))
                If y.Value = "Strongly Agree" Then
                   y.Value = 1
                End If
                If y.Value = "Agree" Then
                   y.Value = 2
                End If
                If y.Value = "Neutral" Then
                   y.Value = 3
                End If
                If y.Value = "Disagree" Then
                   y.Value = 4
                End If
                If y.Value = "Strongly Disagree" Then
                   y.Value = 5
                End If
        Next y

     
End Sub
