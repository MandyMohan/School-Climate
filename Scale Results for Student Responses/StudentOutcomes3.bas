Attribute VB_Name = "StudentOutcomes3"
Sub Aspiration()
Dim L As Long
Dim L1 As Long
Dim m As Long
Dim avg As Double, sch As Double
Dim oavg As Double, StdDev As Double
Dim last As Long
Dim lst As Long
Dim sht As String
Dim g As Integer
Dim w As Long
Dim SumSq As Single
Dim i As Integer
Dim sarray As Variant
Dim scharray As Variant
Dim thisarray As Variant
Dim avgarray As Variant
Dim v As Long
Dim s As Long



'change filter column in the following code
last = ThisWorkbook.Sheets("Sheet1").Cells(Rows.Count, "F").End(xlUp).Row
L1 = last - 1

ReDim thisarray(1 To L1) As Long
ReDim avgarray(1 To L1) As Double

v = 1
s = 2
m = 1
g = 2

For v = 1 To L1
  thisarray(v) = Sheets("Sheet1").Range("BP" & s)
  Sheets("Mean Scores").Range("V1").Value = "Student Outcomes:Academic Aspirations"
  Sheets("Mean Scores").Range("V" & s).Value = thisarray(v)
  s = s + 1
Next v
  
oavg = Application.WorksheetFunction.Sum(thisarray) / L1

For i = 1 To L1
    SumSq = SumSq + (thisarray(i) - oavg) ^ 2
Next i
 
StdDev = Sqr(SumSq / L1)

lst = ActiveWorkbook.Sheets("Sheet1").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Sheet1").Range("DL2:DL" & lst)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    
    L = ActiveWorkbook.Sheets("TransformData").Cells(Rows.Count, "F").End(xlUp).Row - 1
    ReDim sarray(1 To L) As Long
    ReDim scharray(1 To L) As Double
    
    For m = 1 To L
      sarray(m) = Sheets("TransformData").Range("BP" & g)
      g = g + 1
    Next m
    
      
    avg = Application.WorksheetFunction.Sum(sarray) / L
    
    sch = Round((avg - oavg) / StdDev + 10, 1)
    Sheets("Score Results").Range("A22").Value = "Academic Aspirations"
    Sheets("Score Results").Range("B22").Value = sch
                                                     
     'Save workbook
        ActiveWorkbook.Save
        
    'Close workbook
    ActiveWorkbook.Close
    
    m = 1
    g = 2
Next x

End Sub



