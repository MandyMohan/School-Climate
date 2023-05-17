Attribute VB_Name = "TeacherExpectation2"
Sub Instructional()
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
Dim v1 As Long
Dim v2 As Long
Dim s As Long

'change filter column in the following code
last = ThisWorkbook.Sheets("Sheet1").Cells(Rows.Count, "F").End(xlUp).Row
L1 = last - 1

ReDim thisarray(1 To L1) As Long
ReDim avarray(1 To L1) As Variant
ReDim avgarray(1 To L1) As Double

v = 1
v1 = 1
s = 2
m = 1
m1 = 1
g = 2


For v = 1 To L1
  thisarray(v) = Application.WorksheetFunction.Sum(Sheets("Sheet1").Range("CG" & s & ":CK" & s))
  w = Application.WorksheetFunction.Count(Sheets("Sheet1").Range("CG" & s & ":CK" & s))
  If thisarray(v) <> 0 Then
    avarray(v) = thisarray(v) / w
    avgarray(v1) = avarray(v)
    v1 = v1 + 1
  Else
    avarray(v) = ""
  End If
  Sheets("Mean Scores").Range("J1").Value = "Teachers Expectation: Instructional Practices"
  Sheets("Mean Scores").Range("J" & s).Value = avarray(v)
  s = s + 1
Next v

v2 = v1 - 1
ReDim Preserve avgarray(1 To v2) As Double

oavg = Application.WorksheetFunction.Sum(avgarray) / v2

For i = 1 To v2
    SumSq = SumSq + (avgarray(i) - oavg) ^ 2
Next i
 
StdDev = Sqr(SumSq / v2)

lst = ActiveWorkbook.Sheets("Sheet1").Cells(Rows.Count, "DL").End(xlUp).Row
For Each x In ActiveWorkbook.Sheets("Sheet1").Range("DL2:DL" & lst)
    Workbooks.Open "C:\Users\" & Environ("username") & "\Documents\School Climate\" & x.Value & " School Climate Students Report 2022.xlsx"
    
    L = ActiveWorkbook.Sheets("TransformData").Cells(Rows.Count, "F").End(xlUp).Row - 1
    ReDim sarray(1 To L) As Long
    ReDim scharray(1 To L) As Double
    ReDim scarray(1 To L) As Variant
    
    For m = 1 To L
      sarray(m) = Application.WorksheetFunction.Sum(Sheets("TransformData").Range("CG" & g & ":CK" & g))
      w = Application.WorksheetFunction.Count(Sheets("TransformData").Range("CG" & g & ":CK" & g))
      If sarray(m) <> 0 Then
        scarray(m) = sarray(m) / w
        scharray(m1) = scarray(m)
        m1 = m1 + 1
      End If
      g = g + 1
    Next m
      
    ReDim Preserve avgarray(1 To m1) As Double
    m2 = m1 - 1
      
    avg = Application.WorksheetFunction.Sum(scharray) / m2
    
    sch = Round((avg - oavg) / StdDev + 10, 1)
    Sheets("Score Results").Range("A10").Value = "Instructional Practices"
    Sheets("Score Results").Range("B10").Value = sch
                                                     
     'Save workbook
        ActiveWorkbook.Save
        
    'Close workbook
    ActiveWorkbook.Close
    
    m = 1
    m1 = 1
    g = 2
Next x

End Sub






