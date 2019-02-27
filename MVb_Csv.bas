Attribute VB_Name = "MVb_Csv"
Option Explicit

Function CvCsv$(A)
Select Case True
Case IsStr(A): CvCsv = """" & A & """"
Case IsDte(A): CvCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
Case Else: CvCsv = IIf(IsNull(A), "", A)
End Select
End Function

Function CsvzDr$(A)
If Sz(A) = 0 Then Exit Function
Dim O$(), U&, J&, V
U = UB(A)
ReDim O(U)
For Each V In A
    O(J) = CvCsv(V)
    J = J + 1
Next
CsvzDr = Join(O, ",")
End Function
