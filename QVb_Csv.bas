Attribute VB_Name = "QVb_Csv"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Csv."
Private Const Asm$ = "QVb"

Function CvCsv$(A)
Select Case True
Case IsStr(A): CvCsv = """" & A & """"
Case IsDte(A): CvCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
Case Else: CvCsv = IIf(IsNull(A), "", A)
End Select
End Function

Function CsvzDr$(A)
If Si(A) = 0 Then Exit Function
Dim O$(), U&, J&, V
U = UB(A)
ReDim O(U)
For Each V In A
    O(J) = CvCsv(V)
    J = J + 1
Next
CsvzDr = Join(O, ",")
End Function
