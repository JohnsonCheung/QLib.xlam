Attribute VB_Name = "MxSeq"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSeq."

Function CvIntAy(A) As Integer()
On Error Resume Next
CvIntAy = A
End Function

Function CvLngAy(A) As Long()
On Error Resume Next
CvLngAy = A
End Function
