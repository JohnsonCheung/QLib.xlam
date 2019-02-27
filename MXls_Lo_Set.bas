Attribute VB_Name = "MXls_Lo_Set"
Option Explicit

Function SetLoNm(A As ListObject, LoNm$) As ListObject
If LoNm <> "" Then
    If Not HasLo(A, LoNm) Then
        A.Name = LoNm
    End If
End If
Set SetLoNm = A
End Function


