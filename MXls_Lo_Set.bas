Attribute VB_Name = "MXls_Lo_Set"
Option Explicit
Function WszLo(A As ListObject) As Worksheet
Set WszLo = A.Parent
End Function

Function LoSetNm(A As ListObject, LoNm$) As ListObject
If LoNm <> "" Then
    If Not HasLo(WszLo(A), LoNm) Then
        A.Name = LoNm
    Else
        Inf CSub, "Lo"
    End If
End If
Set LoSetNm = A
End Function


