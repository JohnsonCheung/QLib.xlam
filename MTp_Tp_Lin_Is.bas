Attribute VB_Name = "MTp_Tp_Lin_Is"
Option Explicit
Function IsDDRmkLin(A$) As Boolean
Dim L$: L = LTrim(A)
If L <> "" Then
    If HasPfx(L, "--") Then
        IsDDRmkLin = True
    End If
End If
End Function
