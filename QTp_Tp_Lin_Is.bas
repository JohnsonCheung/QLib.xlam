Attribute VB_Name = "QTp_Tp_Lin_Is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MTp_Tp_Lin_Is."
Private Const Asm$ = "QTp"
Function IsDDRmkLin(A$) As Boolean
Dim L$: L = LTrim(A)
If L <> "" Then
    If HasPfx(L, "--") Then
        IsDDRmkLin = True
    End If
End If
End Function
