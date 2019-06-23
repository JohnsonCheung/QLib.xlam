Attribute VB_Name = "QTp_B_Is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MTp_Tp_Lin_Is."
Private Const Asm$ = "QTp"
Function IsLinDDRmk(A$) As Boolean
Dim L$: L = LTrim(A)
If L <> "" Then
    If HasPfx(L, "--") Then
        IsLinDDRmk = True
    End If
End If
End Function
