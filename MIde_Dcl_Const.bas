Attribute VB_Name = "MIde_Dcl_Const"
Option Explicit

Function ShfConst(O) As Boolean
ShfConst = ShfX(O, "Const")
End Function

Function HasConstNm(A As CodeModule, ConstNm$) As Boolean
Dim J%
For J = 1 To A.CountOfDeclarationLines
    If HitConstNm(A.Lines(J, 1), ConstNm) Then HasConstNm = True: Exit Function
Next
End Function

Function ConstNmzSrcLin$(SrcLin)
Dim L$: L = RmvMdy(SrcLin)
If ShfConst(L) Then ConstNmzSrcLin = TakNm(L)
End Function

