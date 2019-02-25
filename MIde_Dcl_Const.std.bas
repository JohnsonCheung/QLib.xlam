Attribute VB_Name = "MIde_Dcl_Const"
Option Explicit

Function ShtConst(O) As Boolean
ShtConst = ShfX(O, "Const")
End Function

Function HasMdConstNm(A As CodeModule, ConstNm$) As Boolean
Dim J%
For J = 1 To A.CountOfDeclarationLines
    If ConstNmLin(A.Lines(J, 1)) = ConstNm Then HasMdConstNm = True: Exit Function
Next
End Function

Function ConstNmLin$(A)
Dim L$: L = RmvMdy(A)
If ShtConst(L) Then ConstNmLin = TakNm(L)
End Function

