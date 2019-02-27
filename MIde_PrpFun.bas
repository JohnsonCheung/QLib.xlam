Attribute VB_Name = "MIde_PrpFun"
Option Explicit
Dim Info$()
Function IsPrpFunLin(Lin) As Boolean
Dim L$, B$
L = RmvMdy(Lin)
If ShfMthTy(L) <> "Function" Then Exit Function
If ShfNm(L) = "" Then Exit Function
ShfMthChr L
IsPrpFunLin = Left(L, 2) = "()"
End Function

Function PrpFunLnoAy(A As CodeModule) As Long()
Dim J&, L
For Each L In Src(A)
    J = J + 1
    If IsPrpFunLin(L) Then PushI PrpFunLnoAy, J
Next
End Function


Sub EnsPjFunMd(Md As CodeModule, Optional WhatIf As Boolean)
Dim L
For Each L In Itr(PrpFunLnoAy(Md))
    EnsPrpFunMdLno Md, L, WhatIf
Next
End Sub

Sub EnsPrpFunPj(Pj As VBProject, Optional WhatIf As Boolean)
Dim I
Erase Info
'For Each I In MdItr(Pj)
'    EnsPjFunMd CvMd(I), WhatIf
'Next
Brw Info
End Sub

Sub EnsPrpFun()
EnsPjFunMd CurMd
End Sub

Private Sub EnsPrpFunMdLno(A As CodeModule, Lno, Optional WhatIf As Boolean)
Dim OldLin$
Dim NewLin$
    OldLin = A.Lines(Lno, 1)
    NewLin = Replace(A.Lines(Lno, 1), "Function", "Property Get")
If Not WhatIf Then A.ReplaceLine Lno, NewLin
PushI Info, "EnsPrpFun:EnsPrpFunMdLno NewLin: " & OldLin
PushI Info, "                 OldLin: " & NewLin
End Sub
