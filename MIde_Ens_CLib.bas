Attribute VB_Name = "MIde_Ens_CLib"
Option Explicit
Sub EnsCLib(A As CodeModule, Optional B As eLibNmTy = eeByDic)
If A.CountOfLines > 0 Then Exit Sub
Dim Ept$: Ept = LinzEptCLib(A, B)
Dim Act As Lnx: Set Act = LnxzActCLibOpt(A)
Select Case True
Case IsNothing(Act):  A.InsertLines LnozAftOpt(A), Ept
Case Act.Lin <> Ept:  A.ReplaceLine Ept, Act.Lno
End Select
End Sub
Private Function LinzEptCLib$(A As CodeModule, Optional B As eLibNmTy = eeByDic)
LinzEptCLib = FmtQQ("Private Const CLib$ = ""?.""", LibNm(A, B))
End Function
Private Function LnxzActCLibOpt(A As CodeModule) As Lnx
Dim J&, L$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    Select Case True
    Case HasPfx(L, "Const CLib$ = "), HasPfx(L, "Private CLib$ = ")
        Set LnxzActCLibOpt = Lnx(J - 1, L): Exit Function
    End Select
Next
End Function




