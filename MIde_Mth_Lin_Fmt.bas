Attribute VB_Name = "MIde_Mth_Lin_Fmt"
Option Explicit
Function NArg(MthLin$) As Byte
NArg = Si(SplitComma(MthPm(MthLin)))
End Function

Function ArgNy(MthLin$) As String()
ArgNy = NyzOy(ArgSy(MthLin))
End Function
Function ArgSy(Lin$) As String()
ArgSy = SplitCommaSpc(MthPm(Lin))
End Function

Function Arg(ArgStr$) As Arg
Dim L$: L = ArgStr
Const Opt$ = "Optional"
Const PA$ = "ParamArray"
Dim TyChr$
Set Arg = New Arg
With Arg
    .IsOpt = ShfPfxSpc(L, Opt)
    .IsPmAy = ShfPfxSpc(L, PA)
    .Nm = ShfNm(L)
    .TyChr = ShfChr(L, "!@#$%^&")
    .IsAy = ShfPfx(L, "()") = "()"
End With
End Function

Function ArgNyzArgAy(A() As Arg) As String()
ArgNyzArgAy = NyzOy(A)
End Function

Private Sub Z()
Exit Sub
'Lin_NArg
'Lin_PmNy A
'MthMthPmzLinStr
'MthLinPmAy
'MthPm
'MthPmAyNy
'MthPmSz
'MthPmTyAsTyNm
'MthPmTyShtNm
'MthPmUB
'PushMthPm
End Sub

