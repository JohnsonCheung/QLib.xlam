Attribute VB_Name = "QIde_Mth_Lin_Fmt"
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Fmt."
Private Const Asm$ = "QIde"
Function NArg(MthLin) As Byte
NArg = Si(SplitComma(MthPm(MthLin)))
End Function

Function ArgNy(MthLin) As String()
ArgNy = NyzOy(ArgSy(MthLin))
End Function
Function ArgSy(Lin) As String()
ArgSy = SplitCommaSpc(MthPm(Lin))
End Function
Function ArgSfx$(ArgStr)

End Function
Function Arg(ArgStr) As Arg
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

