Attribute VB_Name = "MIde_Mth_Lin_Sig_Pm"
Option Explicit

Function NMthArg(MthLin) As Byte
NMthArg = Si(SplitComma(MthPm(MthLin)))
End Function

Function ArgNy(MthLin) As String()
ArgNy = NyzOy(ArgAy(MthLin))
End Function

Function ArgAy(MthLin) As Arg()
Dim P$()
    P = SplitComma(StrBetBkt(MthLin))
If Si(P) = 0 Then Exit Function
Dim O() As Arg
    Dim U%: U = UB(P)
    ReDim O(U)
    Dim J%
    For J = 0 To U
        Set O(J) = Arg(P(J))
    Next
ArgAy = O
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
'Lin_NMthArg
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

