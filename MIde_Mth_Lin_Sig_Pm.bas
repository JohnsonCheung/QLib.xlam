Attribute VB_Name = "MIde_Mth_Lin_Sig_Pm"
Option Explicit

Function NPrm(MthLin) As Byte
NPrm = Sz(SplitComma(TakBetBkt(MthLin)))
End Function
Function ArgNm$(Arg)
ArgNm = RmvPfxAySpc(Arg, TermAy("Optional Paramarray [By Val] [By Ref]"))
End Function
Function ArgNy(MthLin) As String()
Dim Arg
For Each Arg In Itr(SplitComma(TakBetBkt(MthLin)))
    PushI ArgNy, ArgNm(Arg)
Next
End Function

Function ArgAy(MthLin$) As Arg()
Dim P$()
    P = SplitComma(TakBetBkt(MthLin))
If Sz(P) = 0 Then Exit Function
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

Function ArgNyArgAy(A() As Arg) As String()
Dim J%
For J = 0 To UB(A)
    PushI ArgNyArgAy, A(J).Nm
Next
End Function

Function ArgTy$(A As Arg)
Dim B$
With A
    If .IsAy Or .IsPmAy Then B = "()"
    If .TyChr <> "" Then ArgTy = ArgTyNmTyChr(.TyChr) & B: Exit Function
    If .AsTy = "" Then
        ArgTy = "Variant" & B
    Else
        ArgTy = .AsTy & B
    End If
End With
End Function


Private Sub Z()
Exit Sub
'Lin_NPrm
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
Function MthPmStr$(MthLin)
If IsMthLin(MthLin) Then
    MthPmStr = TakBetBkt(MthLin)
End If
End Function

