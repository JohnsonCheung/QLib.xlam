Attribute VB_Name = "QIde_Mth_Pm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Pm_Arg."
Private Const Asm$ = "QIde"
Public Const DoczShtArg$ = "It is string from Arg"
Public Const DoczArg$ = "It is Sy.  It is splitting of MthPm"
Public Const DoczArgSy$ = "It Array of Arg"
Public Const DoczSset$ = "String-Aset"
Public Const DoczArgTy$ = "It is a string defining the type of an arg.  Eg, Dim A() as integer => ArgTy[Integer()].  Dim A%() => ArgTy[%()]"
Type Arg
    Nm As String
    IsOpt As Boolean
    IsPmAy As Boolean
    IsAy As Boolean
    TyChr As String
    AsTy As String
    DftVal As String
End Type
Type Args: N As Byte: Ay() As Arg: End Type
Function MthPm$(MthLin)
MthPm = BetBktMust(MthLin, CSub)
End Function

Property Get ArgAsetP() As Aset
Set ArgAsetP = ArgAsetzP(CPj)
End Property

Function ArgAsetzP(P As VBProject) As Aset
Set ArgAsetzP = New Aset
Dim L$, I
'For Each I In MthLinAyzP(A)
    L = I
    'ArgAsetzPj.PushAy ArgSy(L)
'Next
End Function

Private Sub Z_ArgAsetP()
ArgAsetP.Srt.Vc
End Sub

Function DimItmzArg$(Arg$)
DimItmzArg = BefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " =")
End Function

Function ShfArgPfx$(OLin$)
Select Case True
Case ShfTerm(OLin, "Optional "):   ShfArgPfx = "?"
Case ShfTerm(OLin, "Paramarray "): ShfArgPfx = ".."
End Select
End Function
Function ArgTy$(AftNm$)

End Function
Function ShtArg$(Arg$)
Dim L$:     L = Arg
Dim Pfx$:     Pfx = ShfArgPfx(L)
Dim Ty$: Ty = ArgTy(L)
'ShtArg = Pfx & Nm & Ty
End Function

Function FmtPm(Pm$, Optional IsNoBkt As Boolean) 'Pm is wo bkt.
Dim A$: A = Replace(Pm, "Optional ", "?")
Dim B$: B = Replace(A, " As ", ":")
Dim C$: C = Replace(B, "ParamArray ", "...")
If IsNoBkt Then
    FmtPm = C
Else
    FmtPm = QuoteSq(C)
End If
End Function

Function ShtRetTyAsetInVbe() As Aset
Set ShtRetTyAsetInVbe = ShtRetTyAsetzV(CVbe)
End Function

Function ShtRetTyAsetzV(A As Vbe) As Aset
Set ShtRetTyAsetzV = ShtRetTyAset(MthLinAyzV(A))
End Function

Function ShtRetTyAset(MthLinAy$()) As Aset
Set ShtRetTyAset = AsetzAy(ShtRetTyAy(MthLinAy))
End Function

Function ShtRetTyAy(MthLinAy$()) As String()
Dim MthLin, I
For Each I In Itr(MthLinAy)
    MthLin = I
    PushI ShtRetTyAy, ShtRetTyzLin(MthLin)
Next
End Function

Function ShtRetTyzLin(MthLin)
Dim A$: A = MthLinRec(MthLin).ShtRetTy
ShtRetTyzLin = A
If LasChr(A) = ":" Then Stop
End Function

Function ShtRetTy$(TyChr$, RetTy$, IsRetVal As Boolean, Optional ExlColon As Boolean)
Dim O$, Colon$
Colon = IIf(ExlColon, "", ":")
Select Case True
Case Not IsRetVal
Case TyChr = "" And RetTy = "": O = Colon & "Variant"
Case TyChr = "" And RetTy <> "": O = Colon & RetTy
Case RetTy <> "": Thw CSub, "TyChr and RetTy should both have value", "TyChr RetTy", TyChr, RetTy
Case Else: O = TyChr
End Select
ShtRetTy = O
End Function


Function ArgAyzPmAy(PmAy$()) As String()
Dim Pm, Arg
For Each Pm In Itr(PmAy)
    For Each Arg In Itr(SplitCommaSpc(Pm))
        PushI ArgAyzPmAy, Arg
    Next
Next
End Function

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

Function BrwrzRet$(Ret$)
Select Case True
Case IsTyChr(FstChr(Ret)): BrwrzRet = "Brw"
Case Ret = "Drs": BrwrzRet = "BrwDrs"
Case Else: BrwrzRet = "Brw"
End Select
End Function
Function DclAsAndChrAyzRet(Ret) As S1S2
Dim DclAs$, DclChrAy$
Select Case True
Case IsTyChr(FstChr(Ret)): DclChrAy = Ret
Case HasSfx(Ret, "()")
    Dim TyChr$: TyChr = TyChrzTyNm(RmvSfx(Ret, "()"))
    If TyChr = "" Then
        DclAs = " As " & Ret
    Else
        DclChrAy = TyChr & "()"
    End If
Case Else: DclAs = " As " & Ret
End Select
DclAsAndChrAyzRet = S1S2(DclAs, DclChrAy)
End Function
Function RetAs$(Ret)
If IsTyChr(FstChr(Ret)) Then
    RetAs = TyNmzTyChr(FstChr(Ret)) & RmvFstChr(Ret)
    Exit Function
End If
If TyChrzTyNm(Ret) <> "" Then Exit Function
RetAs = " As " & Ret
End Function
Function RetAszL$(MthLin$)
Dim L$: L = AftBkt(MthLin)
If Not ShfAs(L) Then Exit Function
Dim O$: O = ShfNm(L)
If ShfBkt(L) Then O = O & "()"
RetAszL = O
End Function
Function RetAszRet$(Ret)
RetAszRet = RetAs(Ret)
End Function
Function RetAszDclSfx$(DclSfx)
If DclSfx = "" Then Exit Function
Dim B$
Dim F$: F = FstChr(DclSfx)
If IsTyChr(F) Then
    If Len(DclSfx) = 1 Then Exit Function
    B = RmvFstChr(DclSfx): If B <> "()" Then Stop
    RetAszDclSfx = " As " & TyNmzTyChr(F) & "()"
    Exit Function
End If
If TyChrzTyNm(DclSfx) <> "" Then Exit Function
If Left(DclSfx, 4) <> " As " Then Stop
RetAszDclSfx = DclSfx
End Function
Function TyChrzDclSfx$(DclSfx)
If Len(DclSfx) = 1 Then
    If IsTyChr(DclSfx) Then TyChrzDclSfx = DclSfx
End If
End Function
Function TyChrzRet$(Ret)
If Len(Ret) = 1 And IsTyChr(Ret) Then TyChrzRet = Ret: Exit Function
Dim O$: O = TyChrzTyNm(Ret): If O <> "" Then TyChrzRet = O: Exit Function
End Function
Function ArgSfxzRet$(Ret)
'Ret is either FunRetTyChr (in Sht-TyChr) or
'              FunRetAs    (The Ty-Str without As)
Select Case True
Case IsTyChr(FstChr(Ret)): ArgSfxzRet = Ret
Case HasSfx(Ret, "()") And TyChrzTyNm(RmvSfx(Ret, "()")) <> "": ArgSfxzRet = TyChrzTyNm(RmvSfx(Ret, "()")) & "()"
Case Else: ArgSfxzRet = " As " & Ret
End Select
End Function
Function Arg(ArgStr) As Arg
Dim L$: L = ArgStr
Const Opt$ = "Optional"
Const PA$ = "ParamArray"
Dim TyChr$
With Arg
    .IsOpt = ShfPfxSpc(L, Opt)
    .IsPmAy = ShfPfxSpc(L, PA)
    .Nm = ShfNm(L)
    .TyChr = ShfChr(L, "!@#$%^&")
    .IsAy = ShfPfx(L, "()") = "()"
End With
End Function

Function ArgSfxy(ArgAy$()) As String()
Dim ArgStr
For Each ArgStr In Itr(ArgAy)
    PushI ArgSfxy, ArgSfx(ArgStr)
Next
End Function



