Attribute VB_Name = "QIde_Mth_Pm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Pm_Arg."
Private Const Asm$ = "QIde"
Public Const DoczShtArg$ = "It is string from Arg"
Public Const DoczArg$ = "It is Str.  It is splitting of MthPm"
Public Const DoczArgAy$ = "It Array of Arg"
Public Const DoczSset$ = "String-Aset"
Public Const DoczDArg$ = "Dt of Arg: Nm IsByVal IsOpt IsPmAy IsAy TyChr AsTy DftVal"
Public Const DoczArgTy$ = "It is a string defining the type of an arg.  Eg, Dim A() as integer => ArgTy[Integer()].  Dim A%() => ArgTy[%()]"

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
    FmtPm = QteSq(C)
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
NArg = Si(SplitComma(BetBkt(MthLin)))
End Function

Function RetAs$(Ret)
If IsTyChr(FstChr(Ret)) Then
    RetAs = TyNmzTyChr(FstChr(Ret)) & RmvFstChr(Ret)
    Exit Function
End If
If TyChrzTyNm(Ret) <> "" Then Exit Function
RetAs = " As " & Ret
End Function

Private Sub Z_RetAszL()
Dim L, A As S1S2s: For Each L In MthLinAyP
    PushS1S2 A, S1S2(RetAszL(L), L)
Next
BrwS1S2s A
End Sub

Function RetAszL$(MthLin)
Dim A$: A = AftBkt(MthLin)
Dim B$: B = BefOrAll(A, ":")
Dim C$: C = BefOrAll(B, "'")
RetAszL = RmvPfx(C, "As ")
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
Select Case True
Case Left(DclSfx, 4) = " As ":      RetAszDclSfx = DclSfx
Case Left(DclSfx, 6) = "() As ":    RetAszDclSfx = Mid(DclSfx, 3) & "()"
Case Else: Stop
End Select
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

Sub Z_DArg()
Dim Mth$(): Mth = StrCol(DMthP, "MthLin")
Dim L, Dry(): For Each L In Itr(Mth)
    PushIAy Dry, DryArg(L)
Next
BrwDrs Drs(FnyArg, Dry)
End Sub
Private Function DryArgzMthLinAy(MthLinAy$()) As Variant()
Dim L: For Each L In Itr(MthLinAy)
    PushIAy DryArgzMthLinAy, DryArg(L)
Next
End Function

Private Function DryArg(MthLin) As Variant()
Dim Pm$: Pm = BetBkt(MthLin)
Dim A$(): A = SplitCommaSpc(Pm)
Dim Mthn$: Mthn = MthnzLin(MthLin)
Dim Arg, Dry(): For Each Arg In Itr(A)
    PushI DryArg, DrArg(Arg, Mthn)
Next
End Function
Private Function FnyArg() As String()
FnyArg = SyzSS("Mthn Nm IsOpt IsByVal IsPmAy IsAy TyChr AsTy DftVal")
End Function

Function DArgzMthLinAy(MthLinAy$()) As Drs
DArgzMthLinAy = Drs(FnyArg, DryArgzMthLinAy(MthLinAy))
End Function

Function DArgzM(M As CodeModule) As Drs
Dim L$(): L = StrCol(DMthzM(M), "MthLin")
     DArgzM = Drs(FnyArg, DryArgzMthLinAy(L))
End Function

Function DArgzP(P As VBProject) As Drs
Dim M$(): M = StrCol(DMthzP(P), "MthLin")
     DArgzP = Drs(FnyArg, DryArgzMthLinAy(M))
End Function

Function DArgP() As Drs
DArgP = DArgzP(CPj)
End Function

Function DArg(MthLin) As Drs
DArg = DrszFF("Mthn Nm IsOpt IsByVal IsPmAy IsAy TyChr RetAs DftVal", DryArg(MthLin))
End Function
Function ArgSfx$(Arg)
Dim L$: L = Arg
ShfPfxSpc L, "Optional"
ShfPfxSpc L, "ByVal"
ShfPfxSpc L, "ParamArray"
If ShfNm(L) = "" Then Thw CSub, "Arg is invalid", "Arg", Arg
ArgSfx = ShfDclSfx(L)
End Function
Function ArgSfxy(Argy$()) As String()
Dim Arg: For Each Arg In Itr(Argy)
    PushI ArgSfxy, ArgSfx(Arg)
Next
End Function

Function DrArg(Arg, Mthn$) As Variant()
Dim L$: L = Arg
Dim IsOpt   As Boolean:   IsOpt = ShfPfxSpc(L, "Optional")
Dim IsByVal As Boolean: IsByVal = ShfPfxSpc(L, "ByVal")
Dim IsPmAy  As Boolean:  IsPmAy = ShfPfxSpc(L, "ParamArray")
Dim Nm$:                     Nm = ShfNm(L)
Dim TyChr$:               TyChr = ShfChr(L, "!@#$%^&")
Dim IsAy    As Boolean:    IsAy = ShfBkt(L)
    If TyChr = "" Then
        If ShfAs(L) Then
            Dim AsTy$: AsTy = " As " & ShfDotNm(L)
            IsAy = ShfBkt(L)
        End If
    End If
    If L <> "" Then
        If Not ShfPfx(L, " = ") Then Stop
        Dim DftVal$: DftVal = L
    End If
DrArg = Array(Mthn, Nm, IsOpt, IsPmAy, IsAy, TyChr, AsTy, DftVal)
End Function

