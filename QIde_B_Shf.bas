Attribute VB_Name = "QIde_B_Shf"
Option Explicit
Option Compare Text
':Arg:    It is :s: splitting :MthPm:
':Sset:   It is :Aset: with all ele is :s:
':DoArg:   It is :Drs:Nm IsByVal IsOpt IsPmAy IsAy TyChr:C AsTy DftVal:
':ArgTy:  It is :s: coming :Arg: which dfn the type of an :Arg:
'         Eg, Dim A() as integer: :Arg:   => "A() As Integer"
'                                 :ArgTy: => "%()"
Public Const MthTyChrLis$ = "!@#$%^&"
Type AA
AA() As String
End Type
Function ShfMthTy$(OLin$)
Dim O$: O = TakMthTy(OLin$)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfTermAftAs$(OLin$)
If Not ShfTerm(OLin, "As") Then Exit Function
ShfTermAftAs = ShfT1(OLin)
End Function

Function ShfShtMdy$(OLin$)
ShfShtMdy = ShtMdy(ShfMdy(OLin))
End Function

Function ShfShtMthTy$(OLin$)
ShfShtMthTy = ShtMthTy(ShfMthTy(OLin))
End Function

Function ShfShtMthKd$(OLin$)
ShfShtMthKd = ShtMthKdzShtMthTy(ShtMthTy(ShfMthTy(OLin)))
End Function

Function ShfMdy$(OLin$)
Dim O$
O = MthMdy(OLin):
ShfMdy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfKd$(OLin$)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function

Function ShfMthSfx$(OLin$)
ShfMthSfx = ShfChr(OLin, "#!@#$%^&")
End Function
Private Sub Z_ShfBef()
Dim L$, Sep$, EptL$
GoSub T0
Exit Sub
T0:
    L = "aaa.bbb"
    Sep = "."
    Ept = "aaa"
    EptL = ".bbb"
    GoTo Tst
Tst:
    Act = ShfBef(L, Sep)
    C
    If L <> EptL Then Stop
    Return
End Sub

Function ShfBef$(OLin$, Sep)
Dim P%: P = InStr(OLin, Sep)
If P = 0 Then Exit Function
ShfBef = Left(OLin, P - 1)
OLin = Mid(OLin, P + Len(Sep) - 1)
End Function

Function ShfBefOrAll$(OLin$, Sep$, Optional NoTrim As Boolean)
Dim P%: P = InStr(OLin, Sep)
If P = 0 Then
    If NoTrim Then
        ShfBefOrAll = OLin
    Else
        ShfBefOrAll = Trim(OLin)
    End If
    OLin = ""
    Exit Function
End If
ShfBefOrAll = Bef(OLin, Sep, NoTrim)
OLin = Aft(OLin, Sep, NoTrim)
End Function
Function ShfDotNm$(OLin$)
OLin = LTrim(OLin)
Dim O$: O = TakDotNm(OLin): If O = "" Then Exit Function
ShfDotNm = O
OLin = RmvPfx(OLin, O)
End Function
Function ShfNm$(OLin$)
Dim O$: O = Nm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function

Function ShfRmk$(OLin$)
Dim L$
L = LTrim(OLin)
If FstChr(L) = "'" Then
    ShfRmk = Mid(L, 2)
    OLin = ""
End If
End Function

Function TakMthKd$(Lin)
TakMthKd = PfxzAy(Lin, MthKdAy)
End Function

Function TakMthTy$(Lin)
TakMthTy = PfxzAy(Lin, MthTyAy)
End Function

Function RmvMdy$(Lin)
RmvMdy = LTrim(RmvPfxSySpc(Lin, MthMdyAy))
End Function

Function RmvMthTy$(Lin)
RmvMthTy = RmvPfxSySpc(Lin, MthTyAy)
End Function

Function ShfAs(OLin$) As Boolean
ShfAs = ShfTermX(OLin, "As")
End Function

Function IsTyChr(A) As Boolean
If Len(A) <> 1 Then Exit Function
IsTyChr = HasSubStr(MthTyChrLis, A)
End Function
Function TyChrzTyNm$(TyNm)
Select Case TyNm
Case "String":   TyChrzTyNm = "$"
Case "Integer":  TyChrzTyNm = "%"
Case "Long":     TyChrzTyNm = "&"
Case "Double":   TyChrzTyNm = "#"
Case "Single":   TyChrzTyNm = "!"
Case "Currency": TyChrzTyNm = "@"
End Select
End Function

Function TyNmzTyChr$(TyChr$)
Dim O$
Select Case TyChr
Case "": O = "Variant"
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Thw CSub, "Invalid TyChr", "TyChr VdtTyChrLis", TyChr, MthTyChrLis
End Select
TyNmzTyChr = O
End Function

Function RmvTyChr$(S)

RmvTyChr = RmvLasChrzzLis(S, MthTyChrLis)
End Function

Function ShfDclSfx$(OLin$)
Dim O$: O = ShfTyChr(OLin)
If O <> "" Then
    ShfDclSfx = O & IIf(ShfBkt(OLin), "()", "")
    Exit Function
End If
Dim Bkt$:
    If ShfBkt(OLin) Then
        Bkt = "()"
    End If
If ShfAs(OLin) Then
    Dim DNm$: DNm = ShfDotNm(OLin):
    ShfDclSfx = Bkt & " As " & DNm
    If DNm = "" Then Stop
Else
    ShfDclSfx = Bkt
End If
End Function
Function ShfTyChr$(OLin$)
ShfTyChr = ShfChr(OLin, MthTyChrLis)
End Function

Function TyChr$(Lin)
If IsLinMth(Lin) Then TyChr = TakTyChr(RmvMthn3(Lin))
End Function

Function TakTyChr$(S)
TakTyChr = TakChr(S, MthTyChrLis)
End Function

Function MthTyChr$(Lin)
MthTyChr = TakTyChr(RmvNm(RmvMthTy(RmvMdy(Lin))))
End Function

Function ShfCnstLin$(Src$())
If Si(Src) = 0 Then Exit Function
Dim L$: L = RmvMdy(Src(0))
'IsLinCnst
Stop
'ShfCnstLin = ShfT1(OLin) = "Const"
End Function

Function ShfDim(OLin$) As Boolean
ShfDim = ShfT1(OLin) = "Dim"
End Function

Function ArgMdyAy() As String()
Static X As Boolean, Y
If Not X Then
    Dim O$()
    PushI O, "ByVal"
    PushI O, "ByRef"
    PushI O, "Optional ByVal"
    PushI O, "Optional ByRef"
    PushI O, "Optional"
    PushI O, "ParamArray"
    Y = O
End If
ArgMdyAy = O
End Function

Function ShfArgMdy$(OArg$)
Dim A$: A = ShfPfxAyS(OArg, ArgMdyAy)
Select Case A
Case "ByVal":: ShfArgMdy = "*"
Case "ByRef":
Case "Optional ByVal": ShfArgMdy = "?*"
Case "Optional ByRef", "Optional": ShfArgMdy = "?"
Case "ParamArray": ShfArgMdy = ".."
End Select
End Function
Private Function DclSfxzAftNm$(AftNm$)
If AftNm = "" Then Exit Function
Dim L$: L = AftNm
Select Case True
Case L = "As Boolean":: DclSfxzAftNm = "-"
Case L = "As Boolean()": DclSfxzAftNm = "-()"
Case Else
    ShfPfx L, "As "
    DclSfxzAftNm = L
End Select
End Function
Sub Z_ShtArg()
Dim Arg$: Arg = "Optional UseVc As Boolean"
'Debug.Print ShtArg(Arg)
Debug.Print ShtArg("Optional NoBlnkStr As Boolean")
End Sub
Sub Z_ShtArgAy()
Dim A As S12s
Dim Arg: For Each Arg In AwDist(ArgAyP)
    PushS12 A, S12(Arg, ShtArg(Arg))
Next
BrwS12s A
End Sub
Function ShtArgAy(ArgAy$()) As String()
Dim Arg: For Each Arg In Itr(ArgAy)
    PushI ShtArgAy, ShtArg(Arg)
Next
End Function

Function ShtArg$(Arg)
'Ret : :ShtArg @@
':ShtArg: :S ! Fmt: <Mdy><Nm><Sep><Sfx><Dft>
'            ! *Dft* :S  Fmt: =<D>-or-Blnk
'            ! *Sep* :C  Blnk-or-:
'            ! *Mdy* :S  One-of-ArgMdyAy
Dim L$: L = Arg
Dim Mdy:  Mdy = ShfArgMdy(L)
Dim Nm$: Nm = ShfNm(L):     If Nm = "" Then Exit Function
Dim AftNm$, D$: AsgBrk1 L, "=", AftNm, D
Dim Sfx$: Sfx = DclSfxzAftNm(AftNm)
Dim Sep$: Sep = DclSfxSep(Sfx)
Dim Dft$: If D <> "" Then Dft = "=" & D
Dim Ty$: Ty = DclSfxzAftNm(AftNm):   If IsEmpty(Ty) Then Exit Function
ShtArg = Mdy & Nm & Sep & Sfx & Dft
End Function
Function DclSfxSep$(DclSfx$)
If DclSfx = "" Then Exit Function
Dim C$: C = FstChr(DclSfx)
Dim A%: A = Asc(C)
If IsAscUCas(A) Then DclSfxSep = ":"
End Function
Function DclSfx$(DclItm$)
DclSfx = DclSfxzAftNm(RmvNm(DclItm))
End Function

Function DclSfxzArg$(Arg$)
DclSfxzArg = DclSfx(DclItm(Arg))
End Function

Function DclItm$(Arg$)
DclItm = BefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " =")
End Function
Function PmNm$(Arg$)
PmNm = RmvLasChrzzLis(BefOrAll(DclItm(Arg), "="), MthTyChrLis)
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

Function RetAsAyP() As String()
'RetAsAyP = DistColoS(DoMth, "RetAs")
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
Dim L, A As S12s: For Each L In MthLinAyP
    PushS12 A, S12(RetAszL(L), L)
Next
BrwS12s A
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
Case Left(DclSfx, 4) = " As ":   RetAszDclSfx = DclSfx
Case Left(DclSfx, 6) = "() As ": RetAszDclSfx = Mid(DclSfx, 3) & "()"
Case DclSfx = "()":              RetAszDclSfx = " As Variant()"
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

Private Sub Z_DoArg()
Dim Mth$(): Mth = StrCol(DoPubMth, "MthLin")
Dim L, Dy(): For Each L In Itr(Mth)
    PushIAy Dy, DyoArg(L)
Next
BrwDrs Drs(FoArg, Dy)
End Sub

Private Function DyoArgzMthL(MthLinAy$()) As Variant()
Dim L: For Each L In Itr(MthLinAy)
    PushIAy DyoArgzMthL, DyoArg(L)
Next
End Function

Private Function DyoArg(MthLin) As Variant()
Dim Pm$: Pm = BetBkt(MthLin)
Dim A$(): A = SplitCommaSpc(Pm)
Dim Mthn$: Mthn = MthnzLin(MthLin)
Dim Arg$, Dy(), ArgNo%: For ArgNo = 1 To Si(A)
    Arg = A(ArgNo - 1)
    PushI DyoArg, DroArg(Arg, ArgNo, Mthn)
Next
End Function
Private Function FoArg() As String()
FoArg = SyzSS("Mthn No Nm IsOpt IsByVal IsPmAy IsAy TyChr AsTy DftVal")
End Function

Function DoArgzMthL(MthLinAy$()) As Drs
DoArgzMthL = Drs(FoArg, DyoArgzMthL(MthLinAy))
End Function

Function DArgzM(M As CodeModule) As Drs
Dim L$(): L = StrCol(DoMthzM(M), "MthLin")
     DArgzM = Drs(FoArg, DyoArgzMthL(L))
End Function

Function DArgzP(P As VBProject) As Drs
Dim M$(): M = StrCol(DoMthzP(P), "MthLin")
     DArgzP = Drs(FoArg, DyoArgzMthL(M))
End Function

Function DArgP() As Drs
DArgP = DArgzP(CPj)
End Function

Function DoArg(MthLin) As Drs
DoArg = DrszFF("Mthn Nm ArgNo IsOpt IsByVal IsPmAy IsAy TyChr RetAs DftVal", DyoArg(MthLin))
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

Function DroArg(Arg$, ArgNo%, Mthn$) As Variant()
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
DroArg = Array(Mthn, ArgNo, Nm, IsOpt, IsByVal, IsPmAy, IsAy, TyChr, AsTy, DftVal)
End Function

Function DoArgV() As Drs
DoArgV = DoArgzMthL(MthLinAyV)
End Function

Private Sub Z_MthRetTy()
'Dim MthLin
'Dim A$:
'MthLin = "Function MthPm(MthPm$) As MthPm"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = False
'Ass A.TyChr = ""
'
'MthLin = "Function MthPm(MthPm$) As MthPm()"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = True
'Ass A.TyChr = ""
'
'MthLin = "Function MthPm$(MthPm$)"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = "$"
'
'MthLin = "Function MthPm(MthPm$)"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = ""
End Sub

Function ShfLHS$(OLin$)
Dim L$:                   L = OLin
Dim IsSet As Boolean: IsSet = ShfTermX(L, "Set")
Dim S$:                       If IsSet Then S = "Set "
Dim LHS$:               LHS = ShfDotNm(L)
If FstChr(L) = "(" Then
    LHS = LHS & QteBkt(BetBkt(L))
    L = AftBkt(L)
End If
If Not ShfPfx(L, " = ") Then Exit Function
ShfLHS = S & LHS & " = "
OLin = L
End Function

Function ShfLRHS(OLin$) As Variant()
Dim L$:     L = OLin
Dim LHS$: LHS = ShfLHS(L)
With Brk1(L, "'")
    Dim RHS$:  RHS = .S1
              OLin = "   ' " & .S2
End With
ShfLRHS = Array(LHS, RHS)
End Function

