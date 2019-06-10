VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ATmp__AlignMthDimzML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Friend Function CmlAct(CmlEpt As Drs, Md As CodeModule) As Drs
Dim CV$(): CV = StrColzDrs(CmlEpt, "V")
Dim Act$(): Act = MthLinAyzM(Md)
Insp CSub, "CmlAct", "CV Act", CV, Act
Stop
End Function

Friend Function CmlActV(CmlAct1 As Drs, CmPfx$) As Drs
'Ret      : L ActL V ' RmvPfx CmlPfx from V @@
'Fm CmlAct : L ActL V
'Ret       : RmvPfx CmPfx from V
Dim Dry(), Dr, IV%, V$
AsgIx CmlAct1, "V", IV
For Each Dr In Itr(CmlAct1.Dry)
    Dr(IV) = RmvPfx(Dr(IV), CmPfx)
    PushI Dry, Dr
Next
CmlActV = Drs(CmlAct1.Fny, Dry)
'Insp "QIde_B_MthOp.CmlActV", "Inspect", "Oup(CmlActV) D1 CmPfx", FmtDrs(CmlActV), D1, CmPfx: Stop
End Function
Friend Function MlVSfx(Ml$) As Drs
'Ret      : Ret V Sfx ! the MthLin's pm V Sfx @@
'Fm Ml # MthLin
'Ret V Sfx ! the MthLin's pm V Sfx
Dim Pm$: Pm = BetBkt(Ml)
Dim P, V$, Sfx$, Dry(), L$
For Each P In Itr(TrimAy(SplitComma(Pm)))
    L = P
    V = ShfNm(L)
    Sfx = L
    PushI Dry, Array(V, Sfx)
Next
MlVSfx = DrszFF("V Sfx", Dry)
'Insp "QIde_B_MthOp.MlVSfx", "Inspect", "Oup(MlVSfx) Ml", FmtDrs(MlVSfx), Ml: Stop
End Function

Friend Function CmlMthRet(CmlDclPm As Drs) As Drs
'Fm  CmlDclPm : V Sfx Expr Mthn CallPm DclPm
'Ret          : V Sfx Expr Mthn CallPm DclPm TyChr RetAs @@
Dim Dr, Sfx$, TyChr$, RetAs$, I%, Dry()
I = IxzAy(CmlDclPm.Fny, "Sfx")
For Each Dr In Itr(CmlDclPm.Dry)
    Sfx = Dr(I)
    TyChr = TyChrzDclSfx(Sfx)
    RetAs = RetAszDclSfx(Sfx)
    PushI Dr, TyChr
    PushI Dr, RetAs
    PushI Dry, Dr
Next
CmlMthRet = AddColzFFDry(CmlDclPm, "TyChr RetAs", Dry)
'BrwDrs CmlMthRet: Stop
'Insp "QIde_B_MthOp.CmlMthRet", "Inspect", "Oup(CmlMthRet) CmlDclPm", FmtDrs(CmlMthRet), FmtDrs(CmlDclPm): Stop
End Function
Friend Function CmlDclPm(CmlCallg As Drs, CmlVSfx As Drs) As Drs
'Fm  CmlCallg : V Sfx Expr                   ! It is subset of McR123 where Expr is a calling1 or calling2.
'               Mthn CallgPm                 ! calling1 is Expr = CmPfx & V              No Pm yet
'                                            ! calling2 is HasSfx(Expr, CmPfx & V & "("  with Pm
'Ret          : V Sfx Expr Mthn CallPm DclPm @@
Const CmPfx$ = "XH_"
Dim Dry(), Dr, DclPm$, CallgPm$, Expr$
For Each Dr In Itr(CmlCallg.Dry)
    Expr = Dr(2)
    CallgPm = BetBkt(Expr)
    DclPm = WDclPm(CallgPm, CmlVSfx)
    PushI Dr, DclPm
    PushI Dry, Dr
Next
CmlDclPm = AddColzFFDry(CmlCallg, "DclPm", Dry)
'BrwDrs CmlDclPm: Stop
'Insp "QIde_B_MthOp.CmlDclPm", "Inspect", "Oup(CmlDclPm) CmlCallg CmlVSfx", FmtDrs(CmlDclPm), FmtDrs(CmlCallg), FmtDrs(CmlVSfx): Stop
End Function
Private Function WDclPm$(CallgPm$, CmlVSfx As Drs)
Dim O$(), Sfx$, P
For Each P In Itr(TrimAy(SplitComma(CallgPm)))
    Sfx = ValzColEq(CmlVSfx, "Sfx", "V", P)
    PushI O, P & Sfx
Next
WDclPm = JnCommaSpc(O)
'Insp CSub, "Finding DclPm(CallgPm, CmlFmMc)", "CallgPm CmlFmMc WDclPm", CallgPm, FmtDrs(CmlFmMc), WDclPm: Stop
End Function
Friend Function CrFill(CrVPR3 As Drs) As Drs
'Fm  CrVPR3 : V P R1 R2 R3
'Ret        : V P R1 R2 R3 FP FR1 FR2 @@
'Fm CrVPR3 : V P R1 R2 R3
'Ret       : V P R1 R2 R3 FP FR1 FR2 ' For each dist v return the Filler of P R1 R2 as FP FR1 FR2
If NoReczDrs(CrVPR3) Then Exit Function
Dim V: For Each V In DistCol(CrVPR3, "V")
    Dim A As Drs: A = DrswColEq(CrVPR3, "V", V)   ' V P R1 R2 R3 ! Sam V
    Dim B As Drs: B = WRmk_Fill(A)                   ' V P R1 R2 R3 FP FR1 FR2
    Dim O As Drs: O = AddDrs(O, B)
Next
CrFill = O
'Insp "QIde_B_MthOp.CrFill", "Inspect", "Oup(CrFill) CrVPR3", FmtDrs(CrFill), FmtDrs(CrVPR3): Stop
End Function
Friend Function CrFst(CrFill As Drs) As Drs
'Fm  CrFill : V P R1 R2 R3 FP FR1 FR2
'Ret        : V P R1 R2 R3 FP FR1 FR2 Fst @@
'Fm CrFill : V P R1 R2 R3 FR1 FR2
'Ret       : V P R1 R2 R3 FR1 FR2 Fst
If NoReczDrs(CrFill) Then Exit Function
Dim Dr, LasV$, LasP$, Dry(): For Each Dr In CrFill.Dry
    Dim V$: V = Dr(0)
    Dim P$: P = Dr(1)
    If V <> LasV Or P <> LasP Then
        PushI Dr, True
        LasV = V
        LasP = P
    Else
        PushI Dr, False
    End If
    PushI Dry, Dr
Next
CrFst = AddColzFFDry(CrFill, "Fst", Dry)
'Insp "QIde_B_MthOp.CrFst", "Inspect", "Oup(CrFst) CrFill", FmtDrs(CrFst), FmtDrs(CrFill): Stop
End Function
Friend Function CrEpt(CmPfx$, CrRmk As Drs) As S1S2s
'Fm  CrRmk : V P R1 R3 R3 FP FR1 FR2 Fst Rmk
'Ret       : Mthn RmkLines                   ! RmkLines is find by each V in CrVPR3 & Mthn = V & CmPfx @@
'Fm CrRmk : V P R1 R3 R3 FP FR1 FR2 Fst Rmk
If NoReczDrs(CrRmk) Then Exit Function
Dim A As Drs: A = SelDrs(CrRmk, "V Rmk")
Dim Dr, Ly$(): For Each Dr In A.Dry
    PushI Ly, CmPfx & Dr(0) & " " & Dr(1)
Next
Dim D As Dictionary: Set D = Dic(Ly, vbCrLf)
Dim O As Dictionary: Set O = AddSfxzDic(D, " @@")
CrEpt = S1S2szDic(O)
'Insp "QIde_B_MthOp.CrEpt", "Inspect", "Oup(CrEpt) CmPfx CrRmk", FmtS1S2s(CrEpt), CmPfx, FmtDrs(CrRmk): Stop
End Function
Friend Function CrRmk(CrFst As Drs) As Drs
'Fm  CrFst : V P R1 R2 R3 FP FR1 FR2 Fst
'Ret       : V P R1 R3 R3 FP FR1 FR2 Fst Rmk @@
'Fm CrFst : V P R1 R2 R3 FP FR1 FR2 Fst
'Ret      : V P R1 R2 R3 FP FR1 FR2 Fst Rmk
If NoReczDrs(CrFst) Then Exit Function
Dim Dr, Dry(): For Each Dr In CrFst.Dry
    Dim Lin$: Lin = WRmk_Lin(Dr)
    PushI Dr, Lin
    PushI Dry, Dr
Next
CrRmk = AddColzFFDry(CrFst, "Rmk", Dry)
'Insp "QIde_B_MthOp.CrRmk", "Inspect", "Oup(CrRmk) CrFst", FmtDrs(CrRmk), FmtDrs(CrFst): Stop
End Function

Private Function WRmk_Fill(A As Drs) As Drs
'Fm A : V P R1 R2 R3 ! Sam V
'Ret  : V FP FR1 FR2 FR3
Dim FP%: FP = WdtzAy(StrCol(A, "P"))
Dim FR1%: FR1 = WdtzAy(StrCol(A, "R1"))
Dim FR2%: FR2 = WdtzAy(StrCol(A, "R2"))
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    PushIAy Dr, Array(FP, FR1, FR2)
    PushI Dry, Dr
Next
WRmk_Fill = AddColzFFDry(A, "FP FR1 FR2", Dry)
End Function
Friend Function CrVPR3(CrPm As Drs, CrVR3 As Drs, MlVSfx As Drs) As Drs
'Fm  CrPm   : V CallgPm
'Fm  CrVR3  : V R1 R2 R3
'Fm  MlVSfx : Ret V Sfx    ! the MthLin's pm V Sfx
'Ret        : V P R1 R2 R3 @@
Dim VP, R, P, ODry(), V$, CallgPm$, PmR4(), ODr(), A As Drs
For Each VP In Itr(CrPm.Dry)
    V = VP(0)
    CallgPm = VP(1)
    For Each P In Itr(TrimAy(SplitComma(CallgPm)))
        If HasColEq(MlVSfx, "V", P) Then
            ODr = Array(V, P, "InpPm", "", "")
        Else
            A = ColEqExlEqCol(CrVR3, "V", P)
            For Each R In Itr(A.Dry)
                ODr = Array(V, P, R(0), R(1), R(2))
                PushI ODry, ODr
            Next
        End If
    Next
    '-- Ret
    A = ColEqExlEqCol(CrVR3, "V", V)
    For Each R In Itr(A.Dry)
        ODr = Array(V, "*Ret", R(0), R(1), R(2))
        PushI ODry, ODr
    Next
Next
CrVPR3 = DrszFF("V P R1 R2 R3", ODry)
'BrwDrs3 CrPm, CrVR3, CrVPR3, NN:="CrPm CrVR3 Oup-CrVPR3": Stop
'Insp "QIde_B_MthOp.CrVPR3", "Inspect", "Oup(CrVPR3) CrPm CrVR3 MlVSfx", FmtDrs(CrVPR3), FmtDrs(CrPm), FmtDrs(CrVR3), FmtDrs(MlVSfx): Stop
End Function
Friend Function CrVR3(McR123 As Drs) As Drs
'Fm  McR123 : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
'Ret        : V R1 R2 R3 @@
Dim A As Drs: A = SelDrs(McR123, "V IsRmk R1 R2 R3")
Dim V$, IsRmk As Boolean, R1$, R2$, R3$, Dry(), Dr, X$
Dim LasV$: LasV = A.Dry(0)(0)
For Each Dr In Itr(A.Dry)
    AsgAp Dr, V, IsRmk, R1, R2, R3
    If IsRmk Then
        X = LasV
    Else
        X = V
        LasV = V
    End If
    Select Case True
    Case R1 <> "", R2 <> "", R3 <> ""
        PushI Dry, Array(X, R1, R2, R3)
    End Select
Next
CrVR3 = DrszFF("V R1 R2 R3", Dry)
'BrwDrs2 McR123, CrVR3: Stop
'Insp "QIde_B_MthOp.CrVR3", "Inspect", "Oup(CrVR3) McR123", FmtDrs(CrVR3), FmtDrs(McR123): Stop
End Function


Friend Function McVSfx(McTRmk As Drs) As Drs
'Fm  McTRmk : L Gpno MthLin IsRmk ! For each gp, the front rmk lines are TopRmk,
'                                 ! rmv them
'Ret        : L Gpno IsRmk
'             V Sfx Rst @@
Const CmPfx$ = "XB_"
Dim Dr, Dry()
For Each Dr In Itr(McTRmk.Dry)
    Dim Av(): Av = WDrVSfxRst(Dr(2))
    PushIAy Dr, Av
    PushI Dry, Dr
Next
McVSfx = AddColzFFDry(McTRmk, "V Sfx Rst", Dry)
'BrwDrs McVSfx: Stop
'Insp "QIde_B_MthOp.McVSfx", "Inspect", "Oup(McVSfx) McTRmk", FmtDrs(McVSfx), FmtDrs(McTRmk): Stop
End Function

Friend Function McDcl(McVSfx As Drs) As Drs
'Fm  McVSfx : L Gpno IsRmk
'             V Sfx Rst
'Ret        : L Gpno MthLin IsRmk
'             V Sfx Dcl Rst       ! Add Dcl from V & Sfx @@
Const CmPfx$ = "XA_"
Dim MthLin$, IxMthLin%, Dr, Dry(), IGpno
AsgIx McVSfx, "MthLin", IxMthLin
For Each IGpno In AywDist(IntAyzDrsC(McVSfx, "Gpno"))
    Dim A As Drs: A = ColEq(McVSfx, "Gpno", IGpno) ' L Gpno MthLin IsRmk V Sfx Rst ! Sam Gpno
    Dim B As Drs: B = WDcl(A) ' L Gpno MthLin IsRmk V Sfx Dcl Rst ! Adding Dcl using V Sfx
    Dim O As Drs: O = AddDrs(O, B)
Next
McDcl = O
'Insp "QIde_B_MthOp.McDcl", "Inspect", "Oup(McDcl) McVSfx", FmtDrs(McDcl), FmtDrs(McVSfx): Stop
End Function

Private Function WDcl(A As Drs) As Drs
'Fm A :      L Gpno MthLin IsRmk V Sfx Rst
'Ret McDclI: L Gpno MthLin IsRmk V Sfx Dcl Rst @@
Dim IxV%, IxSfx%: AsgIx A, "V Sfx", IxV, IxSfx
Dim B As Drs: B = DrswColPfx(A, "Sfx", " As ")
Dim C$(): C = StrColzDrs(B, "V")
Dim WAsV%: WAsV = WdtzAy(C)
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dim V$: V = Dr(IxV)
    Dim Sfx$: Sfx = Dr(IxSfx)
    Dim Dcl$: Dcl = WDclzV(V, WAsV, Sfx)
    Dim Rst$: Rst = Pop(Dr)
    PushIAy Dr, Array(Dcl, Rst)
    PushI Dry, Dr
Next
Dim Fny$(): Fny = AddAy(AyeLasEle(A.Fny), Array("Dcl", "Rst"))
WDcl = Drs(Fny, Dry)
End Function
Private Function WAlign(A As Drs) As Drs
Dim Dr
Dim L&, Gpno%, MthLin$, Dcl$, LHS$, Expr$, R1$, R2$, R3$, F0%, FDcl%, FLHS%, FExpr%, FR1%, FR2%
Dim T0$, TDcl$, TL$, TR$, TR1$, TR2$, TR3$, Align$, Dry()
For Each Dr In Itr(A.Dry)
    AsgAp Dr, L, Gpno, MthLin, Dcl, LHS, Expr, R1, R2, R3, F0, FDcl, FLHS, FExpr, FR1, FR2
    T0 = Space(F0)
    TDcl = Dcl & Space(FDcl)
    TL = Space(FLHS) & LHS
    TR = Expr & Space(FExpr)
    TR1 = R1 & Space(FR1)
    TR2 = R2 & Space(FR2)
    TR3 = R3
    Align = RTrim(T0 & TDcl & TL & TR & TR1 & TR2 & TR3)
    PushI Dry, Array(L, Gpno, MthLin, Align)
Next
WAlign = DrszFF("L Gpno MthLin Align", Dry)
'BrwDrs2 A, WAlign: Stop
End Function
Private Function WDclzV$(V$, WAsV%, Sfx$)
If V = "" Then Exit Function
Dim O$, S$
If HasPfx(Sfx, " As ") Then S = Space(WAsV - Len(V))
WDclzV = "Dim " & V & S & Sfx & ": "
End Function
Friend Function McLR(McDcl As Drs) As Drs
'Fm  McDcl : L Gpno MthLin IsRmk
'            V Sfx Dcl Rst          ! Add Dcl from V & Sfx
'Ret       : L Gpno MthLin IsRmk
'            V Sfx Dcl LHS Expr Rst ! Add LHS Expr from Rst @@
Dim Dr: For Each Dr In Itr(McDcl.Dry)
    Dim L$: L = Pop(Dr)
    Dim Rst$:        Rst = L
    Dim LHS$, Expr$:       AsgAp ShfLRHS(Rst), LHS, Expr
    If Rst <> "" Then
        If Not IsVbRmk(Rst) Then Stop
    End If
                      Dr = AddAy(Dr, Array(LHS, Expr, Rst))
    Dim Dry():             PushI Dry(), Dr
Next
Dim Fny$(): Fny = AyeLasEle(McDcl.Fny)
Fny = AddAy(Fny, SyzSS("LHS Expr Rst"))
McLR = Drs(Fny, Dry)
'Insp "QIde_B_MthOp.McLR", "Inspect", "Oup(McLR) McDcl", FmtDrs(McLR), FmtDrs(McDcl): Stop
End Function

Friend Function McR123(McLR As Drs) As Drs
'Fm  McLR : L Gpno MthLin IsRmk
'           V Sfx Dcl LHS Expr Rst      ! Add LHS Expr from Rst
'Ret      : L Gpno MthLin IsRmk
'           V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst @@
Dim Dr: For Each Dr In Itr(McLR.Dry)
    Dim Rst$:          Rst = LasEle(Dr)
    If FstChr(Rst) <> "'" Then Stop
    Dim R1$, R2$, R3$:       AsgBrkBet Rst, "#", "!", R1, R2, R3
                        R1 = Trim(RmvFstChr(R1))
                             If R2 <> "" Then R2 = " # " & R2
                             If R3 <> "" Then R3 = " ! " & R3
                             If R1 <> "" Or R2 <> "" Or R3 <> "" Then R1 = " ' " & R1
    Dim Dry():               PushI Dry, AddAy(AyeLasEle(Dr), Array(R1, R2, R3))
Next
Dim Fny$(): Fny = AddAy(AyeLasEle(McLR.Fny), SyzSS("R1 R2 R3"))
McR123 = Drs(Fny, Dry)
'Insp "QIde_B_MthOp.McR123", "Inspect", "Oup(McR123) McLR", FmtDrs(McR123), FmtDrs(McLR): Stop
End Function
Private Function WDrVSfxRst(MthLin) As Variant()
Dim V$, Sfx$, Rst$
    Rst = Trim(MthLin)
    Select Case True
    Case ShfTermX(Rst, "Dim")
        V = ShfNm(Rst)
        Sfx = ShfDclSfx(Rst)
        If HasPfx(Sfx, ",") Then Stop
        Rst = Trim(RmvPfxAll(Rst, ":"))
    End Select
WDrVSfxRst = Array(V, Sfx, Rst)
End Function

Private Function WF0(A As Drs) As Drs
Dim IxMthLin%: IxMthLin = IxzAy(A.Fny, "MthLin")
Dim IxV%: IxV = IxzAy(A.Fny, "V")
Dim Dr:              Dr = A.Dry(0)
Dim V$:               V = Dr(IxV)
Dim MthLin$:     MthLin = Dr(IxMthLin)
Dim T$:               T = LTrim(MthLin)
Dim F0%:             F0 = IIf(V = "", 0, Len(MthLin) - Len(T))
               WF0 = AddColzCV(A, "F0", F0)
End Function

Friend Function McFill(McR123 As Drs) As Drs
'Fm  McR123 : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
'Ret        : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS Expr
'             F0 FSfx FExpr FR1 FR2       ! Adding F* @@
Const CmPfx$ = "XD_"
Dim Gpno%(): Gpno = AywDist(IntAyzDrsC(McR123, "Gpno"))
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim A As Drs: A = ColEq(McR123, "Gpno", IGpno)
    Dim B As Drs: B = WF0(A)
    Dim C As Drs: C = AddColzFiller(B, "Dcl LHS Expr R1 R2")
    Dim O As Drs: O = AddDrs(O, C)
Next
McFill = O
'Insp "QIde_B_MthOp.McFill", "Inspect", "Oup(McFill) McR123", FmtDrs(McFill), FmtDrs(McR123): Stop
End Function

Friend Function De(Mc As Drs) As Drs
'Fm  Mc   : L MthLin # Mc.
'Ret      : L MthLin # Dbl-Eq @@
Dim Dr, Dry()
For Each Dr In Itr(Mc.Dry)
    If Left(LTrim(Dr(1)), 3) = "'==" Then PushI Dry, Dr
Next
De.Fny = Mc.Fny
De.Dry = Dry

'Insp "QIde_B_MthOp.De", "Inspect", "Oup(De) Mc", FmtDrs(De), FmtDrs(Mc): Stop
End Function

Friend Function DeLNewO(McDblEqRmk As Drs) As Drs
'Fm  De   : L MthLin    # Dbl-Eq
'Ret      : L NewL OldL @@
'BrwDrs McDblEqRmk: Stop
Dim Dr, Dry(), OldL$, NewL$, L&, A$
For Each Dr In Itr(McDblEqRmk.Dry)
    L = Dr(0)
    OldL = Dr(1)
    A = Left(Trim(OldL), 120)
    NewL = A & Dup("=", 120 - Len(A))
    If OldL <> NewL Then
        Push Dry, Array(L, NewL, OldL)
    End If
Next
DeLNewO = LNewO(Dry)

'Insp "QIde_B_MthOp.DeLNewO", "Inspect", "Oup(DeLNewO) De", FmtDrs(DeLNewO), FmtDrs(De): Stop
End Function
Friend Function McGp(McCln As Drs) As Drs
'Fm  McCln : L MthLin    # Mc-Cln. ! must Dim | Asg | Rmk(but not 'If 'Insp, '==). Cln to Align
'Ret       : Gpno MthLin           ! with L in seq will be one gp @@
Dim Dr, LasL&, Gpno%, L&, Dry(), J%
For Each Dr In McCln.Dry
    L = Dr(0)
    If LasL + 1 <> L Then
        Gpno = Gpno + 1
    End If
    LasL = L
    PushI Dry, Array(L, Gpno, Dr(1))
Next
McGp = DrszFF("L Gpno MthLin", Dry)
'Insp "QIde_B_MthOp.McGp", "Inspect", "Oup(McGp) McCln", FmtDrs(McGp), FmtDrs(McCln): Stop
End Function

Friend Function McTRmk(McRmk As Drs) As Drs
'Fm  McRmk : L Gpno MthLin IsRmk ! a column IsRmk is added
'Ret       : L Gpno MthLin IsRmk ! For each gp, the front rmk lines are TopRmk,
'                                ! rmv them @@
' L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! For each gp, the front rmk lines are TopRmk, rmv them
Dim IGpno%, MaxGpno, A As Drs, B As Drs, O As Drs
MaxGpno = MaxzAy(IntAyzDrsC(McRmk, "Gpno"))
For IGpno = 1 To MaxGpno
    A = ColEq(McRmk, "Gpno", IGpno)
    B = McTRmkI(A)
    O = AddDrs(O, B)
Next
McTRmk = O
'Insp "QIde_B_MthOp.McTRmk", "Inspect", "Oup(McTRmk) McRmk", FmtDrs(McTRmk), FmtDrs(McRmk): Stop
End Function
Friend Function McTRmkI(A As Drs) As Drs
' Fm  A :    L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! All Gpno are eq
' Ret : L Gpno MthLin IsRmk ! Rmk TopRmk
McTRmkI.Fny = A.Fny
Dim J%
    Dim Dr
    For Each Dr In Itr(A.Dry)
        If Not Dr(3) Then GoTo Fnd
        J = J + 1
    Next
    Exit Function
Fnd:
    For J = J To UB(A.Dry)
        PushI McTRmkI.Dry, A.Dry(J)
    Next
End Function


Friend Function McRmk(McGp As Drs) As Drs
'Fm  McGp : Gpno MthLin         ! with L in seq will be one gp
'Ret      : L Gpno MthLin IsRmk ! a column IsRmk is added @@
'Ret McRmk L Gpno MthLin IsRmk    #Mth-Cxt-isRmk
Dim Dr
For Each Dr In McGp.Dry
    PushI Dr, FstChr(LTrim(Dr(2))) = "'"
    Push McRmk.Dry, Dr
Next
McRmk.Fny = AddFF(McGp.Fny, "IsRmk")
'Insp "QIde_B_MthOp.McRmk", "Inspect", "Oup(McRmk) McGp", FmtDrs(McRmk), FmtDrs(McGp): Stop
End Function


Friend Function CmStr$(CmNew$(), McR123 As Drs, CmMdy$, CmPfx$)
'Fm  CmNew  :                             ! The new ChdMthNy to be created.
'Fm  McR123 : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
'Ret        :                             ! Mth-Str to be append to CmMd @@
Dim A As Drs: A = ColEqSel(McR123, "IsRmk", False, "V Sfx")
Dim CmNm, O$(), Sfx, V$, TyChr$, RetAs$
'BrwDrs McR123: Stop
For Each CmNm In Itr(CmNew)
    V = RmvPfx(CmNm, CmPfx)
    Sfx = ValzColEq(A, "Sfx", "V", V): If Not IsStr(Sfx) Then Stop
    TyChr = TyChrzDclSfx(Sfx)
    RetAs = RetAszDclSfx(Sfx)
    PushI O, ""
    PushI O, FmtQQ("? Function ??()?", CmMdy, CmNm, TyChr, RetAs)
    PushI O, "End Function"
Next
CmStr = JnCrLf(O)
'Insp "QIde_B_MthOp.CmStr", "Inspect", "Oup(CmStr) CmNew McR123 CmMdy CmPfx", CmStr, CmNew, FmtDrs(McR123), CmMdy, CmPfx: Stop
End Function

Friend Function CmlEpt(CmlMthRet As Drs, CmMdy$) As Drs
'Fm  CmlMthRet : V Sfx Expr Mthn CallPm DclPm TyChr RetAs
'Ret           : V Mthn EptL @@
Dim Dr, Dry(), Nm$, Ty$, Pm$, Ret$, V$, EptL$, INm%, ITy%, IPm%, IRet%, IV%
AsgIx CmlMthRet, "Mthn TyChr DclPm RetAs V", INm, ITy, IPm, IRet, IV
'BrwDrs CmlMthRet: Stop
For Each Dr In Itr(CmlMthRet.Dry)
    Nm = Dr(INm)
    Ty = Dr(ITy)
    Pm = Dr(IPm)
    Ret = Dr(IRet)
    V = Dr(IV)
    EptL = FmtQQ("? Function ??(?)?", CmMdy, Nm, Ty, Pm, Ret)
    PushI Dry, Array(V, Nm, EptL)
Next
CmlEpt = DrszFF("V Mthn EptL", Dry)
'BrwDrs CmlEpt: Stop
'Insp "QIde_B_MthOp.CmlEpt", "Inspect", "Oup(CmlEpt) CmlMthRet CmMdy", FmtDrs(CmlEpt), FmtDrs(CmlMthRet), CmMdy: Stop
End Function

Friend Function CmlCallg(CmlFmMc As Drs, CmlCallgPfx$) As Drs
'Ret      : V Sfx Expr   ! It is subset of McR123 where Expr is a calling1 or calling2.
'           Mthn CallgPm ! calling1 is Expr = CmPfx & V              No Pm yet
'                        ! calling2 is HasSfx(Expr, CmPfx & V & "("  with Pm @@
If CmlCallgPfx = "" Then Stop
Dim ODry()
    Dim Dr
    For Each Dr In Itr(CmlFmMc.Dry)
        Dim V$: V = Dr(0)
        Dim Sfx$: Sfx = Dr(1)
        Dim Expr$: Expr = Dr(2)
        Dim ExprPfx$: ExprPfx = CmlCallgPfx & V & "("
        If HasPfx(Expr, ExprPfx) Then
            Dim Mthn$
            If CmlCallgPfx = "F." Then
                Mthn = V
            Else
                Mthn = CmlCallgPfx & V
            End If
            Dim CallgPm$: CallgPm = BetBkt(Expr)
            PushI Dr, Mthn
            PushI Dr, CallgPm
            PushI ODry, Dr
        End If
    Next
CmlCallg = DrszFF("V Sfx Expr Mthn CallgPm", ODry)
'BrwDrs CmlCallg: Stop
'Insp "QIde_B_MthOp.CmlCallg", "Inspect", "Oup(CmlCallg) CmlFmMc CmlCallgPfx", FmtDrs(CmlCallg), FmtDrs(CmlFmMc), CmlCallgPfx: Stop
End Function

Friend Function McNew(Mc As Drs, McDim As Drs) As String()
'BrwDrs2 Mc, Mc, NN:="Mc McDim", Tit:="Use McDim to Upd Mc to become NewL": Stop
If JnSpc(McDim.Fny) <> "L OldL NewL" Then Stop
Dim A As Drs: A = SelDrs(McDim, "L NewL")
Dim B As Dictionary: Set B = DiczDrsCC(A)
Dim O$()
    Dim Dr, L&, MthLin$
    For Each Dr In Mc.Dry
        L = Dr(0)
        MthLin = Dr(1)
        If B.Exists(L) Then
            PushI O, B(L)
        Else
            PushI O, MthLin
        End If
    Next
McNew = O
End Function
Friend Function BsLNewO(Bs As Drs, BsVPfx As Dictionary, Mdn$, MlNm$) As Drs
'Fm  Bs   : L BsLin            ! FstTwoChr = '@
'Fm  MlNm :         # Ml-Name. @@
Const CmPfx$ = "XF_"
Dim Dr, Dry(), S$, Lin$, L&
For Each Dr In Itr(Bs.Dry)
    L = Dr(0)
    Lin = Dr(1)
    S = WBsStmt(Lin, BsVPfx, Mdn, MlNm)
    PushI Dry, Array(L, S, Lin)
Next
BsLNewO = DrszFF("L NewL OldL", Dry)
'Insp "QIde_B_MthOp.BsLNewO", "Inspect", "Oup(BsLNewO) Bs VSfx Mdn MlNm", FmtDrs(BsLNewO), FmtDrs(Bs), VSfx, Mdn, MlNm: Stop
End Function

Private Function WBsStmt$(BsLin, VSfx As Dictionary, Mdn$, MlNm$)
If Left(BsLin, 2) <> "'@" Then Thw CSub, "BsLin is always begin with '@", "BsLin", BsLin
Dim NN$: NN = Trim(RmvPfx(BsLin, "'@"))
Dim E$: E = InspExprLis(NN, VSfx)
WBsStmt = InspStmt(NN, E, Mdn, MlNm)
End Function

Friend Function OCrtMb(CmMd As CodeModule, MbNew As Drs)
'Fm NewMb : Cm NewMbL @@
Dim Dr: For Each Dr In Itr(MbNew.Dry)
    Dim Lin$: Lin = Dr(1)
    
    Dim Mthn$:  Mthn = Dr(0)
    Dim MLno&: MLno = MthLnozMM(CmMd, Mthn)
    Dim ELno&: ELno = MthELno(CmMd, MLno)
    Dim L&: L = ELno
    
    CmMd.InsertLines L, Lin
Next
End Function

Friend Function MbEpt(MbLis As Drs, Mdn$) As Drs
'Fm  CmLis : Mthn MthLin
'Ret       : Mthn MthLin MbStmt @@
Dim Dr, Dry()
For Each Dr In Itr(MbLis.Dry)
    Dim MthLin$: MthLin = LasEle(Dr)
    Dim MbStmt$: MbStmt = "'" & InspMthStmt(MthLin, Mdn) & ": Stop"
    PushI Dr, MbStmt
    PushI Dry, Dr
Next
MbEpt = AddColzFFDry(MbLis, "MbStmt", Dry)
'Insp "QIde_B_MthOp.MbEpt", "Inspect", "Oup(MbEpt) CmLis Mdn", FmtDrs(MbEpt), FmtDrs(CmLis), Mdn: Stop
End Function

Friend Function MbAct(Cm$(), Md As CodeModule) As Drs
'Ret      : L Mthn OldL ! OldL is MbStmt @@
Dim A As Drs: A = DMthE(Md)             ' L E Mdy Ty Mthn MthLin
Dim B As Drs: B = DrswColIn(A, "Mthn", Cm)
Dim Dr, Dry(): For Each Dr In Itr(B.Dry)
    Dim E&:           E = Dr(1)
    Dim L&:           L = E - 1          ' ! The Lno of MbStmt
    Dim Mthn$:     Mthn = Dr(4)
    Dim MbStmt$: MbStmt = Md.Lines(L, 1)
    Select Case True
    Case HasPfx(MbStmt, "'Insp "), HasPfx(MbStmt, "Insp ")
        PushI Dry, Array(L, Mthn, MbStmt)
    End Select
Next
MbAct = DrszFF("L Mthn OldL", Dry)
'BrwDrs MbAct: Stop
'Insp "QIde_B_MthOp.MbAct", "Inspect", "Oup(MbAct) Cm CmMd", FmtDrs(MbAct), Cm, Mdn(CmMd): Stop
End Function

Friend Function Bs(Mc As Drs) As Drs
'Fm  McCln : L MthLin # Mc-Cln. ! must Dim | Asg | Rmk(but not 'If 'Insp, '==). Cln to Align
'Ret       : L BsLin            ! FstTwoChr = '@ @@
Dim Dr, Dry()
For Each Dr In Itr(Mc.Dry)
    If HasPfx(Dr(1), "'@") Then PushI Dry, Dr
Next
Bs = DrszFF("L BsLin", Dry)
'Insp "QIde_B_MthOp.Bs", "Inspect", "Oup(Bs) McCln", FmtDrs(Bs), FmtDrs(McCln): Stop
End Function
Friend Function McAlign(McFill As Drs) As Drs
'Fm  McFill : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS Expr
'             F0 FSfx FExpr FR1 FR2 ! Adding F*
'Ret        : L Align               ! Bld the new Align @@
Const CmPfx$ = "XE_"
If NoReczDrs(McFill) Then Stop
Dim A As Drs: A = SelDrs(McFill, "L Gpno MthLin Dcl LHS Expr R1 R2 R3 F0 FDcl FLHS FExpr FR1 FR2")
Dim Gpno: Gpno = DistCol(McFill, "Gpno")
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim B As Drs: B = ColEq(A, "Gpno", IGpno)
    Dim C As Drs: C = WAlign(B)
    Dim O As Drs: O = AddDrs(O, C)
Next
McAlign = O
'Insp "QIde_B_MthOp.McAlign", "Inspect", "Oup(McAlign) McFill", FmtDrs(McAlign), FmtDrs(McFill): Stop
End Function

Friend Function SfLNewO(NoSf As Boolean, SfLno&, SfLin$, CclsNm$) As Drs
'Ret      : Only one or no line @@
If NoSf Then SfLNewO = EmpLNewO: Exit Function
Dim NewL$: NewL = "Static F As New " & CclsNm
If SfLin = NewL Then SfLNewO = EmpLNewO: Exit Function
SfLNewO = LNewO(Av(Array(SfLno, NewL, SfLin)))
'Insp "QIde_B_MthOp.SfLNewO", "Inspect", "Oup(SfLNewO) NoSf SfLno SfLin CclsNm", FmtDrs(SfLNewO), NoSf, SfLno, SfLin, CclsNm: Stop
End Function

Friend Sub OEnsCcls(NoSf As Boolean, HasCcls As Boolean, Md As CodeModule, CclsNm$)
'Ret #Ens-Chd-Cls.
If NoSf Then Exit Sub
If HasCcls Then Exit Sub
EnsCmpzPTN PjzM(Md), vbext_ct_ClassModule, CclsNm
End Sub

Friend Function CmPfx$(NoSf As Boolean, McLy$())
If Not NoSf Then Exit Function
CmPfx = StrValzCnstLy(McLy, "CmPfx")
'Insp "QIde_B_MthOp.CmPfx", "Inspect", "Oup(CmPfx) NoSf McLy", CmPfx, NoSf, McLy: Stop
End Function
Friend Function CmPfxEr(CmPfx$, Md As CodeModule) As Boolean
If CmPfx = "" Then Exit Function
Dim L, Cnt%, V$
For Each L In Itr(Src(Md))
    V = StrValzCnstn(L, "CmPfx")
    If V = "" Then GoTo X
    If V = CmPfx Then Cnt = Cnt + 1
    If Cnt >= 2 Then
        Debug.Print "CmPfx(?) is duplicated in Md(?)", CmPfx, Mdn(Md)
        CmPfxEr = True
    End If
    If V <> CmPfx Then
        If HasPfx(V, CmPfx) Then
            Debug.Print FmtQQ("Given CmPfx(?), there is another CmPfx(?) contains the given.  This is not allowed", CmPfx, V)
            CmPfxEr = True
            Exit Function
        End If
    End If
X:
Next
'Insp "QIde_B_MthOp.CmPfxEr", "Inspect", "Oup(CmPfxEr) CmPfx Md", CmPfxEr, CmPfx, Mdn(Md): Stop
End Function

Friend Function WRmk$(FR1%, FR2%, R1$, R2$, R3$)
If R1 = "" And R2 = "" And R3 = "" Then Exit Function
Dim A$, B$, C$
A = R1 & Space(FR1)
If R2 = "" Then
    B = "  " & Space(FR2)
Else
    B = " #" & R2 & Space(FR2)
End If
If R3 <> "" Then
    C = " ! " & R3
End If
WRmk = RTrim(A & B & C)
End Function


Friend Function CmLin(CmDta As Drs, CmMdy$) As String()
'Fm  CmDta  V TyChr Pm RetAs
'Ret CmLin MthLin                     ! MthLin is always a function @@
Dim Dr, CmNm$, MthPfx$, Pm$, TyChr$, RetAs$
For Each Dr In Itr(CmDta.Dry)
    CmNm = Dr(0)
    TyChr = Dr(1)
    Pm = Dr(2)
    RetAs = Dr(4)
    Dim MthLin$: MthLin = FmtQQ("? Function ???(?)?", CmMdy, CmNm, TyChr, Pm, RetAs)
    PushI CmLin, MthLin
Next
End Function

Friend Function XSelf(SkpChkSelf As Boolean, Md As CodeModule, Mln$) As Boolean
If SkpChkSelf Then Exit Function
Dim O As Boolean
O = Mdn(Md) = "QIde_Base_MthOp" And Mln = "AlignMthDimzML"
If O Then Inf CSub, "Self aligning"
XSelf = O
End Function

Friend Function XPm(Md As CodeModule, MthLno&) As Boolean
XPm = True
If IsNothing(Md) Then Debug.Print "Md is nothing": Exit Function
If MthLno <= 0 Then Debug.Print "MthLno <= 0": Exit Function
XPm = False
End Function

Friend Function McCln(Mc As Drs) As Drs
'Fm  Mc   : L MthLin # Mc.
'Ret      : L MthLin # Mc-Cln. ! must Dim | Asg | Rmk(but not 'If 'Insp, '==). Cln to Align @@
Dim Dr, Dry(), L$, Yes As Boolean
For Each Dr In Itr(Mc.Dry)
    L = Trim(Dr(1))
    Yes = False
    Select Case True
    Case HasPfx(L, "'")
        L = LTrim(RmvFstChr(L))
        Select Case True
        Case HasPfxss(L, "If Stop Insp == Brw")
        Case Else: Yes = True
        End Select
    Case IsAlignableDim(L), IsAsgStmt(L)
        Yes = True
    End Select
    If Yes Then PushI Dry, Dr
Next
McCln = Drs(Mc.Fny, Dry)
'BrwDrs McCln: Stop
'Insp "QIde_B_MthOp.McCln", "Inspect", "Oup(McCln) Mc", FmtDrs(McCln), FmtDrs(Mc): Stop
End Function

Friend Sub OUpdCr(CrChg As S1S2s, CmMd As CodeModule)
Dim J&, Ay() As S1S2
Ay = CrChg.Ay
For J = 0 To CrChg.N - 1
    With Ay(J)
    RplMthRmk CmMd, .S1, .S2
    End With
Next
End Sub

Private Sub OUpdCrOneV(V, NewRmk$, CmMd As CodeModule, CmPfx$)
Dim Mthn$: Mthn = CmPfx & V
RplMthRmk CmMd, Mthn, NewRmk
End Sub

Private Function WRmk_Lin$(Dr)
Dim V$, P$, R1$, R2$, R3$, FP%, FR1%, FR2%, Fst As Boolean
AsgAp Dr, V, P, R1, R2, R3, FP, FR1, FR2, Fst
Const CmPfx$ = "WRmk_"
Dim FmRet$:     FmRet = WRmk_FmRet(Fst, P)
Dim NmColon$: NmColon = WRmk_NmColon(Fst, P, FP)
Dim R$:             R = WRmk_RmkL(R1, R2, R3, FR1, FR2)
            WRmk_Lin = Trim(FmRet & NmColon & R)
End Function
Private Function WRmk_FmRet$(Fst As Boolean, P$)
Dim O$
Select Case True
Case Fst And P = "*Ret": O = "'Ret "
Case Fst:                O = "'Fm  "
Case Else:               O = "'    "
End Select
WRmk_FmRet = O
End Function
Private Function WRmk_NmColon$(Fst As Boolean, P$, FP%)
Dim O$
Select Case True
Case Fst And P = "*Ret": O = Space(FP%) & " : "
Case Fst:                O = AlignL(P, FP%) & " : "
Case Else:               O = Space(FP + 3)
End Select
WRmk_NmColon = O
End Function

Private Function WRmk_RmkL$(R1$, R2$, R3$, FR1%, FR2%)
If R1 <> "" Then
    If Left(R1, 3) <> " ' " Then Stop
    R1 = Mid(R1, 4)
End If
WRmk_RmkL = RTrim(AlignL(R1, FR1 - 3) & AlignL(R2, FR2) & R3)
End Function

Friend Function CrChg(CrEpt As S1S2s, CrAct As S1S2s) As S1S2s
'Fm  CrEpt : Mthn RmkLines ! RmkLines is find by each V in CrVPR3 & Mthn = V & CmPfx
'Ret       : Mthn RmkLines ! Only those need to change @@
Dim J&, Ay() As S1S2, R As LnoLines, O As S1S2s
Ay = CrEpt.Ay
For J = 0 To CrEpt.N - 1
    With Ay(J)
    Dim A As StrOpt: A = SomS2(.S1, CrAct)
    If A.Som Then
        If .S2 <> A.Str Then
            PushS1S2 O, Ay(J)
        End If
    Else
        PushS1S2 O, Ay(J)
    End If
    End With
Next
CrChg = O
'Insp "QIde_B_MthOp.CrChg", "Inspect", "Oup(CrChg) CrEpt CrAct", FmtS1S2s(CrChg), FmtS1S2s(CrEpt), FmtS1S2s(CrAct): Stop
End Function

Friend Function CmEpt(McR123 As Drs, ExprPfx$) As String()
'Fm  McR123  : L Gpno MthLin IsRmk
'              V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
'Fm  ExprPfx :                             ! Either F. or CmPfx.  It used to detect the Expr is a calling cm expr
'Ret         :                             ! It is from V and  {V} = {CmPfx}{Expr}.
'                                          ! They will be used create new chd mth @@
'Ret ! It is from V and Expr=V
Dim A As Drs: A = ColEq(McR123, "IsRmk", False)
Dim CV$(), CExpr$()
    AsgCol A, "V Expr", CV, CExpr
    If Si(CV) = 0 Then Exit Function

Dim V, J%
For Each V In CV
    If ExprPfx & V = CExpr(J) Then
        PushI CmEpt, CV(J)  '<=====
    End If
    J = J + 1
Next
'Brw CmEpt: Stop
'Insp "QIde_B_MthOp.CmEpt", "Inspect", "Oup(CmEpt) McR123 ExprPfx", CmEpt, FmtDrs(McR123), ExprPfx: Stop
End Function


