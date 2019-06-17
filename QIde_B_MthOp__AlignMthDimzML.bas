VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QIde_B_MthOp__AlignMthDimzML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Friend Function CmlAct(CmlEpt As Drs, M As CodeModule) As Drs
Dim CV$(): CV = StrColzDrs(CmlEpt, "V")
Dim Act$(): Act = StrCol(DMth(M), "MthLin")
Insp CSub, "CmlAct", "CV Act", CV, Act
Stop
End Function

Friend Function MlVSfx(Ml$) As Drs
'Ret : Ret V Sfx ! the MthLin's pm V Sfx @@
Dim Pm$: Pm = BetBkt(Ml)
Dim P, V$, Sfx$, Dry(), L$
For Each P In Itr(SyzTrim(SplitComma(Pm)))
    L = RmvPfx(P, "ByVal ")
    L = RmvPfx(L, "Optional ")
    V = ShfNm(L)
    Sfx = L
    PushI Dry, Array(V, Sfx)
Next
MlVSfx = DrszFF("V Sfx", Dry)
'Insp "QIde_B_MthOp.MlVSfx", "Inspect", "Oup(MlVSfx) Ml", FmtDrs(MlVSfx), Ml: Stop
End Function

Friend Function CmlMthRet(CmlDclPm As Drs) As Drs
'Fm CmlDclPm : V Sfx RHS CmNm Pm DclPm             ! use [CmlVSfx] & [Pm] to bld [DclPm]
'Ret         : V Sfx RHS CmNm Pm DclPm TyChr RetAs @@
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
Friend Function CmlDclPm(CmlPm As Drs, CmlVSfx As Drs) As Drs
'Fm CmlPm : V Sfx RHS CmNm Pm
'Ret      : V Sfx RHS CmNm Pm DclPm ! use [CmlVSfx] & [Pm] to bld [DclPm] @@
Dim IxPm%: AsgIx CmlPm, "Pm", IxPm
Dim Dr, Dry(): For Each Dr In Itr(CmlPm.Dry)
    Dim Pm$: Pm = Dr(IxPm)
    Dim DclPm$: DclPm = XDclPm(Pm, CmlVSfx)
    PushI Dr, DclPm
    PushI Dry, Dr
Next
CmlDclPm = AddColzFFDry(CmlPm, "DclPm", Dry)
'Insp "QIde_B_MthOp.CmlDclPm", "Inspect", "Oup(CmlDclPm) CmlPm CmlVSfx", FmtDrs(CmlDclPm), FmtDrs(CmlPm), FmtDrs(CmlVSfx): Stop
End Function
Private Function XDclPm$(Pm$, CmlVSfx As Drs)
Dim O$(), Sfx$, P
For Each P In Itr(SyzSS(Pm))
    Sfx = ValzColEq(CmlVSfx, "Sfx", "V", P)
    PushI O, P & Sfx
Next
XDclPm = JnCommaSpc(O)
'Insp CSub, "Finding DclPm(CallgPm, CmlFmMc)", "CallgPm CmlFmMc XDclPm", CallgPm, FmtDrs(CmlFmMc), XDclPm: Stop
End Function

Friend Function CrEpt(CrJn As Drs) As S1S2s
'Fm  CrJn : V Rmk CmNm
'Ret      : CmNm RmkLines ! RmkLines is find by each V in CrVpr & Mthn = V & CmPfx @@
Dim A As Drs: A = SelDrs(CrJn, "CmNm Rmk")
Dim Dr, Ly$(): For Each Dr In Itr(A.Dry)
    PushI Ly, Dr(0) & " " & Dr(1)
Next
Dim D As Dictionary: Set D = Dic(Ly, vbCrLf)
Dim O As Dictionary: Set O = AddSfxzDic(D, " @@")
CrEpt = S1S2szDic(O)
'Insp "QIde_B_MthOp.CrEpt", "Inspect", "Oup(CrEpt) CrJn", FmtS1S2s(CrEpt), FmtDrs(CrJn): Stop
End Function

Private Function XVyzPm(Pm() As Vr) As String()
Dim J%: For J = 0 To SiVr(Pm) - 1
    Push XVyzPm, Pm(J).V
Next
End Function
Private Sub XR12zR123Ay(R() As R123, OR1$(), OR2$())
Dim J%: For J = 0 To SiR123(R) - 1
    PushI OR1, R(J).R1
    PushI OR2, R(J).R2
Next
End Sub

Private Sub XR12zPm(Pm() As Vr, OR1$(), OR2$())
Dim J%: For J = 0 To SiVr(Pm) - 1
    Dim R1$(), R2$(): XR12zR123Ay Pm(J).R, R1, R2
    OR1 = SyzAdd(OR1, R1)
    OR2 = SyzAdd(OR2, R2)
Next
End Sub

Private Function XR12(A As Vpr, OR1$(), OR2$())
Dim P1$(), P2$(): XR12zPm A.Pm, P1, P2
Dim R1$(), R2$(): XR12zR123Ay A.Ret, R1, R2
OR1 = SyzAdd(P1, R1)
OR2 = SyzAdd(P2, R2)
End Function

Private Sub XW12(A As Vpr, OW1%, OW2%)
Dim R1$(), R2$(): XR12 A, R1, R2
OW1 = WdtzAy(R1)
OW2 = WdtzAy(R2)
End Sub

Private Function XRmkzPmAy(Pm() As Vr, WV%, W1%, W2%) As String()
Dim J%: For J = 0 To SiVr(Pm) - 1
    Dim R$(): R = XRmkzPm(Pm(J), WV, W1, W2)
    PushIAy XRmkzPmAy, R
Next
End Function

Private Function XRmkzPm(Pm As Vr, WV%, W1%, W2%) As String()
Dim J%: For J = 0 To SiR123(Pm.R) - 1
    Dim L$
    If J = 0 Then
        L = "Fm " & Align(Pm.V, WV) & " : "
    Else
        L = Space(3 + WV + 3)
    End If
    PushI XRmkzPm, XRmkzLin(L, Pm.R(J), W1, W2)
Next
End Function

Private Function XRmkzLin$(Lbl$, R As R123, W1%, W2%)
'Fm Lbl : Aligned | no ' | with [ : ]
'Fm W1  : wdt of R1
'Fm W2  : wdt of R2
'Ret    : rmk lin
If Left(R.R1, 3) <> " ' " Then Stop
Dim R0$: R0 = "'" & Lbl
Dim R1$: R1 = Mid(AlignL(R.R1, W1), 4)
Dim R2$: R2 = AlignL(R.R2, W2)
XRmkzLin = RTrim(R0 & R1 & R2 & R.R3)
End Function

Private Function XRmkzRet(Ret() As R123, WV%, W1%, W2%) As String()
Dim J%: For J = 0 To SiR123(Ret) - 1
    Dim W%: W = 3 + WV
    Dim Lbl$
        If J = 0 Then
            Lbl = AlignL("Ret", W) & " : "
        Else
            Lbl = Space(W + 3)
        End If
    PushI XRmkzRet, XRmkzLin(Lbl, Ret(J), W1, W2)
Next
End Function

Private Function XRmkL$(Vpr As Vpr)
Dim Vy$():    Vy = XVyzPm(Vpr.Pm) ' All parameter name
Dim WV%:      WV = WdtzAy(Vy)       ' Wdt of all parameter
Dim W1%, W2%:      XW12 Vpr, W1, W2
Dim R1$():    R1 = XRmkzPmAy(Vpr.Pm, WV, W1%, W2%)
Dim R2$():    R2 = XRmkzRet(Vpr.Ret, WV, W1%, W2%)
           XRmkL = RTrim(JnCrLf(SyzAdd(R1, R2))) & " @@"
End Function

Friend Function CrVRmk(CrVpr() As Vpr) As S1S2s
'Fm CrVpr :
'Ret      : V RmkLines @@
Dim J%: For J = 0 To SiVpr(CrVpr) - 1
    Dim Vpr As Vpr: Vpr = CrVpr(J)
    Dim V$:         V = Vpr.V
    Dim RmkL$: RmkL = XRmkL(Vpr)
    PushS1S2 CrVRmk, S1S2(V, RmkL)
Next
'Insp "QIde_B_MthOp.CrVRmk", "Inspect", "Oup(CrVRmk) CrVpr", FmtS1S2s(CrVRmk), "NoFmtr(() As Vpr)": Stop
End Function

Private Function XVrzV(A() As Vr, V$) As Vr
Dim J%: For J = 0 To SiVr(A) - 1
    If A(J).V = V Then
        XVrzV = A(J)
        Exit Function
    End If
Next
End Function

Private Function XVrAyzPm(Pm$, CrVrAy() As Vr) As Vr()
Dim P: For Each P In Itr(SyzSS(Pm))
    Dim V$: V = P
    Dim M As Vr: M = XVrzV(CrVrAy, V)
    If M.V = P Then
        PushVr XVrAyzPm, M.V, M.R
    End If
Next
End Function

Private Function XVpr(V$, Pm$, CrVrAy() As Vr) As Vpr
Dim O As Vpr
O.V = V
O.Ret = XVrzV(CrVrAy, V).R
O.Pm = XVrAyzPm(Pm, CrVrAy)
XVpr = O
End Function

Friend Function CrVpr(CrPm As Drs, CrVrAy() As Vr) As Vpr()
'Fm CrPm : V Pm
'Ret     : Sam itm as CrPm @@
Dim VPm As Dictionary: Set VPm = DiczDrsCC(CrPm, "V Pm")
Dim Dr: For Each Dr In Itr(CrPm.Dry)
    Dim V$:        V = Dr(0)
    Dim Pm$:      Pm = Dr(1)
    Dim M As Vpr:  M = XVpr(V, Pm, CrVrAy)
    PushVpr CrVpr, M
Next
If NReczDrs(CrPm) <> SiVpr(CrVpr) Then Stop
'Insp "QIde_B_MthOp.CrVpr", "Inspect", "Oup(CrVpr) CrPm CrVrAy", "NoFmtr(Vpr())", FmtDrs(CrPm), "NoFmtr(() As Vr)": Stop
End Function

Friend Function CrWiRmk(CrSel As Drs) As Drs
'Fm CrSel : V R1 R2 R3
'Ret      : V R1 R2 R3 ! rmv those R1 2 3 are blank @@
Dim IxR1%, IxR2%, IxR3%
AsgIx CrSel, "R1 R2 R3", IxR1, IxR2, IxR3
Dim Dr, Dry(): For Each Dr In Itr(CrSel.Dry)
    Dim R1$: R1 = Dr(IxR1)
    Dim R2$: R2 = Dr(IxR2)
    Dim R3$: R3 = Dr(IxR3)
    If R1 <> "" Or R2 <> "" Or R3 <> "" Then
        PushI Dry, Dr
    End If
Next
CrWiRmk = Drs(CrSel.Fny, Dry)
'Insp "QIde_B_MthOp.CrWiRmk", "Inspect", "Oup(CrWiRmk) CrSel", FmtDrs(CrWiRmk), FmtDrs(CrSel): Stop
End Function

Friend Function CrVr(CrWiRmk As Drs) As Vr()
'Fm CrWiRmk : V R1 R2 R3 ! rmv those R1 2 3 are blank @@
ThwIf_DifFF CrWiRmk, "V R1 R2 R3", CSub
Dim LasV$, M() As R123, Fst As Boolean
Fst = True
Dim Dr: For Each Dr In Itr(CrWiRmk.Dry)
    Dim V$: V = Dr(0)
    Dim R1$: R1 = Dr(1)
    Dim R2$: R2 = Dr(2)
    Dim R3$: R3 = Dr(3)
    If Fst Then
        Fst = False
        LasV = V
    Else
        If V <> "" Then
            PushVr CrVr, LasV, M
            LasV = V
            Erase M
        End If
    End If
    PushR123 M, R1, R2, R3
Next
PushVr CrVr, LasV, M
'Insp "QIde_B_MthOp.CrVr", "Inspect", "Oup(CrVr) CrWiRmk", "NoFmtr(Vr())", FmtDrs(CrWiRmk): Stop
End Function

Friend Function McVSfx(McInsp As Drs) As Drs
'Fm McInsp : L Gpno MthLin IsRmk ! If las lin is rmk and is 'Insp, exl it.
'Ret       : L Gpno MthLin IsRmk
'          : V Sfx Rst @@
Dim Dr, Dry()
For Each Dr In Itr(McInsp.Dry)
    Dim Av(): Av = WDrVSfxRst(Dr(2))
    PushIAy Dr, Av
    PushI Dry, Dr
Next
McVSfx = AddColzFFDry(McInsp, "V Sfx Rst", Dry)
'Insp "QIde_B_MthOp.McVSfx", "Inspect", "Oup(McVSfx) McInsp", FmtDrs(McVSfx), FmtDrs(McInsp): Stop
End Function

Friend Function McDcl(McVSfx As Drs) As Drs
'Fm McVSfx : L Gpno MthLin IsRmk
'          : V Sfx Rst
'Ret       : L Gpno MthLin IsRmk
'          : V Sfx Dcl Rst       ! Add Dcl from V & Sfx @@
Dim MthLin$, IxMthLin%, Dr, Dry(), IGpno
AsgIx McVSfx, "MthLin", IxMthLin
For Each IGpno In AywDist(IntCol(McVSfx, "Gpno"))
    Dim A As Drs: A = ColEq(McVSfx, "Gpno", IGpno) ' L Gpno MthLin IsRmk V Sfx Rst ! Sam Gpno
    Dim B As Drs: B = XDcl(A) ' L Gpno MthLin IsRmk V Sfx Dcl Rst ! Adding Dcl using V Sfx
    Dim O As Drs: O = AddDrs(O, B)
Next
McDcl = O
'Insp "QIde_B_MthOp.McDcl", "Inspect", "Oup(McDcl) McVSfx", FmtDrs(McDcl), FmtDrs(McVSfx): Stop
End Function
Private Function XWAs%(V$(), Sfx$())
Dim C$(), J%: For J = 0 To UB(V)
    Select Case True
    Case HasPfx(Sfx(J), " As "):   Push C, V(J)
    Case HasPfx(Sfx(J), "() As "): Push C, V(J) & "()"
    End Select
Next
XWAs = WdtzAy(C)
End Function
Private Function XDcl(A As Drs) As Drs
'Fm A :      L Gpno MthLin IsRmk V Sfx Rst
'Ret McDclI: L Gpno MthLin IsRmk V Sfx Dcl Rst @@
Dim V$():     V = StrCol(A, "V")
Dim Sfx$(): Sfx = StrCol(A, "Sfx")
Dim WAs%:   WAs = XWAs(V, Sfx)
Dim Dr, J%, Dry(): For Each Dr In Itr(A.Dry)
    Dim Dcl$: Dcl = XDclzV(V(J), WAs, Sfx(J))
    Dim Rst$: Rst = Pop(Dr)
    PushIAy Dr, Array(Dcl, Rst)
    PushI Dry, Dr
    J = J + 1
Next
Dim Fny$(): Fny = AyzAdd(AyeLasEle(A.Fny), Array("Dcl", "Rst"))
XDcl = Drs(Fny, Dry)
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
Private Function XDclzV$(V$, WAs%, Sfx$)
If V = "" Then Exit Function
Dim O$
Select Case True
Case HasPfx(Sfx, " As "):   O = AlignL(V, WAs) & Sfx
Case HasPfx(Sfx, "() As "): O = AlignL(V & "()", WAs) & RmvPfx(Sfx, "()")
Case Else:                  O = V & Sfx
End Select
XDclzV = "Dim " & O & ": "
'Debug.Print XDclzV; WAs; QuoteSq(Sfx); "<": Stop
End Function
Friend Function McLR(McDcl As Drs) As Drs
'Fm McDcl : L Gpno MthLin IsRmk
'         : V Sfx Dcl Rst         ! Add Dcl from V & Sfx
'Ret      : L Gpno MthLin IsRmk
'         : V Sfx Dcl LHS RHS Rst ! Add LHS Expr from Rst @@
Dim Dr: For Each Dr In Itr(McDcl.Dry)
    Dim L$: L = Pop(Dr)
    Dim Rst$:        Rst = L
    Dim LHS$, Expr$:       AsgAp ShfLRHS(Rst), LHS, Expr
    If Rst <> "" Then
        If Not IsVbRmk(Rst) Then Stop
    End If
                      Dr = AyzAdd(Dr, Array(LHS, Expr, Rst))
    Dim Dry():             PushI Dry(), Dr
Next
Dim Fny$(): Fny = AyeLasEle(McDcl.Fny)
Fny = AyzAdd(Fny, SyzSS("LHS RHS Rst"))
McLR = Drs(Fny, Dry)
'Insp "QIde_B_MthOp.McLR", "Inspect", "Oup(McLR) McDcl", FmtDrs(McLR), FmtDrs(McDcl): Stop
End Function

Friend Function McR123(McLR As Drs) As Drs
'Fm McLR : L Gpno MthLin IsRmk
'        : V Sfx Dcl LHS RHS Rst       ! Add LHS Expr from Rst
'Ret     : L Gpno MthLin IsRmk
'        : V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst @@
Dim Dr: For Each Dr In Itr(McLR.Dry)
    Dim Rst$:          Rst = LasEle(Dr)
    If FstChr(Rst) <> "'" Then Stop
    Dim R1$, R2$, R3$:       AsgBrkBet Rst, "#", "!", R1, R2, R3
                        R1 = Trim(RmvFstChr(R1))
                             If R2 <> "" Then R2 = " # " & R2
                             If R3 <> "" Then R3 = " ! " & R3
                             If R1 <> "" Or R2 <> "" Or R3 <> "" Then R1 = " ' " & R1
    Dim Dry():               PushI Dry, AyzAdd(AyeLasEle(Dr), Array(R1, R2, R3))
Next
Dim Fny$(): Fny = AyzAdd(AyeLasEle(McLR.Fny), SyzSS("R1 R2 R3"))
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
'Fm McR123 : L Gpno MthLin IsRmk
'          : V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
'Ret       : L Gpno MthLin IsRmk
'          : V Sfx Dcl LHS Expr
'          : F0 FSfx FExpr FR1 FR2       ! Adding F* @@
Dim Gpno%(): Gpno = AywDist(IntCol(McR123, "Gpno"))
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim A As Drs: A = ColEq(McR123, "Gpno", IGpno)
    Dim B As Drs: B = WF0(A)
    Dim C As Drs: C = AddColzFiller(B, "Dcl LHS RHS R1 R2")
    Dim O As Drs: O = AddDrs(O, C)
Next
McFill = O
'Insp "QIde_B_MthOp.McFill", "Inspect", "Oup(McFill) McR123", FmtDrs(McFill), FmtDrs(McR123): Stop
End Function

Friend Function De(Mc As Drs) As Drs
'Fm Mc : L MthLin # Mc.
'Ret   : L MthLin # Dbl-Eq | Dbl-Dash @@
Dim Dr, Dry(), L$
For Each Dr In Itr(Mc.Dry)
    L = LTrim(Dr(1))
    Select Case Left(L, 3)
    Case "'==", "'--": PushI Dry, Dr
    End Select
Next
De.Fny = Mc.Fny
De.Dry = Dry
'Insp "QIde_B_MthOp.De", "Inspect", "Oup(De) Mc", FmtDrs(De), FmtDrs(Mc): Stop
End Function

Friend Function DeLNewO(De As Drs) As Drs
'Fm De : L MthLin    # Dbl-Eq | Dbl-Dash
'Ret   : L NewL OldL @@
Dim Dr, Dry(): For Each Dr In Itr(De.Dry)
    Dim L&:       L = Dr(0)
    Dim OldL$: OldL = Dr(1)
    Dim C$:       C = Mid(LTrim(OldL), 2, 1)
    Dim NewL$: NewL = Left(OldL, 120) & Dup(C, 120 - Len(OldL))
    If OldL <> NewL Then
        Push Dry, Array(L, NewL, OldL)
    End If
Next
DeLNewO = LNewO(Dry)
'Insp "QIde_B_MthOp.DeLNewO", "Inspect", "Oup(DeLNewO) De", FmtDrs(DeLNewO), FmtDrs(De): Stop
End Function
Friend Function McGp(McCln As Drs) As Drs
'Fm McCln : L MthLin      # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align
'Ret      : L Gpno MthLin           ! with L in seq will be one gp @@
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
'Fm McRmk : L Gpno MthLin IsRmk ! a column IsRmk is added
'Ret      : L Gpno MthLin IsRmk ! For each gp, the front rmk lines are TopRmk, rmv them @@
Dim IGpno%, MaxGpno, A As Drs, B As Drs, O As Drs
MaxGpno = MaxzAy(IntCol(McRmk, "Gpno"))
For IGpno = 1 To MaxGpno
    A = ColEq(McRmk, "Gpno", IGpno)
    B = McTRmkI(A)
    O = AddDrs(O, B)
Next
McTRmk = O
'Insp "QIde_B_MthOp.McTRmk", "Inspect", "Oup(McTRmk) McRmk", FmtDrs(McTRmk), FmtDrs(McRmk): Stop
End Function

Friend Function McInsp(McTRmk As Drs) As Drs
'Fm McTRmk : L Gpno MthLin IsRmk ! For each gp, the front rmk lines are TopRmk, rmv them
'Ret       : L Gpno MthLin IsRmk ! If las lin is rmk and is 'Insp, exl it. @@
McInsp = McTRmk
If NoReczDrs(McTRmk) Then Exit Function
Dim Dr: Dr = LasEle(McTRmk.Dry)
Dim IxMthLin%: IxMthLin = IxzAy(McTRmk.Fny, "MthLin")
Dim L$: L = Dr(IxMthLin)
If IsVbRmk(L) Then
    Dim A$: A = Left(LTrim(RmvFstChr(LTrim(L))), 4)
    If A = "Insp" Then
        Pop McInsp.Dry
    End If
End If
'Insp "QIde_B_MthOp.McInsp", "Inspect", "Oup(McInsp) McTRmk", FmtDrs(McInsp), FmtDrs(McTRmk): Stop
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
'Fm McGp : L Gpno MthLin       ! with L in seq will be one gp
'Ret     : L Gpno MthLin IsRmk ! a column IsRmk is added @@
Dim Dr: For Each Dr In Itr(McGp.Dry)
    PushI Dr, FstChr(LTrim(Dr(2))) = "'"
    Push McRmk.Dry, Dr
Next
McRmk.Fny = AddFF(McGp.Fny, "IsRmk")
'Insp "QIde_B_MthOp.McRmk", "Inspect", "Oup(McRmk) McGp", FmtDrs(McRmk), FmtDrs(McGp): Stop
End Function

Friend Function CdNewCm$(CmNew As Drs, CmMdy$)
'Ret :  ! Cd to be append to CmMd @@
Dim Dr, O$(): For Each Dr In Itr(CmNew.Dry)
    Dim CmNm$:   CmNm = Dr(2)
    Dim Sfx$:    ' Sfx = ValzColEq(A, "Sfx", "V", V): If Not IsStr(Sfx) Then Stop
    Dim TyChr$: TyChr = TyChrzDclSfx(Sfx)
    Dim RetAs$: RetAs = RetAszDclSfx(Sfx)
    PushI O, ""
    PushI O, FmtQQ("? Function ??()?", CmMdy, CmNm, TyChr, RetAs)
    PushI O, "End Function"
Next
CdNewCm = JnCrLf(O)
'Insp "QIde_B_MthOp.CdNewCm", "Inspect", "Oup(CdNewCm) CmNew CmMdy", CdNewCm, FmtDrs(CmNew), CmMdy: Stop
End Function

Friend Function CmlEpt(CmlMthRet As Drs, CmMdy$) As Drs
'Fm CmlMthRet : V Sfx RHS CmNm Pm DclPm TyChr RetAs
'Ret          : V CmNm EptL
'             : L Mdy Ty Mthn MthLin @@
Dim Dr, Dry(), Nm$, Ty$, Pm$, Ret$, V$, EptL$, INm%, ITy%, IPm%, IRet%, IV%
AsgIx CmlMthRet, "CmNm TyChr DclPm RetAs V", INm, ITy, IPm, IRet, IV
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
CmlEpt = DrszFF("V CmNm EptL", Dry)
'BrwDrs CmlEpt: Stop
'Insp "QIde_B_MthOp.CmlEpt", "Inspect", "Oup(CmlEpt) CmlMthRet CmMdy", FmtDrs(CmlEpt), FmtDrs(CmlMthRet), CmMdy: Stop
End Function

Private Function XPm$(RHS$, CmNm$)
If CmNm = "" Then Exit Function
Dim O$
If HasSubStr(RHS, "(") Then
    O = BetBkt(RHS)
Else
    O = RmvT1(RHS)
End If
XPm = JnSpc(AyTrim(SplitComma(O)))
End Function

Friend Function CmlPm(CmEpt As Drs) As Drs
'Fm CmEpt : V Sfx LHS RHS CmNm ! All CmNm has val
'Ret      : V Sfx RHS CmNm Pm @@
Dim IxRHS%, IxCmNm%: AsgIx CmEpt, "RHS", IxRHS, IxCmNm
Dim Dr, ODry(): For Each Dr In Itr(CmEpt.Dry)
    Dim RHS$: RHS = Dr(IxRHS)
    Dim CmNm$: CmNm = Dr(IxCmNm)
    PushI Dr, XPm(RHS, CmNm)
    PushI ODry, Dr
Next
CmlPm = AddColzFFDry(CmEpt, "Pm", ODry)
'Insp "QIde_B_MthOp.CmlPm", "Inspect", "Oup(CmlPm) CmEpt", FmtDrs(CmlPm), FmtDrs(CmEpt): Stop
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
Friend Function BsLNewO(Bs As Drs, VSfx As Dictionary, Mdn$, MlNm$) As Drs
'Fm Bs   : L BsLin            ! FstTwoChr = '@
'Fm MlNm :         # Ml-Name. @@
Dim Dr, Dry(), S$, Lin$, L&
For Each Dr In Itr(Bs.Dry)
    L = Dr(0)
    Lin = Dr(1)
    S = WBsStmt(Lin, VSfx, Mdn, MlNm)
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

Friend Function MbEpt(CmLis As Drs, Mdn$) As Drs
'Fm CmLis : Mthn MthLin
'Ret      : Mthn MthLin MbStmt @@
Dim Dr, Dry()
For Each Dr In Itr(CmLis.Dry)
    Dim MthLin$: MthLin = LasEle(Dr)
    Dim MbStmt$: MbStmt = "'" & InspMthStmt(MthLin, Mdn) & ": Stop"
    PushI Dr, MbStmt
    PushI Dry, Dr
Next
MbEpt = AddColzFFDry(CmLis, "MbStmt", Dry)
'Insp "QIde_B_MthOp.MbEpt", "Inspect", "Oup(MbEpt) CmLis Mdn", FmtDrs(MbEpt), FmtDrs(CmLis), Mdn: Stop
End Function

Friend Function MbAct(Cm$(), CmMd As CodeModule) As Drs
'Ret : L Mthn OldL ! OldL is MbStmt @@
Dim A As Drs: A = DMthe(CmMd)             ' L E CmMdy Ty Mthn MthLin
Dim B As Drs: B = ColIn(A, "Mthn", Cm)
Dim Dr, Dry(): For Each Dr In Itr(B.Dry)
    Dim E&:           E = Dr(1)
    Dim L&:           L = E - 1          ' ! The Lno of MbStmt
    Dim Mthn$:     Mthn = Dr(4)
    Dim MbStmt$: MbStmt = CmMd.Lines(L, 1)
    Select Case True
    Case HasPfx(MbStmt, "'Insp "), HasPfx(MbStmt, "Insp ")
        PushI Dry, Array(L, Mthn, MbStmt)
    End Select
Next
MbAct = DrszFF("L Mthn OldL", Dry)
'BrwDrs MbAct: Stop
'Insp "QIde_B_MthOp.MbAct", "Inspect", "Oup(MbAct) Cm CmMd", FmtDrs(MbAct), Cm, Mdn(CmMd): Stop
End Function

Friend Function Bs(McCln As Drs) As Drs
'Fm McCln : L MthLin # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align
'Ret      : L BsLin            ! FstTwoChr = '@ @@
Dim Dr, Dry()
For Each Dr In Itr(McCln.Dry)
    If HasPfx(Dr(1), "'@") Then PushI Dry, Dr
Next
Bs = DrszFF("L BsLin", Dry)
'Insp "QIde_B_MthOp.Bs", "Inspect", "Oup(Bs) McCln", FmtDrs(Bs), FmtDrs(McCln): Stop
End Function
Friend Function McAlign(McFill As Drs) As Drs
'Fm McFill : L Gpno MthLin IsRmk
'          : V Sfx Dcl LHS Expr
'          : F0 FSfx FExpr FR1 FR2 ! Adding F*
'Ret       : L Align               ! Bld the new Align @@
If NoReczDrs(McFill) Then Stop
Dim A As Drs: A = SelDrs(McFill, "L Gpno MthLin Dcl LHS RHS R1 R2 R3 F0 FDcl FLHS FRHS FR1 FR2")
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
'Ret : Only one or no line @@
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

Friend Function IsPmEr(Md As CodeModule, MthLno&) As Boolean
IsPmEr = True
If IsNothing(Md) Then Debug.Print "Md is nothing": Exit Function
If MthLno <= 0 Then Debug.Print "MthLno <= 0": Exit Function
IsPmEr = False
End Function

Friend Function McCln(Mc As Drs) As Drs
'Fm Mc : L MthLin # Mc.
'Ret   : L MthLin # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align @@
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
    Case IsLinzSngDimColon(L), IsLinzAsg(L)
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
    EnsMthRmk CmMd, .S1, .S2
    End With
Next
End Sub

Private Function WRmk_Lin$(Dr)
Dim V$, P$, R1$, R2$, R3$, FP%, FR1%, FR2%, Fst As Boolean
AsgAp Dr, V, P, R1, R2, R3, FP, FR1, FR2, Fst
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
'Fm CrEpt : CmNm RmkLines
'Ret      : CmNm RmkLines ! Only those need to change @@
Dim J&, Ay() As S1S2, O As S1S2s
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

Private Function XCmNm$(RHS$, V$, HasSf As Boolean, MlNmDD$)
If RHS = "AlignMthDimzML__X" Then Stop
Dim RDotNm$: RDotNm = TakDotNm(RHS)
Select Case True
Case HasSf And RDotNm = "F." & V:           XCmNm = V
Case RDotNm = MlNmDD & V, RDotNm = "X" & V: XCmNm = RDotNm
End Select
End Function

Friend Function CmNm(CmV As Drs, WiSf As Boolean, MlNmDD$) As Drs
'Fm CmV : V Sfx LHR RHS
'Ret    : V Sfx LHS RHS CmNm ! som CmNm may be blank @@
Dim IxV%, IxRHS%
AsgIx CmV, "V RHS", IxV, IxRHS
Dim Dr, Dry(): For Each Dr In Itr(CmV.Dry)
    Dim V$:     V = Dr(IxV)
    Dim RHS$: RHS = Dr(IxRHS)
    PushI Dr, XCmNm(RHS, V, WiSf, MlNmDD)
    PushI Dry, Dr
Next
CmNm = AddColzFFDry(CmV, "CmNm", Dry)
'Insp "QIde_B_MthOp.CmNm", "Inspect", "Oup(CmNm) CmV WiSf MlNmDD", FmtDrs(CmNm), FmtDrs(CmV), WiSf, MlNmDD: Stop
End Function

Friend Function CmLHS(CmV As Drs) As Drs
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS ! where V & ' = ' = LHS @@
Dim IxV%, IxLHS%
AsgIx CmV, "V LHS", IxV, IxLHS
Dim Dr, Dry(): For Each Dr In Itr(CmV.Dry)
    Dim V$:     V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxLHS)
    If V & " = " = LHS Then Push CmLHS.Dry, Dr
Next
CmLHS.Fny = CmV.Fny
End Function

