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
Private Function WIsChdFun(V$, Expr$) As Boolean
WIsChdFun = HasPfx(Expr, "F." & V)
End Function
Friend Function Clc(McBrk As Drs) As Drs
Dim A As Drs: A = DrswColEq(McBrk, "IsRmk", False)
Dim ODry()
    Dim CV$(), CExpr$()
        AsgCol A, "V Expr", CV, CExpr
    Dim J%, Expr$, V$
    For J = 0 To UB(CV)
        V = CV(J)
        Expr = CExpr(J)
        If WIsChdFun(V, Expr) Then
            PushI ODry, Array(V, Expr)
        End If
    Next
Clc = DrszFF("ChdLinNm Calling", ODry)
End Function
Friend Function CmlLNewO(CmlJn As Drs) As Drs
'Fm CmlJn :  V EptL Lno ActL HasAct
Dim A As Drs: A = DrswColEq(CmlJn, "HasAct", False)
If HasReczDrs(A) Then Stop
'BrwDrs CmlJn: Stop
Dim B As Drs: B = DrseCeqC(CmlJn, "EptL ActL")
CmlLNewO = SelDrs(B, "Lno EptL ActL")
CmlLNewO.Fny = SyzSS("L NewL OldL")
BrwDrs CmlLNewO: Stop
End Function

Friend Function CmlAct(CmlEpt As Drs, Md As CodeModule) As Drs
Dim CV$(): CV = StrColzDrs(CmlEpt, "V")
Dim Act$(): Act = MthLinAyzM(Md)
Insp CSub, "CmlAct", "CV Act", CV, Act
Stop

End Function
Friend Function MlVSfx(Ml$) As Drs
'Ret V Sfx
Dim Pm$: Pm = BetBkt(Ml)
Dim P, V$, Sfx$, Dry(), L$
For Each P In Itr(TrimAy(SplitComma(Pm)))
    L = P
    V = ShfNm(L)
    Sfx = L
    PushI Dry, Array(V, Sfx)
Next
MlVSfx = DrszFF("V Sfx", Dry)
End Function

Friend Function CmlMthRet(CmlDclPm As Drs) As Drs
'Fm CmlDclPm : ..Sfx..
'Ret CmlMthRet : TyChr RetAs
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
End Function
Friend Function CmlDclPm(CmlCallg As Drs, CmlVSfx As Drs) As Drs
'Fm CmlCallg :
'Fm CmlVSfx : VSfx
'Ret CmlDclPm :
Dim Dry(), Dr, DclPm$, CallgPm$, Expr$
For Each Dr In Itr(CmlCallg.Dry)
    Expr = Dr(2)
    CallgPm = BetBkt(Expr)
    DclPm = WDclPm(CallgPm, CmlVSfx)
    PushI Dr, DclPm
    PushI Dry, Dr
Next
CmlDclPm = AddColzFFDry(CmlCallg, "CmlDclPm", Dry)
'BrwDrs CmlDclPm: Stop
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

Friend Function CrEpt(McBrk As Drs) As Drs
End Function

Friend Function McBrk(McTRmk As Drs) As Drs
'Fm  McTRmk: L Gpno MthLin IsRmk #Mc-TopRmk ! For each gp, the front rmk lines are TopRmk, rmv them
'Ret McBrk : L Gpno MthLin IsRmk            ! Brk the MthLin into V Sfx Expr Rmk
'            V Sfx Dcl LHS Expr             ! If there is no asg stmt LHS and Expr will be same as V
'            R1 R2 R3                       ! in this case, a new ChdMth will be created
Dim MaxGpno%: MaxGpno = MaxzAy(IntAyzDrsC(McTRmk, "Gpno"))
Dim IGpno%, A As Drs, O As Drs, B As Drs
If NoReczDrs(McTRmk) Then Stop
For IGpno = 1 To MaxGpno
    A = DrswColEq(McTRmk, "Gpno", IGpno)
    If HasReczDrs(A) Then
        B = McBrkI(A)
        O = AddDrs(O, B)
    End If
Next
McBrk = O
End Function

Friend Function McBrkI(A As Drs) As Drs
'Fm  McTRmk L Gpno MthLin IsRmk    ! For each gp, the front rmk lines are TopRmk, rmv them
'Ret McBrkI L Gpno MthLin IsRmk    ! Gpno is same
'           V Sfx Dcl LHS Expr
'           R1 R2 R3

Dim MthLin$, IsRmk, Dr, WAs%, Dry(), IxMthLin%, IxIsRmk%
AsgIx A, "MthLin IsRmk", IxMthLin, IxIsRmk
WAs = WMcBrkIWAs(A) ' The wdt of all As-V
For Each Dr In Itr(A.Dry)
    MthLin = Dr(IxMthLin)
    IsRmk = Dr(IxIsRmk)
    If IsRmk Then
        PushIAy Dr, McBrkIRmk(RmvFstChr(MthLin))
    Else
        PushIAy Dr, McBrkIDim(MthLin, WAs) ' The MthLin is a DimLin and it may have no assignment stmt.
    End If
    PushI Dry, Dr
Next
McBrkI = Drs(AddFF(A.Fny, "V Sfx Dcl LHS Expr R1 R2 R3"), Dry)
'Insp CSub, "Adding BrK (breaking mth lin into 6 columns", "Bef-break Aft-break", FmtDrs(A), FmtDrs(McBrkI):  Stop
End Function
Private Function WMcBrkIWAs%(A As Drs)
' Ret The wdt of all As-V
Dim B$(): B = StrColzDrs(DrswColEqSel(A, "IsRmk", False, "MthLin"), "MthLin")
Dim C$():
    Dim I
    For Each I In Itr(B)
        PushNonBlank C, Bef(I, " As ")
    Next
Dim D$(): D = RmvPfxzAy(C, "Dim ")
WMcBrkIWAs = WdtzAy(D)
'Thw CSub, "Find WAs: The wdt of all As-V", "Src B C D WAs", FmtDrs(A), B, C, D, WMcBrkIWAs
End Function

Friend Function McBrkIDim(DimLin$, WAs%) As Variant()
'Fm DimLin
'Fm WAs ! It is the Wdt of As-V, it may be zero, than the DimLin should not be an As-line.
'Ret Dcl V LHS Expr R1 R2 R3
Dim A As S3, L$, V$, Sfx$, Dcl$, LHS$, Expr$, R1$, R2$, R3$
A = BrkBet(DimLin, ":", "'")

'V, L <= A.A
    L = A.A
    If ShfT1(L) <> "Dim" Then Stop
    V = ShfNm(L): If V = "" Then Stop

'Sfx, L <= L
    If ShfTerm(L, "As") Then
        Sfx = " As " & ShfNm(L)
    Else
        Sfx = ShfTyChr(L)
    End If
    If ShfPfx(L, "()") Then Sfx = Sfx & "()"

'Dcl <= V, Sfx, WAs
    If HasPfx(Sfx, " As ") Then
        Dcl = V & Space(WAs - Len(V)) & Sfx
    Else
        Dcl = V & Sfx
    End If

'LHS, Expr <= A.B, V
    If A.B = "" Then
        LHS = V
        Expr = V
    Else
        With Brk(A.B, "=") 'Asume must have =, otherwise break
            LHS = .S1
            Expr = .S2
        End With
    End If
    
'Rmk
    AsgBrkBet A.C, vbPround, vbExcM, R1, R2, R3

McBrkIDim = Array(V, Sfx, Dcl, LHS, Expr, R1, R2, R3)
End Function

Friend Function McBrkIRmk(RmkLin$)
'Ret Dcl V LHS Expr R1 R2 R3"
Dim R1$, R2$, R3$
AsgBrkBet RmkLin, vbPround, vbExcM, R1, R2, R3
McBrkIRmk = Array("", "", "", "", "", R1, R2, R3)
End Function

Friend Function McFill(McBrk As Drs) As Drs
'A     : L Gpno MthLin IsRmk V Sfx Dcl LHS Expr R1 R2 R3
'Ret   : Add 6 Columns: F0 FDcl FLHS FExpr FR1 FR2 for each gp
If NoReczDrs(McBrk) Then Stop
Dim Gpno%(), O As Drs, IGpno, A As Drs, B As Drs, C As Drs
Gpno = AywDist(IntAyzDrsC(McBrk, "Gpno"))
For Each IGpno In Itr(Gpno)
    A = DrswColEq(McBrk, "Gpno", IGpno)
    B = McF0(A)
    C = AddColzFiller(B, "Dcl LHS Expr R1 R2")
    O = AddDrs(O, C)
Next
'BrwDrs O, ShwZer:=True: Stop
McFill = O
End Function
Friend Function DblEqDta() As Drs

End Function
Friend Function ODblEq() As Unt

End Function

Friend Function McDblEqRmk(Mc As Drs) As Drs
'Fm Mc L Lin
Dim Dr, Dry()
For Each Dr In Itr(Mc.Dry)
    If Left(LTrim(Dr(1)), 3) = "'==" Then PushI Dry, Dr
Next
McDblEqRmk.Fny = Mc.Fny
McDblEqRmk.Dry = Dry

'Insp CSub, "Finding DblEq LnoAy", "Mth-Cxt LnoAyDblEq", FmtDrs(Mc), FmtDrs(McDblEqRmk): Stop
End Function

Friend Function McDblEqLNewO(McDblEqRmk As Drs) As Drs
'Fm Mc L Lin
'Fm Mc L Lin
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
McDblEqLNewO = LNewO(Dry)

'Insp CSub, "Finding DblEq LnoAy", "Mth-Cxt LnoAyDblEq", FmtDrs(Mc), FmtDrs(McDblEqRmk): Stop
End Function
Friend Function McGp(McCln As Drs) As Drs
'Fm  Xl
'Fm  McCln L MthLin  #Mth-Cxt-Spc  ! no spc line & las lin
'Ret L Gpno MthLin #Mth-Cxt-Gpno ! with L in seq will be one gp
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
End Function

Friend Function McTRmk(McRmk As Drs) As Drs
' L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! For each gp, the front rmk lines are TopRmk, rmv them
Dim IGpno%, MaxGpno, A As Drs, B As Drs, O As Drs
MaxGpno = MaxzAy(IntAyzDrsC(McRmk, "Gpno"))
For IGpno = 1 To MaxGpno
    A = DrswColEq(McRmk, "Gpno", IGpno)
    B = McTRmkI(A)
    O = AddDrs(O, B)
Next
McTRmk = O
End Function
Friend Function McTRmkI(A As Drs) As Drs
' Fm  A     L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! All Gpno are eq
' Ret McTRmkI L Gpno MthLin IsRmk ! Rmk TopRmk
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
'Ret McRmk L Gpno MthLin IsRmk    #Mth-Cxt-isRmk
Dim Dr
For Each Dr In McGp.Dry
    PushI Dr, FstChr(LTrim(Dr(2))) = "'"
    Push McRmk.Dry, Dr
Next
McRmk.Fny = AddFF(McGp.Fny, "IsRmk")
End Function
Private Function WCmDta3(MthLin$) As Dictionary
Dim MthPm$: MthPm = BetBkt(MthLin)
Set WCmDta3 = New Dictionary
Dim P
For Each P In Itr(TrimAy(Split(MthPm, ",")))
    WCmDta3.Add TakNm(P), P
Next
'BrwDic WCmDta3: Stop
End Function
Function AddColzFFDry(A As Drs, FF$, NewDry()) As Drs
AddColzFFDry = Drs(AddFF(A.Fny, FF), NewDry)
End Function
Private Function WCmDta1(A As Drs) As Drs
Dim Dr, ODry()
For Each Dr In Itr(A.Dry)
    PushI Dr, BetBkt(Dr(1))
    PushI ODry, Dr
Next
WCmDta1 = AddColzFFDry(A, "Pm", ODry)
End Function
Private Function WCmDta4$(Pm$, Dic_V_Dcl1 As Dictionary)
Dim P, O$()
For Each P In Itr(TrimAy(Split(Pm, ",")))
    PushI O, Dic_V_Dcl1(P)
Next
WCmDta4 = JnCommaSpc(O)
End Function
Private Function WCmDta21$(Dcl)
Dim O$: O = Dcl
While HasSubStr(O, "  As ")
    O = Replace(O, "  As ", " As ")
Wend
WCmDta21 = O
End Function
Private Function WCmDta2(A As Drs) As Drs
'Fm A      V Expr Dcl Pm
'Ret WCmDta2 V Expr Dcl Pm Dcl1
Dim ODry(), Dr
For Each Dr In Itr(A.Dry)
    PushI Dr, WCmDta21(Dr(2))
    PushI ODry, Dr
Next
WCmDta2 = Drs(AddFF(A.Fny, "Dcl1"), ODry)
End Function
Private Sub WCmDta5(Dcl1$, OTyChr$, ORetAs$)
Dim A$: A = RmvNm(Dcl1)
Select Case True
Case HasPfx(A, " As ")
    OTyChr = ""
    ORetAs = A
Case HasSfx(A, "()")
    OTyChr = ""
    ORetAs = " As " & TyNmzTyChr(TakTyChr(A)) & "()"
Case Else
    OTyChr = A
    ORetAs = ""
End Select
End Sub

Friend Function CmStr$(CmNew$(), McBrk As Drs, CmMdy$, CmPfx$)
Dim A As Drs: A = DrswColEqSel(McBrk, "IsRmk", False, "V Sfx")
Dim CmNm, O$(), Sfx, V$
For Each CmNm In Itr(CmNew)
    V = RmvPfx(CmNm, CmPfx)
    Sfx = ValzColEq(A, "Sfx", "V", V): If Not IsStr(Sfx) Then Stop
    PushI O, ""
    PushI O, FmtQQ("? Function ??", CmMdy, CmNm, Sfx)
    PushI O, "End Function"
Next
CmStr = JnCrLf(O)
End Function

Friend Function CmlEpt(CmlMthRet As Drs, CmMdy$) As Drs
'Fm CmlMthRet : V Sfx Expr CmNm CallgPm CmlDclPm TyChr RetAs
'Ret CmlEpt   : V EptL
Dim Dr, Dry(), Nm$, Ty$, Pm$, Ret$, V$, EptL$, INm%, ITy%, IPm%, IRet%, IV%
AsgIx CmlMthRet, "CmNm TyChr CmlDclPm RetAs V", INm, ITy, IPm, IRet, IV
'BrwDrs CmlMthRet: Stop
For Each Dr In Itr(CmlMthRet.Dry)
    Nm = Dr(INm)
    Ty = Dr(ITy)
    Pm = Dr(IPm)
    Ret = Dr(IRet)
    V = Dr(IV)
    EptL = FmtQQ("? Function ??(?)?", CmMdy, Nm, Ty, Pm, Ret)
    PushI Dry, Array(V, EptL)
Next
CmlEpt = DrszFF("V EptL", Dry)
'BrwDrs CmlEpt: Stop
End Function

Friend Function CmlCallg(CmlFmMc As Drs, CmlCallgPfx$) As Drs
'Fm CmlFmMc : V Sfx Expr
'Ret CmlCallg : V Sfx Expr CmNm CallgPm
If CmlCallgPfx = "" Then Stop
Dim ODry()
    Dim Dr
    For Each Dr In Itr(CmlFmMc.Dry)
        Dim V$: V = Dr(0)
        Dim Sfx$: Sfx = Dr(1)
        Dim Expr$: Expr = Dr(2)
        Dim ExprPfx$: ExprPfx = CmlCallgPfx & V & "("
        If HasPfx(Expr, ExprPfx) Then
            Dim CmNm$
            If CmlCallgPfx = "F." Then
                CmNm = V
            Else
                CmNm = CmlCallgPfx & V
            End If
            Dim CallgPm$: CallgPm = BetBkt(Expr)
            PushI Dr, CmNm
            PushI Dr, CallgPm
            PushI ODry, Dr
        End If
    Next
CmlCallg = DrszFF("V Sfx Expr CmNm CallgPm", ODry)
'BrwDrs CmlCallg: Stop
End Function

Friend Function CmDta1(Xnc As Boolean, Cmo$(), McBrk As Drs, Ml$) As Drs
'Ret CmDta Chdn Pm
If Xnc Then Exit Function
Dim A As Drs: A = DrswColEqSel(McBrk, "IsRmk", False, "V Expr Dcl") 'V Expr Dcl
Dim B As Drs: B = WCmDta1(A) 'V Expr Dcl Pm
Dim C As Drs: C = WCmDta2(B) 'V Expr Dcl Pm Dcl1
Dim D As Drs: D = SelDrs(C, "V Dcl1")
Dim E As Dictionary: Set E = DiczDrsCC(D) 'Dic_V_Dcl1
Dim F As Dictionary: Set F = WCmDta3(Ml)    'Dic_MthPm_Dcl
Dim G As Dictionary: Set G = AddDic(E, F)
Dim CV$(), CPm$(), CDcl1$()
    AsgCol C, "V Pm Dcl1", CV, CPm, CDcl1
Dim ODry()
    Dim J%
    For J = 0 To UB(CV)
        Dim Dcl1$: Dcl1 = CDcl1(J)
        Dim TyChr$, RetAs$
            WCmDta5 Dcl1, TyChr, RetAs
        PushI ODry, Array(CV(J), TyChr, WCmDta4(CPm(J), G), RetAs)
    Next
CmDta1 = DrszFF("Chdn TyChr MthPm RetAs", ODry)
'BrwDrs CmDta: Stop
End Function

Private Function WCclsNm$(Md As CodeModule, Mln$)
WCclsNm = Mdn(Md) & "__" & Mln
End Function
Friend Function Ccls(NoSf As Boolean, Md As CodeModule, Mln$) As CodeModule
If NoSf Then Exit Function
Set Ccls = MdzPN(PjzM(Md), WCclsNm(Md, Mln))
End Function


Friend Function McNew(Mc As Drs, McDim As Drs) As String()
'Fm Mc L MthLin #Mth-Cxt.
'Fm McDim L MthLin DimLin #Mth-Cxt-newDimLin.
'Ret McNew L MthLin #Mth-Cxt-New.  ! After after

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

Friend Function McDim(McFill As Drs) As Drs
'Fm  Xl
'Fm  McFill : L Gpno MthLin IsRmk V Sfx Dcl LHS Expr R1 R2 R3
'Ret Mco    : L OldL NewL                              ! #Mth-Cxt-Oup.  Bld the new DimLin
'BrwDrs McFill: Stop
If NoReczDrs(McFill) Then Stop
Dim Dr, L&, MthLin$, IsRmk As Boolean, DimLin$, ODry()
For Each Dr In McFill.Dry
    L = Dr(0)
    MthLin = Dr(2)
    IsRmk = Dr(3)
    If IsRmk Then
        DimLin = McDimRmk(Dr)
    Else
        DimLin = McDimLin(Dr)
    End If
    PushI ODry, Array(L, DimLin, MthLin)
Next
McDim = DrszFF("L NewL OldL", ODry) 'Som may be OldL=NewL
'BrwDrs McDim: Stop
End Function

Friend Function McDimLin$(McFillDr)
'Fm  McFill : L Gpno MthLin IsRmk V Sfx Dcl LHS Expr R1 R2 R3 F0 FDcl FLHS FExpr FR1 FR2 ! Upd F0..FR1
'Ret DimLin$
'Brw LyzNNAv("L Gpno MthLin IsRmk Dcl V LHS Expr R1 R2 R3 F0 FDcl FLHS FExpr FR1 FR2", CvAv(McFillDr)): Stop
Dim L&, Gpno%, MthLin$, IsRmk, V$, Sfx$, Dcl$, LHS$, Expr$, R1$, R2$, R3$, F0%, FDcl%, FLHS%, FExpr%, FR1%, FR2%
AsgAp McFillDr, _
    L&, Gpno%, MthLin$, IsRmk, V$, Sfx$, Dcl$, LHS$, Expr$, R1$, R2$, R3$, F0%, FDcl%, FLHS%, FExpr%, FR1%, FR2%

Dim S0$:       S0 = Space(F0)
Dim SDcl$:   SDcl = Space(FDcl)
Dim SLHS$:   SLHS = Space(FLHS)
Dim SExpr$: SExpr = Space(FExpr)
Dim SR1$:     SR1 = Space(FR1)
Dim SR2$:     SR2 = Space(FR2)
'
Dim OD$:  OD = Dcl
Dim OL$:  OL = SDcl & SLHS & LHS
Dim OE$:  OE = Expr & SExpr
Dim OM$:  OM = XRmk(FR1, FR2, R1, R2, R3)
If OM = "" Then
    McDimLin = RTrim(FmtQQ("?Dim ?: ? = ?", S0, OD, OL, OE))
Else
    McDimLin = RTrim(FmtQQ("?Dim ?: ? = ? ' ?", S0, OD, OL, OE, OM))
End If
End Function
Friend Function McDimRmk$(McFillDr)
'Fm  McFill    : L Gpno MthLin IsRmk V Sfx Dcl LHS Expr R1 R2 R3 F0 FDcl FLHS FExpr FR1 FR2
'Ret McDimRmk$
Dim L&, Gpno%, MthLin$, IsRmk, V$, Sfx$, Dcl$, LHS$, Expr$, R1$, R2$, R3$, F0%, FDcl%, FLHS%, FExpr%, FR1%, FR2%
AsgAp McFillDr, _
    L&, Gpno%, MthLin$, IsRmk, V$, Sfx$, Dcl$, LHS$, Expr$, R1$, R2$, R3$, F0%, FDcl%, FLHS%, FExpr%, FR1%, FR2%

'
Dim OS$:  OS = Space(11 + FLHS + FDcl + FExpr)
Dim OM$:  OM = XRmk(FR1, FR2, R1, R2, R3)
McDimRmk = Trim(FmtQQ("'??", OS, OM))
End Function

Friend Function OEnsSf(NoSf As Boolean, Md As CodeModule, MthLno&, Mln$) As Unt
If NoSf Then Exit Function
Dim Lno&: Lno = SrcLnozNxt(Md, MthLno)
Dim OldL$: OldL = Md.Lines(Lno, 1)
Dim NewL$: NewL = "Static F As New " & WCclsNm(Md, Mln)
If OldL <> NewL Then
    Stop
    Md.ReplaceLine Lno, NewL
End If
End Function

Friend Function OEnsCcls(NoSf As Boolean, Md As CodeModule, Mln$) As Unt
'Ret #Ens-Chd-Cls.
If NoSf Then Exit Function
EnsCmpzPTN PjzM(Md), vbext_ct_ClassModule, WCclsNm(Md, Mln)
End Function

Friend Function CmPfx$(NoSf As Boolean, McLy$())
If Not NoSf Then Exit Function
CmPfx = StrValzCnstLy(McLy, "CmPfx")
End Function

Friend Function XRmk$(FR1%, FR2%, R1$, R2$, R3$)
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
XRmk = RTrim(A & B & C)
End Function


Friend Function CmLin(CmDta As Drs, CmMdy$) As String()
'Fm  CmDta  V TyChr Pm RetAs
'Ret CmLin MthLin                     ! MthLin is always a function
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
'Fm  Mth MthLin ! mth Mct
'Ret Mth MthLin ! Rmv non-Dim & non-Rmk lin
Dim Dr, Dry(), L$, Yes As Boolean
For Each Dr In Itr(Mc.Dry)
    L = Trim(Dr(1))
    Yes = False
    If HasPfx(L, "'") Then
        L = LTrim(RmvFstChr(L))
        Select Case True
        Case HasPfx(L, "If"), HasPfx(L, "Stop"), HasPfx(L, "Insp"), HasPfx(L, "=="), HasPfx(L, "Brw")
        Case Else: Yes = True
        End Select
    Else
        Yes = IsAlignableDim(L)
    End If
    If Yes Then PushI Dry, Dr
Next
McCln = Drs(Mc.Fny, Dry)
'BrwDrs McCln: Stop
End Function

Friend Function McF0(A As Drs) As Drs
'Fm   : L Gpno MthLin V Sfx Expr R1 R2 R3
'Ret  : Add  column-filler-F0  ! Sfx-0-of-McF0 is add-column-filler-0-F0
Dim Dr: Dr = A.Dry(0)
Dim F0%: F0 = WF0zDimLin(Dr(2))
Dim ODry()
For Each Dr In Itr(A.Dry)
    PushI Dr, F0
    PushI ODry, Dr
Next
McF0 = Drs(AddFF(A.Fny, "F0"), ODry)
End Function


Friend Function CrLNewO(CrJn As Drs) As Drs
End Function

Friend Function CrAct(CmEpt$(), Md As CodeModule) As Drs

End Function
Friend Function OUpdCr(CrDta As Drs) As Drs

End Function

Friend Function CmEpt(McBrk As Drs) As String()
'Ret ! It is from V and Expr=V
Dim A As Drs: A = DrswColEq(McBrk, "IsRmk", False)
Dim CV$(), CExpr$()
    AsgCol A, "V Expr", CV, CExpr
    If Si(CV) = 0 Then Exit Function

Dim V, J%
For Each V In CV
    If CExpr(J) = CV(J) Then
        PushI CmEpt, CV(J)  '<=====
    End If
    J = J + 1
Next
'Brw CmEpt: Stop
End Function

Friend Function WF0zDimLin%(DimLin)
If T1(DimLin) <> "Dim" Then Stop
Dim L$: L = LTrim(DimLin)
WF0zDimLin = Len(DimLin) - Len(L)
End Function


