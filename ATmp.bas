Attribute VB_Name = "ATmp"
Option Explicit
Option Compare Text
Private Sub Z_IsSngDimCol()
Dim L$
GoSub T0
Exit Sub
T0:
    L = "Dim IsDesAy() As Boolean: IsDesAy = XIsDesAy(Ay)"
    GoTo Tst
Tst:
    Act = IsLinSngDimColon(L)
    C
    Stop
    Return
End Sub

Function IsLinSngDimColon(L) As Boolean
'Ret true if L is Single-Dim-Colon: one V aft Dim and Colon aft DclSfx
Dim Lin$: Lin = L
If Not ShfDim(Lin) Then Exit Function
If ShfNm(Lin) = "" Then Exit Function
ShfBkt Lin
ShfDclSfx Lin
IsLinSngDimColon = FstChr(Lin) = ":"
End Function

Sub Z_IsLinSngDimColon()
Dim L
'GoSub T0
GoSub Z
Exit Sub
T0:
    L = "Dim A As Access.Application: Set A = DftAcs(Acs)"
    Ept = True
    GoTo Tst
Tst:
    Act = IsLinSngDimColon(L)
    If Act <> Ept Then Stop
    Return
Z:
    Dim A As New Aset
    For Each L In SrczP(CPj)
        L = Trim(L)
        If T1(L) = "Dim" Then
            Dim S$: S = IIf(IsLinSngDimColon(L), "1", "0")
            A.PushItm S & " " & L
        End If
    Next
    A.Srt.Vc
    Return
End Sub

Sub AlignMthDimzML(M As CodeModule, MthLno&, Optional Rpt As EmRpt, Optional IsUpdSelf As Boolean)
Dim D1 As Drs, D2 As Drs
'== Exit if parameter error ============================================================================================
Dim IsErPm As Boolean: IsErPm = XIsErPm(M, MthLno)
Dim OErPm:                      If IsErPm Then Exit Sub            ' #Is-Parameter-er. ! M-isnothg | MthLno<=0
Dim Ml$:                   Ml = ContLinzML(M, MthLno)
Dim IsUpd  As Boolean:  IsUpd = IsUpdzRpt(Rpt)
Dim MlNm$:               MlNm = Mthn(Ml)                           '                                        # Ml-Name.
Dim IsSelf As Boolean: IsSelf = XIsSelf(IsUpd, IsUpdSelf, M, MlNm)
Dim OErSelf:                    If IsSelf Then Exit Sub            ' #Is-Self-aligning-er.                                  ! Mdn<>'QIde...' & MlNm<>'AlignMthDimzML
Dim Mc     As Drs:         Mc = DMthCxtzML(M, MthLno)              ' L MthLin                               # Mth-Context.
                    If NoReczDrs(Mc) Then Exit Sub

'== Align DblEqRmk (De) ================================================================================================
'   When a rmk lin begins with '== or '-- or '.., expand it to 120 = or - or .
Dim De      As Drs:      De = XDe(Mc)                         ' L MthLin    # Dbl-Eq | Dbl-Dash | Dbl-Dot
Dim DeLNewO As Drs: DeLNewO = XDeLNewO(De)                    ' L NewL OldL
Dim OUpdDe:                   If IsUpd Then RplLin M, DeLNewO

'== Align Mth Cxt ======================================================================================================
Dim McCln As Drs: McCln = XMcCln(Mc) ' L MthLin # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align
If NoReczDrs(McCln) Then Exit Sub
Dim McGp    As Drs:    McGp = XMcGp(McCln)     ' L Gpno MthLin               ! with L in seq will be one gp
Dim McRmk   As Drs:   McRmk = XMcRmk(McGp)     ' L Gpno MthLin IsRmk         ! a column IsRmk is added
Dim McTRmk  As Drs:  McTRmk = XMcTRmk(McRmk)   ' L Gpno MthLin IsRmk         ! For each gp, the front rmk lines are TopRmk, rmv them
Dim McInsp  As Drs:  McInsp = XMcInsp(McTRmk)  ' L Gpno MthLin IsRmk         ! If las lin is rmk and is 'Insp, exl it.
Dim McVSfx  As Drs:  McVSfx = XMcVSfx(McInsp)  ' L Gpno MthLin IsRmk
                                               ' V Sfx Rst
Dim McDcl   As Drs:   McDcl = XMcDcl(McVSfx)   ' L Gpno MthLin IsRmk
                                               ' V Sfx Dcl Rst               ! Add Dcl from V & Sfx
Dim McLR    As Drs:    McLR = XMcLR(McDcl)     ' L Gpno MthLin IsRmk
                                               ' V Sfx Dcl LHS RHS Rst       ! Add LHS Expr from Rst
Dim McLREmp As Drs: McLREmp = XMcLREmp(McLR)   ' L Gpno MthLin IsRmk
                                               ' V Sfx Dcl LHS RHS Rst       ! set LHS & RHS to V = CmNm if V<>"" and LHS="" and RHS=""
Dim McR123  As Drs:  McR123 = XMcR123(McLREmp) ' L Gpno MthLin IsRmk
                                               ' V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
Dim McFill  As Drs:  McFill = XMcFill(McR123)  ' L Gpno MthLin IsRmk
                                               ' V Sfx Dcl LHS Expr
                                               ' F0 FSfx FExpr FR1 FR2       ! Adding F*
Dim McAlign As Drs: McAlign = XMcAlign(McFill) ' L Align                     ! Bld the new Align

                         D1 = DrseCeqC(McAlign, "MthLin Align")
                         D2 = DrszSel(D1, "L Align MthLin")
Dim McLNewO As Drs: McLNewO = LNewO(D2.Dry)
Dim OAlignCm:                 If IsUpd Then RplLin M, McLNewO

'== Gen Bs (Brw-Stmt) ==================================================================================================
Dim Bs      As Drs:              Bs = XBs(McCln)                                ' L BsLin ! Fst2Chr = '@
Dim Bs1     As Drs:             Bs1 = ColEqSel(McR123, "IsRmk", False, "V Sfx")
Dim Bs2     As Drs:             Bs2 = ColNe(Bs1, "V", "")
Dim VSfx    As Dictionary: Set VSfx = DiczDrsCC(Bs2)
Dim Mdn$:                       Mdn = MdnzM(M)
Dim BsLNewO As Drs:         BsLNewO = XBsLNewO(Bs, VSfx, Mdn, MlNm)
Dim OUpdBs:                           If IsUpd Then RplLin M, BsLNewO

'== Crt Chd-Mth (Cm)====================================================================================================
Dim CmSel  As Drs:   CmSel = ColEq(McR123, "IsRmk", False)
Dim CmV    As Drs:     CmV = DrszSel(CmSel, "V Sfx LHS RHS")   ' V Sfx LHS RHS
Dim CmDot  As Drs:   CmDot = XCmDot(CmV)                       ' V Sfx LHS RHS DotNm
Dim CmNmDD As Drs:  CmNmDD = XCmNmDD(CmDot, MlNm)              ' V Sfx LHS RHS DotNm CmNmDD       ! DD : cm is DblDash xxx_xxx
Dim CmNmX  As Drs:   CmNmX = XCmNmX(CmNmDD)                    ' V Sfx LHS RHS DotNm CmNmDD CmNmX ! X  : cm is Xmmm
Dim CmNm   As Drs:    CmNm = XCmNm(CmNmX)                      ' V Sfx LHS RHS CmNm
Dim CmEpt  As Drs:   CmEpt = ColNe(CmNm, "CmNm", "")           ' Sfx CmNm
Dim CmEptNm$():    CmEptNm = StrCol(CmEpt, "CmNm")             ' CmNm                             ! It is ept mth ny.  They will be used create new chd mth
Dim CmActNm$():    CmActNm = MthNyzM(M)                        ' CmNm                             ! It is from chd cls of given md
Dim CmNewNm$():    CmNewNm = MinusAy(CmEptNm, CmActNm)         ' CmNm                             ! The new ChdMthNy to be created.
Dim CmNew  As Drs:   CmNew = DrswIn(CmEpt, "CmNm", CmNewNm)
Dim CdNewCm$:      CdNewCm = XCdNewCm(CmNew)                   '                                  ! Cd to be append to M
Dim OCrtCm:                  If IsUpd Then ApdLines M, CdNewCm

'== Upd Chd-Mth-Lin (Cml) ==============================================================================================
'   If the calling pm has been changed, the chd-mth-lin will be updated.
Dim MlVSfx    As Drs:    MlVSfx = XMlVSfx(Ml)                              ' Ret V Sfx                           ! the MthLin's pm V Sfx
                             D1 = DrszSel(CmV, "V Sfx")
Dim CmlVSfx   As Drs:   CmlVSfx = DrszAdd(MlVSfx, D1)
Dim CmlPm     As Drs:     CmlPm = XCmlPm(CmEpt)                            ' V Sfx RHS CmNm Pm
Dim CmlDclPm  As Drs:  CmlDclPm = XCmlDclPm(CmlPm, CmlVSfx)                ' V Sfx RHS CmNm Pm DclPm             ! use [CmlVSfx] & [Pm] to bld [DclPm]
Dim CmlMthRet As Drs: CmlMthRet = XCmlMthRet(CmlDclPm)                     ' V Sfx RHS CmNm Pm DclPm TyChr RetAs
Dim CmlEpt    As Drs:    CmlEpt = XCmlEpt(CmlMthRet)                       ' V CmNm EptL
                             D1 = DMth(M)                                  ' L Mdy Ty Mthn MthLin
                             D1 = ColEq(D1, "Mdy", "Prv")
Dim CmlAct    As Drs:    CmlAct = DrszSelAs(D1, "L Mthn:CmNm MthLin:ActL") ' L CmNm ActL
Dim CmlJn     As Drs:     CmlJn = DrszJn(CmlEpt, CmlAct, "CmNm", "L ActL") ' V CmNm EptL L ActL                  ! som EptL & ActL may eq
                             D2 = DrseCeqC(CmlJn, "EptL ActL")             ' V CmNm EptL L ActL                  ! All EptL & ActL are diff
Dim CmlLNewO  As Drs:  CmlLNewO = DrszSelAs(D2, "L EptL:NewL ActL:OldL")   ' L NewL OldL
Dim OUpdCml:                      If IsUpd Then RplLin M, CmlLNewO

'== Rpl Mth-Brw (Mb)====================================================================================================
'   Des: Mth-Brw is a remarked Insp-stmt in each las lin of cm.  It insp all the inp oup
'   Lgc: Fnd-and-do MbLNewO
'        Fnd-and-do NewMb
'BrwDrs CmlEpt: Stop
Dim CmLis   As Drs:   CmLis = DrszSelAs(CmlEpt, "CmNm:Mthn EptL:MthLin") ' Mthn MthLin
Dim MbEpt   As Drs:   MbEpt = XMbEpt(CmLis, Mdn)                         ' Mthn MthLin MbStmt
Dim Cm$():               Cm = StrCol(CmLis, "Mthn")
Dim MbAct   As Drs:   MbAct = XMbAct(Cm, M)                              ' L Mthn OldL               ! OldL is MbStmt
Dim MbJn    As Drs:    MbJn = DrszJn(MbEpt, MbAct, "Mthn", "OldL L")     ' Mthn MthLin MbStmt OldL L
Dim MbSel   As Drs:   MbSel = DrszSelAs(MbJn, "L MbStmt:NewL OldL")      ' L NewL OldL
Dim MbLNewO As Drs: MbLNewO = DrseCeqC(MbSel, "NewL OldL")
Dim OUpdMb:                   If IsUpd Then RplLin M, MbLNewO

'== Crt Mth-Brw (Mb)====================================================================================================
                     D1 = LDrszJn(MbEpt, MbAct, "Mthn", "L", "HasAct") ' Mthn MthLin MbStmt L HasAct
                     D2 = ColEq(D1, "HasAct", False)                   ' Mthn MthLin MbStmt L HasAct
Dim MbNew As Drs: MbNew = DrszSelAs(D2, "Mthn MbStmt:NewL")
Dim OCrtMb:               If IsUpd Then XOCrtMb M, MbNew

'== Upd Chd-Rmk (Cr) ===================================================================================================

'-- Fnd CrEpt as Drs ---------------------------------------------------------------------------------------------------
'-- Fnd #Ept      : CmNm RmkLines  ! The expected chd mth rmk lines=====================================================
'   #Fm1-McR123   : V R1 R2 R3     ! The rmk lines of each variable
'   #Fm2-CmlDclPm : V Pm           ! The v is calling chd mth is using what pm
'   #Fm3-CmlEpt   : V CmNm         ! The v is calling what chd mth

'-- Fnd #WiRmk    : V R1 R2 R3     ! all rec will have at least 1 rmk (R1..3 som not blank).----------------------------
'   Fm  McR123
Dim CrSel   As Drs:   CrSel = DrszSel(McR123, "V R1 R2 R3") ' V R1 R2 R3
Dim CrSelV  As Drs:  CrSelV = DrszFillLasIfB(CrSel, "V")    ' V R1 R2 R3 ! Fill those blank col-V by las val
Dim FF$:                 FF = "R1 R2 R3"
Dim Sy$():               Sy = SyzAp("", "", "")
Dim CrWiRmk As Drs: CrWiRmk = DrseVy(CrSelV, FF, Sy)        ' V R1 R2 R3 ! Rmv those rec with all R1..3 are blank

'-- Fnd #Vpr1    : V P R1 R2 R3 IsRet   ! Each V | P having what rmk.  IsRet is True------------------------------------
'   Fm  CmlDclPm : V Pm                 ! The var calling chd mth is using what Pm
'   Fm  #WiRmk   : V R1 R2 R3           ! Each var having what Rmk
Dim CrVpm  As Drs:  CrVpm = DrszSelAs(CmlDclPm, "V Pm:P")    ' V P                ! each V is calling what Pm. Pm is SS.
Dim CrVp   As Drs:   CrVp = DrszSplitSS(CrVpm, "P")          ' V P                ! Brk P-SS into muli P
Dim CrVpr  As Drs:  CrVpr = DrszJn(CrVp, CrWiRmk, "P:V", FF) ' V P R1 R2 R3
Dim CrVpr1 As Drs: CrVpr1 = DrszAddCV(CrVpr, "IsRet", False) ' V P R1 R2 R3 IsRet ! All IsRet is FALSE

'-- Fnd #Ret     : V P R1 R2 R3 IsRet (P="" | IsRet=True) --------------------------------------------------------------
'   Fm  #WiRmk   : V R1 R2 R3
                       FF = "V P R1 R2 R3 IsRet"
Dim CrEmpP As Drs: CrEmpP = DrszSelAlwE(CrWiRmk, FF)               ' V P R1 R2 R3 IsRet ! All P & IsRet is empty
Dim CrRet  As Drs:  CrRet = DrszUpdCC(CrEmpP, "P IsRet", "", True) ' V P R1 R2 R3 IsRet ! All P is '' & IsRet is TRUE

'-- Fnd #RmkL : V Rmk           ! each V can map to CmNm ---------------------------------------------------------------
'   Fm  #Vrp1 : V P R1 R2 R3    ! P is pm | IsRet = false
'   Fm  #Ret :  V P R1 R2 R3    ! P is '' | IsRet = true
Dim CrMge As Drs: CrMge = DrszAdd(CrVpr1, CrRet)         ' V P R1 R2 R3 IsRet ! adding CrVpr & CrRet
Dim CrAli As Drs: CrAli = DrszAli(CrMge, "V", "P R1 R2") ' V P R1 R2 IsRet    ! P R1..3 are aligned (always hav sam len)
If False Then
    Dim CrInspAli As Drs:      CrInspAli = XCrInspAli(CrAli)
    Dim Db        As Database:    Set Db = TmpDb
    CrtTzDrs Db, "#InspAli", CrInspAli
    BrwT Db, "#InspAli"
    Stop
End If
Dim CrFst  As Drs:  CrFst = DrszAddFst(CrAli, "V P") ' V P R1 R2 R3 IsRet Fst     ! P R1..3 are aligned (always hav sam len & las chr is [.]
Dim CrRmk  As Drs:  CrRmk = XCrRmk(CrFst)            ' V P R1 R2 R3 IsRet Fst Rmk ! Bld Rmk from P R1 R2 R3 & Fst
Dim CrRmkL As Drs: CrRmkL = DrszSel(CrRmk, "V Rmk")  ' V Rmk

'-- Fnd #Ept : CmNm RmkLines :S1S2s ! each @CmNm should have waht @RmkLines---------------------------------------------
'   Fm  CmlEpt:V CmNm
'   Fm  RmkL  :
Dim CrVCm    As Drs:             CrVCm = DrszSel(CmlEpt, "V CmNm")     ' V CmNm
Dim CrV$():                        CrV = StrCol(CrVCm, "V")            ' V             ! all V have chd mth
Dim CrVRmkCm As Drs:          CrVRmkCm = DrswIn(CrRmkL, "V", CrV)      ' V Rmk         ! all V has chd mth
Dim CrVRmk   As S1S2s:          CrVRmk = S1S2szDrs(CrVRmkCm)           ' V Rmk
Dim CrVRmkS  As S1S2s:         CrVRmkS = AddS2Sfx(CrVRmk, " @@")
Dim CrVCmD   As Dictionary: Set CrVCmD = DiczDrsCC(CrVCm)
Dim CrEpt    As S1S2s:           CrEpt = MapS1(CrVRmkS, CrVCmD)        ' CmNm RmkLines

'== Upd Chd-Rmk (Cr) ===================================================================================================
'   If any of the calling pm's rmk is changed, the chd-mth-rmk will be updated
Dim CrAct As S1S2s: CrAct = MthRmkzNy(M, Cm)
Dim CrChg As S1S2s: CrChg = XCrChg(CrEpt, CrAct)           ' CmNm RmkLines ! Only those need to change
Dim OUpdCr:                 If IsUpd Then XOUpdCr CrChg, M

'== Rpt <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Rpt:
If IsRptzRpt(Rpt) Then
    Insp CSub, "Changes", _
        "EmRpt DblEqRmk Align BrwStmt " & _
        "Crt-Chd-Mth Rpl-Mth-Brw Crt-Mth-Brw Upd-Chd-MthLin Rfh-Chd-Mth-Rmk", _
        StrzRpt(Rpt), _
        LinzDrs(DeLNewO), LinzDrs(McLNewO), LinzDrs(BsLNewO), _
        CdNewCm, LinzDrs(MbLNewO), LinzDrs(MbNew), LinzDrs(CmlLNewO), FmtS1S2s(CrChg)
End If
'Insp CSub, "Cr", "CrVpr CrS1S2s", LinzDrs(CrVpr), FmtS1S2s(CrS1S2s): Stop
End Sub
Private Function XCrInspAli(CrAli As Drs) As Drs
'Fm CrAli : V P R1 R2 IsRet ! P R1..3 are aligned (always hav sam len) @@
'Fm CrAli: V P R1 R2 R3 IsRet
'Ret     : ..                 WP W1 W2 W3
If IsNeFF(CrAli, "V P R1 R2 R3 IsRet") Then Stop
Dim Dr, Dry(): For Each Dr In Itr(CrAli.Dry)
    Dim WP%: WP = Len(Dr(1))
    Dim W1%: W1 = Len(Dr(2))
    Dim W2%: W2 = Len(Dr(3))
    Dim W3%: W3 = Len(Dr(4))
    PushIAy Dr, Array(WP, W1, W2, W3)
    PushI Dry, Dr
Next
XCrInspAli = DrszAddFF(CrAli, "WP W1 W2 W3", Dry)
'Insp "QIde_B_AlignMth.XCrInspAli", "Inspect", "Oup(XCrInspAli) CrAli", LinzDrs(XCrInspAli), LinzDrs(CrAli): Stop
End Function

Sub AlignMthDim(Optional Rpt As EmRpt)
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Sub
AlignMthDimzML M, CMthLno, Rpt:=Rpt
End Sub

Private Function WF0(A As Drs) As Drs
Dim IxMthLin%: IxMthLin = IxzAy(A.Fny, "MthLin")
Dim IxV%: IxV = IxzAy(A.Fny, "V")
Dim Dr:              Dr = A.Dry(0)
Dim V$:               V = Dr(IxV)
Dim MthLin$:     MthLin = Dr(IxMthLin)
Dim T$:               T = LTrim(MthLin)
Dim F0%:             F0 = IIf(V = "", 0, Len(MthLin) - Len(T))
               WF0 = DrszAddCV(A, "F0", F0)
End Function

Private Function XMcFill(McR123 As Drs) As Drs
'Fm McR123 : L Gpno MthLin IsRmk
'            V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
'Ret       : L Gpno MthLin IsRmk
'            V Sfx Dcl LHS Expr
'            F0 FSfx FExpr FR1 FR2       ! Adding F* @@
Dim Gpno%(): Gpno = AywDist(IntCol(McR123, "Gpno"))
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim A As Drs: A = ColEq(McR123, "Gpno", IGpno)
    Dim B As Drs: B = WF0(A)
    Dim C As Drs: C = DrszAddFiller(B, "Dcl LHS RHS R1 R2")
    Dim O As Drs: O = DrszAdd(O, C)
Next
XMcFill = O
'Insp "QIde_B_AlignMth.XMcFill", "Inspect", "Oup(XMcFill) McR123", LinzDrs(XMcFill), LinzDrs(McR123): Stop
End Function

Private Function XDe(Mc As Drs) As Drs
'Fm Mc : L MthLin # Mth-Context.
'Ret   : L MthLin # Dbl-Eq | Dbl-Dash | Dbl-Dot @@
Dim Dr, Dry(): For Each Dr In Itr(Mc.Dry)
    Dim L$: L = LTrim(Dr(1))
    If FstChr(L) = "'" Then
        L = Left(RmvFstChr(L), 2)
        Select Case L
        Case "==", "--", "..": PushI Dry, Dr
        End Select
    End If
Next
XDe.Fny = Mc.Fny
XDe.Dry = Dry
'Insp "QIde_B_AlignMth.XDe", "Inspect", "Oup(XDe) Mc", LinzDrs(XDe), LinzDrs(Mc): Stop
End Function

Private Function XDeLNewO(De As Drs) As Drs
'Fm De : L MthLin    # Dbl-Eq | Dbl-Dash | Dbl-Dot
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
XDeLNewO = LNewO(Dry)
'Insp "QIde_B_AlignMth.XDeLNewO", "Inspect", "Oup(XDeLNewO) De", LinzDrs(XDeLNewO), LinzDrs(De): Stop
End Function


Private Function XMcGp(McCln As Drs) As Drs
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
XMcGp = DrszFF("L Gpno MthLin", Dry)
'Insp "QIde_B_AlignMth.XMcGp", "Inspect", "Oup(XMcGp) McCln", LinzDrs(XMcGp), LinzDrs(McCln): Stop
End Function

Private Function XMcTRmk(McRmk As Drs) As Drs
'Fm McRmk : L Gpno MthLin IsRmk ! a column IsRmk is added
'Ret      : L Gpno MthLin IsRmk ! For each gp, the front rmk lines are TopRmk, rmv them @@
Dim IGpno%, MaxGpno, A As Drs, B As Drs, O As Drs
MaxGpno = MaxzAy(IntCol(McRmk, "Gpno"))
For IGpno = 1 To MaxGpno
    A = ColEq(McRmk, "Gpno", IGpno)
    B = XMcTRmkI(A)
    O = DrszAdd(O, B)
Next
XMcTRmk = O
'Insp "QIde_B_AlignMth.XMcTRmk", "Inspect", "Oup(XMcTRmk) McRmk", LinzDrs(XMcTRmk), LinzDrs(McRmk): Stop
End Function


Private Function XMcInsp(McTRmk As Drs) As Drs
'Fm McTRmk : L Gpno MthLin IsRmk ! For each gp, the front rmk lines are TopRmk, rmv them
'Ret       : L Gpno MthLin IsRmk ! If las lin is rmk and is 'Insp, exl it. @@
XMcInsp = McTRmk
If NoReczDrs(McTRmk) Then Exit Function
Dim Dr: Dr = LasEle(McTRmk.Dry)
Dim IxMthLin%: IxMthLin = IxzAy(McTRmk.Fny, "MthLin")
Dim L$: L = Dr(IxMthLin)
If IsLinVbRmk(L) Then
    Dim A$: A = Left(LTrim(RmvFstChr(LTrim(L))), 4)
    If A = "Insp" Then
        Pop XMcInsp.Dry
    End If
End If
'Insp "QIde_B_AlignMth.XMcInsp", "Inspect", "Oup(XMcInsp) McTRmk", LinzDrs(XMcInsp), LinzDrs(McTRmk): Stop
End Function

Private Function XMcTRmkI(A As Drs) As Drs
' Fm  A :    L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! All Gpno are eq
' Ret : L Gpno MthLin IsRmk ! Rmk TopRmk
XMcTRmkI.Fny = A.Fny
Dim J%
    Dim Dr: For Each Dr In Itr(A.Dry)
        If Not Dr(3) Then GoTo Fnd
        J = J + 1
    Next
    Exit Function
Fnd:
    For J = J To UB(A.Dry)
        PushI XMcTRmkI.Dry, A.Dry(J)
    Next
End Function


Private Function XMcRmk(McGp As Drs) As Drs
'Fm McGp : L Gpno MthLin       ! with L in seq will be one gp
'Ret     : L Gpno MthLin IsRmk ! a column IsRmk is added @@
Dim Dr: For Each Dr In Itr(McGp.Dry)
    PushI Dr, FstChr(LTrim(Dr(2))) = "'"
    Push XMcRmk.Dry, Dr
Next
XMcRmk.Fny = FnyzAddFF(McGp.Fny, "IsRmk")
'Insp "QIde_B_AlignMth.XMcRmk", "Inspect", "Oup(XMcRmk) McGp", LinzDrs(XMcRmk), LinzDrs(McGp): Stop
End Function

Private Function XCdNewCm$(CmNew As Drs)
'Ret :  ! Cd to be append to M @@
Dim A As Drs: A = DrszSel(CmNew, "Sfx CmNm")
If IsNeFF(CmNew, "V Sfx LHS RHS CmNm") Then Stop
Dim Dr, O$(): For Each Dr In Itr(A.Dry)
    Dim Sfx$:   Sfx = Dr(0)
    Dim CmNm$: CmNm = Dr(1)
    Dim TyChr$: TyChr = TyChrzDclSfx(Sfx)
    Dim RetAs$: RetAs = RetAszDclSfx(Sfx)
    PushI O, ""
    PushI O, FmtQQ("Private Function ??()?", CmNm, TyChr, RetAs)
    PushI O, "End Function"
Next
XCdNewCm = JnCrLf(O)
'Insp "QIde_B_AlignMth.XCdNewCm", "Inspect", "Oup(XCdNewCm) CmNew", XCdNewCm, LinzDrs(CmNew): Stop
End Function


Private Function XCmlEpt(CmlMthRet As Drs) As Drs
'Fm CmlMthRet : V Sfx RHS CmNm Pm DclPm TyChr RetAs
'Ret          : V CmNm EptL
'               L Mdy Ty Mthn MthLin @@
Dim Dr, Dry(), Nm$, Ty$, Pm$, Ret$, V$, EptL$, INm%, ITy%, IPm%, IRet%, IV%
AsgIx CmlMthRet, "CmNm TyChr DclPm RetAs V", INm, ITy, IPm, IRet, IV
'BrwDrs CmlMthRet: Stop
For Each Dr In Itr(CmlMthRet.Dry)
    Nm = Dr(INm)
    Ty = Dr(ITy)
    Pm = Dr(IPm)
    Ret = Dr(IRet)
    V = Dr(IV)
    EptL = FmtQQ("Private Function ??(?)?", Nm, Ty, Pm, Ret)
    PushI Dry, Array(V, Nm, EptL)
Next
XCmlEpt = DrszFF("V CmNm EptL", Dry)
'BrwDrs CmlEpt: Stop
'Insp "QIde_B_AlignMth.XCmlEpt", "Inspect", "Oup(XCmlEpt) CmlMthRet", LinzDrs(XCmlEpt), LinzDrs(CmlMthRet): Stop
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

Private Function XCmlPm(CmEpt As Drs) As Drs
'Fm CmEpt : Sfx CmNm
'Ret      : V Sfx RHS CmNm Pm @@
Dim IxRHS%, IxCmNm%: AsgIx CmEpt, "RHS", IxRHS, IxCmNm
Dim Dr, ODry(): For Each Dr In Itr(CmEpt.Dry)
    Dim RHS$: RHS = Dr(IxRHS)
    Dim CmNm$: CmNm = Dr(IxCmNm)
    PushI Dr, XPm(RHS, CmNm)
    PushI ODry, Dr
Next
XCmlPm = DrszAddFF(CmEpt, "Pm", ODry)
'Insp "QIde_B_AlignMth.XCmlPm", "Inspect", "Oup(XCmlPm) CmEpt", LinzDrs(XCmlPm), LinzDrs(CmEpt): Stop
End Function

Private Function XMcNew(Mc As Drs, McDim As Drs) As String()
'BrwDrs2 Mc, Mc, NN:="Mc McDim", Tit:="Use McDim to Upd Mc to become NewL": Stop
If JnSpc(McDim.Fny) <> "L OldL NewL" Then Stop
Dim A As Drs: A = DrszSel(McDim, "L NewL")
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
XMcNew = O
End Function

Private Function XBsLNewO(Bs As Drs, VSfx As Dictionary, Mdn$, MlNm$) As Drs
'Fm Bs   : L BsLin            ! Fst2Chr = '@
'Fm MlNm :         # Ml-Name. @@
Dim Dr, Dry(), S$, Lin$, L&
For Each Dr In Itr(Bs.Dry)
    L = Dr(0)
    Lin = Dr(1)
    S = WBsStmt(Lin, VSfx, Mdn, MlNm)
    PushI Dry, Array(L, S, Lin)
Next
XBsLNewO = DrszFF("L NewL OldL", Dry)
'Insp "QIde_B_AlignMth.XBsLNewO", "Inspect", "Oup(XBsLNewO) Bs VSfx Mdn MlNm", LinzDrs(XBsLNewO), LinzDrs(Bs), VSfx, Mdn, MlNm: Stop
End Function

Private Function WBsStmt$(BsLin, VSfx As Dictionary, Mdn$, MlNm$)
If Left(BsLin, 2) <> "'@" Then Thw CSub, "BsLin is always begin with '@", "BsLin", BsLin
Dim NN$: NN = Trim(RmvPfx(BsLin, "'@"))
Dim E$: E = InspExprLis(NN, VSfx)
WBsStmt = InspStmt(NN, E, Mdn, MlNm)
End Function

Private Function XOCrtMb(M As CodeModule, MbNew As Drs)
'Fm NewMb : Cm NewMbL @@
Dim Dr: For Each Dr In Itr(MbNew.Dry)
    Dim Lin$: Lin = Dr(1)
    
    Dim Mthn$:  Mthn = Dr(0)
    Dim MLno&: MLno = MthLnozMM(M, Mthn)
    Dim ELno&: ELno = MthELno(M, MLno)
    Dim L&: L = ELno
    
    M.InsertLines L, Lin
Next
End Function


Private Function XMbEpt(CmLis As Drs, Mdn$) As Drs
'Fm CmLis : Mthn MthLin
'Ret      : Mthn MthLin MbStmt @@
Dim Dr, Dry()
For Each Dr In Itr(CmLis.Dry)
    Dim MthLin$: MthLin = LasEle(Dr)
    Dim MbStmt$: MbStmt = "'" & InspMthStmt(MthLin, Mdn) & ": Stop"
    PushI Dr, MbStmt
    PushI Dry, Dr
Next
XMbEpt = DrszAddFF(CmLis, "MbStmt", Dry)
'Insp "QIde_B_AlignMth.XMbEpt", "Inspect", "Oup(XMbEpt) CmLis Mdn", LinzDrs(XMbEpt), LinzDrs(CmLis), Mdn: Stop
End Function

Private Function XMbAct(Cm$(), M As CodeModule) As Drs
'Ret : L Mthn OldL ! OldL is MbStmt @@
Dim A As Drs: A = DMthe(M)             ' L E CmMdy Ty Mthn MthLin
Dim B As Drs: B = DrswIn(A, "Mthn", Cm)
Dim Dr, Dry(): For Each Dr In Itr(B.Dry)
    Dim E&:           E = Dr(1)
    Dim L&:           L = E - 1          ' ! The Lno of MbStmt
    Dim Mthn$:     Mthn = Dr(4)
    Dim MbStmt$: MbStmt = M.Lines(L, 1)
    Select Case True
    Case HasPfx(MbStmt, "'Insp "), HasPfx(MbStmt, "Insp ")
        PushI Dry, Array(L, Mthn, MbStmt)
    End Select
Next
XMbAct = DrszFF("L Mthn OldL", Dry)
'BrwDrs MbAct: Stop
'Insp "QIde_B_AlignMth.XMbAct", "Inspect", "Oup(XMbAct) Cm M", LinzDrs(XMbAct), Cm, Mdn(M): Stop
End Function


Private Function XBs(McCln As Drs) As Drs
'Fm McCln : L MthLin # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align
'Ret      : L BsLin            ! Fst2Chr = '@ @@
Dim Dr, Dry()
For Each Dr In Itr(McCln.Dry)
    If HasPfx(Dr(1), "'@") Then PushI Dry, Dr
Next
XBs = DrszFF("L BsLin", Dry)
'Insp "QIde_B_AlignMth.XBs", "Inspect", "Oup(XBs) McCln", LinzDrs(XBs), LinzDrs(McCln): Stop
End Function
Private Function XMcAlign(McFill As Drs) As Drs
'Fm McFill : L Gpno MthLin IsRmk
'            V Sfx Dcl LHS Expr
'            F0 FSfx FExpr FR1 FR2 ! Adding F*
'Ret       : L Align               ! Bld the new Align @@
If NoReczDrs(McFill) Then Stop
Dim A As Drs: A = DrszSel(McFill, "L Gpno MthLin Dcl LHS RHS R1 R2 R3 F0 FDcl FLHS FRHS FR1 FR2")
Dim Gpno: Gpno = DistCol(McFill, "Gpno")
Dim IGpno: For Each IGpno In Itr(Gpno)
    Dim B As Drs: B = ColEq(A, "Gpno", IGpno)
    Dim C As Drs: C = WAlign(B)
    Dim O As Drs: O = DrszAdd(O, C)
Next
XMcAlign = O
'Insp "QIde_B_AlignMth.XMcAlign", "Inspect", "Oup(XMcAlign) McFill", LinzDrs(XMcAlign), LinzDrs(McFill): Stop
End Function


Private Function XWRmk$(FR1%, FR2%, R1$, R2$, R3$)
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
XWRmk = RTrim(A & B & C)
End Function

Private Function XCmLin(CmDta As Drs, CmMdy$) As String()
'Fm  CmDta  V TyChr Pm RetAs
'Ret CmLin MthLin                     ! MthLin is always a function @@
Dim Dr, CmNm$, MthPfx$, Pm$, TyChr$, RetAs$
For Each Dr In Itr(CmDta.Dry)
    CmNm = Dr(0)
    TyChr = Dr(1)
    Pm = Dr(2)
    RetAs = Dr(4)
    Dim MthLin$: MthLin = FmtQQ("? Function ???(?)?", CmMdy, CmNm, TyChr, Pm, RetAs)
    PushI XCmLin, MthLin
Next
End Function

Private Function XIsSelf(IsUpd As Boolean, IsUpdSelf As Boolean, M As CodeModule, MlNm$) As Boolean
'Fm MlNm :  # Ml-Name. @@
If Not IsUpd Then Exit Function
If IsUpdSelf Then Exit Function
Dim O As Boolean
O = Mdn(M) = "QIde_B_AlignMth" And MlNm = "AlignMthDimzML"
If O Then Inf CSub, "Self aligning"
XIsSelf = O
'Insp "QIde_B_AlignMth.XIsSelf", "Inspect", "Oup(XIsSelf) IsUpd IsUpdSelf M MlNm", XIsSelf, IsUpd, IsUpdSelf, Mdn(M), MlNm: Stop
End Function

Private Function XIsErPm(M As CodeModule, MthLno&) As Boolean
XIsErPm = True
If IsNothing(M) Then Debug.Print "Md is nothing": Exit Function
If MthLno <= 0 Then Debug.Print "MthLno <= 0": Exit Function
XIsErPm = False
'Insp "QIde_B_AlignMth.XIsErPm", "Inspect", "Oup(XIsErPm) M MthLno", XIsErPm, Mdn(M), MthLno: Stop
End Function

Private Function XMcCln(Mc As Drs) As Drs
'Fm Mc : L MthLin # Mth-Context.
'Ret   : L MthLin # Mc-Cln.      ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align @@
Dim Dr, Dry(), L$, Yes As Boolean
Dim PfxAy$(): PfxAy = SyzSS("If Stop Insp == -- Brw")
For Each Dr In Itr(Mc.Dry)
    L = Trim(Dr(1))
    Yes = False
    Select Case True
    Case HasPfx(L, "'")
        L = LTrim(RmvFstChr(L))
        Select Case True
        Case HasPfxAy(L, PfxAy)
        Case Else: Yes = True
        End Select
    Case IsLinSngDimColon(L), IsLinAsg(L)
        Yes = True
    End Select
    If Yes Then PushI Dry, Dr
Next
XMcCln = Drs(Mc.Fny, Dry)
'BrwDrs McCln: Stop
'Insp "QIde_B_AlignMth.XMcCln", "Inspect", "Oup(XMcCln) Mc", LinzDrs(XMcCln), LinzDrs(Mc): Stop
End Function

Private Sub XOUpdCr(CrChg As S1S2s, M As CodeModule)
Dim J&, Ay() As S1S2
Ay = CrChg.Ay
For J = 0 To CrChg.N - 1
    With Ay(J)
    EnsMthRmk M, .S1, .S2
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


Private Function XCrChg(CrEpt As S1S2s, CrAct As S1S2s) As S1S2s
'Fm CrEpt : CmNm RmkLines
'Ret      : CmNm RmkLines ! Only those need to change @@
Dim J&, Ay() As S1S2, O As S1S2s
Ay = CrEpt.Ay
For J = 0 To CrEpt.N - 1
    With Ay(J)
    Dim A As StrOpt: A = FstS2(.S1, CrAct)
    If A.Som Then
        If .S2 <> A.Str Then
            PushS1S2 O, Ay(J)
        End If
    Else
        PushS1S2 O, Ay(J)
    End If
    End With
Next
XCrChg = O
'Insp "QIde_B_AlignMth.XCrChg", "Inspect", "Oup(XCrChg) CrEpt CrAct", FmtS1S2s(XCrChg), FmtS1S2s(CrEpt), FmtS1S2s(CrAct): Stop
End Function

Private Function XCmNmDD(CmDot As Drs, MlNm$) As Drs
'Fm CmDot : V Sfx LHS RHS DotNm
'Fm MlNm  :                            # Ml-Name.
'Ret      : V Sfx LHS RHS DotNm CmNmDD            ! DD : cm is DblDash xxx_xxx @@
'Fm CmDot : V Sfx RHS DotNm
'Ret      : V Sfx RHS DotNm CmNmDD
Dim MlNmDD$: MlNmDD = MlNm & "__"
Dim Dr, Dry(): For Each Dr In Itr(CmDot.Dry)
    Dim DotNm$:       DotNm = Dr(3)
    Dim V$:               V = Dr(0)
    Dim Hit As Boolean: Hit = DotNm = MlNmDD & V
    Dim DD$:             DD = IIf(Hit, DotNm, "")
                              PushI Dr, DD
                              PushI Dry, Dr
Next
XCmNmDD = DrszAddFF(CmDot, "CmNmDD", Dry)
'Insp "QIde_B_AlignMth.XCmNmDD", "Inspect", "Oup(XCmNmDD) CmDot MlNm", LinzDrs(XCmNmDD), LinzDrs(CmDot), MlNm: Stop
End Function

Private Function XCmNmX(CmNmDD As Drs) As Drs
'Fm CmNmDD : V Sfx LHS RHS DotNm CmNmDD       ! DD : cm is DblDash xxx_xxx
'Ret       : V Sfx LHS RHS DotNm CmNmDD CmNmX ! X  : cm is Xmmm @@
'Fm CmNmDD : V Sfx LHS RHS DotNm CmNmDD
'Ret       : V Sfx LHS RHS DotNm CmNmDD CmNmX
If IsNeFF(CmNmDD, "V Sfx LHS RHS DotNm CmNmDD") Then Stop
Dim Dr, Dry(): For Each Dr In Itr(CmNmDD.Dry)
    Dim DotNm$:       DotNm = Dr(4)
    Dim V$:               V = Dr(0)
    Dim Hit As Boolean: Hit = DotNm = "X" & V
    Dim X$:               X = IIf(Hit, DotNm, "")
                              PushI Dr, X
                              PushI Dry, Dr
Next
XCmNmX = DrszAddFF(CmNmDD, "CmNmX", Dry)
'Insp "QIde_B_AlignMth.XCmNmX", "Inspect", "Oup(XCmNmX) CmNmDD", LinzDrs(XCmNmX), LinzDrs(CmNmDD): Stop
End Function

Private Function XCmDot(CmV As Drs) As Drs
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS DotNm @@
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS DotNm
If IsNeFF(CmV, "V Sfx LHS RHS") Then Stop
Dim Dr, Dry(): For Each Dr In Itr(CmV.Dry)
    Dim RHS$: RHS = Dr(3)
    PushI Dr, TakDotNm(RHS)
    PushI Dry, Dr
Next
XCmDot = DrszAddFF(CmV, "DotNm", Dry)
'BrwDrs CmDot: Stop
'Insp "QIde_B_AlignMth.XCmDot", "Inspect", "Oup(XCmDot) CmV", LinzDrs(XCmDot), LinzDrs(CmV): Stop
End Function

Private Function XCmNm(CmNmX As Drs) As Drs
'Fm CmNmX : V Sfx LHS RHS DotNm CmNmDD CmNmX ! X  : cm is Xmmm
'Ret      : V Sfx LHS RHS CmNm @@
'Fm CmNmX : V Sfx LHS RHS DotNm CmNmDD CmNmX
'Ret      : V Sfx LHS RHS CmNm
Dim IxSfx%, IxDD%, IxX%, IxV%, IxR%, IxL%
AsgIx CmNmX, "V Sfx LHS RHS CmNmDD CmNmX", IxV, IxSfx, IxL, IxR, IxDD, IxX
Dim Dr, Dry(): For Each Dr In Itr(CmNmX.Dry)
    Dim V$:     V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxL)
    Dim RHS$: RHS = Dr(IxR)
    Dim Sfx$: Sfx = Dr(IxSfx)
    Dim DD$:   DD = Dr(IxDD)
    Dim X$:     X = Dr(IxX)
    Dim Nm$: Nm = ""
    Select Case True
    Case DD <> "": Nm = DD
    Case X <> "":  Nm = X
    End Select
    If Nm <> "" Then PushI Dry, Array(V, Sfx, LHS, RHS, Nm)
Next
XCmNm = DrszFF("V Sfx LHS RHS CmNm", Dry)
'Insp "QIde_B_AlignMth.XCmNm", "Inspect", "Oup(XCmNm) CmNmX", LinzDrs(XCmNm), LinzDrs(CmNmX): Stop
End Function


Private Function XCmLHS(CmV As Drs) As Drs
'Fm CmV : V Sfx LHS RHS
'Ret    : V Sfx LHS RHS ! where V & ' = ' = LHS @@
Dim IxV%, IxLHS%
AsgIx CmV, "V LHS", IxV, IxLHS
Dim Dr, Dry(): For Each Dr In Itr(CmV.Dry)
    Dim V$:     V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxLHS)
    If V & " = " = LHS Then Push XCmLHS.Dry, Dr
Next
XCmLHS.Fny = CmV.Fny
End Function

Private Function XMcR123(McLREmp As Drs) As Drs
'Fm McLREmp : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS RHS Rst       ! set LHS & RHS to V = CmNm if V<>"" and LHS="" and RHS=""
'Ret        : L Gpno MthLin IsRmk
'             V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst @@
Dim Dr: For Each Dr In Itr(McLREmp.Dry)
    Dim Rst$:          Rst = LasEle(Dr)
    If FstChr(Rst) <> "'" Then Stop
    Dim R1$, R2$, R3$:       AsgBrkBet Rst, "#", "!", R1, R2, R3
                        R1 = Trim(RmvFstChr(R1))
                             If R2 <> "" Then R2 = " # " & R2
                             If R3 <> "" Then R3 = " ! " & R3
                             If R1 <> "" Or R2 <> "" Or R3 <> "" Then R1 = " ' " & R1
    Dim Dry():               PushI Dry, AyzAdd(AyeLasEle(Dr), Array(R1, R2, R3))
Next
Dim Fny$(): Fny = AyzAdd(AyeLasEle(McLREmp.Fny), SyzSS("R1 R2 R3"))
XMcR123 = Drs(Fny, Dry)
'Insp "QIde_B_AlignMth.XMcR123", "Inspect", "Oup(XMcR123) McLREmp", LinzDrs(XMcR123), LinzDrs(McLREmp): Stop
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



Private Function XMcLR(McDcl As Drs) As Drs
'Fm McDcl : L Gpno MthLin IsRmk
'           V Sfx Dcl Rst         ! Add Dcl from V & Sfx
'Ret      : L Gpno MthLin IsRmk
'           V Sfx Dcl LHS RHS Rst ! Add LHS Expr from Rst @@
Dim Dr: For Each Dr In Itr(McDcl.Dry)
    Dim L$:     L = Pop(Dr)
    Dim Rst$: Rst = L
    Dim LHS$, Expr$:       AsgAp ShfLRHS(Rst), LHS, Expr
    If Rst <> "" Then
        If Not IsLinVbRmk(Rst) Then Stop
    End If
                      Dr = AyzAdd(Dr, Array(LHS, Expr, Rst))
    Dim Dry():             PushI Dry(), Dr
Next
Dim Fny$(): Fny = AyeLasEle(McDcl.Fny)
Fny = AyzAdd(Fny, SyzSS("LHS RHS Rst"))
XMcLR = Drs(Fny, Dry)
'Insp "QIde_B_AlignMth.XMcLR", "Inspect", "Oup(XMcLR) McDcl", LinzDrs(XMcLR), LinzDrs(McDcl): Stop
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
'Debug.Print XDclzV; WAs; QteSq(Sfx); "<": Stop
End Function
Private Function XMcLREmp(McLR As Drs) As Drs
'Fm McLR : L Gpno MthLin IsRmk
'          V Sfx Dcl LHS RHS Rst ! Add LHS Expr from Rst
'Ret     : L Gpno MthLin IsRmk
'          V Sfx Dcl LHS RHS Rst ! set LHS & RHS to V = CmNm if V<>"" and LHS="" and RHS="" @@
'Ret     : L Gpno MthLin IsRmk
'          V Sfx Dcl LHS RHS Rst ! for V<>"", LHS="" and RHS="", set LHS = V and RHS = X@V
Dim IxV%, IxL%, IxR%: AsgIx McLR, "V LHS RHS", IxV, IxL, IxR
Dim Dr, Dry(): For Each Dr In Itr(McLR.Dry)
    Dim V$: V = Dr(IxV)
    Dim LHS$: LHS = Dr(IxL)
    Dim RHS$: RHS = Dr(IxR)
    If V <> "" Then
        If LHS = "" Then
            If RHS = "" Then
                Dr(IxL) = V & " = "
                Dr(IxR) = "X" & V
            End If
        End If
    End If
    PushI Dry, Dr
Next
XMcLREmp = Drs(McLR.Fny, Dry)
'Insp "QIde_B_AlignMth.XMcLREmp", "Inspect", "Oup(XMcLREmp) McLR", LinzDrs(XMcLREmp), LinzDrs(McLR): Stop
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


Private Function XMcVSfx(McInsp As Drs) As Drs
'Fm McInsp : L Gpno MthLin IsRmk ! If las lin is rmk and is 'Insp, exl it.
'Ret       : L Gpno MthLin IsRmk
'            V Sfx Rst @@
Dim Dr, Dry()
For Each Dr In Itr(McInsp.Dry)
    Dim Av(): Av = WDrVSfxRst(Dr(2))
    PushIAy Dr, Av
    PushI Dry, Dr
Next
XMcVSfx = DrszAddFF(McInsp, "V Sfx Rst", Dry)
'Insp "QIde_B_AlignMth.XMcVSfx", "Inspect", "Oup(XMcVSfx) McInsp", LinzDrs(XMcVSfx), LinzDrs(McInsp): Stop
End Function

Private Function XMcDcl(McVSfx As Drs) As Drs
'Fm McVSfx : L Gpno MthLin IsRmk
'            V Sfx Rst
'Ret       : L Gpno MthLin IsRmk
'            V Sfx Dcl Rst       ! Add Dcl from V & Sfx @@
Dim MthLin$, IxMthLin%, Dr, Dry(), IGpno
AsgIx McVSfx, "MthLin", IxMthLin
For Each IGpno In AywDist(IntCol(McVSfx, "Gpno"))
    Dim A As Drs: A = ColEq(McVSfx, "Gpno", IGpno) ' L Gpno MthLin IsRmk V Sfx Rst ! Sam Gpno
    Dim B As Drs: B = XDcl(A) ' L Gpno MthLin IsRmk V Sfx Dcl Rst ! Adding Dcl using V Sfx
    Dim O As Drs: O = DrszAdd(O, B)
Next
XMcDcl = O
'Insp "QIde_B_AlignMth.XMcDcl", "Inspect", "Oup(XMcDcl) McVSfx", LinzDrs(XMcDcl), LinzDrs(McVSfx): Stop
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



Private Function XCmlAct(CmlEpt As Drs, M As CodeModule) As Drs
Dim CV$(): CV = StrColzDrs(CmlEpt, "V")
Dim Act$(): Act = StrCol(DMth(M), "MthLin")
Insp CSub, "CmlAct", "CV Act", CV, Act
Stop
End Function

Private Function XMlVSfx(Ml$) As Drs
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
XMlVSfx = DrszFF("V Sfx", Dry)
'Insp "QIde_B_AlignMth.XMlVSfx", "Inspect", "Oup(XMlVSfx) Ml", LinzDrs(XMlVSfx), Ml: Stop
End Function

Private Function XCmlMthRet(CmlDclPm As Drs) As Drs
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
XCmlMthRet = DrszAddFF(CmlDclPm, "TyChr RetAs", Dry)
'BrwDrs CmlMthRet: Stop
'Insp "QIde_B_AlignMth.XCmlMthRet", "Inspect", "Oup(XCmlMthRet) CmlDclPm", LinzDrs(XCmlMthRet), LinzDrs(CmlDclPm): Stop
End Function
Private Function XCmlDclPm(CmlPm As Drs, CmlVSfx As Drs) As Drs
'Fm CmlPm : V Sfx RHS CmNm Pm
'Ret      : V Sfx RHS CmNm Pm DclPm ! use [CmlVSfx] & [Pm] to bld [DclPm] @@
Dim IxPm%: AsgIx CmlPm, "Pm", IxPm
Dim Dr, Dry(): For Each Dr In Itr(CmlPm.Dry)
    Dim Pm$: Pm = Dr(IxPm)
    Dim DclPm$: DclPm = XDclPm(Pm, CmlVSfx)
    PushI Dr, DclPm
    PushI Dry, Dr
Next
XCmlDclPm = DrszAddFF(CmlPm, "DclPm", Dry)
'Insp "QIde_B_AlignMth.XCmlDclPm", "Inspect", "Oup(XCmlDclPm) CmlPm CmlVSfx", LinzDrs(XCmlDclPm), LinzDrs(CmlPm), LinzDrs(CmlVSfx): Stop
End Function
Private Function XDclPm$(Pm$, CmlVSfx As Drs)
Dim O$(), Sfx$, P
For Each P In Itr(SyzSS(Pm))
    Sfx = ValzColEq(CmlVSfx, "Sfx", "V", P)
    PushI O, P & Sfx
Next
XDclPm = JnCommaSpc(O)
'Insp CSub, "Finding DclPm(CallgPm, CmlFmMc)", "CallgPm CmlFmMc XDclPm", CallgPm, LinzDrs(CmlFmMc), XDclPm: Stop
End Function

Private Function XCrEpt(CrJn As Drs) As S1S2s
'Fm  CrJn : V Rmk CmNm
'Ret      : CmNm RmkLines ! RmkLines is find by each V in CrVpr & Mthn = V & CmPfx @@
Dim A As Drs: A = DrszSel(CrJn, "CmNm Rmk")
Dim Dr, Ly$(): For Each Dr In Itr(A.Dry)
    PushI Ly, Dr(0) & " " & Dr(1)
Next
Dim D As Dictionary: Set D = Dic(Ly, vbCrLf)
Dim O As Dictionary: Set O = AddSfxzDic(D, " @@")
XCrEpt = S1S2szDic(O)
'Insp "QIde_B_MthOp.CrEpt", "Inspect", "Oup(CrEpt) CrJn", FmtS1S2s(CrEpt), LinzDrs(CrJn): Stop
End Function

Private Function XCrRmk(CrFst As Drs) As Drs
'Fm CrFst : V P R1 R2 R3 IsRet Fst     ! P R1..3 are aligned (always hav sam len & las chr is [.]
'Ret      : V P R1 R2 R3 IsRet Fst Rmk ! Bld Rmk from P R1 R2 R3 & Fst @@
If IsNeFF(CrFst, "V P R1 R2 R3 IsRet Fst") Then Stop
Dim V$, P$, R1$, R2$, R3$, IsRet As Boolean, Fst As Boolean, Rmk$
Dim Dr, Dry(): For Each Dr In Itr(CrFst.Dry)
    AsgAp Dr, V, P, R1, R2, R3, IsRet, Fst
    If R1 <> "" Then
        If Left(R1, 3) <> " ' " Then Stop
        R1 = Mid(R1, 4)
    End If
    Dim Lbl$
        Select Case True
        Case Fst And IsRet
            Lbl = "'Ret" & Space(Len(P)) & " : "
        Case Fst:
            Lbl = "'Fm " & P & " : "
        Case Else
            Lbl = "'" & Space(Len(P) + 6)
        End Select
    Rmk = Lbl & R1 & R2 & R3
    PushI Dr, Rmk
    PushI Dry, Dr
Next
XCrRmk = DrszAddFF(CrFst, "Rmk", Dry)
'Insp "QIde_B_AlignMth.XCrRmk", "Inspect", "Oup(XCrRmk) CrFst", LinzDrs(XCrRmk), LinzDrs(CrFst): Stop
End Function

Private Function XCrWiRmk(CrSel As Drs) As Drs
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
XCrWiRmk = Drs(CrSel.Fny, Dry)
'Insp "QIde_B_MthOp.CrWiRmk", "Inspect", "Oup(CrWiRmk) CrSel", LinzDrs(CrWiRmk), LinzDrs(CrSel): Stop
End Function

Private Sub Z()
QIde_B_AlignMth:
End Sub


