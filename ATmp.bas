Attribute VB_Name = "ATmp"
Dim Class1 As New Class1

Sub AlignMthDimzML(Md As CodeModule, MthLno&, Optional SkpChkSelf As Boolean, Optional Rpt As EmRpt)
Static F As New ATmp__AlignMthDimzML
Dim D1 As Drs, D2 As Drs, Dr1()
'== Exit if parameter error ============================================================================================
If F.XPm(Md, MthLno) Then Exit Sub       ' X-Parameter-er. Md-isnothg | MthLno<=0
Dim Ml$:     Ml = ContLinzML(Md, MthLno)
Dim MlNm$: MlNm = Mthn(Ml)               '  # Ml-Name.
If F.XSelf(SkpChkSelf, Md, MlNm) Then Exit Sub ' #X-Self-aligning-er. ! Mdn<>'QIde...' & MlNm<>'AlignMthDimzML
Dim Mc As Drs: Mc = DMthCxtzML(Md, MthLno) ' L MthLin # Mc.
If NoReczDrs(Mc) Then Exit Sub

Dim IsUpd As Boolean: IsUpd = IsUpdzRpt(Rpt)
'== Align DblEqRmk (De) ================================================================================================
'   When a rmk lin begins with '== or '--, expand it to 120 = or -
Dim De      As Drs:      De = F.De(Mc)                         ' L MthLin    # Dbl-Eq
Dim DeLNewO As Drs: DeLNewO = F.DeLNewO(De)                    ' L NewL OldL
Dim OUpdDe:                   If IsUpd Then RplLin Md, DeLNewO

'== Align Mth Cxt ======================================================================================================
Dim McCln As Drs: McCln = F.McCln(Mc) ' L MthLin # Mc-Cln. ! must Dim | Asg | Rmk(but not 'If 'Insp, '==). Cln to Align
If NoReczDrs(McCln) Then Exit Sub

Dim McGp   As Drs:   McGp = F.McGp(McCln)    ' Gpno MthLin            ! with L in seq will be one gp
Dim McRmk  As Drs:  McRmk = F.McRmk(McGp)    ' L Gpno MthLin IsRmk    ! a column IsRmk is added
Dim McTRmk As Drs: McTRmk = F.McTRmk(McRmk)  ' L Gpno MthLin IsRmk    ! For each gp, the front rmk lines are TopRmk,
                                             '                        ! rmv them
Dim McVSfx As Drs: McVSfx = F.McVSfx(McTRmk) ' L Gpno IsRmk
                                             ' V Sfx Rst
Dim McDcl  As Drs:  McDcl = F.McDcl(McVSfx)  ' L Gpno MthLin IsRmk
                                             ' V Sfx Dcl Rst          ! Add Dcl from V & Sfx
Dim McLR   As Drs:   McLR = F.McLR(McDcl)    ' L Gpno MthLin IsRmk
                                             ' V Sfx Dcl LHS Expr Rst ! Add LHS Expr from Rst
                                            

Dim McR123 As Drs: McR123 = F.McR123(McLR) ' L Gpno MthLin IsRmk
                                           ' V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst

Dim McFill  As Drs:  McFill = F.McFill(McR123)  ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS Expr
                                                ' F0 FSfx FExpr FR1 FR2 ! Adding F*
Dim McAlign As Drs: McAlign = F.McAlign(McFill) ' L Align               ! Bld the new Align
                         
                         D1 = DrseCeqC(McAlign, "MthLin Align")
                         D2 = SelDrs(D1, "L Align MthLin")
Dim McLNewO As Drs: McLNewO = LNewO(D2.Dry)
Dim OAlignCm:                 If IsUpd Then RplLin Md, McLNewO

'== Gen Bs (Brw-Stmt) ==================================================================================================
Dim Bs      As Drs:              Bs = F.Bs(McCln)                               ' L BsLin ! FstTwoChr = '@
Dim Bs1     As Drs:             Bs1 = ColEqSel(McR123, "IsRmk", False, "V Sfx")
Dim Bs2     As Drs:             Bs2 = ColNe(Bs1, "V", "")
Dim VSfx    As Dictionary: Set VSfx = DiczDrsCC(Bs2)
Dim Mdn$:                       Mdn = MdnzM(Md)
Dim BsLNewO As Drs:         BsLNewO = F.BsLNewO(Bs, VSfx, Mdn, MlNm)
Dim OUpdBs:                           If IsUpd Then RplLin Md, BsLNewO

'== Ens Static-F-Dcl (Sf) ==============================================================================================
                       D1 = DrswColPfx(Mc, "MthLin", "Static F")
Dim NoSf As Boolean: NoSf = NoReczDrs(D1)
                      Dr1 = D1.Dry(0)
Dim SfLno&:                         If Not NoSf Then SfLno = Dr1(0)
Dim SfLin$:                         If Not NoSf Then SfLin = Dr1(1)
Dim CclsNm$:         CclsNm = IIf(NoSf, "", Mdn & "__" & MlNm)
Dim SfLNewO As Drs: SfLNewO = F.SfLNewO(NoSf, SfLno, SfLin, CclsNm) ' Only one or no line
Dim OEnsSf:                   If IsUpd Then RplLin Md, SfLNewO

'== Ens Chd-Cls (Ccls)==================================================================================================
Dim HasCcls As Boolean: HasCcls = HasCmpzP(PjzM(Md), CclsNm)
Dim OEnsCcls:                     If IsUpd Then F.OEnsCcls NoSf, HasCcls, Md, CclsNm
Dim Ccls As CodeModule:           If Not NoSf Then Set Ccls = PjzM(Md).VBComponents(CclsNm).CodeModule

'== Crt Chd-Mth (Cm)====================================================================================================
Dim McLy$():          McLy = StrColzDrs(Mc, "MthLin")
Dim CmPfx$:          CmPfx = F.CmPfx(NoSf, McLy)
Dim NoCm As Boolean:  NoCm = NoSf And CmPfx = ""
If NoCm Then GoTo Rpt

Dim CmPfxEr As Boolean: CmPfxEr = F.CmPfxEr(CmPfx, Md)
If CmPfxEr Then Exit Sub

Dim CmMd As CodeModule: Set CmMd = IIf(NoSf, Md, Ccls)

Dim CmMdy$:     CmMdy = IIf(NoSf, "Private", "Friend")
Dim ExprPfx$: ExprPfx = IIf(CmPfx = "", "F.", CmPfx)            '  ! Either F. or CmPfx.  It used to detect the Expr is a calling cm expr
Dim CmEpt$():   CmEpt = F.CmEpt(McR123, ExprPfx)                '  ! It is from V and  {V} = {CmPfx}{Expr}.
                                                                '  ! They will be used create new chd mth
Dim CmAct$():   CmAct = MthnyzM(CmMd)                           '  ! It is from chd cls of given md
Dim CmNew$():   CmNew = MinusAy(AddPfxzAy(CmEpt, CmPfx), CmAct) '  ! The new ChdMthNy to be created.
Dim CmStr$:     CmStr = F.CmStr(CmNew, McR123, CmMdy, CmPfx)    '  ! Mth-Str to be append to CmMd
Dim OCrtCm:             If IsUpd Then ApdLines CmMd, CmStr

'== Upd Chd-Mth-Lin (Cml) ==============================================================================================
'   If the calling pm has been changed, the chd-mth-lin will be updated.
Dim CmlCallgPfx$:   CmlCallgPfx = IIf(NoSf, CmPfx, "F.")
Dim CmlFmMc As Drs:     CmlFmMc = ColEqSel(McR123, "IsRmk", False, "V Sfx Expr")

Dim MlVSfx  As Drs:  MlVSfx = F.MlVSfx(Ml)             ' Ret V Sfx ! the MthLin's pm V Sfx
                         D1 = SelDrs(CmlFmMc, "V Sfx")
Dim CmlVSfx As Drs: CmlVSfx = AddDrs(MlVSfx, D1)

Dim CmlCallg As Drs: CmlCallg = F.CmlCallg(CmlFmMc, CmlCallgPfx) ' V Sfx Expr   ! It is subset of McR123 where Expr is a calling1 or calling2.
                                                                 ' Mthn CallgPm ! calling1 is Expr = CmPfx & V              No Pm yet
                                                                 '              ! calling2 is HasSfx(Expr, CmPfx & V & "("  with Pm

Dim CmlDclPm  As Drs:  CmlDclPm = F.CmlDclPm(CmlCallg, CmlVSfx)                    ' V Sfx Expr Mthn CallPm DclPm
Dim CmlMthRet As Drs: CmlMthRet = F.CmlMthRet(CmlDclPm)                            ' V Sfx Expr Mthn CallPm DclPm TyChr RetAs
Dim CmlEpt    As Drs:    CmlEpt = F.CmlEpt(CmlMthRet, CmMdy)                       ' V Mthn EptL
Dim CmlAct    As Drs:    CmlAct = DMth(CmMd)                                       ' L Mdy Ty Mthn MthLin
                             D1 = SelDrszAs(CmlAct, "L MthLin:ActL MthLin:V")      ' L ActL V
Dim CmlActV   As Drs:   CmlActV = F.CmlActV(D1, CmPfx)                             ' L ActL V ' RmvPfx CmlPfx from V
Dim CmlJn     As Drs:     CmlJn = LJnDrs(CmlEpt, CmlActV, "V", "L ActL", "HasAct") ' V EptL L ActL HasAct
                             D1 = ColEq(CmlJn, "HasAct", True)                     ' V EptL L ActL HasAct
                             D2 = DrseCeqC(D1, "EptL ActL")
Dim CmlLNewO  As Drs:  CmlLNewO = SelDrszAs(D2, "L EptL:NewL ActL:OldL")
Dim OUpdCml:                      If IsUpd Then RplLin CmMd, CmlLNewO

'== Rpl Mth-Brw (Mb)====================================================================================================
'   Des: Mth-Brw is a remarked Insp-stmt in each las lin of cm.  It insp all the inp oup
'   Lgc: Fnd-and-do MbLNewO
'        Fnd-and-do NewMb
'BrwDrs CmlEpt: Stop
Dim CmLis   As Drs:   CmLis = SelDrszAs(CmlEpt, "Mthn EptL:MthLin") ' Mthn MthLin
Dim MbEpt   As Drs:   MbEpt = F.MbEpt(CmLis, Mdn)                   ' Mthn MthLin MbStmt
Dim Cm$():               Cm = StrCol(CmLis, "Mthn")
Dim MbAct   As Drs:   MbAct = F.MbAct(Cm, CmMd)                     ' L Mthn OldL               ! OldL is MbStmt
Dim MbJn    As Drs:    MbJn = JnDrs(MbEpt, MbAct, "Mthn", "OldL L") ' Mthn MthLin MbStmt OldL L
Dim MbSel   As Drs:   MbSel = SelDrszAs(MbJn, "L MbStmt:NewL OldL") ' L NewL OldL
Dim MbLNewO As Drs: MbLNewO = DrseCeqC(MbSel, "NewL OldL")
Dim OUpdMb:                   If IsUpd Then RplLin CmMd, MbLNewO

'== Crt Mth-Brw (Mb)====================================================================================================
                     D1 = LJnDrs(MbEpt, MbAct, "Mthn", "L", "HasAct") ' Mthn MthLin MbStmt L HasAct
                     D2 = ColEq(D1, "HasAct", False)                  ' Mthn MthLin MbStmt L HasAct
Dim MbNew As Drs: MbNew = SelDrszAs(D2, "Mthn MbStmt:NewL")
Dim OCrtMb:               If IsUpd Then F.OCrtMb CmMd, MbNew

'== Upd Chd-Rmk (Cr) ===================================================================================================
'   If any of the calling pm's rmk is changed, the chd-mth-rmk will be updated
Dim CrVR3  As Drs:  CrVR3 = F.CrVR3(McR123)               ' V R1 R2 R3
Dim CrPm   As Drs:   CrPm = SelDrs(CmlDclPm, "V CallgPm") ' V CallgPm
Dim CrVPR3 As Drs: CrVPR3 = F.CrVPR3(CrPm, CrVR3, MlVSfx) ' V P R1 R2 R3
Dim CrFill As Drs: CrFill = F.CrFill(CrVPR3)              ' V P R1 R2 R3 FP FR1 FR2
Dim CrFst  As Drs:  CrFst = F.CrFst(CrFill)               ' V P R1 R2 R3 FP FR1 FR2 Fst
Dim CrRmk  As Drs:  CrRmk = F.CrRmk(CrFst)                ' V P R1 R3 R3 FP FR1 FR2 Fst Rmk

Dim CrEpt As S1S2s: CrEpt = F.CrEpt(CmPfx, CrRmk)              ' Mthn RmkLines ! RmkLines is find by each V in CrVPR3 & Mthn = V & CmPfx
Dim CrAct As S1S2s: CrAct = MthRmkzM(CmMd)
Dim CrChg As S1S2s: CrChg = F.CrChg(CrEpt, CrAct)              ' Mthn RmkLines ! Only those need to change
Dim OUpdCr:                 If IsUpd Then F.OUpdCr CrChg, CmMd
Rpt:
If IsRptzRpt(Rpt) Then
    Dim CclsMsg$: CclsMsg = CclsMsg
    Select Case True
    Case NoSf: CclsMsg = "No Static F"
    Case Not NoSf: CclsMsg = "Has Static F & Ccls exist: " & Ccls.Parent.Name
    Case Else: CclsMsg = "Has Static F & Ccls is created: " & Ccls.Parent.Name
    End Select
    Insp CSub, "Changes", _
        "EmRpt DblEqRmk Align BrwStmt EnsSf EnsCcls " & _
        "Crt-Chd-Mth Rpl-Mth-Brw Crt-Mth-Brw Upd-Chd-MthLin Rfh-Chd-Mth-Rmk", _
        RptStr(Rpt), _
        FmtDrs(DeLNewO), FmtDrs(McLNewO), FmtDrs(BsLNewO), FmtDrs(SfLNewO), CclsMsg, _
        CmStr, FmtDrs(MbLNewO), FmtDrs(MbNew), FmtDrs(CmlLNewO), FmtS1S2s(CrChg)
End If
'Insp CSub, "Cr", "CrVPR3 CrS1S2s", FmtDrs(CrVPR3), FmtS1S2s(CrS1S2s): Stop
End Sub

