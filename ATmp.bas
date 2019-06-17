Attribute VB_Name = "ATmp"
Option Explicit
Option Compare Text
Dim Class1 As New Class1

Sub AlignMthDimzML(M As CodeModule, MthLno&, Optional SkpChkSelf As Boolean, Optional Rpt As EmRpt)
Static F As New ATmp__AlignMthDimzML
Dim D1 As Drs, D2 As Drs
'== Exit if parameter error ============================================================================================
                  If F.IsPmEr(M, MthLno) Then Exit Sub       ' X-Parameter-er. M-isnothg | MthLno<=0
Dim Ml$:     Ml = ContLinzML(M, MthLno)
Dim MlNm$: MlNm = Mthn(Ml)              '  # Ml-Name.
                  If F.XSelf(SkpChkSelf, M, MlNm) Then Exit Sub ' #X-Self-aligning-er. ! Mdn<>'QIde...' & MlNm<>'AlignMthDimzML
Dim Mc As Drs: Mc = DMthCxtzML(M, MthLno) ' L MthLin # Mc.
                    If NoReczDrs(Mc) Then Exit Sub

Dim IsUpd As Boolean: IsUpd = IsUpdzRpt(Rpt)

'== Align DblEqRmk (De) ================================================================================================
'   When a rmk lin begins with '== or '--, expand it to 120 = or -
Dim De      As Drs:      De = F.De(Mc)                        ' L MthLin    # Dbl-Eq | Dbl-Dash
Dim DeLNewO As Drs: DeLNewO = F.DeLNewO(De)                   ' L NewL OldL
Dim OUpdDe:                   If IsUpd Then RplLin M, DeLNewO

'== Align Mth Cxt ======================================================================================================
Dim McCln As Drs: McCln = F.McCln(Mc) ' L MthLin # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align
If NoReczDrs(McCln) Then Exit Sub
Dim McGp    As Drs:    McGp = F.McGp(McCln)     ' L Gpno MthLin               ! with L in seq will be one gp
Dim McRmk   As Drs:   McRmk = F.McRmk(McGp)     ' L Gpno MthLin IsRmk         ! a column IsRmk is added
Dim McTRmk  As Drs:  McTRmk = F.McTRmk(McRmk)   ' L Gpno MthLin IsRmk         ! For each gp, the front rmk lines are TopRmk, rmv them
Dim McInsp  As Drs:  McInsp = F.McInsp(McTRmk)  ' L Gpno MthLin IsRmk         ! If las lin is rmk and is 'Insp, exl it.
Dim McVSfx  As Drs:  McVSfx = F.McVSfx(McInsp)  ' L Gpno MthLin IsRmk
                                                ' V Sfx Rst
Dim McDcl   As Drs:   McDcl = F.McDcl(McVSfx)   ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl Rst               ! Add Dcl from V & Sfx
Dim McLR    As Drs:    McLR = F.McLR(McDcl)     ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS RHS Rst       ! Add LHS Expr from Rst
Dim McR123  As Drs:  McR123 = F.McR123(McLR)    ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
Dim McFill  As Drs:  McFill = F.McFill(McR123)  ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS Expr
                                                ' F0 FSfx FExpr FR1 FR2       ! Adding F*
Dim McAlign As Drs: McAlign = F.McAlign(McFill) ' L Align                     ! Bld the new Align

                         D1 = DrseCeqC(McAlign, "MthLin Align")
                         D2 = SelDrs(D1, "L Align MthLin")
Dim McLNewO As Drs: McLNewO = LNewO(D2.Dry)
Dim OAlignCm:                 If IsUpd Then RplLin M, McLNewO

'== Gen Bs (Brw-Stmt) ==================================================================================================
Dim Bs      As Drs:              Bs = F.Bs(McCln)                               ' L BsLin ! FstTwoChr = '@
Dim Bs1     As Drs:             Bs1 = ColEqSel(McR123, "IsRmk", False, "V Sfx")
Dim Bs2     As Drs:             Bs2 = ColNe(Bs1, "V", "")
Dim VSfx    As Dictionary: Set VSfx = DiczDrsCC(Bs2)
Dim Mdn$:                       Mdn = MdnzM(M)
Dim BsLNewO As Drs:         BsLNewO = F.BsLNewO(Bs, VSfx, Mdn, MlNm)
Dim OUpdBs:                           If IsUpd Then RplLin M, BsLNewO

                       D1 = ColPfx(Mc, "MthLin", "Static F")
Dim NoSf As Boolean: NoSf = NoReczDrs(D1)
If Not NoSf Then
    '== Ens Static-F-Dcl (Sf) ==========================================================================================
    Dim Dr1:                Dr1 = D1.Dry(0)
    Dim SfLno&:           SfLno = Dr1(0)
    Dim SfLin$:           SfLin = Dr1(1)
    Dim CclsNm$:         CclsNm = IIf(NoSf, "", Mdn & "__" & MlNm)
    Dim SfLNewO As Drs: SfLNewO = F.SfLNewO(NoSf, SfLno, SfLin, CclsNm) ' Only one or no line
    Dim OEnsSf:                   If IsUpd Then RplLin M, SfLNewO
    
    '== Ens Chd-Cls (Ccls)==============================================================================================
    Dim HasCcls As Boolean:     HasCcls = HasCmpzP(PjzM(M), CclsNm)
    Dim OEnsCcls:                         If IsUpd Then F.OEnsCcls NoSf, HasCcls, M, CclsNm
    Dim Ccls    As CodeModule: Set Ccls = PjzM(M).VBComponents(CclsNm).CodeModule
End If

'== Crt Chd-Mth (Cm)====================================================================================================
Dim CmMd  As CodeModule: Set CmMd = IIf(NoSf, M, Ccls)
Dim CmMdy$:                 CmMdy = IIf(NoSf, "Private", "Friend")
Dim CmV   As Drs:             CmV = ColEqSel(McR123, "IsRmk", False, "V Sfx LHS RHS") ' V Sfx LHR RHS
Dim WiSf  As Boolean:        WiSf = Not NoSf
Dim MlNmDD$:               MlNmDD = BefOrAll(MlNm, "__") & "__"
Dim CmNm  As Drs:            CmNm = F.CmNm(CmV, WiSf, MlNmDD)                         ' V Sfx LHS RHS CmNm ! som CmNm may be blank
Dim CmEpt As Drs:           CmEpt = ColNe(CmNm, "CmNm", "")                           ' V Sfx LHS RHS CmNm ! All CmNm has val
Dim CmEptNm$():           CmEptNm = StrCol(CmEpt, "CmNm")                             '                    ! It is ept mth ny.  They will be used create new chd mth
Dim CmActNm$():           CmActNm = MthNyzM(CmMd)                                     '                    ! It is from chd cls of given md
Dim CmNewNm$():           CmNewNm = MinusAy(CmEptNm, CmActNm)                         '                    ! The new ChdMthNy to be created.
Dim CmNew As Drs:           CmNew = ColIn(CmEpt, "CmNm", CmNewNm)
Dim CdNewCm$:             CdNewCm = F.CdNewCm(CmNew, CmMdy)                           '                    ! Cd to be append to CmMd
Dim OCrtCm:                         If IsUpd Then ApdLines CmMd, CdNewCm

'== Upd Chd-Mth-Lin (Cml) ==============================================================================================
'   If the calling pm has been changed, the chd-mth-lin will be updated.
Dim MlVSfx    As Drs:    MlVSfx = F.MlVSfx(Ml)               ' Ret V Sfx                           ! the MthLin's pm V Sfx
                             D1 = SelDrs(CmV, "V Sfx")
Dim CmlVSfx   As Drs:   CmlVSfx = AddDrs(MlVSfx, D1)
Dim CmlPm     As Drs:     CmlPm = F.CmlPm(CmEpt)             ' V Sfx RHS CmNm Pm
Dim CmlDclPm  As Drs:  CmlDclPm = F.CmlDclPm(CmlPm, CmlVSfx) ' V Sfx RHS CmNm Pm DclPm             ! use [CmlVSfx] & [Pm] to bld [DclPm]
Dim CmlMthRet As Drs: CmlMthRet = F.CmlMthRet(CmlDclPm)      ' V Sfx RHS CmNm Pm DclPm TyChr RetAs
Dim CmlEpt    As Drs:    CmlEpt = F.CmlEpt(CmlMthRet, CmMdy) ' V CmNm EptL
                             D1 = DMth(CmMd)                 ' L Mdy Ty Mthn MthLin
                                If NoSf Then D1 = ColEq(D1, "Mdy", "Prv")
                                If WiSf Then D1 = ColEq(D1, "Mdy", "Frd")
Dim CmlAct   As Drs:   CmlAct = SelDrszAs(D1, "L Mthn:CmNm MthLin:ActL") ' L CmNm ActL
Dim CmlJn    As Drs:    CmlJn = JnDrs(CmlEpt, CmlAct, "CmNm", "L ActL")  ' V CmNm EptL L ActL ! som EptL & ActL may eq
                           D2 = DrseCeqC(CmlJn, "EptL ActL")             ' V CmNm EptL L ActL ! All EptL & ActL are diff
Dim CmlLNewO As Drs: CmlLNewO = SelDrszAs(D2, "L EptL:NewL ActL:OldL")   ' L NewL OldL
Dim OUpdCml:                    If IsUpd Then RplLin CmMd, CmlLNewO

'== Rpl Mth-Brw (Mb)====================================================================================================
'   Des: Mth-Brw is a remarked Insp-stmt in each las lin of cm.  It insp all the inp oup
'   Lgc: Fnd-and-do MbLNewO
'        Fnd-and-do NewMb
'BrwDrs CmlEpt: Stop
Dim CmLis   As Drs:   CmLis = SelDrszAs(CmlEpt, "CmNm:Mthn EptL:MthLin") ' Mthn MthLin
Dim MbEpt   As Drs:   MbEpt = F.MbEpt(CmLis, Mdn)                        ' Mthn MthLin MbStmt
Dim Cm$():               Cm = StrCol(CmLis, "Mthn")
Dim MbAct   As Drs:   MbAct = F.MbAct(Cm, CmMd)                          ' L Mthn OldL               ! OldL is MbStmt
Dim MbJn    As Drs:    MbJn = JnDrs(MbEpt, MbAct, "Mthn", "OldL L")      ' Mthn MthLin MbStmt OldL L
Dim MbSel   As Drs:   MbSel = SelDrszAs(MbJn, "L MbStmt:NewL OldL")      ' L NewL OldL
Dim MbLNewO As Drs: MbLNewO = DrseCeqC(MbSel, "NewL OldL")
Dim OUpdMb:                   If IsUpd Then RplLin CmMd, MbLNewO

'== Crt Mth-Brw (Mb)====================================================================================================
                     D1 = LJnDrs(MbEpt, MbAct, "Mthn", "L", "HasAct") ' Mthn MthLin MbStmt L HasAct
                     D2 = ColEq(D1, "HasAct", False)                  ' Mthn MthLin MbStmt L HasAct
Dim MbNew As Drs: MbNew = SelDrszAs(D2, "Mthn MbStmt:NewL")
Dim OCrtMb:               If IsUpd Then F.OCrtMb CmMd, MbNew

'== Upd Chd-Rmk (Cr) ===================================================================================================
'   If any of the calling pm's rmk is changed, the chd-mth-rmk will be updated
Dim CrSel     As Drs:              CrSel = SelDrs(McR123, "V R1 R2 R3") ' V R1 R2 R3
Dim CrWiRmk   As Drs:            CrWiRmk = F.CrWiRmk(CrSel)             ' V R1 R2 R3      ! rmv those R1 2 3 are blank
Dim CrVrAy()  As Vr:              CrVrAy = F.CrVrAy(CrWiRmk)
Dim CrPm      As Drs:               CrPm = SelDrs(CmlDclPm, "V Pm")     ' V Pm
Dim CrVprAy() As Vpr:            CrVprAy = F.CrVprAy(CrPm, CrVrAy)      ' Sam itm as CrPm
Dim CrRmkV    As S1S2s:           CrRmkV = F.CrRmkV(CrVprAy)            ' V RmkLines
Dim CrCmNm    As Drs:             CrCmNm = SelDrs(CmlEpt, "V CmNm")     ' V CmNm
Dim CrVCmNm   As Dictionary: Set CrVCmNm = DiczDrsCC(CrCmNm)
Dim CrEpt     As S1S2s:            CrEpt = MapS1(CrRmkV, CrVCmNm)       ' CmNm RmkLines
If SiVpr(CrVprAy) <> CrEpt.N Then Stop

Dim CrAct As S1S2s: CrAct = MthRmkzNy(CmMd, Cm)
Dim CrChg As S1S2s: CrChg = F.CrChg(CrEpt, CrAct)              ' CmNm RmkLines ! Only those need to change
Dim OUpdCr:                 If IsUpd Then F.OUpdCr CrChg, CmMd

'== Rpt <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
        StrzRpt(Rpt), _
        FmtDrs(DeLNewO), FmtDrs(McLNewO), FmtDrs(BsLNewO), FmtDrs(SfLNewO), CclsMsg, _
        CdNewCm, FmtDrs(MbLNewO), FmtDrs(MbNew), FmtDrs(CmlLNewO), FmtS1S2s(CrChg)
End If
'Insp CSub, "Cr", "CrVpr CrS1S2s", FmtDrs(CrVpr), FmtS1S2s(CrS1S2s): Stop
End Sub

