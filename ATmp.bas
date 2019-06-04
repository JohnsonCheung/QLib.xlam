Attribute VB_Name = "ATmp"
Sub Z2()
Dim MthLno&, M As CodeModule
Set M = Md("ATmp")
MthLno = MthLnozMM(M, "X13A")
QIde_Base_MthOp.AlignMthDimzML M, MthLno
End Sub
Sub X13A()
Const CmPfx$ = "W"
Dim A$: A = A
Dim C$: C = C
Dim D:  D = D
End Sub
Sub AlignMthDimzML(Md As CodeModule, MthLno&, Optional SkpChkSelf As Boolean)
Static F As New QIde_Base_MthOp__AlignMthDimzML
'-- #O  = Oup
'-- #X  = eXit
'-- #Mc = Mth-Cxt
'-- #Ml = Mth-Lin
'-- #Cm = New-Chdfun
'-- #Cr = Chd-Rmk

'== Exit if parameter error ============================================================================================
If F.XPm(Md, MthLno) Then Exit Sub       ' X-Parameter-er. Md-isnothg | MthLno<=0
Dim Ml$:     Ml = ContLinzML(Md, MthLno)
Dim MlNm$: MlNm = Mthn(Ml)               '  #Ml-Name.
If F.XSelf(SkpChkSelf, Md, MlNm) Then Exit Sub ' #X-Self-aligning-er. ! Mdn<>'QIde...' & MlNm<>'AlignMthDimzML

'== Align DblEqRmk =====================================================================================================
Dim Mc           As Drs:           Mc = DMthCxt(Md, MthLno)        ' L MthLin  #Mc.
Dim McDblEqRmk   As Drs:   McDblEqRmk = F.McDblEqRmk(Mc)           ' L MthLin  #Mc-Dbl-Equal-rmk
Dim McDblEqLNewO As Drs: McDblEqLNewO = F.McDblEqLNewO(McDblEqRmk) ' NewL OldL
Dim OUpdDblEqRmk As Unt: OUpdDblEqRmk = RplLin(Md, McDblEqLNewO)

'== Align / GenCalling the main method =================================================================================
Dim McCln As Drs: McCln = F.McCln(Mc) ' L MthLin #Mc-Cln. ! must Dim, Rmk(but not 'If 'Insp, '==). Cln to Align
If NoReczDrs(McCln) Then Exit Sub

Dim McGp   As Drs:   McGp = F.McGp(McCln)
Dim McRmk  As Drs:  McRmk = F.McRmk(McGp)   ' L Gpno MthLin IsRmk     ! a column IsRmk is added
Dim McTRmk As Drs: McTRmk = F.McTRmk(McRmk) ' L Gpno MthLin IsRmk     ! For each gp, the front rmk lines are TopRmk,
'                                                                     ! rmv them
Dim McBrk  As Drs:  McBrk = F.McBrk(McTRmk) ' L Gpno MthLin IsRmk     ! Brk the MthLin into V Sfx Expr Rmk
'                                             V Sfx Dcl LHS Expr      ! If there is no asg stmt LHS and Expr will be same as V
'                                             R1 R2 R3                ! in this case, a new ChdMth will be created
Dim McFill As Drs: McFill = F.McFill(McBrk) ' L Gpno MthLin IsRmk2    ! @Fill@ : FV..FR2
'                                             V Sfx Dcl LHS Expr
'                                             F0 FSfx FExpr FR1 FR2
Dim McDim  As Drs:  McDim = F.McDim(McFill) ' L NewL OldL             ! Bld the new DimLin from @Brk and @Fill

Dim McLNewO  As Drs:  McLNewO = DrseCeqC(McDim, "NewL OldL")
Dim OAlignCm As Unt: OAlignCm = RplLin(Md, McLNewO)
'Brw FmtLNewO(McLNewO, Mc):Exit Sub

'== Optional Ensure <Static F> declaration =============================================================================
'== Optional Ensure Ccls ===============================================================================================
Dim McLy$():                    McLy = StrColzDrs(Mc, "MthLin")
Dim NoSf     As Boolean:        NoSf = Not HasPfxzAy(McLy, "Static F")
Dim OEnsCcls As Unt:        OEnsCcls = F.OEnsCcls(NoSf, Md, MlNm)
Dim OEnsSf   As Unt:          OEnsSf = F.OEnsSf(NoSf, Md, MthLno, MlNm) '  #Ens-StaticF
Dim Ccls     As CodeModule: Set Ccls = F.Ccls(NoSf, Md, MlNm)

'=======================================================================================================================
'== Optional Create ChdFun =============================================================================================
'-- Cm #Chd-Mth the new chd mth to be created
Dim CmPfx$:          CmPfx = F.CmPfx(NoSf, McLy)
Dim NoCm As Boolean:  NoCm = NoSf And CmPfx = ""
If NoCm Then Exit Sub
Dim CmMd As CodeModule: Set CmMd = IIf(NoSf, Md, Ccls)

Dim CmMdy$:         CmMdy = IIf(NoSf, "Private", "Friend")
Dim CmEpt$():       CmEpt = F.CmEpt(McBrk)                          ' It is from V and Expr=F.{V}
Dim CmAct$():       CmAct = MthnyzM(CmMd)                           ' It is from chd cls of given md
Dim CmNew$():       CmNew = MinusAy(AddPfxzAy(CmEpt, CmPfx), CmAct) ' The new ChdMthNy to be created.
Dim CmStr$:         CmStr = F.CmStr(CmNew, McBrk, CmMdy, CmPfx)     ' Mth-Str to be append to CmMd
Dim OCrtCm As Unt: OCrtCm = ApdLines(CmMd, CmStr)

'== Upd ChdMthLin ======================================================================================================
'--Cml = Child-Mth-Lin to be updated
Dim CmlCallgPfx$:   CmlCallgPfx = IIf(NoSf, CmPfx, "F.")
Dim CmlFmMc As Drs:     CmlFmMc = DrswColEqSel(McBrk, "IsRmk", False, "V Sfx Expr")

    Dim D1      As Drs:      D1 = F.MlVSfx(Ml)
    Dim D2      As Drs:      D2 = SelDrs(CmlFmMc, "V Sfx")
    Dim CmlVSfx As Drs: CmlVSfx = AddDrs(D1, D2)

Dim CmlCallg   As Drs:   CmlCallg = F.CmlCallg(CmlFmMc, CmlCallgPfx)
Dim CmlDclPm   As Drs:   CmlDclPm = F.CmlDclPm(CmlCallg, CmlVSfx)
Dim CmlMthRet As Drs: CmlMthRet = F.CmlMthRet(CmlDclPm)
Dim CmlEpt     As Drs:     CmlEpt = F.CmlEpt(CmlMthRet, CmMdy)                                   ' V EptL
Dim CmlAct     As Drs:     CmlAct = Drs_MthLinzM(CmMd)                                            ' Mdn Lno Mthn MthLin
Dim CmlJn      As Drs:      CmlJn = LJnDrs(CmlEpt, CmlAct, "V:Mthn", "Lno MthLin:ActL", "HasAct")
Dim CmlLNewO As Drs: CmlLNewO = F.CmlLNewO(CmlJn)
Dim OUpdCml  As Unt:  OUpdCml = RplLin(CmMd, CmlLNewO)
Exit Sub
'== Update Chd Mth's Rmk ===============================================================================================
Dim CrEpt   As Drs:   CrEpt = F.CrEpt(McBrk)                   ' V EptR
Dim CrAct   As Drs:   CrAct = F.CrAct(CmEpt, CmMd)             ' V ActR L
Dim CrJn    As Drs:    CrJn = JnDrs(CrAct, CrEpt, "V", "EptR") ' V ActR L EptR
Dim CrLNewO As Drs: CrLNewO = F.CrLNewO(CrJn)
Dim OUpdCr  As Unt:  OUpdCr = RplLin(CmMd, CrLNewO)
End Sub
