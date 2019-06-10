Attribute VB_Name = "QIde_B_MthOp"
Option Compare Text
Option Explicit
Private Const Asm$ = ""
Private Const CMod$ = "MIde_Mth_Op."
Enum EmRpt
    EiRptOnly
    EiUpdAndRpt
    EiUpdOnly
    EiPushOnly
    EiUpdAndPush
End Enum
Function RptStr$(Rpt As EmRpt)
Dim O$
Select Case True
Case Rpt = EiRptOnly: O = "*RptOnly"
Case Rpt = EiUpdAndRpt: O = "*UpdAndRpt"
Case Rpt = EiUpdOnly: O = "*UpdOnly"
Case Else: O = "EmRptEr(" & Rpt & ")"
End Select
RptStr = O
End Function
Function IsPushzRpt(Rpt As EmRpt) As Boolean
Select Case True
Case Rpt = EiPushOnly, Rpt = EiUpdAndPush: IsPushzRpt = True
End Select
End Function
Function IsRptzRpt(Rpt As EmRpt) As Boolean
Select Case True
Case Rpt = EiUpdAndRpt, Rpt = EiRptOnly: IsRptzRpt = True
End Select
End Function
Function IsUpdzRpt(Rpt As EmRpt) As Boolean
Select Case True
Case Rpt = EiUpdAndRpt, Rpt = EiUpdOnly: IsUpdzRpt = True
End Select
End Function

Sub Z3()
Dim M As CodeModule: Set M = Md("QDao_Lnk_ErzLnk")
Dim L&: L = MthLnozMM(M, "ErzLnk")
AlignMthDimzML M, L
End Sub
Sub ZZ_IsAsgStmt()
Dim Dry(), L
    For Each L In SrczP(CPj)
        PushI Dry, Array(IIf(IsAsgStmt(L), "*", ""), L)
    Next
BrwDrs DrszFF("IsAsgStmt Lin", Dry)
End Sub
Function IsAsgStmt(Lin) As Boolean
Dim L$: L = Lin
Dim Nm$: Nm = ShfNm(L)
If Nm = "" Then Exit Function
IsAsgStmt = HasPfx(L, " = ")
End Function
Function IsAlignableDim(Lin) As Boolean
If T1(Lin) <> "Dim" Then Exit Function
Dim CommaP%, ColonP%, EqP%, A$, B$, C$
CommaP = InStr(Lin, ",")
ColonP = InStr(Lin, ":")
EqP = InStr(Lin, "=")
Select Case True
Case ColonP > 0 And CommaP > 0 And ColonP > CommaP
Case ColonP = 0 And CommaP > 0
Case ColonP > 0 And EqP > 0 And EqP > ColonP
    A = Bet(Lin, ":", "=")
    B = LTrim(RmvPfx(RmvPfx(A, "'"), "Set"))
    C = Dimn(Lin)
    If C = B Then
        IsAlignableDim = True
    End If
Case Else: IsAlignableDim = True
End Select
End Function
Sub AlignMthDimXX()
Dim A As New Aset, L
For Each L In SrczP(CPj)
    If T1(L) = "Dim" Then
        If IsAlignableDim(L) Then
            A.PushItm "1 " & LTrim(L)
        Else
            A.PushItm "0 " & LTrim(L)
        End If
    End If
Next
A.Srt.Vc
End Sub

Sub AlignMthDimzML(Md As CodeModule, MthLno&, Optional SkpChkSelf As Boolean, Optional Rpt As EmRpt)
Static F As New QIde_B_MthOp__AlignMthDimzML
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

Sub AlignMthDim(Optional Rpt As EmRpt)
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Sub
AlignMthDimzML M, CMthLno, Rpt:=Rpt
End Sub

Function AddColzCV(A As Drs, C$, V) As Drs
Dim Dr, Dry()
For Each Dr In Itr(A.Dry)
    PushI Dr, V
    PushI Dry, Dr
Next
AddColzCV = AddColzFFDry(A, C, Dry)
End Function

Function AddColzFFDry(A As Drs, FF$, NewDry()) As Drs
AddColzFFDry = Drs(AddFF(A.Fny, FF), NewDry)
End Function

Function AddColzFiller(A As Drs, CC$) As Drs
Dim O As Drs: O = A
Dim C
For Each C In SyzSS(CC)
    O = AddColzFillerC(O, C)
Next
AddColzFiller = O
End Function

Private Function AddColzFillerC(A As Drs, C) As Drs
'Fm   A
'Fm   C #ColumnNm.
'Ret  Drs{ <A> {F<C>} } ! Add a new column {F<C>} add end which is Filler-column
'                       ! Filler column of a given-column-A is an integer-columns with value = MaxWdt(col-A) - Len(cur-value-of-col-A)
If NoReczDrs(A) Then Stop
Dim W%: W = WdtzAy(StrColzDrs(A, C))
Dim I%: I = IxzAy(A.Fny, C)
Dim ODry(): ODry = A.Dry
Dim Dr, J&
For Each Dr In Itr(ODry)
    PushI Dr, W - Len(Dr(I))
    ODry(J) = Dr
    J = J + 1
Next
AddColzFillerC = Drs(AddFF(A.Fny, "F" & C), ODry)
End Function

Sub RmvMthzMNn(M As CodeModule, Mthnn)
Dim I
For Each I In TermAy(Mthnn)
    RmvMthzMN M, I
Next
End Sub

Sub RmvMth(M As CodeModule, Mthn)
RmvMthzMN M, Mthn
End Sub

Sub RmvMthzMN(M As CodeModule, Mthn)
With MthSC(M, Mthn)
    If .S2 > 0 Then M.DeleteLines .S2, .C2
    If .S1 > 0 Then M.DeleteLines .S1, .C1
End With
End Sub

Sub CpyMthAs(M As CodeModule, Mthn, AsMthn)
If HasMthzM(M, AsMthn) Then Inf CSub, "AsMth exist.", "Mdn FmMth AsMth", Mdn(M), Mthn, AsMthn: Exit Sub
End Sub
Sub BrwMd(Md As CodeModule)
If Md.CountOfLines = 0 Then BrwStr "No lines in Md[" & Mdn(Md) & "]": Exit Sub
Brw Src(Md), "Md" & Mdn(Md)
End Sub
Private Sub ZZ_RmvMthzMN()
Dim Md As CodeModule
Const Mthn$ = "ZZRmv1"
Dim Bef$(), Aft$()
Crt:
        Set Md = TmpMod
        ApdLines Md, LineszVbl("|'sdklfsdf||'dsklfj|Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
Tst:
        Bef = Src(Md)
        RmvMthzMN Md, Mthn
        Aft = Src(Md)

Insp:   Insp CSub, "RmvMth Test", "Bef RmvMth Aft", Bef, Mthn, Aft
Rmv:    RmvMd Md
End Sub


Sub MovMthzNM(Mthn, ToMdn)
MovMthzMNM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzMNM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Mthn, Md, ToMd
RmvMthzMN Md, Mthn
End Sub

Function EmpFunLines$(FunNm)
EmpFunLines = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function EmpSubLines$(Subn)
EmpSubLines = FmtQQ("Sub ?()|End Sub", Subn)
End Function
Sub AddSub(Subn)
ApdLines CMd, EmpSubLines(Subn)
JmpMth Subn
End Sub

Sub AddFun(FunNm)
ApdLines CMd, EmpFunLines(FunNm)
JmpMth FunNm
End Sub

Sub Z1()
ZZ_AlignMthDimzML
End Sub

Sub Z11()
Const Mdn$ = "QIde_B_MthOp"
Const Mthn$ = "AlignMthDimzML"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&: L = MthLnozMM(M, Mthn)
AlignMthDimzML M, L
End Sub

Sub ZZ_AlignMthDimzML()
Const TMthn$ = "AlignMthDimzML"
Const TMdn$ = "QIde_B_MthOp"
Const TCclsNm$ = TMdn & "__" & TMthn
Const TmMdn$ = "ATmp"
Const TmCclsNm$ = TmMdn & "__" & TMthn

Dim FmM As CodeModule, ToM As CodeModule, M As CodeModule, MthLno&, Lno&
Dim OldL$, NewL$, NewLy$()
Dim S1 As Boolean, S2 As Boolean
'Cpy Mth
    EnsMod CPj, TmMdn
    Set FmM = Md(TMdn)
    Set ToM = Md(TmMdn)
    'NewL
        NewL = MthLineszM(FmM, TMthn)
        NewLy = SplitCrLf(NewL)
        If Not HasPfx(NewLy(1), "Static F") Then Stop
        NewLy(1) = "Static F As New ATmp__AlignMthDimzML"
        NewL = JnCrLf(NewLy)
    
    'Rpl
        S1 = RplMth(ToM, TMthn, NewL)
        'Debug.Print "CpyMth: "; S1

'Cpy Md
    EnsCls CPj, TmCclsNm
    Set FmM = Md(TCclsNm)
    Set ToM = Md(TmCclsNm)
    S2 = CpyMd(FmM, ToM)
    'Debug.Print "CpyMd: "; S2
    If S1 Or S2 Then MsgBox "Copied": Exit Sub

'Align
    Set M = Md(TMdn)
    MthLno = MthLnozMM(M, TMthn)
    ATmp.AlignMthDimzML M, MthLno, SkpChkSelf:=True, Rpt:=EiUpdAndRpt
End Sub

Function CpyMd(FmM As CodeModule, ToM As CodeModule) As Boolean
CpyMd = RplMd(ToM, SrcLines(FmM))
End Function

Function CpyMth(Mthn, FmM As CodeModule, ToM As CodeModule) As Boolean
Dim NewL$
'NewL
    NewL = MthLineszM(FmM, Mthn)
'Rpl
    CpyMth = RplMth(ToM, Mthn, NewL)
'
'Const CSub$ = CMod & "CpyMdMthToMd"
'Dim Nav(): ReDim Nav(2)
'GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
'If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
'If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
'ToMd.AddFromString MthLineszM(Md, Mthn)
'RmvMth Md, Mthn
'If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", Mthn, Mdn(Md), Mdn(ToMd))
'Exit Sub
'BldNav:
'    Nav(0) = "FmMd Mth ToMd"
'    Nav(1) = Mdn(Md)
'    Nav(2) = Mthn
'    Nav(3) = Mdn(ToMd)
'    Return

End Function
Function CpyMthAsVer(M As CodeModule, Mthn, Ver%) As Boolean
'Ret True if copied
Dim VerMthn$, NewL$, L$, OldL$
If Not HasMthzM(M, Mthn) Then Inf CSub, "No from-mthn", "Md Mthn", Mdn(M), Mthn: Exit Function
VerMthn = Mthn & "_Ver" & Ver
'NewL
    L = MthLineszM(M, Mthn)
    NewL = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    CpyMthAsVer = RplMth(M, VerMthn, NewL)

End Function

Sub ZZ_RplMthRmk()
'GoSub Z1
GoSub Z3
Exit Sub
Z1:
    Dim M As CodeModule: Set M = TmpMod
    M.AddFromString "Sub AAXX()" & vbCrLf & "End Sub"
    RplMthRmk M, "AAXX", "'skldfsd" & vbCrLf & "'sldkfjsdlfk"
    Return
Z2:
    Set M = Md("TmpMod20190605_231101")
    RplMthRmk M, "AAXX", RplVBar("'sldkfjsd|'slkdfj|slkdfj|'sldkfjsdf|'sdf")
    Return
Z3:
    Set M = Md("TmpMod20190605_231101")
    RplMthRmk M, "AAXX", RplVBar("'a|'bb|'cfsdfdsc")
    Return
End Sub

Sub RplMthRmk(M As CodeModule, Mthn, NewRmk$)
RplLines M, MthRmk(M, Mthn), NewRmk
End Sub
Sub RplMthRmkzS12(M As CodeModule, NewRmk As S1S2s)
Dim Ay() As S1S2: Ay = NewRmk.Ay
Dim J%
For J = 0 To NewRmk.N - 1
    RplMthRmk M, Ay(J).S1, Ay(J).S2
Next
End Sub
Sub DltLines(M As CodeModule, Lno&, OldLines$)
If OldLines = "" Then Exit Sub
If Lno = 0 Then Exit Sub
Dim Cnt&: Cnt = LinCnt(OldLines)
If M.Lines(Lno, Cnt) <> OldLines Then Thw CSub, "OldL <> ActL", "OldL ActL", OldLines, M.Lines(Lno, Cnt)
M.DeleteLines Lno, Cnt
Debug.Print FmtQQ("DltLines: Lno(?) Cnt(?)", Lno, Cnt)
D BoxLy(SplitCrLf(OldLines))
End Sub
Sub RplLines(M As CodeModule, Old As LnoLines, NewLines$)
Dim Act$: Act = Old.Lines
Dim Lno&: Lno = Old.Lno
If Lno = 0 Then Debug.Print "RplLines: Lno=0": Exit Sub
If Act = NewLines Then Debug.Print "RplLines: Same": Exit Sub
DltLines M, Lno, Act
If NewLines <> "" Then
    'Debug.Print "RplLines: Rplaced"
    M.InsertLines Lno, NewLines
End If
End Sub

