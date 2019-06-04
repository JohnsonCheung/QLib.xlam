Attribute VB_Name = "QIde_Base_MthOp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Op."
Sub Z()
AlignMthDim
End Sub
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

Dim CmlCallg  As Drs:  CmlCallg = F.CmlCallg(CmlFmMc, CmlCallgPfx)
Dim CmlDclPm  As Drs:  CmlDclPm = F.CmlDclPm(CmlCallg, CmlVSfx)
Dim CmlMthRet As Drs: CmlMthRet = F.CmlMthRet(CmlDclPm)
Dim CmlEpt    As Drs:    CmlEpt = F.CmlEpt(CmlMthRet, CmMdy)                                    ' V EptL
Dim CmlAct    As Drs:    CmlAct = Drs_MthLinzM(CmMd)                                            ' Mdn Lno Mthn MthLin
Dim CmlJn     As Drs:     CmlJn = LJnDrs(CmlEpt, CmlAct, "V:Mthn", "Lno MthLin:ActL", "HasAct")
Dim CmlLNewO  As Drs:  CmlLNewO = F.CmlLNewO(CmlJn)
Dim OUpdCml   As Unt:   OUpdCml = RplLin(CmMd, CmlLNewO)
Exit Sub
'== Update Chd Mth's Rmk ===============================================================================================
Dim CrEpt   As Drs:   CrEpt = F.CrEpt(McBrk)                   ' V EptR
Dim CrAct   As Drs:   CrAct = F.CrAct(CmEpt, CmMd)             ' V ActR L
Dim CrJn    As Drs:    CrJn = JnDrs(CrAct, CrEpt, "V", "EptR") ' V ActR L EptR
Dim CrLNewO As Drs: CrLNewO = F.CrLNewO(CrJn)
Dim OUpdCr  As Unt:  OUpdCr = RplLin(CmMd, CrLNewO)
End Sub

Sub AlignMthDim()
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Sub
AlignMthDimzML M, CMthLno
End Sub

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

Sub ZZ_AlignMthDimzML()
Const TMthn$ = "AlignMthDimzML"
Const TMdn$ = "QIde_Base_MthOp"
Const TCclsNm$ = TMdn & "__" & TMthn
Const TmMdn$ = "ATmp"
Const TmCclsNm$ = TmMdn & "__" & TMthn

Dim FmM As CodeModule, ToM As CodeModule, M As CodeModule, MthLno&
Dim S1 As Boolean, S2 As Boolean
'Cpy Mth
    EnsMod CPj, TmMdn
    Set FmM = Md(TMdn)
    Set ToM = Md(TmMdn)
    S1 = CpyMth(TMthn, FmM, ToM)
    Debug.Print "CpyMth: "; S1
'Cpy Md
    EnsCls CPj, TmCclsNm
    Set FmM = Md(TCclsNm)
    Set ToM = Md(TmCclsNm)
    S2 = CpyMd(FmM, ToM)
    Debug.Print "CpyMd: "; S2
    If S1 Or S2 Then MsgBox "Copied": Exit Sub

'Align
    Set M = Md(TMdn)
    MthLno = MthLnozMM(M, TMthn)
    ATmp.AlignMthDimzML M, MthLno, SkpChkSelf:=True
End Sub

Function CpyMd(FmM As CodeModule, ToM As CodeModule) As Boolean
CpyMd = RplMdzML(ToM, SrcLines(FmM))
End Function


Function CpyMth(Mthn, FmM As CodeModule, ToM As CodeModule) As Boolean
Dim NewL$
'NewL
    NewL = MthLineszMN(FmM, Mthn)
'Rpl
    CpyMth = RplMth(ToM, Mthn, NewL)
'
'Const CSub$ = CMod & "CpyMdMthToMd"
'Dim Nav(): ReDim Nav(2)
'GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
'If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
'If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
'ToMd.AddFromString MthLineszMN(Md, Mthn)
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
    L = MthLineszMN(M, Mthn)
    NewL = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    CpyMthAsVer = RplMth(M, VerMthn, NewL)

End Function


