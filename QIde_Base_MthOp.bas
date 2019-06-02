Attribute VB_Name = "QIde_Base_MthOp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Op."

Sub Z1()
ZZ_AlignMthDimzML
End Sub

Private Sub ZZ_AlignMthDimzML()
Dim Md As CodeModule, MthLno&
Set Md = MdzPN(CPj, "QIde_Base_MthOp")
MthLno = MthLnozMM(Md, "AlignMthDimzML")
AlignMthDimzML Md, MthLno
End Sub

Sub AlignMthDim()
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Sub
AlignMthDimzML M, CMthLno
End Sub

Sub AlignMthDimzML(Md As CodeModule, MthLno&)
Static F As New QIde_Base_MthOp__AlignMthDimzML

Dim Xp   As Boolean:  Xp = F.Xp(Md, MthLno):     '  #eXit-Parameter-er.
Dim Xnf  As Boolean: Xnf = F.Xnf(Xp, Md, MthLno) '

'== Align / GenCalling the main method ==================================
Dim Ml$:              Ml = F.Ml(Xp, Md, MthLno) '  #Mth-Lin.
Dim Mln$:            Mln = F.Mln(Xp, Ml)        '  #Mth-Lin-Name.
Dim Mlf  As Boolean: Mlf = F.Mlf(Xp, Ml)        '  #Mth-Lin-isFun.
Dim Mlc$:            Mlc = F.Mlc(Xp, Mlf, Ml)   '  #Mth-Lin-tyChr.
Dim Mlr$:            Mlr = F.Mlr(Xp, Mlf, Ml)   '  #Mth-Lin-Retty.

'== Align / GenCalling the main method ==================================
Dim Mc   As Drs:       Mc = F.Mc(Xp, Md, MthLno)  ' L MthLin              #Mth-Cxt.
Dim Xc   As Boolean:   Xc = F.Xc(Xp, Mc)          '                                           ! eXit-Cnt-of-mc-is-zero-er
Dim Mcc  As Drs:      Mcc = F.Mcc(Xc, Mc)         ' L MthLin              #Mth-Cxt-Cln.       ! non-Dim, non-Rmk are removed. Cln to Align
Dim Xa   As Boolean:   Xa = F.Xa(Xc, Mcc)         '                       #eXit-Alignable-er. ! No alignable lines
Dim Mcg  As Drs:      Mcg = F.Mcg(Xa, Mcc)        ' L Gpno MthLin         #Mth-Cxt-Gp.        ! L in seq will be one gp & Rmv the RmkLin
Dim Mcr  As Drs:      Mcr = F.Mcr(Xa, Mcg)        ' L Gpno MthLin IsRmk   #Mth-Cxt-isRmk
Dim Mctr As Drs:     Mctr = F.Mctr(Xa, Mcr)       ' L Gpno MthLin IsRmk   #Mth-Cxt-TopRmk     ! For each gp, the front rmk lines are TopRmk,
'                                                                                             ! rmv them
Dim Mcb  As Drs:      Mcb = F.Mcb(Xa, Mctr) ' L Gpno MthLin IsRmk   #Mth-Cxt-Brk.       ! Brk the MthLin into V Sfx Expr Rmk
'                                                   Dcl V Expr R1 R2 R3
Dim Mcf  As Drs:      Mcf = F.Mcf(Xa, Mcb)        ' L Gpno MthLin IsRmk   #Mth-Cxt-Filler.    ! Add5Col FV..FR2
'                                                   V Sfx Expr R1 R2
'                                                   F0 FSfx FExpr FR1 FR2
Dim Mcd  As Drs:      Mcd = F.Mcd(Xa, Mcf)  ' L MthLin DimLin       #Mth-Cxt-DimLin     ! Bld the new DimLin
Dim Mco  As Drs:      Mco = F.Mco(Xa, Mcd)        ' L DimLin              #Oup-Mth-Cxt.       ! Keep MthLin<>DimLin
Dim Omc  As Unt:      Omc = F.Omc(Xa, Md, Mco)    '

'== Create ChdFun ========================================================
'== Upd ChdFunMthLin =====================================================
Dim Nce$():          Nce = F.Nce(Xa, Mcb) '  #New-Chdmth-Ept. ! It is from V and Expr=F.{V}
Dim Nca$():          Nca = F.Nca(Xa, Md)  '  #New-Chdmth-Act. ! It is from chd cls of given md
Dim Ncn$():          Ncn = MinusAy(Nce, Nca)     '  #New-Chdmth-New. ! The new ChdMthNy to be created.
Dim Xnn As Boolean:  Xnn = F.Xnn(Xa, Ncn) '#eXit-No-New-chd-mth
Dim Ncd   As Drs:  Ncd = F.Ncd(Xnn, Ncn, Mcb, Ml) ' Chdn TyChr MthPm RetAs #New-Chd-Dta-to-bld-mth-lin
BrwDrs Ncd: Stop
Dim Nco$():      Nco = F.Nco(Xnn, Ncd)  ' MthLin  #Oup-Chd-mthlin-Create. ! EndLin is always End Function
Dim Onc  As Unt: Onc = F.Onc(Xnn, Nco) '         #Oup-New-Chdfun.        ! Create the new chd funs

'== Update Chd Fun's rmk =================================================
Dim Cre  As Drs: Cre = F.Cre(Xa, Mcc) ' ChdFun Pm R1 R2 #Chdmth-Rmk-Ept.
Dim Cra  As Drs: Cra = F.Cra          ' ChdFun Arg      #Chdmth-Rmk-Act.
Dim Ocr  As Drs: Ocr = F.Ocr          ' ChdFun Pm R1 R2 #Oup-Chdmth-Rmk.
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

Sub RmvMthzMNn(A As CodeModule, Mthnn, Optional WiTopRmk As Boolean)
Dim I
For Each I In TermAy(Mthnn)
    RmvMthzMN A, I
Next
End Sub

Sub RmvMth(A As CodeModule, Mthn)
RmvMthzMN A, Mthn
End Sub

Sub RmvMthzMN(A As CodeModule, Mthn)
DltLinzF A, MthFeiszMN(A, Mthn, WiTopRmk:=True)
End Sub

Private Sub Z_RmvMthzMN()
Const N$ = "ZZModule"
Dim Md As CodeModule, Mthn
GoSub Crt
'RmvMdEndBlankLin Md
Tst:
    RmvMthzMN Md, Mthn
    If Md.CountOfLines <> 0 Then Stop
    Return
Crt:
    Set Md = TmpMod
    RmvMd Md
    ApdLines Md, LineszVbl("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Private Sub ZZ()
End Sub
Sub CpyMth(Md As CodeModule, Mthn, ToMd As CodeModule, Optional IsSilent As Boolean)
Const CSub$ = CMod & "CpyMdMthToMd"
Dim Nav(): ReDim Nav(2)
GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
ToMd.AddFromString MthLineszMN(Md, Mthn)
RmvMth Md, Mthn
If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", Mthn, Mdn(Md), Mdn(ToMd))
Exit Sub
BldNav:
    Nav(0) = "FmMd Mth ToMd"
    Nav(1) = Mdn(Md)
    Nav(2) = Mthn
    Nav(3) = Mdn(ToMd)
    Return
End Sub

Sub MovMthzNM(Mthn, ToMdn)
MovMthzMNM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzMNM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Md, Mthn, ToMd
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

