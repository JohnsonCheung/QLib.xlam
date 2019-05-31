Attribute VB_Name = "QIde_Mth_MthOp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Op."
Sub Z1()
AlignMthDim
End Sub

Sub AlignMthDim()
AlignMthDimzML CMd, CMthLno
End Sub

Private Function XamdB(M As CodeModule, MthLno&) As Drs
XamdB = DMthCxt(M, MthLno)
End Function

Private Sub AlignMthDimzML(M As CodeModule, MthLno&)
                  If IsNothing(M) Then Debug.Print "No CMd": Exit Sub
Dim A&:       A = MthLno      '                                       !
Dim B As Drs: B = XamdB(M, A) ' L MthLin                              ! mth cxt
Dim C As Drs: C = XamdC(B)    ' L MthLin                              ! mth cxt only lin is alignable
Dim D As Drs: D = XamdD(C)    ' L Gpno MthLin                         ! with L in seq will be one gp
Dim E As Drs: E = XamdE(D)    ' L Gpno MthLin V Sfx Expr Rmk1 Rmk2    ! Brk the MthLin in V Sfx Expr Rmk
Dim F%:       F = XamdF(E)
Dim G As Drs: G = XamdG(F, E) ' L Gpno MthLin V Sfx Expr Rmk1 Rmk2 F0..FRmk1 ! Add4Col  FV..FRmk1
Dim H As Drs: H = XamdH(G)    ' L MthLin DimLin                              ! Bld the new DimLin
Dim I As Drs: I = XamdI(H)    ' L DimLin            ! Keep MthLin<>DimLin
BrwDrs I
Stop
'                  XamdUpd M, I
End Sub

Private Function XamdI(H As Drs) As Drs
'Fm : L MthLin DimLin
'Ret: L Dim ! Keep MthLin<>DimLin
Dim Dr, Dry()
For Each Dr In Itr(H.Dry)
    If Dr(1) <> Dr(2) Then
        PushI Dry, Array(Dr(0), Dr(2))
    End If
Next
XamdI = DrszFF("L DimLin", Dry)
End Function

Private Function XamdH(G As Drs) As Drs
'      0 1    2      3 4   5    6    7    8  9  10   11    12
'Fm  ' L Gpno MthLin V Sfx Expr Rmk1 Rmk2 F0 FV FSfx FExpr FRmk1 ! Upd F0..FRmk1
'Ret ' L DimLin                              ! Bld the new DimLin
Dim Dr
Dim IV%: IV = IxzAy(G.Fny, "V")
For Each Dr In Itr(G.Dry)
    Dim L&:           L = Dr(0)
    Dim V$:           V = Dr(IV)
    Dim Sfx$:       Sfx = Dr(IV + 1)
    Dim Expr$:     Expr = Dr(IV + 2)
    Dim Rmk1$:     Rmk1 = Dr(IV + 3)
    Dim Rmk2$:     Rmk2 = Dr(IV + 4)
    Dim F0$:         F0 = Dr(IV + 5)
    Dim FV$:         FV = Dr(IV + 6)
    Dim FSfx$:     FSfx = Dr(IV + 7)
    Dim FExpr$:   FExpr = Dr(IV + 8)
    Dim FRmk1$:   FRmk1 = Dr(IV + 9)
    Dim R$:           R = XRmk(FRmk1, Rmk1, Rmk2)
    Dim DimLin$: DimLin = FmtQQ("?Dim ??: ??? = ?? ?", F0, V, Sfx, FV, FSfx, V, Expr, FExpr, R)
    Dim ODry(): PushI ODry, Array(L, DimLin)
Next
XamdH = DrszFF("L DimLin", ODry)
End Function

Private Function XRmk$(FRmk1$, R1$, R2$)
Select Case Rmk2 = ""
Case True: XRmk = "' " & R1
Case Else: XRmk = "' " & R1 & Frmk & "! " & R2
End Select
End Function

Private Function XamdF%(A As Drs)
'Fm ' L Gpno V Sfx Expr Rmk1 Rmk2 F0 FV..FRmk1 ! Adding columns EmpCol FV..FRmk1
'Ret' MaxGpno%
If NoReczDrs(A) Then Exit Function
XamdF = MaxzAy(IntAyzDrsC(A, "Gpno"))
End Function

Private Function XF0zDimLin$(DimLin)
If T1(DimLin) <> "Dim" Then Stop
Dim L$: L = LTrim(DimLin)
XF0zDimLin = Space(Len(DimLin) - Len(L))
End Function

Private Function XamdE(A As Drs) As Drs
'Fm : ' L Gpno MthLin                  ! with L in seq will be one gp
'Ret: ' L Gpno MthLin V Sfx Expr Rmk1 Rmk2 ! Brk the MthLin in V Sfx Expr Rmk F0 is spc bef Dim for fst lin
Dim Dry(), Dr
For Each Dr In Itr(A.Dry)
    Push Dry, AddAy(Dr, XamdEi(Dr(2)))
Next
Dim FF$: FF = "L Gpno MthLin V Sfx Expr Rmk1 Rmk2"
XamdE = DrszFF(FF, Dry)
End Function

Private Function XamdEi(MthLin) As Variant()
Dim L$:           L = MthLin
If ShfT1(L) <> "Dim" Then Stop
Dim V$:       V = ShfNm(L)
Dim Sfx$:   Sfx = ShfBef(L, ":")
Dim V1$:     V1 = ShfBef(L, "=")
Dim Expr$: Expr = ShfBefOrAll(L, vbSngQuote)
Dim Rmk1$: Rmk1 = ShfBefOrAll(L, vbExcM)
Dim Rmk2$: Rmk2 = L
If HasPfx(Sfx, "As ") Then Sfx = " " & Sfx
If V <> V1 Then Stop
XamdEi = Array(V, Sfx, Expr, Rmk1, Rmk2)
End Function

Private Function XamdD(A As Drs) As Drs
'Fm  L MthLin      ! mth cxt only lin is alignable
'Ret L Gpno MthLin ! with L in seq will be one gp
Dim Dr, LasL&, Gpno%, L&, Dry()
For Each Dr In Itr(A.Dry)
    L = Dr(0)
    If LasL + 1 <> L Then
        Gpno = Gpno + 1
    End If
    LasL = L
    PushI Dry, Array(L, Gpno, Dr(1))
Next
XamdD = DrszFF("L Gpno MthLin", Dry)
End Function

Private Function XamdC(A As Drs) As Drs
'Fm  : L MthLin ! mth cxt
'Ret : L MthLin ! mth cxt only lin is alignable
Dim Dr, Dry()
For Each Dr In Itr(A.Dry)
    If XIsAlignable(Dr(1)) Then
        PushI Dry, Dr
    End If
Next
XamdC = Drs(A.Fny, Dry)
End Function
Private Function XIsAlignable(Lin) As Boolean
Dim L$, N1$, N2$
L = Lin
If ShfT1(L) <> "Dim" Then Exit Function
N1 = TakNm(L)
N2 = Bet(L, ":", "=")
If N1 <> N2 Then Exit Function
XIsAlignable = True
End Function
Sub ZZ_XamdG()
Dim MaxGpno%, A As Drs, Act As Drs, Ept As Drs
MaxGpno = 1
ZZ:
    GoSub Dta1
    GoTo Tst
Tst:
    Act = XamdG(MaxGpno, A)
    BrwDrs Act: Stop
    If Not IsEqDrs(Act, Ept) Then Stop
    Return
Dta1:
    MaxGpno = 1
    Erase XX
    X "|----|------|------------------------------------------------------------------------------------------------------|---|---------|-------------|----------------------------------------------|----------------------------------|"
    X "| L  | Gpno | MthLin                                                                                               | V | Sfx     | Expr        | Rmk1                                         | Rmk2                             |"
    X "|----|------|------------------------------------------------------------------------------------------------------|---|---------|-------------|----------------------------------------------|----------------------------------|"
    X "| 17 | 1    | Dim A&:       A = MthLno      '                                       !                              | A | &       | MthLno      |                                              |                                  |"
    X "| 18 | 1    | Dim B As Drs: B = XamdB(M, A) ' L MthLin                              ! mth cxt                      | B |  As Drs | XamdB(M, A) | L MthLin                                     | mth cxt                          |"
    X "| 19 | 1    | Dim C As Drs: C = XamdC(B)    ' L MthLin                              ! mth cxt only lin is alignabl | C |  As Drs | XamdC(B)    | L MthLin                                     | mth cxt only lin is alignable    |"
    X "| 20 | 1    | Dim D As Drs: D = XamdD(C)    ' L Gpno MthLin                         ! with L in seq will be one gp | D |  As Drs | XamdD(C)    | L Gpno MthLin                                | with L in seq will be one gp     |"
    X "| 21 | 1    | Dim E As Drs: E = XamdE(D)    ' L Gpno MthLin V Sfx Expr Rmk1 Rmk2    ! Brk the MthLin in V Sfx Expr | E |  As Drs | XamdE(D)    | L Gpno MthLin V Sfx Expr Rmk1 Rmk2           | Brk the MthLin in V Sfx Expr Rmk |"
    X "| 22 | 1    | Dim F%:       F = XamdF(E)                                                                           | F | %       | XamdF(E)    |                                              |                                  |"
    X "| 23 | 1    | Dim G As Drs: G = XamdG(F, E) ' L Gpno MthLin V Sfx Expr Rmk1 Rmk2 F0..FRmk1 ! Add4Col  FV..FRmk1    | G |  As Drs | XamdG(F, E) | L Gpno MthLin V Sfx Expr Rmk1 Rmk2 F0..FRmk1 | Add4Col  FV..FRmk1               |"
    X "| 24 | 1    | Dim H As Drs: H = XamdH(G)    ' L MthLin DimLin                              ! Bld the new DimLin    | H |  As Drs | XamdH(G)    | L MthLin DimLin                              | Bld the new DimLin               |"
    X "| 25 | 1    | Dim I As Drs: I = XamdI(H)    ' L DimLin            ! Keep MthLin<>DimLin                            | I |  As Drs | XamdI(H)    | L DimLin                                     | Keep MthLin<>DimLin              |"
    X "|----|------|------------------------------------------------------------------------------------------------------|---|---------|-------------|----------------------------------------------|----------------------------------|"
    A = DrszFmtg(XX)
    Erase XX
    Return
End Sub

Private Function XamdG(MaxGpno%, A As Drs) As Drs
'A     : L Gpno MthLin V Sfx Expr Rmk1 Rmk2
'Ret   : Add 5 Columns: FV FSfx FExpr FRmk1 for each gp
Dim O As Drs, IGpno%, Fny$()
O = A
For IGpno = 1 To MaxGpno
    O = XamdGi(MaxGpno, O, IGpno)
Next
Fny = Sy(A.Fny, SyzSS("F0 FV FSfx FExpr FRmk1"))
XamdG = Drs(Fny, O.Dry)
End Function

Private Function XamdGi(MaxGpno%, A As Drs, IGpno%) As Drs
'Dry   : L Gpno MthLin V Sfx Expr Rmk1 Rmk2
'IGpno : Updating the Gpno=IGpno
'Ret   : Add 5 columns F0 FV FSfx FExpr FRmk1
Dim B As Drs, CV$(), CSfx$(), CExpr$(), CRmk1$()
B = DrswColEq(A, "Gpno", IGpno)
    IntoColApzDrs B, "V Sfx Expr Rmk1", CV, CSfx, CExpr, CRmk1
Dim WV%:       WV = WdtzAy(CV)
Dim WSfx%:   WSfx = WdtzAy(CSfx)
Dim WExpr%: WExpr = WdtzAy(CExpr)
Dim WRmk1%: WRmk1 = WdtzAy(CRmk1)
Dim ODry(): ODry = A.Dry
Dim IV%: IV = IxzAy(A.Fny, "V")
Dim Dr: Dr = ODry(0)
Dim F0$: F0 = XF0zDimLin(Dr(2))
Dim J%
For Each Dr In Itr(ODry)
    Dim Gpno%:   Gpno = Dr(1)
                        If Gpno <> IGpno Then GoTo X
    Dim V$:         V = Dr(IV)
    Dim Sfx$:     Sfx = Dr(IV + 1)
    Dim Expr$:   Expr = Dr(IV + 2)
    Dim Rmk1$:   Rmk1 = Dr(IV + 3)
    Dim FV$:       FV = Space(WV - Len(V))
    Dim FSfx$:   FSfx = Space(WSfx - Len(Sfx))
    Dim FExpr$: FExpr = Space(WExpr - Len(Expr))
    Dim FRmk1$: FRmk1 = Space(WRmk1 - Len(Rmk1))
    Dim Dr1():    Dr1 = AddAy(Dr, Array(F0, FV, FSfx, FExpr, FRmk1))
              ODry(J) = Dr1
X:
    J = J + 1
Next
XamdGi = Drs(A.Fny, ODry)
End Function
Private Sub XamdUpd(Md As CodeModule, O As Drs)
'O: L DimLin
Dim Dr
For Each Dr In Itr(O.Dry)
    Md.ReplaceLine Dr(0), Dr(1)
Next
End Sub
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

